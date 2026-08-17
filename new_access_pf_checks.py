# -*- coding: utf-8 -*-
"""
PF BP-PY-ZY — логика 1:1 с Access (q_*_Joined_1_Checks, q_*_Joined_3_BillToByINN).
Итоговый Excel — парами (3801_3803, 3802_3804, 3805_3806).
"""
from __future__ import annotations

import argparse
import asyncio
import os
import re
import sys
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None  # type: ignore

from parallel_io import async_io, gather_limited, parallel_enabled, shutdown_executor, worker_count
from staging_db import staging_active, start_staging, stop_staging

from build_checks import (
    BASE_DIR,
    BILL_TO_COLS,
    BLOCKED_ORBLK,
    FAKE_INN,
    OUTPUT_DIR,
    PAIRS,
    _CYR,
    _col,
    _excel_sheet_name_safe,
    _get_file,
    _mismatch_sheet_label,
    _align_merge_keys,
    _nc,
    _zfill10,
    _norm,
    _ns,
    _prep_bill_to_sheet,
    _attach_comment_om,
    _prep_mismatch_sheet,
    _write_mismatch_sheet,
    _write_bill_to_sheet,
    _write_exception_sheet,
    _read_base,
    _read_partner,
    _read_all_partners,
    _so_col,
    _unique_excel_sheet_name,
    collect_and_persist_global_exception,
    exception_keys,
    load_runtime_paths_dict,
    _norm_cust,
    to_export_orblk_names,
)

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# Как в Access q_*_Joined_1 (без «самлинг» — его нет в SQL)
ACCESS_NAME_BLACKLIST = (
    "sampling",
    "сэмплинг",
    "самплинг",
    "семплинг",
    "dummy",
    "дамми",
    "внутреннее",
    "внп",
    "bf_",
    "test",
    "тест",
)

_BILL_TO_B_SIDE = [
    "BP SO", "BP Customer", "BP Name", "BP Tax Number 1", "BP CGrp",
    "BP Grp4", "BP Search Term 2", "BP SDst", "BP A7", "BP OrBlk1", "BP OrBlk2",
]


def dedupe_base_access(df: pd.DataFrame) -> pd.DataFrame:
    """q_*_Base_DelDup: GROUP BY Customer, First() = первая строка группы (кроме Name — кириллица)."""
    if df.empty:
        return df

    rows = []
    for _, grp in df.groupby("Customer", sort=False, dropna=False):
        # Access First([field]) — значение из первой физической строки группы, не «первое непустое».
        row = grp.iloc[0].to_dict()
        if "Name" in grp.columns:
            names = [s.strip() for s in grp["Name"].fillna("").astype(str) if s.strip()]
            cyr = [n for n in names if _CYR.search(n)]
            row["Name"] = min(cyr) if cyr else (min(names) if names else "")
        rows.append(row)
    result = pd.DataFrame(rows)
    for c in df.columns:
        if c not in result.columns:
            result[c] = ""
    out = result[df.columns].reset_index(drop=True)
    return out.drop_duplicates(subset=["Customer"], keep="first").reset_index(drop=True)


def _effective_so(df: pd.DataFrame, folder_so: str) -> pd.Series:
    so = _ns(df["SO"]) if "SO" in df.columns else pd.Series("", index=df.index, dtype=object)
    if "_folder" in df.columns:
        so = so.where(so.ne(""), _ns(df["_folder"]))
    if folder_so:
        so = so.where(so.ne(""), folder_so)
    return so


def _rows_for_sorg_sheet(df: pd.DataFrame, so_key: str) -> pd.DataFrame:
    """Строки для листа пары: _folder (номер папки SO) важнее колонки SO из выгрузки."""
    if df is None or df.empty or not so_key:
        return pd.DataFrame()
    key = _norm(so_key)
    if "_folder" in df.columns:
        by_folder = _ns(df["_folder"]).eq(key)
        if by_folder.any():
            return df[by_folder].copy()
    return df[_effective_so(df, key).eq(key)].copy()


def apply_exception_access(base_dd: pd.DataFrame, exc_keys: set[tuple[str, str]], folder_so: str) -> pd.DataFrame:
    """Access: LEFT JOIN Exception … WHERE e.Customer IS NULL (до checks, на base)."""
    if base_dd.empty or not exc_keys:
        return base_dd
    so = _effective_so(base_dd, folder_so)
    cu = _nc(base_dd["Customer"])
    keep = pd.Series(
        [(_norm(s), _norm_cust(c)) not in exc_keys for s, c in zip(so, cu)],
        index=base_dd.index,
    )
    return base_dd[keep].copy()


def _master_lookup(base_dd: pd.DataFrame) -> pd.DataFrame:
    """Строки Base DelDup для LEFT JOIN d ON pf.KTONR = d.Customer."""
    ob = _col(base_dd, "OrBlk", "OrBlk 1")
    ob1 = _col(base_dd, "OrBlk1", "OrBlk 1")
    cols = ["Customer", "Name", "Tax Number 1"]
    if ob:
        cols.append(ob)
    if ob1:
        cols.append(ob1)
    present = [c for c in cols if c in base_dd.columns]
    return base_dd[present].drop_duplicates(subset=["Customer"], keep="first").copy()


def attach_partner_access(
    base_dd: pd.DataFrame,
    partner_df: pd.DataFrame | None,
    prefix: str,
    *,
    master_lookup: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Access:
      FROM base b LEFT JOIN (
        SELECT pf.KUNNR AS Customer, pf.KTONR AS {prefix}, d.*
        FROM partner pf LEFT JOIN Base_DelDup d ON pf.KTONR = d.Customer
      ) ON b.Customer = sub.Customer
    """
    result = _align_merge_keys(base_dd.copy())
    lookup_src = master_lookup if master_lookup is not None else result
    lookup_src = _align_merge_keys(lookup_src.copy()) if lookup_src is not None else result
    ob = _col(lookup_src, "OrBlk", "OrBlk 1")
    ob1 = _col(lookup_src, "OrBlk1", "OrBlk 1")
    pcols = [prefix, f"{prefix} Name", f"{prefix} Tax Number 1", f"{prefix} OrBlk1", f"{prefix} OrBlk2"]

    if partner_df is None or partner_df.empty:
        # Access: LEFT JOIN к подзапросу BP/PY/ZY без строки → NULL (пусто), не Customer.
        for c in pcols:
            result[c] = ""
        return result

    lookup = _master_lookup(lookup_src)
    rename_d = {"Customer": "_lk_cust", "Name": "_lk_name", "Tax Number 1": "_lk_tax"}
    if ob:
        rename_d[ob] = "_lk_ob1"
    if ob1:
        rename_d[ob1] = "_lk_ob2"
    d = lookup.rename(columns=rename_d)
    if "_lk_cust" in d.columns:
        d["_lk_cust"] = _nc(d["_lk_cust"])

    if "KUNNR" not in partner_df.columns or "KTONR" not in partner_df.columns:
        kunnr = _col(partner_df, "KUNNR", "Customer", "Sold-to")
        ktonr = _col(partner_df, "KTONR", "KUNN2", "Partner", prefix)
        if not kunnr or not ktonr:
            for c in pcols:
                result[c] = ""
            return result
        pf = partner_df[[kunnr, ktonr]].rename(columns={kunnr: "KUNNR", ktonr: "KTONR"}).copy()
    else:
        pf = partner_df[["KUNNR", "KTONR"]].copy()
    pf["KUNNR"] = _nc(pf["KUNNR"])
    pf["KTONR"] = _nc(pf["KTONR"])
    pf = pf[pf["KUNNR"].ne("") & pf["KTONR"].ne("")].copy()
    pf = pf.rename(columns={"KUNNR": "Customer", "KTONR": prefix})

    sub = pf.merge(d, left_on=prefix, right_on="_lk_cust", how="left")
    rename_sub = {
        "_lk_name": f"{prefix} Name",
        "_lk_tax": f"{prefix} Tax Number 1",
        "_lk_ob1": f"{prefix} OrBlk1",
        "_lk_ob2": f"{prefix} OrBlk2",
    }
    sub = sub.rename(columns={k: v for k, v in rename_sub.items() if k in sub.columns})
    sub = sub.drop(columns=["_lk_cust"], errors="ignore")
    for c in pcols:
        if c not in sub.columns:
            sub[c] = ""

    sub = sub[["Customer"] + pcols]

    result = result.drop(columns=[c for c in pcols if c in result.columns], errors="ignore")
    # Access: LEFT JOIN без Distinct — возможны many-to-many
    result = result.merge(sub, on="Customer", how="left")

    miss_mask = result[prefix].fillna("").astype(str).str.strip().eq("")
    if miss_mask.any() and not pf.empty:
        pf_z = pf.drop_duplicates(subset=["Customer"], keep="first").copy()
        pf_z["_z"] = _zfill10(pf_z["Customer"])
        miss = result.loc[miss_mask, ["Customer"]].copy()
        miss["_z"] = _zfill10(miss["Customer"])
        filled = miss.merge(pf_z.drop(columns=["Customer"]), on="_z", how="left")
        filled.index = miss.index
        got = filled[prefix].fillna("").astype(str).str.strip().ne("")
        if got.any():
            for c in pcols:
                if c in filled.columns:
                    result.loc[got.index[got], c] = filled.loc[got, c].to_numpy()

    for c in pcols:
        if c not in result.columns:
            result[c] = ""
        result[c] = (
            result[c].fillna("")
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": "", "<NA>": ""})
        )

    matched = int(result[prefix].ne("").sum()) if prefix in result.columns else 0
    miss_n = len(result) - matched
    print(
        f"[new_access] {prefix}: join {matched}/{len(result)} "
        f"(партнёрских строк {len(pf)})",
        flush=True,
    )
    if miss_n:
        sample = result.loc[result[prefix].eq(""), "Customer"].astype(str).head(15).tolist()
        print(
            f"[new_access] {prefix}: нет номера у {miss_n} клиентов, примеры: {sample}",
            flush=True,
        )

    return result


def merge_all_partners_access(
    base_dd: pd.DataFrame,
    bp_df: pd.DataFrame | None,
    py_df: pd.DataFrame | None,
    zy_df: pd.DataFrame | None,
    *,
    master_lookup: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Цепочка LEFT JOIN BP → PY → ZY (Access может размножить строки)."""
    lk = master_lookup if master_lookup is not None else base_dd
    result = attach_partner_access(base_dd, bp_df, "BP", master_lookup=lk)
    result = attach_partner_access(result, py_df, "PY", master_lookup=lk)
    result = attach_partner_access(result, zy_df, "ZY", master_lookup=lk)
    return result


def apply_filters_access(df: pd.DataFrame) -> pd.DataFrame:
    """Фильтры Access (Exception уже применён на base)."""
    if df.empty:
        return df

    name_s = df["Name"].fillna("").astype(str).str.lower() if "Name" in df.columns else pd.Series("", index=df.index)
    mask_name = ~name_s.apply(lambda n: any(b in n for b in ACCESS_NAME_BLACKLIST))

    cgrp_s = df["CGrp"].fillna("").astype(str).str.upper().str.strip() if "CGrp" in df.columns else pd.Series("", index=df.index)
    mask_cgrp = cgrp_s != "Z"

    def _ob(col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series("", index=df.index)
        t = df[col].fillna("").astype(str).str.strip().str.upper()
        return t.replace({"NAN": "", "<NA>": "", "NONE": ""})

    mask_orblk = ~_ob("OrBlk").isin(BLOCKED_ORBLK) & ~_ob("OrBlk1").isin(BLOCKED_ORBLK)

    tax_s = df["Tax Number 1"].fillna("").astype(str).str.strip() if "Tax Number 1" in df.columns else pd.Series("", index=df.index)
    mask_inn = ~tax_s.apply(lambda t: any(f in t for f in FAKE_INN))

    return df[mask_name & mask_cgrp & mask_orblk & mask_inn].copy()


def _load_sorg_from_excel(
    folder: str,
) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    fp = BASE_DIR / folder
    f_base = _get_file(fp, "*Base*.xlsx")
    if not f_base:
        return pd.DataFrame(), None, None, None

    def _file_mb(path: Path | None) -> str:
        if not path or not path.is_file():
            return "?"
        return f"{path.stat().st_size / (1024 * 1024):.1f} MB"

    def _load_base() -> pd.DataFrame:
        print(f"[new_access] SO {folder}: читаю Base {f_base.name} ({_file_mb(f_base)})…", flush=True)
        t0 = time.perf_counter()
        df = dedupe_base_access(_read_base(f_base, folder))
        print(f"[new_access] SO {folder}: Base готов, {len(df)} строк, {time.perf_counter() - t0:.0f} с", flush=True)
        return df

    def _load_partner(label: str, pattern: str) -> pd.DataFrame | None:
        print(f"[new_access] SO {folder}: читаю {label}…", flush=True)
        t0 = time.perf_counter()
        df = _read_all_partners(fp, folder, pattern, label, allow_com=True)
        rows = len(df) if df is not None and not df.empty else 0
        print(f"[new_access] SO {folder}: {label} готов, {rows} строк, {time.perf_counter() - t0:.0f} с", flush=True)
        return df

    if parallel_enabled():
        jobs: dict[str, object] = {
            "base": _load_base,
            "BP": lambda: _load_partner("BP", "*BP*.xlsx"),
            "PY": lambda: _load_partner("PY", "*PY*.xlsx"),
            "ZY": lambda: _load_partner("ZY", "*ZY*.xlsx"),
        }
        with ThreadPoolExecutor(max_workers=min(4, len(jobs)), thread_name_prefix=f"so{folder}") as pool:
            futs = {k: pool.submit(fn) for k, fn in jobs.items()}
            base = futs["base"].result()
            bp = futs["BP"].result()
            py = futs["PY"].result()
            zy = futs["ZY"].result()
        return base, bp, py, zy

    base = _load_base()
    return (
        base,
        _load_partner("BP", "*BP*.xlsx"),
        _load_partner("PY", "*PY*.xlsx"),
        _load_partner("ZY", "*ZY*.xlsx"),
    )


def load_sorg(folder: str) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    """DuckDB staging (если включена) или чтение xlsx."""
    if staging_active():
        from staging_db import get_staging

        base, bp, py, zy = get_staging().load_raw(folder)
        if base.empty:
            return pd.DataFrame(), None, None, None
        base = dedupe_base_access(base)
        def _rekey(df: pd.DataFrame | None) -> pd.DataFrame | None:
            if df is None or df.empty or "KUNNR" not in df.columns:
                return df
            out = df.copy()
            out["KUNNR"] = _nc(out["KUNNR"])
            if "KTONR" in out.columns:
                out["KTONR"] = _nc(out["KTONR"])
            return out[out["KUNNR"].ne("") & out["KTONR"].ne("")]
        bp, py, zy = _rekey(bp), _rekey(py), _rekey(zy)
        print(f"[new_access] SO {folder}: из DuckDB — base {len(base)} строк", flush=True)
        return base, bp, py, zy
    return _load_sorg_from_excel(folder)


def add_checks_access(df: pd.DataFrame) -> pd.DataFrame:
    """q_*_Joined_1_Checks — тексты комментариев как в Access (SP, не Cust)."""
    out = df.copy()

    def _s(col: str) -> pd.Series:
        return out[col].fillna("").astype(str).str.strip() if col in out.columns else pd.Series("", index=out.index)

    # Access OrBlk (= export OrBlk1) — Bill-to при M. Проверки до переименования в отчёт.
    orblk = _s("OrBlk").str.upper() if "OrBlk" in out.columns else _s("OrBlk1").str.upper()
    bp_orblk1 = _s("BP OrBlk1").str.upper()
    bp, py, zy = _s("BP"), _s("PY"), _s("ZY")
    cust, tax, bp_tax, bp_name = _s("Customer"), _s("Tax Number 1"), _s("BP Tax Number 1"), _s("BP Name")

    out["Cust OrBlk Bill-to"] = np.where(orblk == "M", "Bill-to", "Ship-to")
    out["BP OrBlk Bill-to"] = np.where(bp_orblk1 == "M", "Bill-to", "Ship-to")
    out["Check BP-PY-ZY"] = np.where((bp == py) & (bp == zy), "TRUE", "FALSE")
    out["Check Tax Number Cust&BP"] = np.where(tax == bp_tax, "TRUE", "FALSE")
    out["Check Cust&BP"] = np.where(cust == bp, "TRUE", "FALSE")

    bp_mismatch = (bp != py) | (bp != zy)
    tax_mismatch = tax != bp_tax
    cust_ne_bp = cust != bp

    comment = pd.Series("", index=out.index)
    comment = comment.mask(bp_mismatch & ~tax_mismatch, "Несоответствие BP-PY-ZY")
    comment = comment.mask(bp_mismatch & tax_mismatch, "Несоответствие BP-PY-ZY; несоответствие ИНН Cust и BP")
    comment = comment.mask(~bp_mismatch & tax_mismatch, "Несоответствие ИНН Cust и BP")
    comment = comment.mask(~bp_mismatch & ~tax_mismatch & (bp_name == ""), "BP помечен на удаление")
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk == "M") & cust_ne_bp,
        "Bill-to прикреплён к другому Bill-to",
    )
    comment = comment.mask(
        ~bp_mismatch & ~tax_mismatch & (bp_name != "") & (orblk != "M") & (bp_orblk1 != "M") & cust_ne_bp,
        "Ship-to прикреплён к BP без OB M",
    )
    out["Comment MD Analyst"] = comment
    return to_export_orblk_names(out)


def build_bill_to_access(df_checks: pd.DataFrame, base_dd: pd.DataFrame) -> pd.DataFrame:
    """
    q_*_Joined_3_BillToByINN:
    Access SQL: INNER JOIN только по Tax Number 1, b.OrBlk = 'M'
    (в черновике docx упомянут ещё SO — в access.txt SO в ON нет).
    Без groupby/drop_duplicates после join — Access размножает строки.
    """
    if df_checks.empty or base_dd.empty:
        return pd.DataFrame()

    pool1 = df_checks[
        (df_checks["Cust OrBlk Bill-to"] == "Ship-to")
        & (df_checks["Check Cust&BP"] == "TRUE")
        & (df_checks["Comment MD Analyst"].fillna("").astype(str).str.strip() == "")
    ].copy()
    if pool1.empty:
        return pd.DataFrame()

    pool1 = pool1.copy()
    if "_folder" in pool1.columns:
        pool1["SO"] = _ns(pool1["_folder"])
    elif "SO" not in pool1.columns:
        pool1["SO"] = ""

    orblk_col = _col(base_dd, "OrBlk", "OrBlk 1")
    if not orblk_col:
        return pd.DataFrame()

    pool2 = base_dd[base_dd[orblk_col].fillna("").astype(str).str.upper().str.strip() == "M"].copy()
    if pool2.empty:
        return pd.DataFrame()

    rename_p2: dict[str, str] = {"Customer": "BP Customer", orblk_col: "BP OrBlk1"}
    so_c = _so_col(pool2)
    if so_c:
        rename_p2[so_c] = "BP SO"
    elif "_folder" in pool2.columns:
        rename_p2["_folder"] = "BP SO"
    for src, dst in {
        "Name": "BP Name",
        "Tax Number 1": "BP Tax Number 1",
        "CGrp": "BP CGrp",
    }.items():
        if src in pool2.columns:
            rename_p2[src] = dst
    g4 = _col(pool2, "Grp4", "GROUP4", "Group4")
    if g4:
        rename_p2[g4] = "BP Grp4"
    ob2 = _col(pool2, "OrBlk1", "OrBlk 1", "OrBlk2", "OrBlk 2")
    if ob2:
        rename_p2[ob2] = "BP OrBlk2"
    pool2 = pool2.rename(columns=rename_p2)

    for col in _BILL_TO_B_SIDE:
        if col not in pool2.columns:
            pool2[col] = ""

    pool1_j = pool1.drop(columns=[c for c in _BILL_TO_B_SIDE if c in pool1.columns], errors="ignore")
    pool1_j["_tax"] = _ns(pool1_j["Tax Number 1"]) if "Tax Number 1" in pool1_j.columns else ""
    pool2["_tax"] = _ns(pool2["BP Tax Number 1"])

    p2_cols = ["_tax"] + [c for c in _BILL_TO_B_SIDE if c in pool2.columns]
    # Access: ON j.[Tax Number 1] = b.[Tax Number 1] (без SO)
    joined = pool1_j.merge(pool2[p2_cols], on="_tax", how="inner")
    joined = joined.drop(columns=["_tax"], errors="ignore")

    if "BP" in joined.columns:
        joined = joined.rename(columns={"BP": "BP now"})

    for col in _BILL_TO_B_SIDE + ["BP SO"]:
        if col not in joined.columns:
            joined[col] = ""

    # BP status — как IIf в Access
    ob1 = joined["BP OrBlk1"].fillna("").astype(str).str.strip().str.upper()
    ob2s = joined["BP OrBlk2"].fillna("").astype(str).str.strip()
    joined["BP status"] = np.select(
        [(ob1 == "M") & (ob2s == ""), (ob1 == "M") & (ob2s.str.upper() == "M")],
        ["Ошибка для OB", "active"],
        default="block",
    )

    status_order = {"active": 0, "block": 1, "Ошибка для OB": 2}
    joined["_st"] = joined["BP status"].map(status_order).fillna(1).astype(int)
    sort_cols = [c for c in ["Tax Number 1", "Customer", "_st", "BP Customer"] if c in joined.columns]
    joined = joined.sort_values(sort_cols, kind="stable").drop(columns=["_st"], errors="ignore")
    return joined.reset_index(drop=True)


def _bill_to_rows_access(bt: pd.DataFrame, so_key: str) -> pd.DataFrame:
    """Лист SOrg: j.[SO] = SOrg (для пары — по _folder / SO из папки выгрузки)."""
    return _rows_for_sorg_sheet(bt, so_key)


def _bill_to_sheet_label(so_token: str) -> str:
    so = _norm(so_token) or "SO"
    return f"{so} Привязка Bill-to по ИНН"


def save_pair_excel(
    pair_name: str,
    errors_df: pd.DataFrame,
    bill_to_df: pd.DataFrame,
    exc_df: pd.DataFrame,
    sorg_folders: list[str],
) -> None:
    pair_dir = OUTPUT_DIR / pair_name
    pair_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%d.%m.%Y")
    out_file = pair_dir / f"Check PF BP-PY-ZY {pair_name} {date_str}.xlsx"

    with pd.ExcelWriter(out_file, engine="openpyxl") as w:
        used: set[str] = set()
        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_mismatch_sheet_label(so_key), used)
            used.add(title)
            sub = _rows_for_sorg_sheet(errors_df, so_key)
            if not sub.empty:
                sub = sub.sort_values(
                    by=[c for c in ("Comment MD Analyst", "Customer") if c in sub.columns],
                    kind="stable",
                )
            prep = _prep_mismatch_sheet(sub)
            if prep.empty:
                _write_mismatch_sheet(
                    w,
                    title,
                    pd.DataFrame({"Сообщение": [f"Нет несоответствий для SO {so_key}"]}),
                )
                print(f"[new_access] лист «{title}»: 0 несоответствий", flush=True)
            else:
                _write_mismatch_sheet(w, title, prep)
                print(f"[new_access] лист «{title}»: {len(prep)} несоответствий", flush=True)

        for folder_so in sorg_folders:
            so_key = _norm(folder_so)
            if not so_key:
                continue
            title = _unique_excel_sheet_name(_bill_to_sheet_label(so_key), used)
            used.add(title)
            grp_raw = _bill_to_rows_access(bill_to_df, so_key)
            grp = _prep_bill_to_sheet(grp_raw) if not grp_raw.empty else pd.DataFrame(columns=BILL_TO_COLS)
            if grp.empty:
                _write_bill_to_sheet(w, title, pd.DataFrame(columns=BILL_TO_COLS))
                print(f"[new_access] лист «{title}»: 0 Bill-to", flush=True)
            else:
                _write_bill_to_sheet(w, title, grp)
                print(f"[new_access] лист «{title}»: {len(grp)} Bill-to", flush=True)

        if not exc_df.empty:
            _write_exception_sheet(w, exc_df)
        else:
            _write_exception_sheet(w, pd.DataFrame({"Сообщение": ["Нет записей-исключений"]}))

    print(f"[new_access] сохранено: {out_file}", flush=True)


def _diagnose_empty_partners(
    folder: str,
    errors: pd.DataFrame,
    bp_df: pd.DataFrame | None,
    py_df: pd.DataFrame | None,
    zy_df: pd.DataFrame | None,
) -> None:
    """По пустым BP/PY/ZY на листе несоответствий: клиент есть в файле как KUNNR или только как KTONR."""
    if errors is None or errors.empty or "Customer" not in errors.columns:
        return

    def _blank(col: str) -> pd.Series:
        if col not in errors.columns:
            return pd.Series(True, index=errors.index)
        return errors[col].fillna("").astype(str).str.strip().eq("")

    empty_all = _blank("BP") & _blank("PY") & _blank("ZY")
    n = int(empty_all.sum())
    if n == 0:
        return
    print(
        f"[new_access] SO {folder}: на несоответствиях {n} строк с пустыми BP+PY+ZY",
        flush=True,
    )
    files = {"BP": bp_df, "PY": py_df, "ZY": zy_df}
    for cust in _nc(errors.loc[empty_all, "Customer"]).drop_duplicates().head(20):
        bits = [f"Customer={cust}"]
        for label, df in files.items():
            if df is None or df.empty:
                bits.append(f"{label}:нет файла")
                continue
            as_kunnr = int((df["KUNNR"] == cust).sum()) if "KUNNR" in df.columns else 0
            as_ktonr = int((df["KTONR"] == cust).sum()) if "KTONR" in df.columns else 0
            bits.append(f"{label}:KUNNR={as_kunnr}/KTONR={as_ktonr}")
        print("  " + " | ".join(bits), flush=True)


def process_sorg(
    folder: str,
    exc_keys: set[tuple[str, str]],
    master_lookup: pd.DataFrame | None = None,
    loaded: tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Access-цепочка для одного SOrg:
    Base → DelDup → Exception → BP/PY/ZY → фильтры → Checks → Errors / Bill-to
    """
    if loaded is None:
        base_raw, bp_df, py_df, zy_df = load_sorg(folder)
    else:
        base_raw, bp_df, py_df, zy_df = loaded
    if base_raw.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    base = apply_exception_access(base_raw, exc_keys, folder)
    if base.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    lk = master_lookup if master_lookup is not None and not master_lookup.empty else base_raw
    merged = merge_all_partners_access(base, bp_df, py_df, zy_df, master_lookup=lk)
    filtered = apply_filters_access(merged)
    if filtered.empty:
        return pd.DataFrame(), pd.DataFrame(), base_raw

    checked = add_checks_access(filtered)
    # Access ErrorsOnly: Comment не пустой; без Distinct и без фильтра по BP Name
    errors = checked[checked["Comment MD Analyst"].fillna("").astype(str).str.strip() != ""].copy()
    bill_to = build_bill_to_access(checked, base_raw)
    _diagnose_empty_partners(folder, errors, bp_df, py_df, zy_df)

    # Диагностика лавины «Ship-to … без OB M»: после OrBlk-swap у BP OrBlk1 должно быть много M.
    bp_ob = (
        filtered["BP OrBlk1"].fillna("").astype(str).str.upper().str.strip()
        if "BP OrBlk1" in filtered.columns
        else pd.Series(dtype=object)
    )
    m_orblk = (
        (base_raw["OrBlk"].fillna("").astype(str).str.upper().str.strip() == "M").sum()
        if "OrBlk" in base_raw.columns
        else -1
    )
    ship_bp = 0
    if not errors.empty and "Comment MD Analyst" in errors.columns:
        ship_bp = int(
            errors["Comment MD Analyst"].astype(str).str.contains("Ship-to прикреплён к BP без OB M", na=False).sum()
        )
    print(
        f"[new_access] SO {folder}: OrBlk(M)={m_orblk} | BP OrBlk1(M)={(bp_ob == 'M').sum()} "
        f"| errors={len(errors)} (Ship-to без OB M={ship_bp}) | Bill-to={len(bill_to)}",
        flush=True,
    )
    if ship_bp > 1000:
        print(
            f"[new_access] SO {folder}: ВНИМАНИЕ — {ship_bp} «Ship-to без OB M». "
            f"Похоже OrBlk не swap’нут (для 3802 нужны колонки OrBlk+OrBlk.1→swap). "
            f"Проверьте git log -1 и что нет старого staging.duckdb.",
            flush=True,
        )

    if os.environ.get("REPORTS_DEBUG"):
        dup = checked.groupby("Customer").size().sort_values(ascending=False)
        print(
            f"[new_access] DEBUG SO {folder}: base_raw={len(base_raw)} base_exc={len(base)} "
            f"merged={len(merged)} filtered={len(filtered)} errors={len(errors)} "
            f"dup_max={dup.max() if len(dup) else 0}",
            flush=True,
        )
        print(
            f"[new_access] DEBUG SO {folder}: Cust OrBlk Bill-to={checked['Cust OrBlk Bill-to'].value_counts().to_dict()} "
            f"BP OrBlk Bill-to={checked['BP OrBlk Bill-to'].value_counts().to_dict()} "
            f"Check Cust&BP={checked['Check Cust&BP'].value_counts().to_dict()}",
            flush=True,
        )

    return errors, bill_to, base_raw


def _merge_sorg_results(
    pair_name: str,
    folders: list[str],
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]],
    exc_df: pd.DataFrame,
    *,
    failed_folders: list[str] | None = None,
) -> bool:
    errors_parts: list[pd.DataFrame] = []
    bill_parts: list[pd.DataFrame] = []
    order = {f: i for i, f in enumerate(folders)}
    done = {f for f, _ in results}
    for folder, (errors, bill_to, base) in sorted(results, key=lambda x: order.get(x[0], 0)):
        if not base.empty:
            print(
                f"[new_access] SO {folder}: ошибок {len(errors)}, Bill-to {len(bill_to)}",
                flush=True,
            )
        if not errors.empty:
            errors_parts.append(errors)
        if not bill_to.empty:
            bill_parts.append(bill_to)

    errors_all = pd.concat(errors_parts, ignore_index=True) if errors_parts else pd.DataFrame()
    if not errors_all.empty:
        errors_all = _attach_comment_om(errors_all, exc_df)
    bill_all = pd.concat(bill_parts, ignore_index=True) if bill_parts else pd.DataFrame()
    save_pair_excel(pair_name, errors_all, bill_all, exc_df, folders)

    missing = failed_folders or [f for f in folders if f not in done]
    if missing:
        print(
            f"[new_access] сохранено (частично): пара {pair_name} — без SO {', '.join(missing)}",
            flush=True,
        )
    return not missing


def _pair_master_lookup(
    loaded: dict[str, tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]],
) -> pd.DataFrame | None:
    bases = [raw[0] for raw in loaded.values() if raw[0] is not None and not raw[0].empty]
    if not bases:
        return None
    return pd.concat(bases, ignore_index=True)


def process_pair(pair_name: str, folders: list[str], exc_df: pd.DataFrame) -> bool:
    """Последовательная обработка (для отладки и --no-parallel)."""
    exc_keys = exception_keys(exc_df)
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]] = []
    failed: list[str] = []
    loaded: dict[str, tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]] = {}
    for folder in folders:
        loaded[folder] = load_sorg(folder)
    pair_lk = _pair_master_lookup(loaded)
    iterator = folders
    if tqdm is not None:
        iterator = tqdm(folders, desc=pair_name, unit="SOrg", leave=True)
    for folder in iterator:
        msg = f"SO {folder}: чтение и проверки…"
        if tqdm is not None and hasattr(iterator, "set_postfix_str"):
            iterator.set_postfix_str(msg)
        else:
            print(f"[new_access] {pair_name}: {msg}", flush=True)
        try:
            results.append(
                (
                    folder,
                    process_sorg(
                        folder,
                        exc_keys,
                        master_lookup=pair_lk,
                        loaded=loaded.get(folder),
                    ),
                )
            )
        except Exception as exc:
            failed.append(folder)
            print(f"[new_access] SO {folder}: ОШИБКА — {exc}", flush=True)
            traceback.print_exc()
    if not results:
        print(
            f"[new_access] пара {pair_name}: нет успешных SO — сохраняю пустой отчёт пары (2 листа SOrg)",
            flush=True,
        )
        save_pair_excel(pair_name, pd.DataFrame(), pd.DataFrame(), exc_df, folders)
        return False
    return _merge_sorg_results(pair_name, folders, results, exc_df, failed_folders=failed)


async def process_pair_async(pair_name: str, folders: list[str], exc_df: pd.DataFrame) -> bool:
    exc_keys = exception_keys(exc_df)
    workers = worker_count(len(folders), default_cap=2)

    if not parallel_enabled() or workers <= 1 or len(folders) <= 1:
        return process_pair(pair_name, folders, exc_df)

    print(
        f"[new_access] {pair_name}: параллельно {len(folders)} SOrg (workers={workers})",
        flush=True,
    )

    async def _load(folder: str) -> tuple[str, object]:
        try:
            return folder, await async_io(load_sorg, folder)
        except Exception as exc:
            print(f"[new_access] SO {folder}: ОШИБКА загрузки — {exc}", flush=True)
            traceback.print_exc()
            return folder, exc

    loaded_raw = await gather_limited([_load(f) for f in folders], limit=workers)
    loaded: dict[str, tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]] = {}
    failed: list[str] = []
    for folder, item in loaded_raw:
        if isinstance(item, Exception):
            failed.append(folder)
        else:
            loaded[folder] = item
    pair_lk = _pair_master_lookup(loaded)

    async def _one(folder: str) -> tuple[str, object]:
        print(f"[new_access] {pair_name}: SO {folder}: чтение и проверки…", flush=True)
        try:
            res = await async_io(
                process_sorg,
                folder,
                exc_keys,
                pair_lk,
                loaded.get(folder),
            )
            return folder, res
        except Exception as exc:
            print(f"[new_access] SO {folder}: ОШИБКА — {exc}", flush=True)
            traceback.print_exc()
            return folder, exc

    raw = await gather_limited([_one(f) for f in loaded], limit=workers)
    results: list[tuple[str, tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]]] = []
    for folder, item in raw:
        if isinstance(item, Exception):
            failed.append(folder)
        else:
            results.append((folder, item))
    if not results:
        print(
            f"[new_access] пара {pair_name}: нет успешных SO — сохраняю пустой отчёт пары (2 листа SOrg)",
            flush=True,
        )
        save_pair_excel(pair_name, pd.DataFrame(), pd.DataFrame(), exc_df, folders)
        return False
    return _merge_sorg_results(pair_name, folders, results, exc_df, failed_folders=failed)


async def _run_all_pairs(jobs: list[tuple[str, list[str]]], exc_df: pd.DataFrame) -> int:
    """Обрабатывает задания по очереди; каждое сохраняется сразу после своих SOrg."""
    errors = 0
    for job_name, folders in jobs:
        label = f"SOrg {job_name}" if len(folders) == 1 and job_name in folders else f"пара {job_name}"
        print(f"[new_access] === {label} ===", flush=True)
        try:
            ok = await process_pair_async(job_name, folders, exc_df)
            if not ok:
                errors += 1
        except Exception as exc:
            errors += 1
            print(f"[new_access] {label}: критическая ошибка — {exc}", flush=True)
            traceback.print_exc()
    return errors


def _all_so_folders() -> list[str]:
    """SOrg, для которых на диске есть папка и *Base*.xlsx."""
    return _existing_sorg_folders()


def _sorg_has_data(folder: str) -> bool:
    fp = BASE_DIR / folder
    if not fp.is_dir():
        return False
    return _get_file(fp, "*Base*.xlsx") is not None


def _existing_sorg_folders(candidates: list[str] | None = None) -> list[str]:
    known = candidates or ["3801", "3802", "3803", "3804", "3805", "3806"]
    return [f for f in known if _sorg_has_data(f)]


def _pair_data_status(folders: list[str]) -> tuple[list[str], list[str]]:
    ok = [f for f in folders if _sorg_has_data(f)]
    missing = [f for f in folders if f not in ok]
    return ok, missing


def _print_data_inventory() -> None:
    on_disk = _existing_sorg_folders()
    print(f"[new_access] SOrg с Base на диске: {', '.join(on_disk) if on_disk else 'нет'}", flush=True)
    for name, cfg in PAIRS.items():
        ok, miss = _pair_data_status(cfg["folders"])
        if not miss:
            print(f"[new_access]   пара {name}: готова ({', '.join(ok)})", flush=True)
        elif ok:
            print(
                f"[new_access]   пара {name}: частично — есть {', '.join(ok)}, нет {', '.join(miss)}",
                flush=True,
            )
        else:
            print(f"[new_access]   пара {name}: нет данных ({', '.join(cfg['folders'])})", flush=True)


def _parse_folders(raw: str) -> list[str]:
    if not raw:
        return []
    seen: set[str] = set()
    out: list[str] = []
    for token in raw.split(","):
        for f in re.findall(r"\d{4}", token.strip()):
            if f not in seen:
                out.append(f)
                seen.add(f)
    return out


def build_jobs(mode: str, folders_csv: str = "") -> list[tuple[str, list[str]]]:
    if mode == "pairs":
        jobs = [(name, cfg["folders"]) for name, cfg in PAIRS.items()]
        raw = folders_csv.strip()
        if raw:
            sel = set(_parse_folders(folders_csv))
            if sel:
                filtered = [(n, folders) for n, folders in jobs if sel.intersection(set(folders))]
                if filtered:
                    jobs = filtered
        return jobs
    if mode == "one_pair":
        name = _resolve_pair_name(folders_csv)
        if not name:
            known = ", ".join(PAIRS.keys())
            raise ValueError(f"Неизвестная пара: {folders_csv!r}. Доступны: {known}")
        pair_folders = PAIRS[name]["folders"]
        ok, miss = _pair_data_status(pair_folders)
        if miss:
            print(
                f"[new_access] пара {name}: нет выгрузок для SO {', '.join(miss)} "
                f"(листы будут пустые)",
                flush=True,
            )
        if not ok:
            raise ValueError(
                f"Пара {name}: нет *Base*.xlsx ни для одной SOrg ({', '.join(pair_folders)}). "
                f"Проверьте папку data."
            )
        return [(name, pair_folders)]
    all_f = _all_so_folders()
    if mode == "single":
        if not all_f:
            raise ValueError("Нет SOrg с *Base*.xlsx на диске — отдельная сборка невозможна.")
        return [(f, [f]) for f in all_f]
    sel = _parse_folders(folders_csv) or all_f
    if mode == "custom_single":
        ready = [f for f in sel if _sorg_has_data(f)]
        skipped = [f for f in sel if f not in ready]
        if skipped:
            print(f"[new_access] пропуск SOrg без Base: {', '.join(skipped)}", flush=True)
        if not ready:
            raise ValueError(f"Нет данных для SOrg: {', '.join(sel)}")
        return [(f, [f]) for f in ready]
    if mode == "custom_group":
        ready = [f for f in sel if _sorg_has_data(f)]
        skipped = [f for f in sel if f not in ready]
        if skipped:
            print(f"[new_access] пропуск SOrg без Base: {', '.join(skipped)}", flush=True)
        if not ready:
            raise ValueError(f"Нет данных для группы SOrg: {', '.join(sel)}")
        name = "_".join(ready)
        return [(f"custom_{name}", ready)]
    raise ValueError(f"Неизвестный режим: {mode}")


def _resolve_pair_name(raw: str) -> str | None:
    """1 / 3801_3803 / 3802 → имя пары из PAIRS."""
    text = raw.strip()
    if not text:
        return None
    if text in PAIRS:
        return text
    pair_names = list(PAIRS.keys())
    if text.isdigit():
        idx = int(text)
        if 1 <= idx <= len(pair_names):
            return pair_names[idx - 1]
    compact = text.replace(" ", "")
    for name in pair_names:
        if name == compact or name in compact:
            return name
    sel = set(_parse_folders(text))
    if sel:
        matches = [
            n for n, cfg in PAIRS.items() if sel.intersection(set(cfg["folders"]))
        ]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            exact = [n for n in matches if sel == set(PAIRS[n]["folders"])]
            if len(exact) == 1:
                return exact[0]
    return None


def _read_menu_line(prompt: str, default: str = "") -> str:
    try:
        raw = input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print(flush=True)
        raise SystemExit(0)
    return raw if raw else default


def _interactive_menu() -> tuple[str, str, bool, bool, int]:
    """Консольное меню режимов (без веба)."""
    all_f = ", ".join(_all_so_folders())
    print("\n=== PF BP-PY-ZY — выбор режима ===", flush=True)
    print("  1  Все 3 пары (3801_3803, 3802_3804, 3805_3806) — 3 файла", flush=True)
    print("  2  Одна пара — 1 файл на выбор:", flush=True)
    for i, (name, cfg) in enumerate(PAIRS.items(), start=1):
        ok, miss = _pair_data_status(cfg["folders"])
        if not miss:
            tag = "данные есть"
        elif ok:
            tag = f"частично, нет {', '.join(miss)}"
        else:
            tag = "НЕТ данных"
        print(f"       {i}  {name} ({', '.join(cfg['folders'])}) — {tag}", flush=True)
    print("  3  Все SOrg по отдельности — по 1 файлу на каждую", flush=True)
    print(f"     ({all_f})", flush=True)
    print("  4  Выбранные SOrg по отдельности (через запятую)", flush=True)
    print("  5  Своя группа SOrg в одном файле", flush=True)
    print("  0  Выход", flush=True)

    choice = _read_menu_line("Выбор [1]: ", "1")
    if choice == "0":
        raise SystemExit(0)

    mode = "pairs"
    folders = ""
    if choice == "1":
        mode = "pairs"
        folders = _read_menu_line(
            "Только пары, содержащие SOrg (Enter = все 3 пары): ",
        )
    elif choice == "2":
        mode = "one_pair"
        default_pair = _resolve_pair_name("1") or list(PAIRS.keys())[0]
        pair_pick = _read_menu_line(
            f"Какая пара [1={default_pair}]: ",
            "1",
        )
        resolved = _resolve_pair_name(pair_pick)
        if not resolved:
            print(f"[new_access] не удалось распознать пару «{pair_pick}» — выход", flush=True)
            raise SystemExit(1)
        ok, miss = _pair_data_status(PAIRS[resolved]["folders"])
        if not ok:
            print(
                f"[new_access] пара {resolved}: нет *Base*.xlsx — выберите другую пару или положите выгрузки",
                flush=True,
            )
            raise SystemExit(1)
        if miss:
            print(
                f"[new_access] внимание: для {resolved} нет SO {', '.join(miss)} — эти листы будут пустые",
                flush=True,
            )
        folders = resolved
        print(f"[new_access] выбрана пара {resolved} ({', '.join(PAIRS[resolved]['folders'])})", flush=True)
    elif choice == "3":
        mode = "single"
        if not _all_so_folders():
            print("[new_access] нет SOrg с Base на диске — режим 3 недоступен", flush=True)
            raise SystemExit(1)
    elif choice == "4":
        mode = "custom_single"
        if not all_f:
            print("[new_access] нет SOrg с Base на диске", flush=True)
            raise SystemExit(1)
        folders = _read_menu_line(f"SOrg через запятую [{all_f}]: ", all_f)
        if not _parse_folders(folders):
            folders = all_f
    elif choice == "5":
        mode = "custom_group"
        if not all_f:
            print("[new_access] нет SOrg с Base на диске", flush=True)
            raise SystemExit(1)
        folders = _read_menu_line(f"SOrg через запятую [{all_f}]: ", all_f)
        if not _parse_folders(folders):
            print("[new_access] не указаны SOrg — выход", flush=True)
            raise SystemExit(1)
    else:
        print(f"[new_access] неизвестный пункт «{choice}» — режим 1 (пары)", flush=True)

    print("\n--- Настройки прогона ---", flush=True)

    print("\n▸ DuckDB staging [Y/n]", flush=True)
    print(
        "  Зачем: один раз прочитать все xlsx и положить в временную БД data/staging.duckdb,\n"
        "  дальше проверки берут данные из БД, а не с диска.",
        flush=True,
    )
    print(
        "  Плюсы включения (Y):\n"
        "    • xlsx не открывается заново на каждой SOrg — меньше I/O, стабильнее прогон\n"
        "    • обход BadZipFile: если pandas не читает файл (часто 3804 BP), на Windows\n"
        "      срабатывает запасное чтение через установленный Excel\n"
        "    • все 6 SOrg грузятся один раз в начале — удобно для полного прогона пар\n"
        "    • в конце staging-таблицы удаляются автоматически",
        flush=True,
    )
    print(
        "  Минусы / когда n:\n"
        "    • первый этап (загрузка в DuckDB) долгий — 10–20+ мин на больших выгрузках\n"
        "    • нужен пакет duckdb (pip install duckdb)\n"
        "    • для быстрой проверки одной SOrg проще читать xlsx напрямую",
        flush=True,
    )
    print("  Рекомендация: Y; для одной пары staging грузит только 2 SOrg.", flush=True)
    staging = _read_menu_line("DuckDB staging? [Y/n]: ", "Y")
    no_staging = staging.lower() in ("n", "no", "0", "н", "нет")

    print("\n▸ Параллельная обработка [Y/n]", flush=True)
    print(
        "  Зачем: в одной задаче (например пара 3801+3803) обе SOrg считаются\n"
        "  одновременно или строго по очереди.",
        flush=True,
    )
    print(
        "  Плюсы включения (Y):\n"
        "    • быстрее на многоядерном ПК — 3801 и 3803 идут параллельно\n"
        "    • для пары из 2 SOrg экономия времени почти в 2 раза на этапе проверок",
        flush=True,
    )
    print(
        "  Минусы / когда n:\n"
        "    • выше расход RAM (две большие таблицы в памяти сразу)\n"
        "    • проще читать лог и отлаживать ошибки по одной SOrg",
        flush=True,
    )
    print("  Рекомендация: Y; при нехватке памяти или ошибках — n.", flush=True)
    parallel = _read_menu_line("Параллельная обработка? [Y/n]: ", "Y")
    no_parallel = parallel.lower() in ("n", "no", "0", "н", "нет")

    workers = 0
    if not no_parallel:
        print("\n▸ Workers (число параллельных SOrg в одной задаче)", flush=True)
        print(
            "  Зачем: ограничить, сколько SOrg одновременно обрабатываются внутри\n"
            "  одной задачи (одной пары или одной группы).",
            flush=True,
        )
        print(
            "  Плюсы / как выбирать:\n"
            "    • 0 (авто) — скрипт сам подберёт (для пары 3801+3803 обычно 2 потока)\n"
            "    • 1 — одна SOrg за раз; как «без параллели», минимум RAM\n"
            "    • 2 — оптимально для пары из двух SOrg (3801 и 3803 одновременно)\n"
            "    • 3–4 — имеет смысл в режиме 4 (своя группа из 3–4 SOrg в одном файле)",
            flush=True,
        )
        print(
            "  Когда менять с 0:\n"
            "    • ПК тормозит / мало RAM → поставьте 1\n"
            "    • большая группа SOrg и мощный ПК → 3 или 4",
            flush=True,
        )
        print("  Рекомендация: 0 (авто) для режима «пары».", flush=True)
        w_raw = _read_menu_line("Workers [0 = авто]: ", "0")
        try:
            workers = max(0, int(w_raw))
        except ValueError:
            workers = 0

    try:
        jobs = build_jobs(mode, folders)
    except ValueError as exc:
        print(f"[new_access] {exc}", flush=True)
        raise SystemExit(1)
    if not jobs:
        print("[new_access] нет задач для выбранного режима", flush=True)
        raise SystemExit(1)

    print(f"\n[new_access] режим={mode}, задач={len(jobs)}:", flush=True)
    for name, fl in jobs:
        print(f"  • {name}: {', '.join(fl)}", flush=True)
    print(flush=True)
    return mode, folders, no_staging, no_parallel, workers


def main() -> int:
    parser = argparse.ArgumentParser(description="PF BP-PY-ZY — логика Access")
    parser.add_argument(
        "--mode",
        choices=["pairs", "one_pair", "single", "custom_single", "custom_group"],
        default=None,
        help="Режим (без --no-menu откроется консольное меню)",
    )
    parser.add_argument(
        "--folders",
        default="",
        help="SOrg через запятую, имя пары (3802_3804) для one_pair, или фильтр пар",
    )
    parser.add_argument(
        "--no-menu",
        action="store_true",
        help="Не показывать меню; использовать --mode и флаги",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=0,
        help="Параллельных SOrg в паре (0 = авто, 1 = последовательно)",
    )
    parser.add_argument("--no-parallel", action="store_true", help="Отключить параллельное чтение и обработку")
    parser.add_argument(
        "--no-staging",
        action="store_true",
        help="Не использовать DuckDB: читать xlsx напрямую (как раньше)",
    )
    args = parser.parse_args()

    paths = load_runtime_paths_dict()
    script_root = Path(__file__).resolve().parent
    local_data = (script_root / "data").resolve()
    data_dir = paths["data_dir"].resolve()
    print(f"[new_access] data_dir:  {paths['data_dir']}", flush=True)
    print(f"[new_access] base_dir:  {paths['base_dir']}", flush=True)
    print(f"[new_access] result:    {paths['output_dir']}", flush=True)
    if data_dir != local_data and script_root not in data_dir.parents:
        print(
            f"[new_access] ВНИМАНИЕ: data_dir не рядом со скриптом ({script_root}).\n"
            f"  Проверьте runtime_paths.json — возможно читаете/пишете не ту папку data.",
            flush=True,
        )

    if not BASE_DIR.exists():
        print(f"[new_access] Нет каталога: {BASE_DIR}", flush=True)
        return 1

    _print_data_inventory()

    use_menu = not args.no_menu and args.mode is None and sys.stdin.isatty()
    if use_menu:
        mode, folders_csv, no_staging, no_parallel, workers = _interactive_menu()
    else:
        mode = args.mode or "pairs"
        folders_csv = args.folders
        no_staging = args.no_staging
        no_parallel = args.no_parallel
        workers = args.workers

    if no_parallel:
        os.environ["REPORTS_PARALLEL"] = "0"
    elif workers > 0:
        os.environ["REPORTS_WORKERS"] = str(workers)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    exc_df = collect_and_persist_global_exception(BASE_DIR, OUTPUT_DIR)

    try:
        jobs = build_jobs(mode, folders_csv)
    except ValueError as exc:
        print(f"[new_access] {exc}", flush=True)
        return 1
    if not jobs:
        print("[new_access] нет задач для выбранного режима", flush=True)
        return 1

    all_sorgs = sorted({so for _, folders in jobs for so in folders})
    par = "вкл" if parallel_enabled() else "выкл"
    w = worker_count(2, default_cap=2)
    staging_mode = "DuckDB" if not no_staging else "xlsx"
    print(
        f"[new_access] режим={mode} | задач={len(jobs)} | "
        f"источник={staging_mode} | parallel={par} workers≈{w}",
        flush=True,
    )
    for job_name, job_folders in jobs:
        print(f"  • {job_name}: {', '.join(job_folders)}", flush=True)
    if mode in ("single", "custom_single"):
        print(
            f"[new_access] отчёты: {OUTPUT_DIR}/<SOrg>/Check PF BP-PY-ZY <SOrg> <дата>.xlsx",
            flush=True,
        )
    elif mode == "custom_group":
        print(
            f"[new_access] отчёт: {OUTPUT_DIR}/custom_<SOrg>/Check PF ... (один файл на группу)",
            flush=True,
        )

    exit_code = 0
    try:
        if not no_staging:
            os.environ["REPORTS_PARALLEL"] = "0"
            start_staging(all_sorgs)
        pair_errors = asyncio.run(_run_all_pairs(jobs, exc_df))
        date_str = datetime.now().strftime("%d.%m.%Y")
        print(f"[new_access] итоговые файлы в {OUTPUT_DIR}:", flush=True)
        for job_name, job_folders in jobs:
            pair_file = OUTPUT_DIR / job_name / f"Check PF BP-PY-ZY {job_name} {date_str}.xlsx"
            mark = "OK" if pair_file.is_file() else "НЕТ ФАЙЛА"
            sorg_note = f" ({len(job_folders)} SOrg)" if len(job_folders) > 1 else ""
            print(f"  [{mark}] {pair_file.name}{sorg_note}", flush=True)
        if pair_errors:
            exit_code = 1
            print(f"[new_access] готово с ошибками в {pair_errors} заданиях", flush=True)
        else:
            print("[new_access] готово", flush=True)
    except Exception:
        traceback.print_exc()
        exit_code = 1
    finally:
        stop_staging()
        shutdown_executor()
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
