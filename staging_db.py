# -*- coding: utf-8 -*-
"""
Временная DuckDB для выгрузок макроса: xlsx → таблицы в начале прогона, DROP в конце.
Обходит BadZipFile в pandas (чтение через Excel COM на Windows при ошибке).
"""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path
from typing import Callable, Iterable

import duckdb
import pandas as pd

from build_checks import (
    BASE_DIR,
    _get_file,
    _get_files,
    _normalize_base_df,
    _normalize_partner_df,
    _read_excel_checked,
    load_runtime_paths_dict,
)

PARTS = ("base", "bp", "py", "zy")
TABLE_PREFIX = "so_"

_active: "StagingDB | None" = None


def staging_path() -> Path:
    return load_runtime_paths_dict()["data_dir"] / "staging.duckdb"


def staging_active() -> bool:
    return _active is not None


def get_staging() -> "StagingDB":
    if _active is None:
        raise RuntimeError("staging DB не инициализирована")
    return _active


def _table_name(folder: str, part: str) -> str:
    return f"{TABLE_PREFIX}{folder}_{part}"


def _file_mb(path: Path | None) -> str:
    if not path or not path.is_file():
        return "?"
    return f"{path.stat().st_size / (1024 * 1024):.1f} MB"


def _read_excel_with_com_fallback(path: Path, *, kind: str, folder: str) -> pd.DataFrame:
    try:
        return _read_excel_checked(path, kind=kind, folder=folder, allow_com=False)
    except Exception as first_exc:
        if sys.platform != "win32":
            raise first_exc
        print(
            f"[staging] SO {folder} {kind}: pandas не прочитал ({type(first_exc).__name__}), "
            f"пробую через Excel…",
            flush=True,
        )
        return _read_excel_checked(path, kind=kind, folder=folder, allow_com=True)


def read_sorg_raw_from_excel(
    folder: str,
) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
    """Нормализованные Base/BP/PY/ZY из xlsx (без DelDup — его делает вызывающий скрипт)."""
    fp = BASE_DIR / folder
    f_base = _get_file(fp, "*Base*.xlsx")
    if not f_base:
        return pd.DataFrame(), None, None, None

    print(f"[staging] SO {folder}: читаю Base {f_base.name} ({_file_mb(f_base)})…", flush=True)
    t0 = time.perf_counter()
    base = _normalize_base_df(_read_excel_with_com_fallback(f_base, kind="Base", folder=folder), folder)
    print(f"[staging] SO {folder}: Base {len(base)} строк, {time.perf_counter() - t0:.0f} с", flush=True)

    def _part(label: str, pattern: str) -> pd.DataFrame | None:
        files = _get_files(fp, pattern)
        if not files:
            print(f"[staging] SO {folder}: {label} не найден", flush=True)
            return None
        parts: list[pd.DataFrame] = []
        t1 = time.perf_counter()
        for path in files:
            print(f"[staging] SO {folder}: читаю {label} {path.name} ({_file_mb(path)})…", flush=True)
            df = _normalize_partner_df(
                _read_excel_with_com_fallback(path, kind=label, folder=folder),
                folder,
                kind=label,
            )
            if df is not None and not df.empty:
                parts.append(df)
        if not parts:
            print(f"[staging] SO {folder}: {label} пустой", flush=True)
            return None
        out = pd.concat(parts, ignore_index=True).drop_duplicates(subset=["KUNNR", "KTONR"], keep="first")
        print(f"[staging] SO {folder}: {label} {len(out)} строк, {time.perf_counter() - t1:.0f} с", flush=True)
        return out

    return base, _part("BP", "*BP*.xlsx"), _part("PY", "*PY*.xlsx"), _part("ZY", "*ZY*.xlsx")


class StagingDB:
    def __init__(self, path: Path) -> None:
        self.path = path
        self._conn: duckdb.DuckDBPyConnection | None = None
        self._lock = threading.Lock()
        self._folders: list[str] = []

    def open(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = duckdb.connect(str(self.path))
        self._drop_staging_tables()

    def _drop_staging_tables(self) -> None:
        assert self._conn is not None
        rows = self._conn.execute(
            """
            SELECT table_name FROM information_schema.tables
            WHERE table_schema = 'main' AND table_name LIKE ?
            """,
            [f"{TABLE_PREFIX}%"],
        ).fetchall()
        for (name,) in rows:
            self._conn.execute(f'DROP TABLE IF EXISTS "{name}"')

    def _write_table(self, name: str, df: pd.DataFrame | None) -> None:
        assert self._conn is not None
        with self._lock:
            if df is None or df.empty:
                self._conn.execute(f'DROP TABLE IF EXISTS "{name}"')
                return
            prepared = df.fillna("").astype(str).replace(
                {"nan": "", "None": "", "<NA>": "", "none": ""}
            )
            self._conn.register("_staging_df", prepared)
            try:
                self._conn.execute(f'CREATE OR REPLACE TABLE "{name}" AS SELECT * FROM _staging_df')
            finally:
                self._conn.unregister("_staging_df")

    def _read_table(self, name: str) -> pd.DataFrame | None:
        assert self._conn is not None
        with self._lock:
            exists = self._conn.execute(
                "SELECT COUNT(*) FROM information_schema.tables WHERE table_name = ?",
                [name],
            ).fetchone()[0]
            if not exists:
                return None
            df = self._conn.execute(f'SELECT * FROM "{name}"').df()
        if df.empty:
            return None
        return df.fillna("").replace({"nan": "", "None": "", "<NA>": "", "none": ""})

    def populate(
        self,
        folders: Iterable[str],
        *,
        loader: Callable[[str], tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]],
    ) -> None:
        for folder in folders:
            if not (BASE_DIR / folder).is_dir():
                print(f"[staging] SO {folder}: папки нет — пропуск", flush=True)
                continue
            base, bp, py, zy = loader(folder)
            if base.empty:
                print(f"[staging] SO {folder}: нет Base — пропуск", flush=True)
                continue
            for part, df in (("base", base), ("bp", bp), ("py", py), ("zy", zy)):
                self._write_table(_table_name(folder, part), df)
            self._folders.append(folder)
            print(f"[staging] SO {folder}: залито в DuckDB", flush=True)

    def load_raw(
        self, folder: str
    ) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]:
        base = self._read_table(_table_name(folder, "base"))
        if base is None:
            return pd.DataFrame(), None, None, None
        return (
            base,
            self._read_table(_table_name(folder, "bp")),
            self._read_table(_table_name(folder, "py")),
            self._read_table(_table_name(folder, "zy")),
        )

    def cleanup(self) -> None:
        if self._conn is None:
            return
        try:
            with self._lock:
                self._drop_staging_tables()
        finally:
            self._conn.close()
            self._conn = None
        print(f"[staging] таблицы удалены ({self.path})", flush=True)


def start_staging(
    folders: Iterable[str],
    *,
    loader: Callable[[str], tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None]] | None = None,
) -> StagingDB:
    global _active
    if _active is not None:
        _active.cleanup()
    path = staging_path()
    print(f"[staging] DuckDB: {path}", flush=True)
    db = StagingDB(path)
    db.open()
    load_fn = loader or read_sorg_raw_from_excel
    unique = []
    seen: set[str] = set()
    for f in folders:
        if f not in seen:
            unique.append(f)
            seen.add(f)
    print(f"[staging] загрузка xlsx → БД: {', '.join(unique)}", flush=True)
    t0 = time.perf_counter()
    db.populate(unique, loader=load_fn)
    print(f"[staging] готово за {time.perf_counter() - t0:.0f} с", flush=True)
    _active = db
    return db


def stop_staging() -> None:
    global _active
    if _active is not None:
        _active.cleanup()
        _active = None
