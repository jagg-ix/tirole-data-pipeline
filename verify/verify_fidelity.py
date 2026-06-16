#!/usr/bin/env python3
r"""verify_fidelity.py -- prove the generated Excel splits and the SQLite DB are
a faithful, lossless image of the Tirole et al. source workbooks.

The pipeline is an ETL chain:

    source_data/<wb>.xlsx
        --ex2csv split-->        pipeline/output/<wb>/<sheet>.xlsx
        --tidy converter-->      pipeline/tidy_output/<group>/<fig>_tidy.csv
        --build_db_from_tidy-->  output/double_slit.sqlite3   (one table per figure)

This script verifies every link with three checks (plus one optional check):

  CHECK 1  SPLIT      every source sheet equals its pipeline/output/<wb>/<sheet>.xlsx
                      split file, cell for cell (exact -- the split copies values,
                      it only normalises column-name spaces to underscores).

  CHECK 3  LOAD       every figure table in output/double_slit.sqlite3 equals its
                      pipeline/tidy_output/.../<fig>_tidy.csv, cell for cell (exact),
                      and dataset_registry.nrows agrees with the table row count.

  CHECK 4  CONSERVE   every numeric value stored in each DB table traces back to a
                      cell in the source workbook, within double-precision round-trip
                      tolerance, after inverting the single declared unit transform
                      (Fig_2g delay: seconds -> femtoseconds, x 1e15).  No value is
                      invented; the tidy reshaping only rearranges source numbers.

  CHECK 2  REPRODUCE  (optional, --reproduce) re-run the whole pipeline into a
                      temporary tree and confirm it reproduces the committed DB
                      table-for-table.  Needs the tidy converters' imports
                      (pandas, openpyxl, xlsxwriter, matplotlib).

The tidy step (CHECK 1 -> CHECK 3) reshapes wide source tables into long form; it is
not shape-preserving by design, so CHECK 4 is the value-level link across it.  Taken
together the checks establish: DB values  subset of  split values  ==  source values.

Exit code is 0 iff every enabled check passes.

Run AFTER the pipeline has produced output/ and pipeline/tidy_output/:

    python build_and_verify_pipeline.py run        # from pipeline/   (writes splits + tidy + a normalized DB)
    python build_db_from_tidy_csv.py \             # from loader/     (writes the canonical per-figure DB)
        --tidy_dir ../pipeline/tidy_output --db ../output/double_slit.sqlite3
    python verify/verify_fidelity.py               # from repo root
"""
from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd

REPO = Path(__file__).resolve().parent.parent
SRC_DIR = REPO / "source_data"
SPLIT_DIR = REPO / "pipeline" / "output"
TIDY_DIR = REPO / "pipeline" / "tidy_output"
DEFAULT_DB = REPO / "output" / "double_slit.sqlite3"

# The five source workbooks the pipeline consumes (matches EXPECTED_FILES in
# build_and_verify_pipeline.py).  source_data/ also ships an underscore-named
# duplicate of the Fig 1 workbook and the paper PDF, neither of which the
# pipeline reads.
SOURCE_WORKBOOKS = [
    "Extended_Data_Fig_1_Source data.xlsx",
    "Extended_Data_Fig_2_Source_Data.xlsx",
    "Extended_Data_Fig_3_Source_data.xlsx",
    "Fig_1_Source_data.xlsx",
    "Fig_2_Source_data.xlsx",
]

# Columns the tidy converters *add* as labels (never sourced from a numeric cell).
# They are text, so pd.to_numeric drops them anyway; listed for clarity.
LABEL_COLUMNS = {"Series", "series", "Pump"}

# Declared, intentional unit transforms applied during tidying.  (table, column) ->
# function mapping a stored DB value back to the source-cell value it came from.
INVERSE_TRANSFORMS = {
    ("Fig_2g", "Delay_fs"): lambda v: v / 1e15,  # seconds -> femtoseconds
}

RTOL = 1e-9
ATOL = 1e-12


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def numeric_cells(path: Path) -> np.ndarray:
    """Every finite numeric value in an .xlsx sheet, reading all physical cells
    (header=None) so header/grid rows are included."""
    df = pd.read_excel(path, header=None)
    vals = pd.to_numeric(pd.Series(df.to_numpy().ravel()), errors="coerce").to_numpy(float)
    return np.sort(vals[np.isfinite(vals)])


def frames_equal(a: pd.DataFrame, b: pd.DataFrame) -> tuple[bool, float, str]:
    """Cell-for-cell equality of two DataFrames (NaN==NaN), numeric to full
    precision.  Returns (ok, max_abs_numeric_diff, reason)."""
    if list(a.columns) != list(b.columns):
        return False, float("nan"), f"column mismatch {list(a.columns)[:4]}... vs {list(b.columns)[:4]}..."
    if a.shape != b.shape:
        return False, float("nan"), f"shape {a.shape} vs {b.shape}"
    maxdiff = 0.0
    for c in a.columns:
        ca, cb = a[c], b[c]
        if pd.api.types.is_numeric_dtype(ca) and pd.api.types.is_numeric_dtype(cb):
            x, y = ca.to_numpy(float), cb.to_numpy(float)
            both_nan = np.isnan(x) & np.isnan(y)
            d = np.where(both_nan, 0.0, np.abs(x - y))
            if np.any(np.isnan(d)):
                return False, float("nan"), f"column {c!r}: NaN pattern differs"
            maxdiff = max(maxdiff, float(d.max()) if d.size else 0.0)
        else:
            sa = ca.astype(object).where(ca.notna(), None).astype(str)
            sb = cb.astype(object).where(cb.notna(), None).astype(str)
            if not sa.equals(sb):
                return False, float("nan"), f"column {c!r}: text values differ"
    return True, maxdiff, ""


def nearest_rel_error(values: np.ndarray, sorted_source: np.ndarray) -> np.ndarray:
    """For each value, the relative distance to the closest cell in sorted_source."""
    if sorted_source.size == 0:
        return np.full(values.shape, np.inf)
    idx = np.clip(np.searchsorted(sorted_source, values), 1, len(sorted_source) - 1)
    lo = np.abs(values - sorted_source[idx - 1])
    hi = np.abs(values - sorted_source[idx])
    nearest = np.minimum(lo, hi)
    denom = np.maximum(ATOL, np.abs(values))
    return nearest / denom


def find_one(root: Path, pattern: str) -> Path | None:
    hits = sorted(root.rglob(pattern))
    return hits[0] if hits else None


# --------------------------------------------------------------------------- #
# checks
# --------------------------------------------------------------------------- #
def check_split() -> tuple[bool, list[dict]]:
    print("\n[CHECK 1] SPLIT  source sheet  ==  pipeline/output/<wb>/<sheet>.xlsx")
    rows, ok_all = [], True
    for fname in SOURCE_WORKBOOKS:
        src = SRC_DIR / fname
        if not src.exists():
            rows.append({"item": fname, "status": "FAIL", "detail": "source workbook missing"})
            ok_all = False
            continue
        wb_stem = src.stem.replace(" ", "_")
        for sheet in pd.ExcelFile(src).sheet_names:
            split = SPLIT_DIR / wb_stem / f"{sheet.replace(' ', '_')}.xlsx"
            label = f"{wb_stem}/{sheet}"
            if not split.exists():
                rows.append({"item": label, "status": "FAIL", "detail": "split file missing -- run the pipeline"})
                ok_all = False
                continue
            src_vals = numeric_cells_sheet(src, sheet)
            split_vals = numeric_cells(split)
            if src_vals.shape != split_vals.shape:
                rows.append({"item": label, "status": "FAIL",
                             "detail": f"value count {src_vals.size} (source) vs {split_vals.size} (split)"})
                ok_all = False
                continue
            maxdiff = float(np.abs(src_vals - split_vals).max()) if src_vals.size else 0.0
            status = "PASS" if maxdiff <= ATOL else "FAIL"
            ok_all &= status == "PASS"
            rows.append({"item": label, "status": status,
                         "detail": f"{src_vals.size} values, max|diff|={maxdiff:.2e}"})
    _print_rows(rows)
    return ok_all, rows


def numeric_cells_sheet(path: Path, sheet: str) -> np.ndarray:
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    vals = pd.to_numeric(pd.Series(df.to_numpy().ravel()), errors="coerce").to_numpy(float)
    return np.sort(vals[np.isfinite(vals)])


def check_load(db_path: Path) -> tuple[bool, list[dict]]:
    print("\n[CHECK 3] LOAD   tidy CSV  ==  table in", db_path.name)
    rows, ok_all = [], True
    con = sqlite3.connect(str(db_path))
    tables = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
    registry = {}
    if "dataset_registry" in tables:
        registry = {r["dataset"]: r for _, r in
                    pd.read_sql("SELECT * FROM dataset_registry", con).iterrows()}
    for t in tables:
        if t == "dataset_registry":
            continue
        db = pd.read_sql(f'SELECT * FROM "{t}"', con)
        csv = find_one(TIDY_DIR, f"{t}_tidy.csv")
        if csv is None:
            rows.append({"item": t, "status": "FAIL", "detail": "no matching tidy CSV"})
            ok_all = False
            continue
        ok, maxdiff, reason = frames_equal(db, pd.read_csv(csv))
        detail = f"{db.shape[0]}x{db.shape[1]}, max|diff|={maxdiff:.2e}" if ok else reason
        if ok and t in registry and int(registry[t]["nrows"]) != len(db):
            ok, detail = False, f"dataset_registry.nrows={registry[t]['nrows']} != {len(db)}"
        ok_all &= ok
        rows.append({"item": t, "status": "PASS" if ok else "FAIL", "detail": detail})
    con.close()
    _print_rows(rows)
    return ok_all, rows


def check_conserve(db_path: Path) -> tuple[bool, list[dict]]:
    print("\n[CHECK 4] CONSERVE  every DB value traces to a source cell "
          f"(rtol={RTOL:g}, declared transforms inverted)")
    rows, ok_all = [], True
    con = sqlite3.connect(str(db_path))
    tables = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name!='dataset_registry' ORDER BY name")]
    for t in tables:
        split = find_one(SPLIT_DIR, f"{t}.xlsx")
        if split is None:
            rows.append({"item": t, "status": "FAIL", "detail": "no matching split xlsx"})
            ok_all = False
            continue
        source = numeric_cells(split)
        db = pd.read_sql(f'SELECT * FROM "{t}"', con)
        collected = []
        for col in db.columns:
            if col in LABEL_COLUMNS:
                continue
            vals = pd.to_numeric(db[col], errors="coerce").to_numpy(float)
            vals = vals[np.isfinite(vals)]
            inv = INVERSE_TRANSFORMS.get((t, col))
            if inv is not None:
                vals = inv(vals)
            collected.append(vals)
        values = np.concatenate(collected) if collected else np.array([])
        rel = nearest_rel_error(values, source)
        n_bad = int((rel > RTOL).sum())
        maxrel = float(rel.max()) if rel.size else 0.0
        status = "PASS" if n_bad == 0 else "FAIL"
        ok_all &= status == "PASS"
        note = " [Delay_fs/1e15]" if t == "Fig_2g" else ""
        rows.append({"item": t, "status": status,
                     "detail": f"{values.size} values, max_rel={maxrel:.2e}, off={n_bad}{note}"})
    con.close()
    _print_rows(rows)
    return ok_all, rows


def check_reproduce(db_path: Path) -> tuple[bool, list[dict]]:
    """Re-run the pipeline into a temp tree and diff the rebuilt DB against the
    committed one, table by table."""
    import os
    import shutil
    import subprocess
    import tempfile

    print("\n[CHECK 2] REPRODUCE  re-run pipeline in a temp tree, diff rebuilt DB")
    rows = []
    tmp = Path(tempfile.mkdtemp(prefix="tirole_repro_"))
    try:
        pipe = tmp / "pipeline"
        pipe.mkdir()
        for s in (REPO / "pipeline").glob("*.py"):
            shutil.copy(s, pipe / s.name)
        env = dict(os.environ, DATA_DIR=str(SRC_DIR), MPLBACKEND="Agg")
        r = subprocess.run([sys.executable, "build_and_verify_pipeline.py", "run"],
                           cwd=pipe, env=env, capture_output=True, text=True)
        if r.returncode != 0:
            tail = (r.stderr or r.stdout).strip().splitlines()[-1:] or ["(no output)"]
            rows.append({"item": "pipeline run", "status": "FAIL", "detail": tail[-1]})
            _print_rows(rows)
            return False, rows
        rebuilt = tmp / "double_slit.sqlite3"
        subprocess.run([sys.executable, str(REPO / "loader" / "build_db_from_tidy_csv.py"),
                        "--tidy_dir", str(pipe / "tidy_output"), "--db", str(rebuilt)],
                       cwd=pipe, env=env, check=True, capture_output=True, text=True)
        con_a, con_b = sqlite3.connect(str(db_path)), sqlite3.connect(str(rebuilt))
        tables = [x[0] for x in con_a.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name!='dataset_registry' ORDER BY name")]
        ok_all = True
        for t in tables:
            a = pd.read_sql(f'SELECT * FROM "{t}"', con_a)
            try:
                b = pd.read_sql(f'SELECT * FROM "{t}"', con_b)
            except Exception:
                rows.append({"item": t, "status": "FAIL", "detail": "absent in rebuilt DB"})
                ok_all = False
                continue
            ok, maxdiff, reason = frames_equal(a, b)
            ok_all &= ok
            rows.append({"item": t, "status": "PASS" if ok else "FAIL",
                         "detail": f"max|diff|={maxdiff:.2e}" if ok else reason})
        con_a.close()
        con_b.close()
        _print_rows(rows)
        return ok_all, rows
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------------------- #
# output
# --------------------------------------------------------------------------- #
def _print_rows(rows: list[dict]) -> None:
    for r in rows:
        mark = "ok " if r["status"] == "PASS" else "XX "
        print(f"    {mark}{r['item']:24s} {r['status']:4s} {r['detail']}")


def main() -> int:
    ap = argparse.ArgumentParser(description="100% fidelity check: source XLSX -> SQLite.")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB,
                    help="canonical per-figure SQLite DB (default: output/double_slit.sqlite3)")
    ap.add_argument("--reproduce", action="store_true",
                    help="also re-run the whole pipeline in a temp tree and diff (CHECK 2)")
    ap.add_argument("--report", type=Path, default=REPO / "verify" / "fidelity_report.csv",
                    help="write a per-item CSV report here")
    args = ap.parse_args()

    if not args.db.exists():
        print(f"ERROR: DB not found: {args.db}\n"
              f"Run the pipeline first (see this file's module docstring).", file=sys.stderr)
        return 2
    if not SPLIT_DIR.exists() or not TIDY_DIR.exists():
        print(f"ERROR: generated artifacts missing under {SPLIT_DIR} / {TIDY_DIR}.\n"
              f"Run the pipeline first (see this file's module docstring).", file=sys.stderr)
        return 2

    print(f"Repo : {REPO}")
    print(f"DB   : {args.db}")

    results, all_rows = {}, []
    for name, ok_rows in [
        ("CHECK 1 SPLIT", check_split()),
        ("CHECK 3 LOAD", check_load(args.db)),
        ("CHECK 4 CONSERVE", check_conserve(args.db)),
    ]:
        ok, rows = ok_rows
        results[name] = ok
        for r in rows:
            all_rows.append({"check": name, **r})

    if args.reproduce:
        ok, rows = check_reproduce(args.db)
        results["CHECK 2 REPRODUCE"] = ok
        for r in rows:
            all_rows.append({"check": "CHECK 2 REPRODUCE", **r})

    args.report.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(all_rows).to_csv(args.report, index=False)

    print("\n" + "=" * 66)
    for name in sorted(results):
        print(f"  {name:20s} {'PASS' if results[name] else 'FAIL'}")
    overall = all(results.values())
    n_items = len(all_rows)
    n_pass = sum(r["status"] == "PASS" for r in all_rows)
    print("-" * 66)
    print(f"  FIDELITY: {'PASS' if overall else 'FAIL'}  "
          f"({n_pass}/{n_items} items, source XLSX -> SQLite, full double precision)")
    print(f"  report: {args.report}")
    print("=" * 66)
    return 0 if overall else 1


if __name__ == "__main__":
    raise SystemExit(main())
