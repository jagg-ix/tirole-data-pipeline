#!/usr/bin/env python3
"""Build a minimal SQLite DB used by catsim from tidy CSVs.

Purpose: Avoid XLSX processing during reproducibility runs.
Input: directory containing tidy_output/**/_tidy.csv files.
Output: sqlite3 DB with a simple schema used by prediction scripts.

This is intentionally conservative: it loads only columns present in the tidy CSVs.
"""

import argparse
import os
import sqlite3
from pathlib import Path

import pandas as pd


def infer_table_name(csv_path: Path) -> str:
    name = csv_path.stem
    # Typical: Fig_2f_tidy.csv -> Fig_2f
    if name.endswith("_tidy"):
        name = name[: -len("_tidy")]
    return name


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tidy_dir", required=True, help="Root tidy_output directory")
    ap.add_argument("--db", required=True, help="Output sqlite3 path")
    args = ap.parse_args()

    tidy_dir = Path(args.tidy_dir)
    db_path = Path(args.db)
    db_path.parent.mkdir(parents=True, exist_ok=True)

    csvs = sorted(tidy_dir.rglob("*_tidy.csv"))
    if not csvs:
        raise SystemExit(f"No *_tidy.csv found under {tidy_dir}")

    if db_path.exists():
        db_path.unlink()

    con = sqlite3.connect(str(db_path))
    cur = con.cursor()

    # Simple registry
    cur.execute(
        """CREATE TABLE dataset_registry (
            dataset TEXT PRIMARY KEY,
            relpath TEXT,
            nrows INTEGER
        )"""
    )

    # Load each tidy CSV into a table named after the figure.
    for csv_path in csvs:
        dataset = infer_table_name(csv_path)
        df = pd.read_csv(csv_path)
        df.to_sql(dataset, con, if_exists="replace", index=False)
        cur.execute(
            "INSERT INTO dataset_registry(dataset, relpath, nrows) VALUES(?,?,?)",
            (dataset, str(csv_path.relative_to(tidy_dir)), int(len(df))),
        )

    con.commit()
    con.close()
    print(f"Wrote DB: {db_path} (tables: {len(csvs)} + dataset_registry)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
