#!/usr/bin/env python3
"""
Build + Verify Pipeline with Environment Variables
- Uses a single OS env var (DATA_DIR) for Excel source file directory
- Detects OS and prints helper instructions if env var missing
- Provides stage-specific functions: split, tidy, load, verify, check, run
"""

import os, sys, platform, sqlite3, pandas as pd
from pathlib import Path
import importlib.util

DB_FILE = "double_slit.sqlite3"
OUTPUT_DIR = "output"
TIDY_DIR = "tidy_output"

# Environment variable: directory containing all Excel source files.
# If not provided, default to the in-repo data_pipeline/source_data directory.
DATA_ENV = "DATA_DIR"

def _default_data_dir() -> str:
    """Return the default Excel source directory bundled with the repo."""
    here = Path(__file__).resolve()
    # .../data_pipeline/user_scripts/build_and_verify_pipeline.py
    # -> .../data_pipeline/source_data
    return str((here.parent.parent / "source_data").resolve())

EXPECTED_FILES = [
    "Extended_Data_Fig_1_Source data.xlsx",
    "Extended_Data_Fig_2_Source_Data.xlsx",
    "Extended_Data_Fig_3_Source_data.xlsx",
    "Fig_1_Source_data.xlsx",
    "Fig_2_Source_data.xlsx",
]

# Script filenames (ASCII-safe)
SCRIPTS = {
    "ex2csv": "ex2csv.py",
    "extend_fig1": "extend-fig1-tidy-files.py",
    "extend_fig2": "extend-fig2-tidy-files.py",
    "extend_fig3": "extend-fig3tidy-files.py",
    "fig1_2_tidy": "extract_tidy_fig1d_and_Fig2a_2g.py",
    "tidy_2a_2g": "tidy-2a-2g.py",
}

# -----------------------
# Env detection
# -----------------------
def check_env():
    """Resolve DATA_DIR.

    If DATA_DIR is not set, fall back to the repo-bundled source_data/ directory.
    """
    data_dir = os.environ.get(DATA_ENV)
    if not data_dir:
        data_dir = _default_data_dir()
    missing = []
    if not data_dir or not Path(data_dir).is_dir():
        missing.append(DATA_ENV)
    return data_dir, missing

def show_helper(missing):
    os_name = platform.system().lower()
    print("\nMissing environment variable(s):")
    for var in missing:
        print(f"  - {var} (expected to point to directory containing Excel source files)")
    default_dir = _default_data_dir()
    print(f"\nDefault used by this repo (if present):\n  {default_dir}\n")
    if "windows" in os_name:
        print("\nWindows (PowerShell):")
        print(f'setx {DATA_ENV} "C:\\path\\to\\data_dir"')
    elif "darwin" in os_name:  # macOS
        print("\nTry with:")
        print(f'export {DATA_ENV}="/Users/you/path/data_dir"')
    elif "linux" in os_name:
        print("\nTry with:")
        print(f'export {DATA_ENV}="/home/you/path/data_dir"')
    else:
        print("\nUnknown OS. Please set environment variable manually.")

# -----------------------
# Helpers
# -----------------------
def load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load module from {file_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

def ensure_dirs():
    Path(OUTPUT_DIR).mkdir(exist_ok=True, parents=True)
    Path(TIDY_DIR).mkdir(exist_ok=True, parents=True)

# -----------------------
# Stage 1: Split workbooks
# -----------------------
def split_workbooks():
    print("=== Step 1: Split workbooks ===")
    data_dir, missing = check_env()
    if missing:
        show_helper(missing)
        sys.exit(1)

    print("using :",data_dir)
    print(OUTPUT_DIR)
    ex2csv_mod = load_module_from_path("ex2csv", SCRIPTS["ex2csv"])
    workbooks = [os.path.join(data_dir, fname) for fname in EXPECTED_FILES]
    ex2csv_mod.split_workbooks(src_dir=data_dir, workbooks=workbooks, output_dir=OUTPUT_DIR)

# -----------------------
# Stage 2: Tidy converters
# -----------------------
def run_tidy_converters():
    print("=== Step 2: Tidy converters ===")
    load_module_from_path("extend_fig1", SCRIPTS["extend_fig1"]).tidy_extended_fig1(
        src_dir=os.path.join(OUTPUT_DIR, "Extended_Data_Fig_1_Source_data"),
        output_dir=os.path.join(TIDY_DIR, "Extended_Fig_1"),
        make_plots=False
    )
    load_module_from_path("extend_fig2", SCRIPTS["extend_fig2"]).tidy_extended_fig2(
        src_dir=os.path.join(OUTPUT_DIR, "Extended_Data_Fig_2_Source_Data"),
        output_dir=os.path.join(TIDY_DIR, "Extended_Fig_2")
    )
    load_module_from_path("extend_fig3", SCRIPTS["extend_fig3"]).tidy_extended_fig3(
        src_dir=os.path.join(OUTPUT_DIR, "Extended_Data_Fig_3_Source_data"),
        output_dir=os.path.join(TIDY_DIR, "Extended_Fig_3")
    )

    # Fig 1d tidy
    fig1d_file = os.path.join(OUTPUT_DIR, "Fig_1_Source_data", "Fig_1d.xlsx")
    out_dir_fig1 = os.path.join(TIDY_DIR, "Fig_1_and_2")
    Path(out_dir_fig1).mkdir(parents=True, exist_ok=True)
    if os.path.exists(fig1d_file):
        mod_fig1 = load_module_from_path("fig1_2_tidy", SCRIPTS["fig1_2_tidy"])
        mod_fig1.tidy_fig_workbook(fig1d_file, out_dir_fig1, make_plots=False)

    # Fig 2a–g tidy
    fig2_dir = os.path.join(OUTPUT_DIR, "Fig_2_Source_data")
    out_dir_fig2 = os.path.join(TIDY_DIR, "Fig_2")
    Path(out_dir_fig2).mkdir(parents=True, exist_ok=True)
    fig2_files = [os.path.join(fig2_dir, f"Fig_2{x}.xlsx") for x in "abcdefg"]
    mod_fig2 = load_module_from_path("tidy_2a_2g", SCRIPTS["tidy_2a_2g"])
    for f in fig2_files:
        if os.path.exists(f):
            mod_fig2.tidy_fig2(f, out_dir_fig2)

# -----------------------
# Stage 3: SQLite loader
# -----------------------
def init_db(conn):
    cur = conn.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS experiments (
        id INTEGER PRIMARY KEY,
        file_name TEXT,
        figure_ref TEXT,
        slit_separation_fs REAL,
        pump_condition TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS time_domain (
        id INTEGER PRIMARY KEY,
        experiment_id INTEGER,
        slit_separation_fs REAL,
        delay_fs REAL,
        reflectivity REAL,
        series TEXT,
        FOREIGN KEY(experiment_id) REFERENCES experiments(id)
    );
    CREATE TABLE IF NOT EXISTS spectra (
        id INTEGER PRIMARY KEY,
        experiment_id INTEGER,
        slit_separation_fs REAL,
        frequency_thz REAL,
        intensity REAL,
        series TEXT,
        FOREIGN KEY(experiment_id) REFERENCES experiments(id)
    );
    """)
    conn.commit()


def _infer_fig2_ab_slit_sep_fs(figure_ref: str):
    """Hard mapping from the paper text: Fig.2A uses S=800 fs; Fig.2B uses S=500 fs.

    These panels correspond to the spectra in tidy files Fig_2a_tidy.csv and Fig_2b_tidy.csv.
    """
    fr = (figure_ref or "").lower()
    if fr == "fig_2a":
        return 800.0
    if fr == "fig_2b":
        return 500.0
    return None

def insert_experiment(conn, file_name, figure_ref=None, slit_sep=None, pump=None):
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO experiments (file_name, figure_ref, slit_separation_fs, pump_condition)
        VALUES (?, ?, ?, ?)
    """, (file_name, figure_ref, slit_sep, pump))
    conn.commit()
    return cur.lastrowid

def load_csv_to_db(conn, csv_file):
    df = pd.read_csv(csv_file)
    fname = Path(csv_file).name
    figure_ref = fname.replace("_tidy.csv", "")

    # Normalize column names (many tidy files use different conventions)
    cols = {c: c.strip() for c in df.columns}
    df = df.rename(columns=cols)

    pump = None
    if "Pump" in df.columns:
        pump = str(df["Pump"].iloc[0])

    # Slit separation:
    # - If the tidy file has a per-row slit separation column (e.g., Fig_2f/Fig_2g), we store
    #   the experiment-level slit separation as NULL and use the per-row value when inserting.
    # - Otherwise we keep a single experiment-level slit separation (e.g., Fig_2a/Fig_2b).
    slit_sep = None
    per_row_slit_col = None
    for key in ["Slit_separation_fs", "slit_separation_fs"]:
        if key in df.columns:
            per_row_slit_col = key
            break
    # Some tidy files use "Slit_separation" without the _fs suffix. In the source paper this
    # axis is in fs, so we treat it as fs as well.
    if per_row_slit_col is None:
        for key in ["Slit_separation", "slit_separation"]:
            if key in df.columns:
                per_row_slit_col = key
                break
    if per_row_slit_col is None:
        # single-valued slit separation (or absent)
        try:
            for key in ["Slit_separation_fs", "Slit_separation", "slit_separation_fs", "slit_separation"]:
                if key in df.columns:
                    slit_sep = float(pd.to_numeric(df[key], errors="coerce").dropna().median())
                    break
        except Exception:
            slit_sep = None
    if slit_sep is None:
        slit_sep = _infer_fig2_ab_slit_sep_fs(figure_ref)

    exp_id = insert_experiment(conn, fname, figure_ref, slit_sep, pump)

    # Time-domain: accept either Delay_fs/Reflectivity or delay_fs/reflectivity
    if {"Delay_fs", "Reflectivity"}.issubset(df.columns) or {"delay_fs", "reflectivity"}.issubset(df.columns):
        if "Delay_fs" not in df.columns and "delay_fs" in df.columns:
            df = df.rename(columns={"delay_fs": "Delay_fs"})
        if "Reflectivity" not in df.columns and "reflectivity" in df.columns:
            df = df.rename(columns={"reflectivity": "Reflectivity"})
        series = df["series"].iloc[0] if "series" in df.columns else None
        cur = conn.cursor()
        # Allow per-row slit separation when present
        if per_row_slit_col is not None and per_row_slit_col in df.columns:
            df["slit_val"] = pd.to_numeric(df[per_row_slit_col], errors="coerce")
            rows = [(exp_id, float(r.slit_val), float(r.Delay_fs), float(r.Reflectivity), series) for r in df.itertuples(index=False) if r.slit_val==r.slit_val]
        else:
            rows = [(exp_id, slit_sep, float(r.Delay_fs), float(r.Reflectivity), series) for r in df.itertuples()]
        cur.executemany("""
            INSERT INTO time_domain (experiment_id, slit_separation_fs, delay_fs, reflectivity, series)
            VALUES (?, ?, ?, ?, ?)
        """, rows)
        conn.commit()

    # Spectra: accept Frequency_THz/Intensity, or Frequency_THz/Counts(_MHz), or Frequencies_(THz) wide-form is handled upstream.
    if "Frequency_THz" in df.columns and ("Intensity" in df.columns or "Counts" in df.columns or "Counts_MHz" in df.columns):
        if "Intensity" not in df.columns and "Counts" in df.columns:
            df = df.rename(columns={"Counts": "Intensity"})
        if "Intensity" not in df.columns and "Counts_MHz" in df.columns:
            df = df.rename(columns={"Counts_MHz": "Intensity"})
        series = df["Series"].iloc[0] if "Series" in df.columns else (df["series"].iloc[0] if "series" in df.columns else None)
        cur = conn.cursor()
        # Allow per-row slit separation when present (Fig_2f/Fig_2g)
        if per_row_slit_col is not None and per_row_slit_col in df.columns:
            df["slit_val"] = pd.to_numeric(df[per_row_slit_col], errors="coerce")
            rows = [(exp_id, float(r.slit_val), float(r.Frequency_THz), float(r.Intensity), series) for r in df.itertuples(index=False) if r.slit_val==r.slit_val]
        else:
            rows = [(exp_id, slit_sep, float(r.Frequency_THz), float(r.Intensity), series) for r in df.itertuples()]
        cur.executemany("""
            INSERT INTO spectra (experiment_id, slit_separation_fs, frequency_thz, intensity, series)
            VALUES (?, ?, ?, ?, ?)
        """, rows)
        conn.commit()

def load_all_csvs():
    print("=== Step 3: Load tidy CSVs into SQLite ===")
    conn = sqlite3.connect(DB_FILE)
    init_db(conn)
    for csv_file in Path(TIDY_DIR).rglob("*.csv"):
        print("Loading:", csv_file)
        try:
            load_csv_to_db(conn, csv_file)
        except Exception as e:
            print("  Failed:", csv_file, e)
    conn.close()

# -----------------------
# Stage 4: Verification
# -----------------------
def verify_db_vs_csv():
    print("=== Step 4: Verify DB vs CSV ===")
    conn = sqlite3.connect(DB_FILE)
    reports = []
    for csv_file in Path(TIDY_DIR).rglob("*.csv"):
        df = pd.read_csv(csv_file)
        fname = Path(csv_file).name
        cur = conn.cursor()
        cur.execute("SELECT id FROM experiments WHERE file_name=?", (fname,))
        row = cur.fetchone()
        if not row:
            reports.append({"file": fname, "status": "Missing in DB", "notes": "No experiment record"})
            continue
        exp_id = row[0]
        status, notes = "PASS", []

        # Time-domain checks
        if {"Delay_fs", "Reflectivity"}.issubset(df.columns):
            cur.execute("SELECT COUNT(*) FROM time_domain WHERE experiment_id=?", (exp_id,))
            db_count = cur.fetchone()[0]
            if db_count != len(df):
                status = "FAIL"
                notes.append(f"time_domain row mismatch CSV={len(df)} DB={db_count}")
            else:
                # Optional sample value check (first row)
                cur.execute("""
                    SELECT delay_fs, reflectivity FROM time_domain
                    WHERE experiment_id=? ORDER BY id ASC LIMIT 1
                """, (exp_id,))
                first_db = cur.fetchone()
                first_csv = (float(df.iloc[0]["Delay_fs"]), float(df.iloc[0]["Reflectivity"]))
                if first_db is None or abs(first_db[0] - first_csv[0]) > 1e-9 or abs(first_db[1] - first_csv[1]) > 1e-9:
                    status = "FAIL"
                    notes.append("time_domain first row mismatch")

        # Spectra checks
        if {"Frequency_THz", "Intensity"}.issubset(df.columns):
            cur.execute("SELECT COUNT(*) FROM spectra WHERE experiment_id=?", (exp_id,))
            db_count = cur.fetchone()[0]
            if db_count != len(df):
                status = "FAIL"
                notes.append(f"spectra row mismatch CSV={len(df)} DB={db_count}")
            else:
                cur.execute("""
                    SELECT frequency_thz, intensity FROM spectra
                    WHERE experiment_id=? ORDER BY id ASC LIMIT 1
                """, (exp_id,))
                first_db = cur.fetchone()
                first_csv = (float(df.iloc[0]["Frequency_THz"]), float(df.iloc[0]["Intensity"]))
                if first_db is None or abs(first_db[0] - first_csv[0]) > 1e-9 or abs(first_db[1] - first_csv[1]) > 1e-9:
                    status = "FAIL"
                    notes.append("spectra first row mismatch")

        reports.append({"file": fname, "status": status, "notes": "; ".join(notes)})

    conn.close()
    df_report = pd.DataFrame(reports)
    df_report.to_csv("verification_report.csv", index=False)
    total = len(df_report)
    passes = (df_report["status"] == "PASS").sum()
    fails = (df_report["status"] == "FAIL").sum()
    missing = (df_report["status"] == "Missing in DB").sum()
    print("Saved verification_report.csv")
    print(f"Summary: {total} files checked — {passes} PASS, {fails} FAIL, {missing} Missing")


# -----------------------
# Summary helper
# -----------------------
def summary_db(out_csv: str = "db_summary.csv"):
    """Write a small summary of the SQLite DB contents.

    This is intended as a quick sanity-check and for downstream "completeness" gates.
    """
    if not Path(DB_FILE).exists():
        print(f"DB not found: {DB_FILE}")
        sys.exit(1)

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()

    def q1(sql: str):
        cur.execute(sql)
        r = cur.fetchone()
        return r[0] if r else None

    total_exps = q1("SELECT COUNT(*) FROM experiments")
    total_spec = q1("SELECT COUNT(*) FROM spectra")
    total_td = q1("SELECT COUNT(*) FROM time_domain")
    distinct_s_spec = q1("SELECT COUNT(DISTINCT slit_separation_fs) FROM spectra WHERE slit_separation_fs IS NOT NULL")
    distinct_s_td = q1("SELECT COUNT(DISTINCT slit_separation_fs) FROM time_domain WHERE slit_separation_fs IS NOT NULL")

    # Per-figure coverage
    rows = []
    cur.execute("SELECT DISTINCT figure_ref FROM experiments ORDER BY figure_ref")
    figs = [r[0] for r in cur.fetchall()]
    for fig in figs:
        cur.execute("SELECT COUNT(*) FROM experiments WHERE figure_ref=?", (fig,))
        n_exp = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM experiments WHERE figure_ref=? AND slit_separation_fs IS NOT NULL", (fig,))
        n_exp_with_s = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM spectra s JOIN experiments e ON e.id=s.experiment_id WHERE e.figure_ref=?", (fig,))
        n_spec = cur.fetchone()[0]
        cur.execute("SELECT COUNT(DISTINCT s.slit_separation_fs) FROM spectra s JOIN experiments e ON e.id=s.experiment_id WHERE e.figure_ref=? AND s.slit_separation_fs IS NOT NULL", (fig,))
        n_spec_s = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM time_domain t JOIN experiments e ON e.id=t.experiment_id WHERE e.figure_ref=?", (fig,))
        n_td = cur.fetchone()[0]
        cur.execute("SELECT COUNT(DISTINCT t.slit_separation_fs) FROM time_domain t JOIN experiments e ON e.id=t.experiment_id WHERE e.figure_ref=? AND t.slit_separation_fs IS NOT NULL", (fig,))
        n_td_s = cur.fetchone()[0]

        rows.append({
            "figure_ref": fig,
            "experiments": n_exp,
            "experiments_with_slit_sep": n_exp_with_s,
            "spectra_rows": n_spec,
            "spectra_distinct_S": n_spec_s,
            "time_domain_rows": n_td,
            "time_domain_distinct_S": n_td_s,
        })

    conn.close()

    df = pd.DataFrame(rows)
    df.to_csv(out_csv, index=False)
    print(f"Wrote {out_csv}")
    print(f"Totals: experiments={total_exps}, spectra_rows={total_spec}, time_domain_rows={total_td}")
    print(f"Distinct S (spectra)={distinct_s_spec}, Distinct S (time_domain)={distinct_s_td}")

# -----------------------
# CLI entrypoint
# -----------------------
def main():
    if len(sys.argv) < 2:
        print("Usage: python build_and_verify_pipeline.py [split|tidy|load|verify|summary|check|run]")
        sys.exit(1)

    cmd = sys.argv[1].strip().lower()
    data_dir, missing = check_env()

    if cmd == "check":
        if missing:
            show_helper(missing)
            sys.exit(1)
        print("Environment variables OK")
        print("DATA_DIR=",data_dir)
        sys.exit(0)

    # For all other commands, require DATA_DIR
    if missing:
        show_helper(missing)
        sys.exit(1)

    ensure_dirs()

    if cmd == "split":
        split_workbooks()
    elif cmd == "tidy":
        run_tidy_converters()
    elif cmd == "load":
        load_all_csvs()
    elif cmd == "verify":
        verify_db_vs_csv()
    elif cmd == "summary":
        summary_db()
    elif cmd == "run":
        split_workbooks()
        run_tidy_converters()
        load_all_csvs()
        verify_db_vs_csv()
    else:
        print(f"Unknown command: {cmd}")
        sys.exit(1)

if __name__ == "__main__":
    main()
