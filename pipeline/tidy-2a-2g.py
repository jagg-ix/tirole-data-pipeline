#!/usr/bin/env python3
"""
Process Fig_2a–2g Excel files into tidy CSVs.
Handles different layouts: frequency series, slit separation vs oscillation,
and matrix-style slit separation × frequency.
"""

import pandas as pd
import os
from pathlib import Path

def tidy_fig2(file_path: str, output_dir: str):
    base = Path(file_path).stem.lower()
    print(f"🔍 Processing {file_path}")
    raw = pd.read_excel(file_path, sheet_name=0, header=None)
    raw = raw.dropna(axis=1, how="all")

    tidy_df = None

    # ─────────────────────────────────────────────
    # Case: Fig_2a–2d → frequency series in columns, rows = series
    # ─────────────────────────────────────────────
    if any(base.startswith(f"fig_2{x}") for x in ["a","b","c","d"]):
        freqs = pd.to_numeric(raw.iloc[0, 1:], errors="coerce")
        # map row indices to series labels
        series_map = {
            1: "raw",
            3: "smooth",
            5: "model"
        }
        records = []
        for row_idx, label in series_map.items():
            if row_idx < len(raw):
                counts = pd.to_numeric(raw.iloc[row_idx, 1:], errors="coerce")
                for f, c in zip(freqs, counts):
                    if pd.notna(f) and pd.notna(c):
                        records.append({"Frequency_THz": f,
                                        "Counts_MHz": c,
                                        "Series": label})
        tidy_df = pd.DataFrame(records)

    # ─────────────────────────────────────────────
    # Case: Fig_2e → slit separation vs oscillation period (raw, model, fit)
    # ─────────────────────────────────────────────
    elif base.startswith("fig_2e"):
        slit_sep = pd.to_numeric(raw.iloc[0, 1:], errors="coerce")
        series_map = {
            1: "raw",
            3: "model",
            5: "fit"
        }
        records = []
        for row_idx, label in series_map.items():
            if row_idx < len(raw):
                vals = pd.to_numeric(raw.iloc[row_idx, 1:], errors="coerce")
                for s, v in zip(slit_sep, vals):
                    if pd.notna(s) and pd.notna(v):
                        records.append({"Slit_separation_fs": s,
                                        "Oscillation_THz": v,
                                        "Series": label})
        tidy_df = pd.DataFrame(records)

    # ─────────────────────────────────────────────
    # Case: Fig_2f → matrix: slit separation × frequency spectrum
    # ─────────────────────────────────────────────
    elif base.startswith("fig_2f"):
        # Layout (header=None):
        #   row0: [label, label, f1, f2, ...]
        #   row1: labels
        #   row2+: [S_fs, NaN, I(f1), I(f2), ...]
        slit_sep = pd.to_numeric(raw.iloc[2:, 0], errors="coerce")
        freqs = pd.to_numeric(raw.iloc[0, 2:], errors="coerce")
        values = raw.iloc[2:, 2:].apply(pd.to_numeric, errors="coerce")
        records = []
        for i, s in enumerate(slit_sep):
            if pd.isna(s):
                continue
            for j, f in enumerate(freqs):
                if j >= values.shape[1] or pd.isna(f):
                    continue
                v = values.iat[i, j]
                if pd.notna(v):
                    records.append({
                        "Slit_separation_fs": float(s),
                        "Frequency_THz": float(f),
                        "Intensity": float(v),
                        "Series": "model",
                    })
        tidy_df = pd.DataFrame(records)

    # ─────────────────────────────────────────────
    # Case: Fig_2g → time-domain heatmap: slit separation × delay
    # ─────────────────────────────────────────────
    elif base.startswith("fig_2g"):
        # Layout (header=None):
        #   row0: [label, t1_s, t2_s, ...]
        #   row1: [label text, NaN, NaN, ...]
        #   row2+: [S_fs, I(t1), I(t2), ...]
        slit_sep = pd.to_numeric(raw.iloc[2:, 0], errors="coerce")
        delays_s = pd.to_numeric(raw.iloc[0, 1:], errors="coerce")
        values = raw.iloc[2:, 1:].apply(pd.to_numeric, errors="coerce")
        records = []
        for i, s in enumerate(slit_sep):
            if pd.isna(s):
                continue
            for j, t_s in enumerate(delays_s):
                if j >= values.shape[1] or pd.isna(t_s):
                    continue
                v = values.iat[i, j]
                if pd.notna(v):
                    records.append({
                        "Slit_separation_fs": float(s),
                        "Delay_fs": float(t_s) * 1e15,
                        "Reflectivity": float(v),
                        "Series": "model",
                    })
        tidy_df = pd.DataFrame(records)

    if tidy_df is not None and not tidy_df.empty:
        os.makedirs(output_dir, exist_ok=True)
        out_file = Path(output_dir) / f"{Path(file_path).stem}_tidy.csv"
        tidy_df.to_csv(out_file, index=False)
        print(f"✅ Saved: {out_file}")
    else:
        print(f"⚠️ Skipped: no tidy data extracted")

if __name__ == "__main__":
    files = [
        "../t/output/Fig_2_Source_data/Fig_2a.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2b.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2c.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2d.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2e.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2f.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2g.xlsx",
    ]
    output_dir = "tidy_output/Fig_2"
    for f in files:
        tidy_fig2(f, output_dir)
