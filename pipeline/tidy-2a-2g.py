#!/usr/bin/env python3
"""
Process Fig_2a–2g Excel files into tidy CSVs.
Handles different layouts: frequency series, slit separation vs oscillation,
and matrix-style slit separation × frequency.
"""

import pandas as pd
import os
from pathlib import Path


def series_from_header(raw: pd.DataFrame) -> str:
    """Read 'model' vs 'measured' off the sheet's own axis labels.

    The workbook qualifies modelled axes explicitly -- 'Frequencies - model
    (THz)' / 'Slit separation - model (fs)' (Fig 2g, Extended Fig 3b), and
    'Frequencies - adiabatic model (THz)' (Extended Fig 3c).  Measured sheets
    carry the bare form, 'Frequencies (THz)' / 'Slit separation (fs)', which is
    also how the measured row of Fig 2e is distinguished from its own '- model'
    counterpart directly below it.

    This label used to be hardcoded per sheet index, which marked Fig 2f as
    'model'.  It is not: its header is unqualified, 135 of its 640 counts at
    S = 800 fs are negative (a model does not emit negative photon counts), and
    it correlates at r = +1.000 with the acknowledged measured Fig 2a -- Fig 2a
    is a 221.0-229.5 THz crop of Fig 2f's full 220.1-242.6 THz window.  The same
    array reached the database correctly labelled 'measured' through
    Extended_Fig_3a, whose converter reads the header the way this now does.
    Mislabelling it hid 27 measured slit separations and 2.7x the bandwidth.
    """
    labels = " ".join(
        str(v).lower()
        for row in (0, 1)
        if row < len(raw)
        for v in raw.iloc[row, :3]
        if pd.notna(v)
    )
    return "model" if "model" in labels else "measured"


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
        label = series_from_header(raw)
        # Layout (header=None):
        #   row0: [label, label, f1, f2, ...]
        #   row1: labels
        #   row2+: [S_fs, NaN, I(f1), I(f2), ...]
        # This is MEASURED data at 27 slit separations, not a model -- see
        # series_from_header().  Fig 2a/2b/2c/2d are narrow crops of it.
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
                        "Series": label,
                    })
        tidy_df = pd.DataFrame(records)

    # ─────────────────────────────────────────────
    # Case: Fig_2g → matrix: slit separation × frequency (modelled spectrum)
    # ─────────────────────────────────────────────
    elif base.startswith("fig_2g"):
        label = series_from_header(raw)
        # Layout (header=None) — identical in shape to Fig_2f:
        #   row0: [label, 'Frequencies - model (THz)', f1, f2, ...]
        #   row1: ['Slit separation - model (fs)', 'Counts (MHz):', NaN, ...]
        #   row2+: [S_fs, NaN, I(f1), I(f2), ...]
        #
        # This sheet was previously read as a time-domain heatmap: its frequency axis
        # was labelled Delay_fs and multiplied by 1e15 as if it were seconds, which put
        # modelled SPECTRA into the time_domain table with corrupted values.
        # Evidence that this axis is a frequency grid stored with an extra 1e-12 factor:
        #   * the source header says so: 'Frequencies - model (THz)', with the row axis
        #     'Slit separation - model (fs)' -- the same shape as the measured Fig 2f;
        #   * rescaled by 1e12 the grid spans 217.391-245.902 THz, which is EXACTLY the
        #     1380 nm - 1220 nm window at c = 3e8 m/s (endpoints 5000/23 and 15000/61),
        #     the near-IR band of the experiment; the measured Fig 2f grid
        #     (220.093-242.643 THz) sits inside it;
        #   * read instead as seconds -> fs (the previous x1e15) the axis would span
        #     217-246 ps, ~300x beyond any delay in the paper (Extended Fig 2: +/-800 fs).
        # Fig 2g and Extended Fig 3b hold byte-identical arrays (main text + extended).
        slit_sep = pd.to_numeric(raw.iloc[2:, 0], errors="coerce")
        freqs = pd.to_numeric(raw.iloc[0, 2:], errors="coerce")
        finite = freqs.dropna()
        if not finite.empty and finite.abs().median() < 1e-6:
            freqs = freqs * 1e12
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
                        "Series": label,
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
