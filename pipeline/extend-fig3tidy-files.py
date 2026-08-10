import os
from pathlib import Path

import pandas as pd


def _normalise_thz(values: pd.Series) -> pd.Series:
    """Return frequencies in THz.

    Extended Fig 3b (and the identical Fig 2g) store the modelled frequency grid with an
    extra 1e-12 factor.  Rescaled by 1e12 it spans 217.391-245.902 THz -- exactly the
    1380 nm - 1220 nm window at c = 3e8 m/s (endpoints 5000/23 and 15000/61), the near-IR
    band of the experiment, containing the measured Fig 2f grid (220.093-242.643 THz).
    The adiabatic model of Extended Fig 3c shares that starting frequency but is a
    different grid (200 points, step 0.1433 THz, vs 240 points, step 0.1193 THz).
    Rescale only when the values are far below any physical optical frequency in THz.
    """
    finite = values.dropna()
    if not finite.empty and finite.abs().median() < 1e-6:
        return values * 1e12
    return values


def _tidy_grid(raw: pd.DataFrame, series_label: str) -> pd.DataFrame:
    """Slit-separation x frequency grid -> long form.

        row0: [_,                      'Frequencies (THz)', f0, f1, ...]
        row1: ['Slit separation (fs)', 'Counts (MHz):',     _,  _,  ...]
        row2+:[S_k,                     _,                  v,  v,  ...]
    """
    freqs = _normalise_thz(pd.to_numeric(raw.iloc[0, 2:], errors="coerce"))
    slit = pd.to_numeric(raw.iloc[2:, 0], errors="coerce")
    values = raw.iloc[2:, 2:].apply(pd.to_numeric, errors="coerce")

    records = []
    for i, s in enumerate(slit):
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
                    "Series": series_label,
                })
    return pd.DataFrame(records)


def _tidy_stacked_pairs(raw: pd.DataFrame) -> pd.DataFrame:
    """Stacked (frequency, counts) row pairs, one pair per rise time.

        row0: ['Frequency - 3.6 fs rise time (THz)', f0, f1, ...]
        row1: ['Counts - 3.6 fs rise time (MHz)',    c0, c1, ...]
        row2: ['Frequency - 7 fs rise time (THz)',   ...]      etc.

    The previous version looked up hard-coded column names such as
    ``"Frequency - 3.6 fs rise time (THz)"``, but ``ex2csv`` rewrites headers with
    underscores, so no lookup matched and the frame fell through to a pass-through.
    Pairing rows positionally is robust to that renaming.
    """
    records = []
    for r in range(0, raw.shape[0] - 1, 2):
        label = str(raw.iloc[r, 0])
        if "frequen" not in label.lower():
            continue
        # 'Frequency - 3.6 fs rise time (THz)' -> '3.6 fs rise time'.  Extended Fig 3d
        # also carries an unqualified 'Frequency (THz)' pair (the measurement) and a
        # 'Frequency - smoothing (THz)' pair, so an empty remainder means the measured
        # trace rather than an unnamed series.
        series = label.replace("Frequency", "").replace("_", " ")
        series = series.replace("-", "").replace("(THz)", "").strip() or "measured"
        freqs = _normalise_thz(pd.to_numeric(raw.iloc[r, 1:], errors="coerce"))
        counts = pd.to_numeric(raw.iloc[r + 1, 1:], errors="coerce")
        for f, c in zip(freqs, counts):
            if pd.notna(f) and pd.notna(c):
                records.append({
                    "Frequency_THz": float(f),
                    "Intensity": float(c),
                    "Series": series,
                })
    return pd.DataFrame(records)


def tidy_extended_fig3(src_dir="output/Extended_Data_Fig_3_Source_data",
                       output_dir="tidy_output/Extended_Fig_3"):
    """
    Reshape Extended Fig 3 tabs into tidy long format.

    3a-3c are slit-separation x frequency grids; 3d stacks one (frequency, counts) row
    pair per rise time.  Both were previously emitted as wide pass-throughs, so all four
    sheets contributed zero rows to the normalised database.

    Emits ``Frequency_THz`` / ``Intensity`` (plus ``Slit_separation_fs`` for the grids)
    so the loader routes them to ``spectra``.
    """
    os.makedirs(output_dir, exist_ok=True)

    for file in sorted(os.listdir(src_dir)):
        if not file.endswith(".xlsx"):
            continue

        file_path = os.path.join(src_dir, file)
        print(f"Processing {file_path}...")

        raw = pd.read_excel(file_path, sheet_name=0, header=None)
        raw = raw.dropna(axis=1, how="all")
        stem = Path(file).stem

        first_col_label = str(raw.iloc[0, 0])
        is_stacked = "frequen" in first_col_label.lower()

        if is_stacked:
            tidy_df = _tidy_stacked_pairs(raw)
        else:
            label = str(raw.iloc[0, 1]) if raw.shape[1] > 1 else ""
            series = "adiabatic model" if "adiabatic" in label.lower() else (
                "model" if "model" in label.lower() else "measured")
            tidy_df = _tidy_grid(raw, series)

        if tidy_df.empty:
            print("  Skipped: no tidy data extracted")
            continue

        out_file = os.path.join(output_dir, f"{stem}_tidy.csv")
        tidy_df.to_csv(out_file, index=False)
        print(f"  Saved {out_file} ({len(tidy_df)} rows)")


if __name__ == "__main__":
    tidy_extended_fig3()
