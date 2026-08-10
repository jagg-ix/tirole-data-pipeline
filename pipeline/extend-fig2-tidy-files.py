import os
from pathlib import Path

import pandas as pd


def tidy_extended_fig2(src_dir="output/Extended_Data_Fig_2_Source_Data",
                       output_dir="tidy_output/Extended_Fig_2"):
    """
    Reshape Extended Fig 2 tabs (2a, 2b) into tidy long format.

    Source layout, read with header=None so nothing is lost to the header row:

        row0: ['Time_(fs)',                      t0, t1, t2, ...]
        row1: ['Fresnel reflection coefficient', v0, v1, v2, ...]

    i.e. a TRANSPOSED two-row series - the x values are spread across columns, not
    down a column.  The previous version assigned ``tidy_df = df`` (a pass-through),
    so the CSV kept 501 columns and one row; the loader looks for ``Delay_fs`` /
    ``Reflectivity`` column pairs and therefore found nothing, leaving Extended Fig 2a
    and 2b absent from the database entirely.

    Emits ``Delay_fs, Reflectivity, Series`` so the loader routes it to ``time_domain``.
    """
    os.makedirs(output_dir, exist_ok=True)

    for file in sorted(os.listdir(src_dir)):
        if not file.endswith(".xlsx"):
            continue

        file_path = os.path.join(src_dir, file)
        print(f"Processing {file_path}...")

        raw = pd.read_excel(file_path, sheet_name=0, header=None)
        raw = raw.dropna(axis=1, how="all")
        if raw.shape[0] < 2:
            print("  Skipped: fewer than two rows")
            continue

        x_label = str(raw.iloc[0, 0])
        y_label = str(raw.iloc[1, 0])
        xs = pd.to_numeric(raw.iloc[0, 1:], errors="coerce")
        ys = pd.to_numeric(raw.iloc[1, 1:], errors="coerce")

        records = [
            {"Delay_fs": float(x), "Reflectivity": float(y), "Series": y_label}
            for x, y in zip(xs, ys)
            if pd.notna(x) and pd.notna(y)
        ]
        if not records:
            print(f"  Skipped: no numeric pairs under {x_label!r}")
            continue

        tidy_df = pd.DataFrame(records)
        out_file = os.path.join(output_dir, f"{Path(file).stem}_tidy.csv")
        tidy_df.to_csv(out_file, index=False)
        print(f"  Saved {out_file} ({len(tidy_df)} rows)")


if __name__ == "__main__":
    tidy_extended_fig2()
