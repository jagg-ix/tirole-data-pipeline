#!/usr/bin/env python3
"""
Process Fig_1d and Fig_2a–2g Excel source files into tidy CSVs.
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

SCRIPTS = { 
    "extend_fig1" : "extend-fig1-tidy-files.py"
    , "extend_fig2" : "extend-fig2-tidy-files.py"
    , "extend_fig3" : "extend-fig3tidy-files.py"
    , "fig1_2_tidy" : "extract_tidy_fig1d_and_Fig2a_2g.py"
    , "tidy_2a_2g" : "tidy-2a-2g.py" 
}


def tidy_fig_workbook(file_path: str, output_dir: str, make_plots: bool = False):
    base_name = Path(file_path).stem
    print(f"Processing: {file_path}")

    raw = pd.read_excel(file_path, sheet_name=0, header=None)
    raw = raw.dropna(axis=1, how="all")

    tidy_df = None

    # ────────────────────────────────────────────────
    # Case: Fig 1d (intensity vs reflectivity, single pump)
    # ────────────────────────────────────────────────
    if "fig_1d" in base_name.lower():
        intensities = pd.to_numeric(raw.iloc[0, 1:], errors="coerce").dropna().values
        refl        = pd.to_numeric(raw.iloc[1, 1:], errors="coerce").dropna().values

        min_len = min(len(intensities), len(refl))
        tidy_df = pd.DataFrame({
            "Intensity_GW_cm2": intensities[:min_len],
            "Reflectivity": refl[:min_len]
        })

    # ────────────────────────────────────────────────
    # Case: Fig 2a–2g (delay × reflectivity matrices)
    # ────────────────────────────────────────────────
    elif "fig_2" in base_name.lower():
        # First row after col0 contains slit separation headers
        slit_seps = pd.to_numeric(raw.iloc[1, 1:], errors="coerce").dropna().values

        # Data block usually starts at row 3 (0-based)
        data = raw.iloc[3:, :]
        delays = pd.to_numeric(data.iloc[:, 0], errors="coerce")
        values = data.iloc[:, 1:].apply(pd.to_numeric, errors="coerce")

        values["Delay_fs"] = delays.values

        long = values.melt(
            id_vars="Delay_fs",
            var_name="col_index",
            value_name="Reflectivity"
        )

        col_map = dict(enumerate(slit_seps))
        long["Slit_separation_fs"] = long["col_index"].map(col_map)

        long = long.drop(columns="col_index")
        tidy_df = long.dropna(subset=["Delay_fs", "Slit_separation_fs", "Reflectivity"])

    if tidy_df is None or tidy_df.empty:
        raise ValueError(f"has no data or unrecognized format")

    # Save tidy CSV
    os.makedirs(output_dir, exist_ok=True)
    out_file = Path(output_dir) / f"{base_name}_tidy.csv"
    tidy_df.to_csv(out_file, index=False)
    print(f"  Saved: {out_file}")

    # Optional quick plot
    if make_plots:
        plt.figure(figsize=(8, 5))
        if "fig_1d" in base_name.lower():
            plt.plot(tidy_df["Intensity_GW_cm2"], tidy_df["Reflectivity"], "o-", lw=1.4)
            plt.xlabel("Intensity (GW/cm²)")
        elif "fig_2" in base_name.lower():
            for sep, g in tidy_df.groupby("Slit_separation_fs"):
                label = f"{sep:.1f} fs" if sep < 100 else f"{sep:.0f} fs"
                plt.plot(g["Delay_fs"], g["Reflectivity"], "o-", lw=1.2, ms=4, label=label)
            plt.xlabel("Delay (fs)")
            plt.legend(frameon=True, fontsize=9)
        plt.ylabel("Reflectivity")
        plt.title(base_name.replace("_", " "))
        plt.grid(True, alpha=0.4, linestyle="--")
        plt.tight_layout()
        plt.show()


if __name__ == "__main__":
    files = [
        "../t/output/Fig_2_Source_data/Fig_2a.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2b.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2c.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2d.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2e.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2f.xlsx",
        "../t/output/Fig_2_Source_data/Fig_2g.xlsx",
        "../t/output/Fig_1_Source_data/Fig_1d.xlsx",
    ]
    output_dir = "tidy_output/Fig_1_and_2"
    for f in files:
        tidy_fig_workbook(f, output_dir, make_plots=False)
