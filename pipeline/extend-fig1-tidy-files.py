def tidy_extended_fig1(
    src_dir="output/Extended_Data_Fig_1_Source_Data",
    output_dir="tidy_output/Extended_Fig_1",
    make_plots=False
):
    import os
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    
    os.makedirs(output_dir, exist_ok=True)
    
    for file in os.listdir(src_dir):
        if not file.lower().endswith((".xlsx", ".xls")):
            continue
            
        file_path = os.path.join(src_dir, file)
        base_name = os.path.splitext(file)[0]
        print(f"Processing: {file}")
        
        try:
            raw = pd.read_excel(file_path, sheet_name=0, header=None)
            
            # ─── Critical fix: prevent .str on integer columns ───
            raw.columns = raw.columns.astype(str)
            
            raw = raw.dropna(axis=1, how="all")   # safe after astype(str)
            
            tidy_df = None
            
            # ────────────────────────────────────────────────
            # Case 1a
            # ────────────────────────────────────────────────
            if "1a" in base_name.lower():
                delay_actual = pd.to_numeric(raw.iloc[0, 1:], errors="coerce").dropna().values
                refl_actual  = pd.to_numeric(raw.iloc[1, 1:], errors="coerce").dropna().values
                delay_model  = pd.to_numeric(raw.iloc[2, 1:], errors="coerce").dropna().values
                refl_model   = pd.to_numeric(raw.iloc[3, 1:], errors="coerce").dropna().values
                
                df_actual = pd.DataFrame({"series": "actual", "delay_fs": delay_actual, "reflectivity": refl_actual})
                df_model  = pd.DataFrame({"series": "modelled", "delay_fs": delay_model, "reflectivity": refl_model})
                tidy_df = pd.concat([df_actual, df_model], ignore_index=True)
            
            # ────────────────────────────────────────────────
            # Case 1c
            # ────────────────────────────────────────────────
            elif "1c" in base_name.lower():
                intensities = pd.to_numeric(raw.iloc[0, 1:], errors="coerce").dropna().values
                refl1 = pd.to_numeric(raw.iloc[1, 1:], errors="coerce").dropna().values
                refl2 = pd.to_numeric(raw.iloc[2, 1:], errors="coerce").dropna().values
                min_len = min(map(len, [intensities, refl1, refl2]))
                df1 = pd.DataFrame({"Intensity_GW_cm2": intensities[:min_len], "Pump": "pump1", "Reflectivity": refl1[:min_len]})
                df2 = pd.DataFrame({"Intensity_GW_cm2": intensities[:min_len], "Pump": "pump2", "Reflectivity": refl2[:min_len]})
                tidy_df = pd.concat([df1, df2], ignore_index=True)
            
            # ────────────────────────────────────────────────
            # Case 1b — simplified robust version
            # ────────────────────────────────────────────────
            elif "1b" in base_name.lower():
                df = raw.copy()
                df = df.replace(['Reflectivity:', 'Reflectivity', ':', ''], np.nan)
                df = df.dropna(how='all')
                
                # Find delay column (first column with mostly numbers after skipping labels)
                delay_col = None
                for col in df.columns:
                    s = pd.to_numeric(df[col], errors='coerce')
                    if s.notna().sum() > len(df) * 0.4 and col != df.columns[-1]:
                        delay_col = col
                        break
                
                if delay_col is None:
                    delay_col = df.columns[0]  # fallback
                
                # Assume slit separations are in the row with most numeric values (excluding delay col)
                numeric_per_row = df.drop(columns=[delay_col]).apply(
                    lambda x: pd.to_numeric(x, errors='coerce').notna().sum(), axis=1
                )
                header_row_idx = numeric_per_row.idxmax()
                
                # Promote that row to column names if reasonable
                if numeric_per_row[header_row_idx] >= 3:
                    slit_vals = pd.to_numeric(df.iloc[header_row_idx], errors='coerce')
                    new_cols = [delay_col if i == df.columns.get_loc(delay_col) else slit_vals.iloc[i]
                                for i in range(len(df.columns))]
                    df.columns = new_cols
                    df = df.drop(df.index[header_row_idx])
                
                # Now melt
                df[delay_col] = pd.to_numeric(df[delay_col], errors='coerce')
                df = df.dropna(subset=[delay_col])
                
                tidy_df = df.melt(
                    id_vars=[delay_col],
                    var_name='Slit_separation_fs',
                    value_name='Reflectivity'
                )
                
                tidy_df['Slit_separation_fs'] = pd.to_numeric(tidy_df['Slit_separation_fs'], errors='coerce')
                tidy_df['Reflectivity']       = pd.to_numeric(tidy_df['Reflectivity'], errors='coerce')
                tidy_df = tidy_df.rename(columns={delay_col: 'Delay_fs'})
                tidy_df = tidy_df.dropna(subset=['Delay_fs', 'Slit_separation_fs', 'Reflectivity'])
                tidy_df = tidy_df.sort_values(['Delay_fs', 'Slit_separation_fs'])
            
            if tidy_df is None or tidy_df.empty:
                print(f"  Skipped — no valid data after cleaning")
                continue
            
            out_file = os.path.join(output_dir, f"{base_name}_tidy.csv")
            tidy_df.to_csv(out_file, index=False)
            print(f"  Saved: {out_file}")
            
            # Plotting block (unchanged)
            if make_plots:
                plt.figure(figsize=(9, 6))
                if "1a" in base_name.lower():
                    for series, g in tidy_df.groupby("series"):
                        plt.plot(g["delay_fs"], g["reflectivity"], "o-", lw=1.4, ms=5, label=series)
                    plt.xlabel("Delay (fs)")
                elif "1b" in base_name.lower():
                    for sep, g in tidy_df.groupby("Slit_separation_fs"):
                        label = f"{sep:.1f} fs" if sep < 100 else f"{sep:.0f} fs"
                        plt.plot(g["Delay_fs"], g["Reflectivity"], "o-", lw=1.3, ms=4.5, label=label)
                    plt.xlabel("Delay (fs)")
                elif "1c" in base_name.lower():
                    for pump, g in tidy_df.groupby("Pump"):
                        plt.plot(g["Intensity_GW_cm2"], g["Reflectivity"], "o-", lw=1.4, ms=5, label=pump)
                    plt.xlabel("Intensity (GW/cm²)")
                plt.ylabel("Reflectivity")
                plt.title(base_name.replace("_", " "))
                plt.legend(frameon=True, fontsize=10)
                plt.grid(True, alpha=0.4, linestyle="--")
                plt.margins(x=0.02, y=0.05)
                plt.tight_layout()
                plt.show()
                
        except Exception as e:
            print(f"  Error processing {file}: {type(e).__name__}: {e}")
            continue


if __name__ == "__main__":
    tidy_extended_fig1()
