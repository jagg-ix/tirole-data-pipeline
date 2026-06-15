import pandas as pd
import os

def tidy_extended_fig3(src_dir="output/Extended_Data_Fig_3_Source_data", 
                       output_dir="tidy_output/Extended_Fig_3"):
    """
    Reshape Extended Fig 3 tabs into tidy format.
    Handles multiple rise times (3.6, 7, 17, 32, 47 fs) for Fig 3d,
    and simpler single-frequency tabs for Fig 3a–3c.
    """
    os.makedirs(output_dir, exist_ok=True)

    for file in os.listdir(src_dir):
        if not file.endswith(".xlsx"):
            continue

        file_path = os.path.join(src_dir, file)
        print(f"Processing {file_path}...")

        # Read Excel file (each has only one sheet)
        xls = pd.ExcelFile(file_path)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Drop unnamed columns
        df = df.loc[:, ~df.columns.str.contains("Unnamed")]

        tidy_df = None

        # Case 1: Extended Fig 3d (multiple rise times)
        if any("fs rise time" in col for col in df.columns):
            rise_times = [3.6, 7, 17, 32, 47]
            tidy_frames = []
            for rt in rise_times:
                freq_col = f"Frequency - {rt} fs rise time (THz)"
                counts_col = f"Counts - {rt} fs rise time (MHz)"
                if freq_col in df.columns and counts_col in df.columns:
                    temp = pd.DataFrame({
                        "Rise_time_fs": rt,
                        "Frequency_THz": df[freq_col],
                        "Counts_MHz": df[counts_col]
                    })
                    tidy_frames.append(temp)
            tidy_df = pd.concat(tidy_frames, ignore_index=True)

        # Case 2: Extended Fig 3a–3c (already tidy or simple pairs)
        else:
            # Normalize column names
            df.columns = [str(col).strip().replace(" ", "_") for col in df.columns]
            tidy_df = df

        # Save tidy CSV
        base_name = os.path.splitext(file)[0]
        out_file = os.path.join(output_dir, f"{base_name}_tidy.csv")
        tidy_df.to_csv(out_file, index=False)
        print(f"  Saved {out_file}")

# Example usage:
tidy_extended_fig3()
