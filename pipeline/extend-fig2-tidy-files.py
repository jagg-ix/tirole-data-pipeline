import pandas as pd
import os

def tidy_extended_fig2(src_dir="output/Extended_Data_Fig_2_Source_Data", 
                       output_dir="tidy_output/Extended_Fig_2"):
    """
    Reshape Extended Fig 2 tabs into tidy format.
    Handles Extended Fig 2a and 2b, which have Time_(fs) plus measurement columns.
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

        # Normalize column names
        df.columns = [str(col).strip().replace(" ", "_") for col in df.columns]

        # Build tidy dataframe
        tidy_df = df

        # Save tidy CSV
        base_name = os.path.splitext(file)[0]
        out_file = os.path.join(output_dir, f"{base_name}_tidy.csv")
        tidy_df.to_csv(out_file, index=False)
        print(f"  Saved {out_file}")

# Example usage:
tidy_extended_fig2()
