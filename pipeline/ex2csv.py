import pandas as pd
import os
from pathlib import Path

def split_workbooks(src_dir, workbooks=None, output_dir="Split_XLSX"):
    """
    Split each workbook into individual XLSX files, one per tab.
    Creates a dedicated output directory for each workbook if it doesn't exist.
    
    Parameters
    ----------
    src_dir : str
        Directory containing the source Excel files.
    workbooks : list of str, optional
        List of workbook filenames or absolute paths to process.
        If None, process all .xlsx in src_dir.
    output_dir : str, optional
        Root directory to save the split files.
    """
    os.makedirs(output_dir, exist_ok=True)

    # If no explicit list, grab all .xlsx files in src_dir
    if workbooks is None:
        workbooks = [os.path.join(src_dir, f) for f in os.listdir(src_dir) if f.endswith(".xlsx")]

    for wb in workbooks:
        # If only a filename was passed, join with src_dir
        wb_path = wb if os.path.isabs(wb) else os.path.join(src_dir, wb)
        print(f"Processing {wb_path}...")
        try:
            xls = pd.ExcelFile(wb_path)
        except Exception as e:
            print(f"  Failed to open {wb_path}: {e}")
            continue

        # Create subfolder per workbook under OUTPUT_DIR
        wb_base = Path(wb_path).stem.replace(" ", "_")
        wb_out_dir = os.path.join(output_dir, wb_base)
        if not os.path.exists(wb_out_dir):
            os.makedirs(wb_out_dir)
            print(f"  Created directory {wb_out_dir}")

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.columns = [str(col).strip().replace(" ", "_") for col in df.columns]

            out_file = os.path.join(wb_out_dir, f"{sheet_name.replace(' ', '_')}.xlsx")
            with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"  Saved {out_file}")

# Example usage (commented out):
workbooks = [
    "Extended_Data_Fig_1_Source data.xlsx",
    "Extended_Data_Fig_2_Source_Data.xlsx",
    "Extended_Data_Fig_3_Source_data.xlsx",
    "Fig_1_Source_data.xlsx",
    "Fig_2_Source_data.xlsx"
]

# split_workbooks(src_dir=data_dir, workbooks=workbooks, output_dir=OUTPUT_DIR)
