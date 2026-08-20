import pandas as pd
from pathlib import PureWindowsPath

input_csv = r"E:/EMKP_work/EMKP_final_rename_rollback.csv"
output_csv = r"E:/EMKP_work/EMKP_final_rename_rollback_with_filenames.csv"

# Change this if you want to extract from "Original Path" instead.
path_column = "Renamed Path"

df = pd.read_csv(input_csv)

if path_column not in df.columns:
    raise ValueError(
        f"Column '{path_column}' was not found. Available columns: {list(df.columns)}"
    )

df["file_name"] = df[path_column].apply(
    lambda x: PureWindowsPath(str(x)).name if pd.notna(x) else ""
)

df["file_name_no_extension"] = df[path_column].apply(
    lambda x: PureWindowsPath(str(x)).stem if pd.notna(x) else ""
)

try:
    df.to_csv(output_csv, index=False)
    print(f"Saved updated CSV to: {output_csv}")
except PermissionError:
    fallback_output_csv = output_csv.replace(".csv", "_new.csv")
    df.to_csv(fallback_output_csv, index=False)
    print(f"Output file was locked. Saved updated CSV to: {fallback_output_csv}")
