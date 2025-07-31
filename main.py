import pandas as pd

def log(message):
    print(f"[LOG] {message}")

def read_excel_flexible(filename):
    combined_data = []

    try:
        # Force openpyxl engine for Excel
        excel_obj = pd.ExcelFile(filename, engine="openpyxl")
        log(f"Opened file: {filename}")

        for sheet in excel_obj.sheet_names:
            try:
                df = pd.read_excel(filename, sheet_name=sheet, dtype=str, engine="openpyxl")
                df = df.fillna("")  # Replace NaN with blanks
                df["SourceSheet"] = sheet
                df["SourceFile"] = filename
                combined_data.append(df)
                log(f"Processed sheet: {sheet} with {len(df)} rows")

            except Exception as e:
                log(f"Error reading sheet '{sheet}': {e}")

    except Exception as e:
        log(f"Error opening file '{filename}': {e}")

    if combined_data:
        return pd.concat(combined_data, ignore_index=True)
    else:
        return pd.DataFrame()

def main():
    log("=== Exception Report Generator Started ===")
    filenames = input("Enter comma-separated Excel file names (e.g., file1.xlsx,file2.xlsx): ")
    files = [f.strip() for f in filenames.split(",") if f.strip()]

    all_data = []
    for file in files:
        df = read_excel_flexible(file)
        if not df.empty:
            all_data.append(df)

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        output_filename = "combined_output.xlsx"
        final_df.to_excel(output_filename, index=False, engine="openpyxl")
        log(f"Output saved as {output_filename} — download from file list.")
    else:
        log("No valid data found in provided files.")

    log("=== Process Complete ===")

if __name__ == "__main__":
    main()
