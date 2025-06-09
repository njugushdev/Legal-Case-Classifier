import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def FilterJuryRows():
    Tk().withdraw()
    filePath = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

    if not filePath:
        print("❌ No file selected.")
        return

    # Read Excel file
    df = pd.read_excel(filePath)

    # Check for required columns
    if "Docket Text" not in df.columns or "Status" not in df.columns:
        print("❌ Columns 'Docket Text' and 'Status' are required.")
        return

    # Filter rows where Docket Text contains "jury" (case-insensitive)
    jury_rows = df[df["Docket Text"].astype(str).str.lower().str.contains("jury", na=False)]

    if jury_rows.empty:
        print("ℹ️ No rows found with 'jury' in Docket Text.")
        return

    # Display results in terminal
    print("\nRows with 'jury' in Docket Text:")
    print("-" * 50)
    for index, row in jury_rows.iterrows():
        print(f"Row {index}:")
        print(f"Docket Text: {row['Docket Text']}")
        print(f"Status: {row['Status']}")
        print("-" * 50)

    # Save filtered rows to Excel
    output_file = "jury_rows.xlsx"
    jury_rows[["Docket Text", "Status"]].to_excel(output_file, index=False)
    print(f"✅ Filtered rows saved to: {output_file}")

if __name__ == "__main__":
    FilterJuryRows()