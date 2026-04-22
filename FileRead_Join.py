from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

# ---- Step 1: get file list ----
folder = Path("/Users/bb5117/Library/CloudStorage/OneDrive-ImperialCollegeLondon/LabWork/Microscope Images/260416")

files = [f.name for f in folder.iterdir() if f.is_file() and not f.name.startswith('.')]
files.sort(key=str.lower)

# ---- Step 2: select Excel file ----
root = tk.Tk()
root.withdraw()

excel_path = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel files", "*.xlsx")]
)

if not excel_path:
    print("No file selected.")
    exit()

print("Selected:", excel_path)

# ---- Step 3: open workbook ----
wb = load_workbook(excel_path)
ws = wb["BlueBox1"]

# ---- Step 4: write filenames starting at J4 ----
start_row = 4
col = "J"

for i, filename in enumerate(files):
    ws[f"{col}{start_row + i}"] = filename

# ---- Step 5: save new file ----
output_path = excel_path.replace(".xlsx", "_updated.xlsx")
wb.save(output_path)

print(f"Saved updated file: {output_path}")