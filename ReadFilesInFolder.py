from pathlib import Path
import pandas as pd

folder = Path("/Users/bb5117/Library/CloudStorage/OneDrive-ImperialCollegeLondon/LabWork/Microscope Images/260416")

files = [f.name for f in folder.iterdir() if f.is_file() and not f.name.startswith('.')]

files.sort()  # 👈 this makes it alphabetical

df = pd.DataFrame(files, columns=["file_name"])

# Export to CSV and Excel
df.to_csv("file_list.csv", index=False)
df.to_excel("file_list.xlsx", index=False)

print("Saved file_list.csv and file_list.xlsx")