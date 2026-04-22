
import os
import pandas as pd

folder = "../../LabWork/Microscope Images/260416/"

files = os.listdir(folder)

df = pd.DataFrame(files, columns=["file_name"])

df.to_csv("file_list.csv", index=False)

print("Saved file_list.csv")