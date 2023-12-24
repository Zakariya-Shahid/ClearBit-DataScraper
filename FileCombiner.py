import pandas as pd
import glob
import os

# combining all xlsx files into one csv file from the current directory
all_files = glob.glob(os.path.join(os.getcwd(), "results", "*.xlsx"))
print(all_files)
df = pd.concat((pd.read_excel(f) for f in all_files), ignore_index=True)
df.to_csv("combined.csv", index=False)
