# Python 3.10+
# Requirements: pandas, openpyxl
# Purpose: Drop ONLY fully identical rows

import pandas as pd
from datetime import datetime
import os

INPUT_FILE = r"C:\Users\FlorianDemir\Desktop\Klarakarbon 12 December 2025\Gapit\Gapit Emissions 12 December 2025 All Together.xlsx"

OUTPUT_FILE = os.path.splitext(INPUT_FILE)[0] + f"_DEDUPED_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

df = pd.read_excel(INPUT_FILE)

before_rows = len(df)

# 🔴 KEY LINE: no subset → all columns are used
df = df.drop_duplicates(keep="first")

after_rows = len(df)

df.to_excel(OUTPUT_FILE, index=False)

print("Done.")
print(f"Rows before: {before_rows}")
print(f"Rows after : {after_rows}")
print(f"Removed    : {before_rows - after_rows}")
print(f"Output saved to: {OUTPUT_FILE}")
