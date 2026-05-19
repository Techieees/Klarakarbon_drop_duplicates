# Python 3.10+
# Requirements: pandas, openpyxl
# Purpose: Remove ONLY fully identical rows (keep first)

import pandas as pd
from datetime import datetime
import os

INPUT_FILE = r"C:\Users\FlorianDemir\Desktop\Klarakarbon 15 December\Klarakarbon YTD october GTN (1).xlsx"

OUTPUT_FILE = os.path.splitext(INPUT_FILE)[0] + f"_DEDUPED_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Load data
df = pd.read_excel(INPUT_FILE)

before_rows = len(df)

# Drop duplicates using ALL columns (exact row match only)
df = df.drop_duplicates(keep="first")

after_rows = len(df)

# Save result
df.to_excel(OUTPUT_FILE, index=False)

print("Done.")
print(f"Rows before: {before_rows}")
print(f"Rows after : {after_rows}")
print(f"Removed    : {before_rows - after_rows}")
print(f"Output saved to: {OUTPUT_FILE}")
