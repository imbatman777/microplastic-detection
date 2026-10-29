"""
Generate sample spectral CSV files for testing the PlastiSense AI web app.
Extracts one real sample per polymer class from the training CSV.

Produces two formats per class:
  1. Full row (with label columns) — matches training CSV format (259 cols)
  2. Pure spectral (256 numeric channels only) — minimal format

Usage:
    python generate_samples.py
"""

import pandas as pd
import os

CSV_PATH = os.path.join(os.path.dirname(__file__), "..", "full_spectra_with_labels.csv")
OUT_DIR  = os.path.join(os.path.dirname(__file__), "samples")
os.makedirs(OUT_DIR, exist_ok=True)

df = pd.read_csv(CSV_PATH)

print(f"Full CSV shape: {df.shape}")
print(f"Columns (first 5): {list(df.columns[:5])}")
print(f"Unique labels (col 1): {df.iloc[:, 1].unique()}")

# Column layout: index | label (PE/PET/PP/PS) | ID | 256 spectral channels
LABEL_COL = df.columns[1]  # 'V0' contains the label string
SPECTRAL_START = 3         # index of first spectral column

for polymer in ['PE', 'PET', 'PP', 'PS']:
    subset = df[df[LABEL_COL] == polymer]
    if subset.empty:
        print(f"⚠ No rows for {polymer}")
        continue

    row = subset.iloc[0:1]  # take the first row as a DataFrame

    # --- Format 1: full row (same as training CSV) ---
    full_path = os.path.join(OUT_DIR, f"spectral_{polymer}_full.csv")
    row.to_csv(full_path, index=False)
    print(f"✓ {full_path}  ({row.shape[1]} columns)")

    # --- Format 2: pure 256-channel spectral ---
    pure = row.iloc[:, SPECTRAL_START:SPECTRAL_START + 256]
    pure_path = os.path.join(OUT_DIR, f"spectral_{polymer}.csv")
    pure.to_csv(pure_path, index=False)
    print(f"✓ {pure_path}  ({pure.shape[1]} columns)")

print("\n✅ All sample CSVs generated in:", OUT_DIR)
