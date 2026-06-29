#!/usr/bin/env python3
"""
convert_xls.py — Hatchwing temperature logger data converter
Usage:  python convert_xls.py yourfile.xlsx
        python convert_xls.py yourfile.xls
Output: data/yourfile.json  (ready to drop into the GitHub repo)

Expects LogTag XLS/XLSX format:
  Columns: Index | Serial # | Device Name | Date | Time | Elapsed | °C | ... | Events
  All loggers interleaved in a single sheet.
"""

import sys
import json
import os
import pandas as pd

def convert(path):
    fname = os.path.basename(path)
    base  = os.path.splitext(fname)[0]
    out_dir = "data"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, base + ".json")

    ext = os.path.splitext(fname)[1].lower()
    engine = "xlrd" if ext == ".xls" else None

    print(f"Reading {fname} ...")
    df = pd.read_excel(path, engine=engine, header=0)

    # Detect column layout: with or without Device Name column
    if df.shape[1] >= 9:
        # New format: Index | Serial | DeviceName | Date | Time | Elapsed | TempC | col8 | Events
        df.columns = ['Index', 'Serial', 'DeviceName', 'Date', 'Time',
                      'Elapsed', 'TempC', 'col8', 'Events'] + list(df.columns[9:])
        use_name = True
    else:
        # Old format: Index | Serial | Date | Time | Elapsed | TempC | col7 | Events | col9
        df.columns = ['Index', 'Serial', 'Date', 'Time', 'Elapsed',
                      'TempC', 'col7', 'Events', 'col9'] + list(df.columns[9:])
        use_name = False

    # Drop phantom header row
    df = df[df['Serial'] != 'Serial #:'].copy()

    # Parse datetime
    df['datetime'] = pd.to_datetime(
        df['Date'] + ' ' + df['Time'],
        format='%m/%d/%Y %I:%M:%S %p',
        errors='coerce'
    )
    df['TempC'] = pd.to_numeric(df['TempC'], errors='coerce')
    df = df.dropna(subset=['datetime', 'TempC', 'Serial'])
    df = df.sort_values('datetime')

    # Build label: prefer Device Name, fall back to Serial
    if use_name and 'DeviceName' in df.columns:
        label_map = df.groupby('Serial')['DeviceName'].first().to_dict()
    else:
        label_map = {s: s for s in df['Serial'].unique()}

    output = {}
    for serial in sorted(df['Serial'].unique()):
        label = str(label_map.get(serial, serial))
        sub   = df[df['Serial'] == serial]
        output[label] = [
            {"t": r['datetime'].strftime('%Y-%m-%dT%H:%M:%S'),
             "v": round(float(r['TempC']), 1)}
            for _, r in sub.iterrows()
        ]
        print(f"  {label}: {len(output[label])} points "
              f"({sub['datetime'].min().date()} → {sub['datetime'].max().date()})")

    with open(out_path, "w") as f:
        json.dump(output, f)

    size_kb = os.path.getsize(out_path) / 1024
    print(f"\nSaved → {out_path}  ({size_kb:.1f} KB)")
    print("Done. Add this file to your GitHub repo under the data/ folder.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_xls.py yourfile.xlsx")
        sys.exit(1)
    convert(sys.argv[1])
