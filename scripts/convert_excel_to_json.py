#!/usr/bin/env python3
import sys, json, pandas as pd
from pathlib import Path

def main():
    if len(sys.argv) < 3:
        print("Usage: convert_excel_to_json.py <excel_path> <out_dir>")
        sys.exit(1)

    excel_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    out_dir.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        raise SystemExit(f"Excel not found: {excel_path}")

    xls = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl')

    def norm_cols(df):
        df = df.copy()
        df.columns = [str(c).strip().replace('
',' ').replace('  ',' ') for c in df.columns]
        return df

    cleaned = {name: norm_cols(df).dropna(how='all') for name, df in xls.items()}
    all_sheets_records = {name: json.loads(df.to_json(orient='records')) for name, df in cleaned.items()}

    # combined
    with (out_dir / 'aircraft_data_combined.json').open('w', encoding='utf-8') as f:
        json.dump(all_sheets_records, f, ensure_ascii=False, indent=2)

    # per-sheet
    for name, records in all_sheets_records.items():
        safe = name.lower().replace(' ', '_')
        with (out_dir / f"aircraft_data_{safe}.json").open('w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)

    print("Wrote JSON to", out_dir)

if __name__ == '__main__':
    main()
