#!/usr/bin/env python3
import os
import sys
import pandas as pd
import re

def parse_filename(filename):
    """
    Expected format:
    AFPC02852862-Tony Duffy-Advantage Finance.pdf
    or
    AFPC02852862 - Tony Duffy - Advantage Finance.pdf

    Returns: (ref, client, lender) or (None, None, None)
    """

    name = filename
    if name.lower().endswith('.pdf'):
        name = name[:-4]

    # Normalize spaces around hyphens
    name = re.sub(r'\s*-\s*', '-', name)

    # Split into parts by hyphens
    parts = [p.strip() for p in name.split('-') if p.strip()]

    if len(parts) < 3:
        # Not enough information to extract all fields
        return None, None, None

    ref = parts[0]
    lender = parts[-1]
    client = " ".join(parts[1:-1])  # Everything between ref and lender

    return ref, client, lender


def main(folder_path):
    if not os.path.isdir(folder_path):
        print(f"ERROR: Folder not found: {folder_path}")
        return 1

    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    rows = []

    for f in files:
        ref, client, lender = parse_filename(f)

        if ref is None and client is None and lender is None:
            print(f"WARNING: Could not parse file name: {f}")

        rows.append({
            'Reference Number': ref if ref else '',
            'Client Name': client if client else '',
            'Lender': lender if lender else '',
            'Original Filename': f
        })

    if not rows:
        print("No PDF files found in the folder.")
        return 1

    df = pd.DataFrame(rows, columns=['Reference Number', 'Client Name', 'Lender', 'Original Filename'])
    out_xlsx = os.path.join(folder_path, 'client_data.xlsx')
    df.to_excel(out_xlsx, index=False)

    print(f"✅ Excel created: {out_xlsx}")
    return 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_filenames.py \"Y:\\FRE\\Affordablitility\\17.10.2025\"")
        sys.exit(1)

    folder = sys.argv[1]
    sys.exit(main(folder))
