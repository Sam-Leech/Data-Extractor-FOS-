#!/usr/bin/env python3
import os
import sys
import pandas as pd
import re

def parse_filename(filename):
    """
    Expecting strings like:
    "Ref: AFPC00762264  Client Name: Roman Patatanian.pdf"
    This returns (ref, client) or (None, None) if it can't parse.
    """
    base = filename
    if base.lower().endswith('.pdf'):
        base = base[:-4]
    # Normalize whitespace
    base = re.sub(r'\s+', ' ', base).strip()
    # Try to find "Ref:" and "Client Name:" parts
    ref = None
    client = None
    # Look for "Ref:" then something, then "Client Name:"
    m = re.search(r'Ref:\s*(.+?)\s+Client Name:\s*(.+)$', base, flags=re.IGNORECASE)
    if m:
        ref = m.group(1).strip()
        client = m.group(2).strip()
        return ref, client
    # fallback: try splitting on "Client Name:"
    if 'Client Name:' in base:
        parts = base.split('Client Name:')
        left = parts[0].replace('Ref:', '').strip()
        right = parts[1].strip()
        if left:
            ref = left
        if right:
            client = right
        return ref, client
    # if nothing matched, attempt a generic split by two spaces
    parts = base.split('  ')
    if len(parts) >= 2:
        left = parts[0].replace('Ref:', '').strip()
        right = parts[1].replace('Client Name:', '').strip()
        return left or None, right or None
    return None, None

def main(folder_path):
    if not os.path.isdir(folder_path):
        print(f"ERROR: Folder not found: {folder_path}")
        return 1

    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    rows = []
    for f in files:
        ref, client = parse_filename(f)
        if ref is None and client is None:
            print(f"WARNING: Could not parse file name, skipping or adding blanks: {f}")
        rows.append({
            'Reference Number': ref if ref else '',
            'Client Name': client if client else '',
            'Lender': 'Oplo Finance',
            'Original Filename': f
        })

    if not rows:
        print("No PDF files found in the folder.")
        return 1

    df = pd.DataFrame(rows, columns=['Reference Number','Client Name','Lender','Original Filename'])
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
