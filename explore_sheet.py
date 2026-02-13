import urllib.request, re, csv, io

SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'
BASE = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}'

# 1) Fetch HTML to find all gids
req = urllib.request.Request(f'{BASE}/edit', headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8', errors='replace')

gids = sorted(set(re.findall(r'gid[=:](\d+)', html)))
print(f"GIDs found in HTML: {gids}")

# 2) Try common gids + found ones
test_gids = list(set(['0'] + gids + ['1005819633', '265498497']))
for gid in test_gids:
    try:
        url = f'{BASE}/export?format=csv&gid={gid}'
        req2 = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        data = urllib.request.urlopen(req2).read().decode('utf-8', errors='replace')
        reader = csv.reader(io.StringIO(data))
        rows = list(reader)
        if rows:
            print(f"\n=== GID {gid}: {len(rows)} rows x {len(rows[0])} cols ===")
            # Print first row (headers)
            print(f"Headers (first 10): {rows[0][:10]}")
            # Print first col values (first 5)
            first_col = [r[0] for r in rows[1:6] if r]
            print(f"First col (rows 1-5): {first_col}")
            
            # Check if this looks like Designazioni
            header_str = ' '.join(rows[0]).lower()
            if 'torneo' in header_str or 'design' in header_str or len(rows[0]) > 20:
                print(">>> POSSIBLE DESIGNAZIONI TAB <<<")
                # Full analysis
                print(f"Total: {len(rows)} rows x {len(rows[0])} cols")
                print(f"ALL headers: {rows[0]}")
                # Check for years
                years = set()
                for h in rows[0]:
                    yr = re.findall(r'20\d{2}', str(h))
                    years.update(yr)
                print(f"Years in headers: {sorted(years)}")
                print(f"Has 2025: {'2025' in years}")
                # First 10 row labels
                labels = [r[0] for r in rows[1:11] if r]
                print(f"Row labels (first 10): {labels}")
                
                # Save CSV
                with open(f'designazioni_gid{gid}.csv', 'w', encoding='utf-8', newline='') as f:
                    f.write(data)
                print(f"Saved to designazioni_gid{gid}.csv")
    except Exception as e:
        print(f"GID {gid}: ERROR - {e}")
