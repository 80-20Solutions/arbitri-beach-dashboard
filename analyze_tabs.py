import urllib.request, re, csv, io, json

SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'

# Get full tab list from htmlview
url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/htmlview'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8', errors='replace')

# Extract all tabs
tabs = re.findall(r'name:\s*"([^"]+)".*?gid:\s*"(\d+)"', html)
print("=== ALL TABS ===")
for name, gid in tabs:
    print(f"  {name}: gid={gid}")

# Now download and analyze each tab
for name, gid in tabs:
    try:
        u = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}'
        r = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
        data = urllib.request.urlopen(r).read().decode('utf-8', errors='replace')
        reader = csv.reader(io.StringIO(data))
        rows = list(reader)
        print(f"\n=== {name} (gid={gid}): {len(rows)} rows x {len(rows[0]) if rows else 0} cols ===")
        if rows:
            print(f"  Headers: {rows[0][:15]}")
            if len(rows[0]) > 15:
                print(f"  ... +{len(rows[0])-15} more columns")
            first_col = [r[0] for r in rows[1:6] if r]
            print(f"  First col (1-5): {first_col}")
            
            # Year analysis
            years = set()
            for h in rows[0]:
                for yr in re.findall(r'20\d{2}', str(h)):
                    years.add(yr)
            if years:
                print(f"  Years in headers: {sorted(years)}")
                print(f"  Has 2025: {'2025' in years}")
        
        # Save CSV
        fname = name.replace(' ', '_') + '.csv'
        with open(fname, 'w', encoding='utf-8', newline='') as f:
            f.write(data)
        print(f"  Saved: {fname}")
    except Exception as e:
        print(f"  ERROR: {e}")
