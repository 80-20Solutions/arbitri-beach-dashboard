import urllib.request, re, csv, io

SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'

# Try htmlview
url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/htmlview'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8', errors='replace')

# Find sheet tabs
matches = re.findall(r'gid=(\d+)[^>]*>([^<]+)<', html)
print('Tab matches:', matches)

matches2 = re.findall(r'id="sheet-button-(\d+)"[^>]*>([^<]*)', html)
print('Sheet buttons:', matches2)

# Search for "Designazioni" or "designaz"
idx = html.lower().find('designaz')
if idx >= 0:
    print(f'Found "designaz" at pos {idx}')
    print(html[max(0,idx-200):idx+200])

# Try brute force common gids
for gid in [0, 1, 2, 1005819633, 265498497, 123456789, 2000000000]:
    try:
        u = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}'
        r = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
        data = urllib.request.urlopen(r).read().decode('utf-8', errors='replace')
        reader = csv.reader(io.StringIO(data))
        rows = list(reader)
        if rows:
            print(f'GID {gid}: {len(rows)}x{len(rows[0])} - headers: {rows[0][:5]}')
    except:
        pass
