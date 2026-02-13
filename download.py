import urllib.request
import ssl
import os

ctx = ssl.create_default_context()

sheets = {
    'foglio1': '1wWqcc8Qa6DRGCG767ZbRViTjYm5cwkrIih6hAMdYTuA',
    'foglio2': '1X3wbqD5cUeWyNCtJ-NBghM2A2XuSYgV1F6SHsf0mP9I',
    'foglio3': '1xdXXqzK3pbaHAJOPxLT7ME9BWr-dRlYAGhnBYMYVJ00',
}

def download(url, fname):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, context=ctx)
    with open(fname, 'wb') as f:
        f.write(resp.read())
    sz = os.path.getsize(fname)
    print(f"OK {fname} ({sz} bytes)")

# Download 3 shared sheets
for name, sid in sheets.items():
    url = f'https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid=0'
    try:
        download(url, f'{name}.csv')
    except Exception as e:
        print(f"FAIL {name}: {e}")

# Download main sheet as XLSX (all tabs)
main_id = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'
try:
    url = f'https://docs.google.com/spreadsheets/d/{main_id}/export?format=xlsx'
    download(url, 'principale.xlsx')
except Exception as e:
    print(f"FAIL principale.xlsx: {e}")

# Also CSV gid=0
try:
    url = f'https://docs.google.com/spreadsheets/d/{main_id}/export?format=csv&gid=0'
    download(url, 'principale_gid0.csv')
except Exception as e:
    print(f"FAIL principale_gid0.csv: {e}")
