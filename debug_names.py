import csv, io, urllib.request

SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'
ORG_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=1005819633'
ANA_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=265498497'

def fetch_csv(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as r:
        return r.read().decode('utf-8')

org_text = fetch_csv(ORG_URL)
ana_text = fetch_csv(ANA_URL)

org_rows = list(csv.reader(io.StringIO(org_text)))
print("=== ORGANICO first 10 rows ===")
for i, r in enumerate(org_rows[:10]):
    print(f"Row {i}: {r[:5]}")

h_idx = next((i for i, r in enumerate(org_rows) if any('Tipo Tecnico' in c for c in r)), -1)
print(f"\nHeader at row {h_idx}: {org_rows[h_idx][:8] if h_idx >= 0 else 'NOT FOUND'}")

if h_idx >= 0:
    org_hdr = org_rows[h_idx]
    nom_idx = next((i for i, h in enumerate(org_hdr) if 'Nominativo' in h), -1)
    org_names = set()
    for r in org_rows[h_idx+1:]:
        if len(r) > nom_idx and r[nom_idx].strip():
            org_names.add(r[nom_idx].strip().lower())
    print(f"\nOrganico names count: {len(org_names)}")
    print("Sample:", list(org_names)[:5])

ana_rows = list(csv.reader(io.StringIO(ana_text)))
print(f"\n=== ANALISI header ===")
print(ana_rows[0])
print(f"Analisi rows: {len(ana_rows)-1}")

ana_names = set()
for r in ana_rows[1:]:
    if r and r[0].strip():
        ana_names.add(r[0].strip().lower())
print(f"Analisi names count: {len(ana_names)}")
print("Sample:", list(ana_names)[:5])

diff = ana_names - org_names
print(f"\nIn analisi but NOT in organico: {len(diff)}")
for n in sorted(diff):
    print(f"  - {n}")
