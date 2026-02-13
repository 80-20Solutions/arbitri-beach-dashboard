import csv
import io
import urllib.request

SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'
ORG_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=1005819633'
ANA_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=265498497'

def fetch_csv(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as r:
        return r.read().decode('utf-8')

# Download
org_text = fetch_csv(ORG_URL)
ana_text = fetch_csv(ANA_URL)

# Parse organico - find header row with "Tipo Tecnico"
org_rows = list(csv.reader(io.StringIO(org_text)))
h_idx = next((i for i, r in enumerate(org_rows) if any('Tipo Tecnico' in c for c in r)), 3)
org_hdr = org_rows[h_idx]
org_data = org_rows[h_idx+1:]

# Get all organico names (arbitri only, not supervisori)
org_names = set()
nom_idx = next(i for i, h in enumerate(org_hdr) if 'Nominativo' in h)
tipo_idx = next(i for i, h in enumerate(org_hdr) if 'Tipo Tecnico' in h)
for r in org_data:
    if len(r) > max(nom_idx, tipo_idx) and r[nom_idx].strip():
        org_names.add(r[nom_idx].strip().lower())

# Parse analisi
ana_rows = list(csv.reader(io.StringIO(ana_text)))
ana_hdr = ana_rows[0]
years = ['2018','2019','2020','2021','2022','2023','2024','2025']

non_organico = []
for r in ana_rows[1:]:
    if not r or not r[0].strip():
        continue
    obj = {h: r[i] if i < len(r) else '' for i, h in enumerate(ana_hdr)}
    nome = (obj.get('Arbitro','') or '').strip()
    if not nome:
        continue
    if nome.lower() not in org_names:
        yr_data = {y: int(obj.get(y,'0') or '0') for y in years}
        totale = int(obj.get('Totale','0') or '0')
        non_organico.append({'nome': nome, 'years': yr_data, 'totale': totale})

# Output
print(f"\n{'='*60}")
print(f"ARBITRI NON IN ORGANICO: {len(non_organico)}")
print(f"{'='*60}\n")

lines = []
lines.append(f"ARBITRI PRESENTI IN ANALISI MA NON IN ORGANICO ({len(non_organico)})")
lines.append("=" * 60)
for a in sorted(non_organico, key=lambda x: -x['totale']):
    yr_str = ' | '.join(f"{y}:{a['years'][y]}" for y in years if a['years'][y] > 0)
    line = f"{a['nome']:40s} Tot: {a['totale']:3d}  [{yr_str}]"
    print(line)
    lines.append(line)

with open('arbitri_non_in_organico.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print(f"\nSalvato in arbitri_non_in_organico.txt")
