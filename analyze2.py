import openpyxl, json

wb = openpyxl.load_workbook('principale.xlsx', read_only=True)

# Check Designazioni tab - structure is transposed: tournaments as columns, arbitri as rows
ws = wb['Designazioni']
rows = list(ws.iter_rows(values_only=True))

print(f"Total rows: {len(rows)}")
print(f"Total columns in row 1: {len(rows[0]) if rows else 0}")

# Row 0 = TIPOLOGIA TORNEO
# Row 1 = DATA
# Row 2 = LUOGO
# Row 3+ = arbitri names

# Get all dates from row 1
dates_row = rows[1]
print(f"\nFirst 3 values in DATA row: {dates_row[:4]}")
print(f"Last 5 values in DATA row: {dates_row[-5:]}")

# Count 2025 dates
count_2025 = 0
dates_2025 = []
for i, d in enumerate(dates_row[1:], 1):  # skip first col which is "DATA"
    s = str(d) if d else ''
    if '2025' in s:
        count_2025 += 1
        dates_2025.append((i, s, rows[0][i] if i < len(rows[0]) else '', rows[2][i] if i < len(rows[2]) else ''))

print(f"\n2025 tournaments found: {count_2025}")
for idx, date, tipo, luogo in dates_2025:
    print(f"  Col {idx}: {date} | {tipo} | {luogo}")

# Check July-Sept 2025
jul_sep = [x for x in dates_2025 if any(m in str(x[1]) for m in ['07/', '08/', '09/', '/7/', '/8/', '/9/', 'Jul', 'Aug', 'Sep', '2025-07', '2025-08', '2025-09'])]
print(f"\nJuly-Sept 2025 tournaments in Designazioni: {len(jul_sep)}")
for x in jul_sep:
    print(f"  {x}")

# Show arbitri names (column A, rows 3+)
print(f"\nArbitri (first col, rows 3-10):")
for r in rows[3:10]:
    print(f"  {r[0]}")

print(f"\nTotal arbitri rows: {len(rows) - 3}")

wb.close()

# Now get GaraIds from the 3 shared sheets
print("\n" + "="*60)
print("SHARED SHEETS - GaraIds and dates summary")
import subprocess
gog = r'C:\Users\KreshOS\go\bin\gog.exe'

sheets = [
    ('foglio1-Luglio', '1wWqcc8Qa6DRGCG767ZbRViTjYm5cwkrIih6hAMdYTuA'),
    ('foglio2-Agosto', '1X3wbqD5cUeWyNCtJ-NBghM2A2XuSYgV1F6SHsf0mP9I'),
    ('foglio3-Settembre', '1xdXXqzK3pbaHAJOPxLT7ME9BWr-dRlYAGhnBYMYVJ00'),
]

for name, sid in sheets:
    r = subprocess.run([gog, 'sheets', 'get', sid, "'Foglio 1'!A1:M500", '--json'],
                       capture_output=True, text=True)
    if r.returncode == 0:
        data = json.loads(r.stdout)
        values = data.get('values', [])
        print(f"\n{name}: {len(values)} rows")
        if values:
            print(f"  Headers: {values[0]}")
            # Unique dates and descriptions
            dates = set()
            descs = set()
            arbitri = set()
            for row in values[1:]:
                if len(row) > 1:
                    dates.add(row[1].split(' ')[0])
                if len(row) > 8:
                    descs.add(row[8].strip())
                if len(row) > 9:
                    arbitri.add(row[9].strip())
            print(f"  Dates: {sorted(dates)}")
            print(f"  Descriptions: {sorted(descs)}")
            print(f"  Arbitri: {sorted(arbitri)}")

import json
