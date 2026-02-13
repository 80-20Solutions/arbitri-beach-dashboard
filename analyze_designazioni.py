import csv, re

with open('Designazioni.csv', encoding='utf-8') as f:
    rows = list(csv.reader(f))

print(f"=== DESIGNAZIONI (gid=21948265) ===")
print(f"Dimensioni: {len(rows)} righe x {len(rows[0])} colonne")
print()

# Structure: Row 0 = tipologia torneo, Row 1 = data, Row 2 = luogo, Row 3+ = arbitri (cognomi)
print("STRUTTURA:")
print(f"  Riga 0 (col 0): '{rows[0][0]}' -> tipologie torneo nelle colonne")
print(f"  Riga 1 (col 0): '{rows[1][0]}' -> date dei tornei")  
print(f"  Riga 2 (col 0): '{rows[2][0]}' -> luoghi dei tornei")
print()

# Tipologie tornei (riga 0, da col 1)
tipologie = [rows[0][i] for i in range(1, len(rows[0])) if rows[0][i].strip()]
print(f"TIPOLOGIE TORNEO (unique): {sorted(set(tipologie))}")
print(f"  Conteggio per tipo:")
from collections import Counter
for tipo, count in Counter(tipologie).most_common():
    print(f"    {tipo}: {count}")
print()

# Date (riga 1, da col 1)  
date = [rows[1][i] for i in range(1, len(rows[1])) if rows[1][i].strip()]
print(f"DATE tornei ({len(date)} valori):")
print(f"  Prima: {date[0] if date else 'N/A'}")
print(f"  Ultima: {date[-1] if date else 'N/A'}")
# Check for 2025
dates_2025 = [d for d in date if '2025' in d]
print(f"  Date 2025: {len(dates_2025)}")
if dates_2025:
    print(f"    Esempi: {dates_2025[:5]}")
print()

# Luoghi (riga 2)
luoghi = [rows[2][i] for i in range(1, len(rows[2])) if rows[2][i].strip()]
print(f"LUOGHI ({len(luoghi)} valori, unique: {len(set(luoghi))})")
print()

# Arbitri (da riga 3)
arbitri = [rows[i][0] for i in range(3, len(rows)) if rows[i][0].strip()]
print(f"ARBITRI ({len(arbitri)}):")
for a in arbitri:
    print(f"  - {a}")
print()

# Check values in cells (what do designations look like?)
print("VALORI CELLE (campione):")
sample_vals = set()
for i in range(3, min(10, len(rows))):
    for j in range(1, min(20, len(rows[i]))):
        v = rows[i][j].strip()
        if v:
            sample_vals.add(v)
print(f"  Valori unici (primi 20 arbitri x 20 tornei): {sorted(sample_vals)[:20]}")
