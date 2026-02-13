import json
import time
import requests
from collections import defaultdict

API_URL = "https://script.google.com/macros/s/AKfycbwMJgjVtb4A9bzQf9-Y3ko-UZnl-DanzfyXFCoTZtjQIJOBU3e2CQ1AED6A5kK028E8/exec"
WORKDIR = r"C:\Users\KreshOS\Desktop\00-arbitr\00-arbitri"

def load_json_sheet(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
    rows = data.get("values", [])
    if not rows:
        return []
    headers = rows[0]
    result = []
    for row in rows[1:]:
        d = {}
        for i, h in enumerate(headers):
            d[h] = row[i] if i < len(row) else ""
        result.append(d)
    return result

def load_json_raw(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
    return data.get("values", [])

def send_updates(updates, tab_label=""):
    total = len(updates)
    print(f"Sending {total} updates {tab_label}...")
    for i in range(0, total, 50):
        batch = updates[i:i+50]
        payload = json.dumps({"updates": batch}).encode("utf-8")
        for attempt in range(3):
            try:
                resp = requests.post(API_URL, data=payload, headers={"Content-Type": "application/json"}, timeout=120)
                print(f"  Batch {i//50+1} ({len(batch)} cells): {resp.status_code} {resp.text[:300]}")
                break
            except Exception as e:
                print(f"  Attempt {attempt+1} failed: {e}")
                if attempt < 2:
                    time.sleep(5)
        time.sleep(1)

def extract_tournaments(source_rows, funzione_filter):
    tournaments = defaultdict(lambda: {"tipo": "", "data": "", "luogo": "", "arbitri": set()})
    for row in source_rows:
        stato = row.get("StatoDesignazione", "").strip()
        funzione = row.get("FunzioneDesignazione", "").strip()
        if stato != "gara accettata" or funzione != funzione_filter:
            continue
        gara = row.get("GaraDescrizione", "").strip()
        impianto = row.get("ImpiantoNome", "").strip()
        key = f"{gara}|||{impianto}"
        t = tournaments[key]
        t["tipo"] = gara
        data_ora = row.get("DataOra", "").strip()
        if data_ora and not t["data"]:
            date_part = data_ora.split(" ")[0] if " " in data_ora else data_ora
            t["data"] = date_part
        luogo = row.get("ImpiantoLocalita", "").strip() or row.get("ImpiantoComune", "").strip()
        if luogo:
            t["luogo"] = luogo
        cognome = row.get("ArbitroCognome", "").strip()
        if cognome:
            t["arbitri"].add(cognome.upper())
    return tournaments

def main():
    # Load source data
    all_source = []
    for f in ["luglio.json", "agosto.json", "settembre.json"]:
        rows = load_json_sheet(f"{WORKDIR}\\{f}")
        print(f"{f}: {len(rows)} rows")
        all_source.extend(rows)
    print(f"Total source rows: {len(all_source)}")
    
    # Debug: show unique StatoDesignazione and FunzioneDesignazione
    stati = set(r.get("StatoDesignazione","").strip() for r in all_source)
    funzioni = set(r.get("FunzioneDesignazione","").strip() for r in all_source)
    print(f"Stati: {stati}")
    print(f"Funzioni: {funzioni}")

    # Load existing Designazioni
    existing = load_json_raw(f"{WORKDIR}\\designazioni.json")
    print(f"Existing: {len(existing)} rows, {max(len(r) for r in existing) if existing else 0} cols")

    row1 = existing[0] if len(existing) > 0 else []
    row2 = existing[1] if len(existing) > 1 else []
    row3 = existing[2] if len(existing) > 2 else []

    # Find last non-empty column
    max_col = 0
    for r in existing[:3]:
        for c in range(len(r)):
            if r[c].strip():
                max_col = max(max_col, c + 1)
    print(f"Last used column: {max_col}")

    # Build existing tournament keys
    existing_keys = set()
    for c in range(1, max(len(row1), len(row2), len(row3))):
        tipo = row1[c].strip() if c < len(row1) else ""
        data = row2[c].strip() if c < len(row2) else ""
        luogo = row3[c].strip() if c < len(row3) else ""
        if tipo or data or luogo:
            existing_keys.add(f"{tipo}|||{data}|||{luogo}")
    
    print(f"Existing tournament keys ({len(existing_keys)}):")
    for k in sorted(existing_keys):
        print(f"  {k}")

    # Build existing arbitri
    existing_arbitri = []
    for r in range(3, len(existing)):
        name = existing[r][0].strip() if existing[r] and len(existing[r]) > 0 else ""
        existing_arbitri.append(name)
    while existing_arbitri and not existing_arbitri[-1]:
        existing_arbitri.pop()
    
    print(f"Existing arbitri ({len(existing_arbitri)}): {existing_arbitri[:10]}...")
    num_existing_rows = 3 + len(existing_arbitri)

    # ============ ARBITRO BEACH ============
    print("\n=== ARBITRO BEACH ===")
    tournaments = extract_tournaments(all_source, "Arbitro Beach")
    print(f"Found {len(tournaments)} tournaments")

    new_tournaments = []
    for key, t in sorted(tournaments.items(), key=lambda x: x[1]["data"]):
        check_key = f"{t['tipo']}|||{t['data']}|||{t['luogo']}"
        if check_key in existing_keys:
            print(f"  SKIP: {t['tipo']} {t['data']} {t['luogo']}")
        else:
            new_tournaments.append(t)
            print(f"  NEW: {t['tipo']} {t['data']} {t['luogo']} ({len(t['arbitri'])} arb: {t['arbitri']})")

    print(f"\nNew tournaments: {len(new_tournaments)}")

    existing_arbitri_upper = {a.upper() for a in existing_arbitri if a}
    all_new_arbitri = set()
    for t in new_tournaments:
        all_new_arbitri.update(t["arbitri"])
    new_arbitri = sorted(a for a in all_new_arbitri if a not in existing_arbitri_upper)
    if new_arbitri:
        print(f"New arbitri: {new_arbitri}")

    full_arbitri = list(existing_arbitri) + new_arbitri

    updates = []
    for i, name in enumerate(new_arbitri):
        row_idx = num_existing_rows + i + 1
        updates.append({"tab": "Designazioni", "row": row_idx, "col": 1, "value": name})

    for t_idx, t in enumerate(new_tournaments):
        col = max_col + t_idx + 1
        updates.append({"tab": "Designazioni", "row": 1, "col": col, "value": t["tipo"]})
        updates.append({"tab": "Designazioni", "row": 2, "col": col, "value": t["data"]})
        updates.append({"tab": "Designazioni", "row": 3, "col": col, "value": t["luogo"]})
        for a_idx, arbitro in enumerate(full_arbitri):
            if arbitro and arbitro.upper() in t["arbitri"]:
                updates.append({"tab": "Designazioni", "row": a_idx + 4, "col": col, "value": "x"})

    print(f"\nTotal updates: {len(updates)}")
    if updates:
        send_updates(updates, "(Arbitro Beach)")

    # ============ SEGNAPUNTI ============
    print("\n=== SEGNAPUNTI ===")
    seg_tournaments = extract_tournaments(all_source, "Segnapunti")
    print(f"Found {len(seg_tournaments)} Segnapunti tournaments")

    if not seg_tournaments:
        print("No Segnapunti data")
        print("\n=== DONE ===")
        return

    seg_list = sorted(seg_tournaments.items(), key=lambda x: x[1]["data"])
    seg_arbitri = sorted(set(a for t in seg_tournaments.values() for a in t["arbitri"]))
    print(f"Segnapunti arbitri ({len(seg_arbitri)}): {seg_arbitri}")

    seg_updates = []
    seg_updates.append({"tab": "Designazioni Segnapunti", "row": 1, "col": 1, "value": "Tipologia"})
    seg_updates.append({"tab": "Designazioni Segnapunti", "row": 2, "col": 1, "value": "Data"})
    seg_updates.append({"tab": "Designazioni Segnapunti", "row": 3, "col": 1, "value": "Luogo"})
    for i, name in enumerate(seg_arbitri):
        seg_updates.append({"tab": "Designazioni Segnapunti", "row": i + 4, "col": 1, "value": name})

    for t_idx, (key, t) in enumerate(seg_list):
        col = t_idx + 2
        seg_updates.append({"tab": "Designazioni Segnapunti", "row": 1, "col": col, "value": t["tipo"]})
        seg_updates.append({"tab": "Designazioni Segnapunti", "row": 2, "col": col, "value": t["data"]})
        seg_updates.append({"tab": "Designazioni Segnapunti", "row": 3, "col": col, "value": t["luogo"]})
        for a_idx, arbitro in enumerate(seg_arbitri):
            if arbitro in t["arbitri"]:
                seg_updates.append({"tab": "Designazioni Segnapunti", "row": a_idx + 4, "col": col, "value": "x"})

    print(f"Total Segnapunti updates: {len(seg_updates)}")
    print("NOTE: Tab 'Designazioni Segnapunti' must exist first!")
    send_updates(seg_updates, "(Segnapunti)")

    print("\n=== DONE ===")

if __name__ == "__main__":
    main()
