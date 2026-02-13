import csv, collections
from collections import defaultdict

all_rows = []
for f in ['foglio1.csv','foglio2.csv','foglio3.csv']:
    with open(f, encoding='utf-8') as fh:
        reader = csv.DictReader(fh)
        for r in reader:
            r['_file'] = f
            all_rows.append(r)

print(f'Totale righe: {len(all_rows)}')

# 1
stati = collections.Counter(r.get('StatoDesignazione','').strip() for r in all_rows)
print('\n=== STATO DESIGNAZIONE ===')
for k,v in stati.most_common():
    print(f'  {repr(k)}: {v}')

# 2
funz = collections.Counter(r.get('FunzioneDesignazione','').strip() for r in all_rows)
print('\n=== FUNZIONE DESIGNAZIONE ===')
for k,v in funz.most_common():
    print(f'  {repr(k)}: {v}')

# 3
gare = collections.Counter(r.get('GaraDescrizione','').strip() for r in all_rows)
print('\n=== GARA DESCRIZIONE ===')
for k,v in gare.most_common():
    print(f'  {repr(k)}: {v}')

# 4
tornei = defaultdict(set)
for r in all_rows:
    key = (r.get('GaraDescrizione','').strip(), r.get('ImpiantoNome','').strip())
    tornei[key].add(r.get('DataOra','')[:10])

print('\n=== TORNEI CON PIU DATE (top 10) ===')
multi = [(k,v) for k,v in tornei.items() if len(v)>1]
multi.sort(key=lambda x: -len(x[1]))
for (gara,imp), dates in multi[:10]:
    print(f'  {gara} @ {imp}: {len(dates)} date -> {sorted(dates)[:5]}')

# 5
acc = stati.get('gara accettata', 0)
rif = stati.get('gara rifiutata', 0)
prop = stati.get('proposta di designazione', 0)
vuote = stati.get('', 0)
print(f'\n=== RIFIUTATE vs ACCETTATE ===')
print(f'  gara accettata: {acc}')
print(f'  gara rifiutata: {rif}')
print(f'  proposta di designazione: {prop}')
print(f'  vuote (senza arbitro): {vuote}')

# Also group by GaraId to see unique tournaments
gara_ids = defaultdict(lambda: {'dates': set(), 'desc': '', 'imp': ''})
for r in all_rows:
    gid = r.get('GaraId','').strip()
    if gid:
        gara_ids[gid]['dates'].add(r.get('DataOra','')[:10])
        gara_ids[gid]['desc'] = r.get('GaraDescrizione','').strip()
        gara_ids[gid]['imp'] = r.get('ImpiantoNome','').strip()

print(f'\n=== TORNEI UNICI (per GaraId) ===')
print(f'  Totale GaraId distinti: {len(gara_ids)}')
for gid, info in sorted(gara_ids.items()):
    print(f'  {gid}: {info["desc"]} @ {info["imp"][:30]} ({sorted(info["dates"])})')
