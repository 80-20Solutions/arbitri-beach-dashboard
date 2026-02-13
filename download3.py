import subprocess, json

gog = r'C:\Users\KreshOS\go\bin\gog.exe'

sheets = [
    ('foglio1', '1wWqcc8Qa6DRGCG767ZbRViTjYm5cwkrIih6hAMdYTuA'),
    ('foglio2', '1X3wbqD5cUeWyNCtJ-NBghM2A2XuSYgV1F6SHsf0mP9I'),
    ('foglio3', '1xdXXqzK3pbaHAJOPxLT7ME9BWr-dRlYAGhnBYMYVJ00'),
    ('principale', '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo'),
]

# Try A1:Z50 with no sheet name prefix
for name, sid in sheets:
    print(f"\n=== {name} ===")
    r = subprocess.run([gog, 'sheets', 'get', sid, 'A1:Z50', '--plain'], 
                       capture_output=True, text=True)
    if r.returncode == 0:
        lines = r.stdout.strip().split('\n')
        for l in lines[:10]:
            print(l[:200])
        print(f"... total {len(lines)} lines")
    else:
        print(f"ERROR: {r.stderr.strip()[:300]}")
    
    # Also try to get sheet names via JSON
    r2 = subprocess.run([gog, 'sheets', 'get', sid, 'A1:A1', '--json'], 
                        capture_output=True, text=True)
    if r2.returncode == 0:
        print(f"JSON sample: {r2.stdout[:300]}")
