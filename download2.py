import subprocess, os

gog = r'C:\Users\KreshOS\go\bin\gog.exe'

sheets = {
    'foglio1': '1wWqcc8Qa6DRGCG767ZbRViTjYm5cwkrIih6hAMdYTuA',
    'foglio2': '1X3wbqD5cUeWyNCtJ-NBghM2A2XuSYgV1F6SHsf0mP9I',
    'foglio3': '1xdXXqzK3pbaHAJOPxLT7ME9BWr-dRlYAGhnBYMYVJ00',
}

# First check gog help for sheets commands
r = subprocess.run([gog, 'help'], capture_output=True, text=True)
print(r.stdout[:2000])
print("STDERR:", r.stderr[:500])
