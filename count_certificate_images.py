import pathlib
path = pathlib.Path('certificate.pdf')
if not path.exists():
    print('missing certificate.pdf')
    raise SystemExit(1)

data = path.read_bytes().decode('latin1', errors='ignore')
print('/Subtype /Image count:', data.count('/Subtype /Image'))
