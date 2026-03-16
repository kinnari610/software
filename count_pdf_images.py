import pathlib

path = pathlib.Path('test_images.pdf')
if not path.exists():
    print('missing pdf')
    raise SystemExit(1)

data = path.read_bytes().decode('latin1', errors='ignore')
count = data.count('/Subtype /Image')
print('/Subtype /Image count:', count)
