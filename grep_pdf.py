import pathlib

path = pathlib.Path('test_images.pdf')
if not path.exists():
    print('missing pdf')
    raise SystemExit(1)

data = path.read_bytes()
text = data.decode('latin1', errors='ignore')
for line in text.split('\n'):
    if '/Image' in line:
        print(line)
        break
else:
    print('no /Image found')
