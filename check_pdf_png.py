import pathlib

p = pathlib.Path('test_images.pdf')
if not p.exists():
    print('missing test_images.pdf')
    raise SystemExit(1)

data = p.read_bytes()
print('PNG signatures in test_images.pdf:', data.count(b'\x89PNG'))
