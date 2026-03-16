import pathlib
import sys

path = pathlib.Path('certificate.pdf')
if not path.exists():
    print('certificate.pdf not found')
    sys.exit(1)

data = path.read_bytes()
count = data.count(b'\x89PNG')
print('PNG signatures found in PDF:', count)
idx = data.find(b'\x89PNG')
print('first PNG at', idx)
if idx != -1:
    print(data[idx:idx+20])
