import base64
from pathlib import Path

for name in ['logo', 'stamp']:
    b64_path = Path(f"{name}.png.b64.txt")
    out_path = Path(f"{name}.png")
    if not b64_path.exists():
        print(f"Missing {b64_path}")
        continue
    b64 = b64_path.read_text(encoding='utf-8')
    b64 = ''.join(b64.split())
    data = base64.b64decode(b64)
    out_path.write_bytes(data)
    print(f"wrote {out_path} ({len(data)} bytes)")
