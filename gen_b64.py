import base64, pathlib
for name in ('logo.png','stamp.png'):
    data = pathlib.Path(name).read_bytes()
    b64 = base64.b64encode(data).decode('ascii')
    out = pathlib.Path(name + '.b64.txt')
    out.write_text(b64)
    print(f"wrote {out} ({len(b64)} chars)")
