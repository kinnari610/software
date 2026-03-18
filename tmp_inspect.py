import app, pprint

tbl = app._extract_no_load_table("2604810-H")
print("rows", len(tbl) if tbl else None)
pprint.pprint(tbl[:10])
