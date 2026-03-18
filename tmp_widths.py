import app

tbl = app._extract_no_load_table('2604810-H')
col_count = max(len(r) for r in tbl)
col_widths = []
for c in range(col_count):
    max_len = max(len(str(r[c])) for r in tbl)
    w = max(40, min(130, max_len * 5))
    col_widths.append(w)
page_width = (8.27*72) - 20 - 20
print('col_count', col_count)
print('raw widths', col_widths)
print('sum', sum(col_widths), 'page', page_width)
if sum(col_widths) > page_width:
    scale = page_width / sum(col_widths)
    col_widths = [w * scale for w in col_widths]
    print('scaled widths', col_widths)
    print('scaled sum', sum(col_widths))
