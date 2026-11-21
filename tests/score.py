"""Scores the reader's bill files against values read by hand: right, wrong but marked, wrong and NOT marked."""
import glob, json, sys
from openpyxl import load_workbook
truth = json.load(open(sys.argv[1]))
right = marked = silent = 0
silent_list = []
for f in sorted(glob.glob(sys.argv[2] + '/bills/*.xlsx')):
    key = '-'.join(f.split('/')[-1].split('-')[:2]).replace('.xlsx', '')
    if key not in truth:
        continue
    wb = load_workbook(f)
    rows = [r for r in list(wb.worksheets[0].iter_rows(values_only=True))[1:]]
    want = truth[key]['lines']
    if len(rows) != len(want):
        print(f'{key}: {len(rows)} lines read, bill has {len(want)}')
    for i, w in enumerate(want):
        r = rows[i] if i < len(rows) else [None] * 15
        flagged = bool(r[14])
        got = {'batch': r[1], 'expiry': r[2], 'qty': r[3], 'rate': r[6]}
        for name, val in zip(('batch', 'expiry', 'qty', 'rate'), w[:4]):
            if val is None:
                continue
            ok = (str(got[name]) == str(val)) if name in ('batch', 'expiry') else got[name] is not None and abs(float(got[name]) - val) < 0.005
            if ok:
                right += 1
            elif flagged:
                marked += 1
                silent_list.append(f'   marked: {key} line {i + 1} {name}: {got[name]} (bill: {val})')
            else:
                silent += 1
                silent_list.append(f'{key} line {i + 1} {name}: {got[name]} (bill: {val})')
    d = {r[0]: r[1] for r in wb.worksheets[1].iter_rows(values_only=True)}
    checks = ' '.join(str(r[1]) for r in wb.worksheets[1].iter_rows(values_only=True) if r[0] == 'Check')
    for name, val in zip(('Total', 'Less discount', 'Rounding', 'Net total'), truth[key]['sums']):
        ok = d.get(name) is not None and abs(float(d[name]) - val) < 0.005
        if ok:
            right += 1
        elif name in checks or 'by hand' in checks or 'not read' in checks:
            marked += 1
            silent_list.append(f'   marked: {key} {name}: {d.get(name)} (bill: {val})')
        else:
            silent += 1
            silent_list.append(f'{key} {name}: {d.get(name)} (bill: {val})')
total = right + marked + silent
print(f'{total} values: {right} right ({100 * right / total:.0f}%), {marked} wrong but marked, {silent} wrong and NOT marked')
for s in silent_list:
    print(s if s.startswith('   marked') else '   NOT MARKED: ' + s)
