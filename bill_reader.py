#!/usr/bin/env python3
"""
Reads distributor bill photos offline into PharmaSathi's import files. Nothing is sent anywhere.

    python bill_reader.py PHOTO_FOLDER [--out FOLDER] [--medicines MEDICINES_EXPORT.xlsx] [--debug]

Writes into PHOTO_FOLDER/bills-read (or --out):
    names.xlsx          every medicine name seen on the bills and the medicine it means. Check the
                        rows with a Note, correct Name/Strength/Unit/Strip size, delete the Note.
                        Your changes are kept on the next run.
    new-medicines.xlsx  the medicines in names.xlsx that the shop does not have yet, in the
                        Medicines import template. Import it before the bills.
    bills/*.xlsx        one purchase-template file per bill, with a Check column.
    summary.txt         one line per photo.

Run it, check names.xlsx, run it again (photos already read are not read again), then import
new-medicines.xlsx and the bill files. --medicines takes a Medicines export from PharmaSathi so
medicines the shop already has are matched, not added again.
"""
import argparse
import itertools
import json
import re
import statistics
from collections import Counter
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

INR_TO_NPR = Decimal('1.6')
CENT = Decimal('0.01')
BILL_SHEET = "Purchase (a distributor's bill)"
BILL_COLUMNS = ['Medicine', 'Batch no.', 'Expiry', 'Qty', 'Unit', 'Strip size', 'Rate', 'Sell price', 'Disc %',
                'Bonus qty', 'Bonus batch', 'Bonus expiry', 'CC %']
MEDICINE_COLUMNS = ['SKU', 'Barcode', 'Name', 'Generic', 'Strength', 'Category', 'Manufacturer', 'Unit', 'Strip size',
                    'Strip cost', 'Strip price', 'Cost price', 'Selling price', 'Reorder level', 'Prescription', 'Active']
NAME_COLUMNS = ['Bill name', 'Name', 'Strength', 'Unit', 'Strip size', 'Bills', 'Distributor', 'Note']
UNITS = ['tablet', 'capsule', 'strip', 'box', 'bottle', 'syrup', 'vial', 'ampoule', 'tube', 'sachet', 'injection',
         'inhaler', 'drops', 'piece', 'pack']  # config/pharmacy.php 'units'
STRIP_WORDS = {'TAB', 'TABS', 'TABLET', 'TABLETS', 'CAP', 'CAPS', 'CAPSULE', 'CAPSULES'}
DOSE = r'[\d.]+(?:MG|MCG|ML|GM|G|IU|%|KG)?(?:/[\d.]+)?(?:MG|MCG|ML|GM|G|IU|%|KG)(?:/[\d.]+(?:MG|MCG|ML|GM|G|IU|%|KG)?)?'
YELLOW = PatternFill('solid', fgColor='FFF2A8')

# Header words -> column, first match wins.
HEADER_KEYS = [
    ('free', ('FREE', 'BONUS', 'SCH')),
    ('name', ('ITEM', 'DESCRIPTION', 'PARTICULAR', 'PRODUCT')),
    ('pack', ('PACK',)),
    ('batch', ('BATCH',)),
    ('expiry', ('EXP',)),
    ('qty', ('QTY', 'QUANTITY')),
    ('rate', ('RATE',)),
    ('amount', ('AMOUNT', 'AMT')),
    ('mrp', ('MRP', 'M.R.P')),
    ('disc', ('DISC',)),
]


# ---------------------------------------------------------------- reading the photo

class Box:
    def __init__(self, points, text, score):
        self.points, self.text, self.score = points, text.strip(), float(score)
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        self.x0, self.x1, self.y0, self.y1 = min(xs), max(xs), min(ys), max(ys)
        self.w, self.h = self.x1 - self.x0, self.y1 - self.y0
        self.cx, self.cy = sum(xs) / 4, sum(ys) / 4
        (ax, ay), (bx, by) = points[0], points[1]
        self.slope = (by - ay) / (bx - ax) if bx != ax else 0.0


def ocr(engine, image):
    r = engine(image)
    if r.txts is None:
        return []
    return [Box([[float(v) for v in p] for p in b], t, s) for b, t, s in zip(r.boxes, r.txts, r.scores)]


def rows_of(boxes):
    """Groups boxes into printed lines, allowing for a photo taken at a slight slant."""
    wide = [b.slope for b in boxes if b.w > 3 * b.h]
    slope = statistics.median(wide) if wide else 0.0
    height = statistics.median(b.h for b in boxes)
    keep = [b for b in boxes if b.h < 2.2 * height]  # tall boxes are pen ticks down a column
    for b in keep:
        b.level = b.cy - slope * b.cx
    keep.sort(key=lambda b: b.level)
    rows = []
    for b in keep:
        if rows and abs(rows[-1][-1].level - b.level) < 0.45 * height:
            rows[-1].append(b)
        else:
            rows.append([b])
    return [sorted(r, key=lambda b: b.x0) for r in rows]


def row_text(row):
    return ' '.join(b.text for b in row).upper()


def words_of(text):
    """The words of a line for matching labels: M.R.P. -> MRP, CC/RATE -> CC RATE."""
    return re.findall(r'[A-Z0-9]+', text.upper().replace('.', ''))


def near(word, target):
    """word starts with target, allowing one misread letter (3ATCH, AMDUNT, BTY for QTY)."""
    head = word[:len(target)]
    if head == target:
        return True
    if len(head) == len(target) and len(target) >= 3:
        return sum(a != b for a, b in zip(head, target)) <= (2 if len(target) >= 6 else 1)
    return one_letter_apart(word, target)


def has_word(text, *targets):
    return any(near(w, t) for w in words_of(text) for t in targets)


HEADER_WORDS = ('QTY', 'RATE', 'AMOUNT', 'BATCH', 'PACK', 'EXP', 'DISC', 'ITEM', 'DESCRIPTION', 'MRP')


def is_header(row):
    """The table's header line: at least four of its usual words, each allowed one misread letter."""
    words = words_of(row_text(row))
    return sum(any(near(w, t) for w in words) for t in HEADER_WORDS) >= 4


def sharpen(image):
    """Evens out light and contrast, so faint ribbon print reads better."""
    import cv2
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return cv2.cvtColor(cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8)).apply(gray), cv2.COLOR_GRAY2BGR)


# Two ways to read a photo: the small recogniser on the photo as taken (best on clear print), and the medium one
# (73 MB) on a sharpened copy (best on faint ribbon print). Both are kept; the bill's own sums pick between them.
READINGS = ('small', 'medium-sharpened')


def engine_for(name, engines):
    if name not in engines:
        from rapidocr import RapidOCR
        from rapidocr.utils.typings import ModelType
        engines[name] = RapidOCR() if name == 'small' else RapidOCR(params={'Rec.model_type': ModelType.MEDIUM})
    return engines[name]


def read_photo(photo, cache, engines):
    """{reading: the bill's printed lines}. What a photo says is kept in the cache, so a second run is quick."""
    stat = photo.stat()
    out, image = {}, None
    for name in READINGS:
        saved = cache / f'{photo.name}.{stat.st_size}-{int(stat.st_mtime)}.{name}.json'
        if saved.exists():
            boxes = [Box(*b) for b in json.loads(saved.read_text())]
            if boxes:
                out[name] = rows_of(boxes)
            continue
        import cv2  # only needed when a photo has not been read before
        if image is None:
            image = cv2.imread(str(photo))
            if image is None:
                raise ValueError('not a picture this tool can open')
        engine = engine_for(name, engines)
        picture = sharpen(image) if name != 'small' else image
        boxes = ocr(engine, picture)
        turns = [None, cv2.ROTATE_180]
        if boxes and sum(b.h > b.w for b in boxes) > len(boxes) / 2:  # text runs sideways
            turns = [cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_90_CLOCKWISE]
            boxes = None
        found = []
        for turn in turns:
            if boxes is None:
                boxes = ocr(engine, picture if turn is None else cv2.rotate(picture, turn))
            if boxes and any(is_header(r) for r in rows_of(boxes)):
                found = boxes
                break
            boxes = None
        saved.write_text(json.dumps([[b.points, b.text, b.score] for b in found]))
        if found:
            out[name] = rows_of(found)
    if not out:
        raise ValueError('could not find the table header (QTY / RATE / AMOUNT) on the bill')
    return out


# ---------------------------------------------------------------- reading values

def money(text):
    t = text.replace(',', '.').replace(' ', '').replace('°', '.')
    if t.count('.') > 1:  # 1,307.50 or 7.967.00: only the last point is the decimal one
        head, _, tail = t.rpartition('.')
        t = head.replace('.', '') + '.' + tail
    t = re.sub(r'[Oo]', '0', t)
    t = re.sub(r'[lI|]', '1', t)
    m = re.search(r'-?\d+(\.\d+)?', t)
    try:
        return Decimal(m.group()) if m else None
    except InvalidOperation:
        return None


def percent(text):
    t = text.replace(' ', '')
    m = re.match(r'^(\d+\.\d\d)8$', t)  # "3.00%" read as "3.008"
    return Decimal(m.group(1)) if m else money(t.rstrip('%'))


def fix_digits(word):
    """2OOMG -> 200MG, 40M6 -> 40MG, 2.SNG -> 2.5MG, Z0RYL -> ZORYL."""
    word = re.sub(r'(?<=\d\.)S', '5', word)
    word = re.sub(r'^([\d.]+)(M6|NG|HG|MO|M9|N6)$', r'\1MG', word)
    if re.fullmatch(r'[\dO]*\d[\dO]*(MG|MCG|ML|GM|G|IU|%)?', word):
        return re.sub(r'O(?=[\dO]*(MG|MCG|ML|GM|G|IU|%)?$)', '0', word)
    if re.fullmatch(r'[\dO]{1,2}[A-Z]{3,}', word) and not re.fullmatch(r'[\d.]+(MG|MCG|ML|GM|IU|KG)', word):
        word = word[:2].translate(str.maketrans('205186', 'ZOSIBG')) + word[2:]
    if re.search(r'[A-Z]{2}', word) and not re.match(r'[\d.]+[A-Z%]', word):
        word = re.sub(r'(?<=[A-Z])0(?=[A-Z])|^0(?=[A-Z])|(?<=[A-Z])0$', 'O', word)
    return word


MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


def expiry(text):
    t = text.replace(' ', '')
    named = re.search(r'(' + '|'.join(MONTHS) + r')[A-Z]*[-/.]?(\d{2,4})', t.upper())
    if named:
        year = int(named.group(2))
        return f'{MONTHS.index(named.group(1)) + 1:02d}/{year + 2000 if year < 100 else year}'
    m = re.search(r'(\d{4})[/\-.](\d{1,2})', t) or re.fullmatch(r'(\d{4})[1lI|\\7](\d{2})', t)  # "2027107": / read as 1 or 7
    if m:
        year, month = int(m.group(1)), int(m.group(2))
    else:
        m = re.search(r'(\d{1,2})[/\-.](\d{2,4})$', t)
        if not m:
            return None
        month, year = int(m.group(1)), int(m.group(2))
        year += 2000 if year < 100 else 0
    if month > 99:
        return None
    return f'{month:02d}/{year}'  # an impossible month (17) is repaired by fix_year, which marks it


def pack(text):
    """'14 TAB' -> (14, 'TAB'); OCR's 8 for B and 0 for O are put right."""
    compact = re.sub(r'[^A-Z0-9]', '', text.upper())
    # "201AB" is "20 TAB": the shortest number that leaves a word like TAB or CAP.
    m = re.match(r'^(\d+?)([1I7T]?[A4][B83RSOD]|C[A4][PF])$', compact) or re.match(r'^(\d+)\s*([A-Z0-9]+)', compact)
    if not m or int(m.group(1)) > 100:
        return None, ''
    word = m.group(2).replace('8', 'B').replace('0', 'O')
    if re.fullmatch(r'C[A4][PF]', word):
        word = 'CAP'
    elif re.fullmatch(r'[TI1J7]?[A4][B3RSOD]', word):  # TAB as faint print reads: 1AR, TA3, A3
        word = 'TAB'
    return int(m.group(1)), word


# ---------------------------------------------------------------- names

def bill_words(text):
    words = [fix_digits(w) for w in re.sub(r'^\d+\.\s*', '', text).upper().split()]
    # TAB / CAP as OCR misreads them: TA3, TA8, IAB, T4B, CAF.
    out = []
    for i, w in enumerate(words):
        last = i == len(words) - 1
        if re.fullmatch(r'[TI7][A4][B83]', w) or (last and i > 0 and re.fullmatch(r'[TI1J7(]?[A4][B83RSOD]', w)):
            w = 'TAB'
        elif re.fullmatch(r'C[A4][PF]', w):
            w = 'CAP'
        out.append(w)
    return out


def is_ditto(text):
    """'- do -' on a bill: the same medicine as the line above."""
    t = text.strip()
    return re.fullmatch(r'[-_~\s]*d[o0][-_~.\s]*', t, re.IGNORECASE) is not None \
        or ('-' in t and re.fullmatch(r'[-_~.\s]*[A-Za-z0-9]{1,2}[-_~.\s]*', t) is not None)  # "- da-", "-0-"


def title(word):
    if len(word) <= 2 and word.isalpha():
        return word
    return word.capitalize()


def split_name(bill_name):
    """'AXCER 90MG TAB' -> ('Axcer', '90mg', 'TAB'). A strength with no unit ('MONOTRATE 10') stays a bare number."""
    words = bill_words(bill_name)
    form = ''
    while len(words) > 1 and words[-1] in STRIP_WORDS:
        form = form or words[-1]
        words.pop()
    strength = ''
    for i in range(len(words) - 1, 0, -1):
        if re.fullmatch(DOSE, words[i]):
            strength = words.pop(i).lower()
            break
    else:
        if len(words) > 1 and re.fullmatch(r'\d+(\.\d+)?', words[-1]):
            strength = words.pop()
    return ' '.join(title(w) for w in words), strength, form


def unit_for(pack_word, form, name):
    n = name.upper()
    word = pack_word or form
    if re.search(r'\bDROPS?\b', n):
        return 'drops'
    if word in ('TAB', 'TABS', 'TABLET', 'TABLETS') or form in ('TAB', 'TABS', 'TABLET', 'TABLETS'):
        return 'tablet'
    if word in ('CAP', 'CAPS', 'CAPSULE', 'CAPSULES') or form in ('CAP', 'CAPS', 'CAPSULE', 'CAPSULES'):
        return 'capsule'
    for pattern, unit in ((r'\b(SYP|SYRUP|SUSP|SUSPENSION)\b', 'syrup'), (r'\b(INJ|INJECTION|CART|CARTRIDGE|PENFILL)\b', 'injection'),
                          (r'\b(CREAM|GEL|OINT|OINTMENT)\b', 'tube'), (r'\b(INHALER|ROTACAP|RESPULES)\b', 'inhaler'),
                          (r'\b(SACHET|POWDER)\b', 'sachet')):
        if re.search(pattern, n):
            return unit
    return {'VIAL': 'vial', 'AMP': 'ampoule', 'AMPOULE': 'ampoule', 'SACHET': 'sachet', 'TUBE': 'tube',
            'PH': 'bottle', 'BTL': 'bottle', 'BOTTLE': 'bottle', 'ML': 'bottle'}.get(word, 'piece')


def bare(strength):
    """A strength printed without its unit: '10', not '10mg'."""
    return not re.search(r'[A-Za-z%]', strength or '')


def key(name, strength):
    return re.sub(r'[^A-Z0-9]', '', name.upper()), re.sub(r'[^0-9.]', '', strength)


def one_letter_apart(a, b):
    if abs(len(a) - len(b)) > 1 or min(len(a), len(b)) < 4:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    short, long_ = sorted((a, b), key=len)
    return any(long_[:i] + long_[i + 1:] == short for i in range(len(long_)))


def same_medicine(a, b):
    """
    Two spellings of one medicine: the same strength digits, the same number of words, and at most one
    word a single letter apart (WINOLAP / MINOLAP). Kept strict so Minolap and Minolap DS, or 10mg and
    20mg, are never joined.
    """
    (an, as_), (bn, bs) = a, b
    if re.sub(r'[^0-9.]', '', as_) != re.sub(r'[^0-9.]', '', bs):
        return False
    aw, bw = an.upper().split(), bn.upper().split()
    if len(aw) != len(bw):
        return False
    diff = [(x, y) for x, y in zip(aw, bw) if x != y]
    return not diff or (len(diff) == 1 and one_letter_apart(*diff[0]))


# ---------------------------------------------------------------- one bill

def columns(header):
    cols = []
    for b in header:
        for k, words in HEADER_KEYS:
            if has_word(b.text, *[w.replace('.', '') for w in words]):
                cols.append((k, b))
                break
    return cols


COLUMN_ORDER = ['name', 'pack', 'batch', 'expiry', 'qty', 'free', 'rate', 'amount', 'mrp', 'disc']


def fill_gaps(header, cols):
    """A header cell no word matched, between two that did with one column missing between them, is that column."""
    known = {id(b): k for k, b in cols}
    cells = sorted(header, key=lambda b: b.cx)
    for i, b in enumerate(cells):
        if id(b) in known or not re.search(r'[A-Za-z]{3}', b.text):
            continue
        left = next((known[id(c)] for c in reversed(cells[:i]) if id(c) in known), None)
        right = next((known[id(c)] for c in cells[i + 1:] if id(c) in known), None)
        if left in COLUMN_ORDER and right in COLUMN_ORDER:
            between = [k for k in COLUMN_ORDER[COLUMN_ORDER.index(left) + 1:COLUMN_ORDER.index(right)]
                       if k not in known.values() and k != 'free']
            if len(between) == 1:
                cols.append((between[0], b))
                known[id(b)] = between[0]
    return cols


def column_for(cell, cols):
    def overlap(h):
        return min(cell.x1, h.x1) - max(cell.x0, h.x0)
    best = max(cols, key=lambda c: overlap(c[1]))
    if overlap(best[1]) > 0:
        return best[0]
    return min(cols, key=lambda c: abs(c[1].cx - cell.cx))[0]


def beside(start, candidates):
    """The candidate on the same printed line as start, to its right, following start's own slant (pages curl)."""
    best, gap = None, 0.6 * start.h
    for c in candidates:
        if c.cx <= start.cx:
            continue
        miss = abs(c.cy - (start.cy + start.slope * (c.cx - start.cx)))
        if miss < gap:
            best, gap = c, miss
    return best


def label_value(boxes, labels, avoid=(), anywhere=False):
    """The amount printed to the right of a label. anywhere: the label may be misread in front ("ILSS DISCOUNT")."""
    numbers = [c for c in boxes if re.fullmatch(r'-?[\d.,]+[.,°]\d\d', c.text.replace(' ', ''))]  # 7.777.30 too
    for b in boxes:
        t = b.text.upper().rstrip(': ')
        hit = any((l in t) if anywhere else t.startswith(l) for l in labels)
        if hit and not any(w in t for w in avoid):
            found = beside(b, numbers)
            if found:
                return money(found.text)
    return None


def bill_info(rows):
    """Supplier, PAN, bill number, date and pay mode from the lines above the table."""
    head = rows[:next((i for i, r in enumerate(rows) if is_header(r)), len(rows))]
    info = {'Supplier': rows[0][0].text if rows else ''}
    dates = []
    for t in (row_text(r) for r in head):
        m = re.search(r'AN\s*[:.]\s*(\d{9})', t)  # "PAN:" is sometimes read as "AN:"
        if m and 'PAN' not in info:  # the supplier's PAN comes first; the shop's own is further down
            info['PAN'] = m.group(1)
        if re.search(r'CREDIT\s*NOTE|GOODS\s*RETURN|DEBIT\s*NOTE', t):
            info['Kind'] = 'Credit note (goods returned)'
        m = re.search(r'(I?NVOICE|BILL|CREDIT\s*NOTE)\s*NO\.?\s*:?\s*([A-Z0-9/\-]+)', t)
        if m and 'Bill no.' not in info:
            info['Bill no.'] = m.group(2)
        dates += re.findall(r'\b(20\d\d/\d\d/\d\d)\b', t)
        m = re.search(r'PAYMODE\s*:?\s*(\w+)', t)
        if m:
            info['Pay mode'] = m.group(1).capitalize()
    if any(re.search(r'CREDIT\s*NOTE|GOODS\s*RETURN', row_text(r)) for r in rows[:len(head) + 3]):
        info['Kind'] = 'Credit note (goods returned)'
    # Nepali (BS) dates are 2070 onwards; the bill prints both, and the shop keeps BS.
    bs = [d for d in dates if d >= '2070' and 1 <= int(d[5:7]) <= 12 and 1 <= int(d[8:]) <= 32]
    if bs or dates:
        info['Bill date'] = (bs or dates)[0]
    return info


# Digits that faint ribbon print turns into one another (8 read as 3, 2 as 7, 9 as 4, ...).
CONFUSED = {'3': '85', '8': '3', '2': '7', '7': '219', '5': '36', '6': '5', '9': '47', '4': '9', '1': '7', '0': '8'}


def variants(value, swaps):
    """value and the numbers it might really be if up to `swaps` digits were misread -> number of digits changed."""
    if value is None:
        return {}
    out, frontier = {value: 0}, {format(value, 'f'): 0}
    for _ in range(swaps):
        nxt = {}
        for text, cost in frontier.items():
            for i, ch in enumerate(text):
                for alt in CONFUSED.get(ch, ''):
                    nxt.setdefault(text[:i] + alt + text[i + 1:], cost + 1)
        for text, cost in nxt.items():
            out.setdefault(Decimal(text), cost)
        frontier = nxt
    return out


def line_options(q, r, a):
    """(digits changed, qty, rate, amount) where qty × rate = amount, fewest changes first."""
    rates, amounts = variants(r, 2), variants(a, 2)
    qtys = variants(q, 1) if q is not None else {}
    opts = set()
    for rv, rc in rates.items():
        if rv <= 0:
            continue
        for av, ac in amounts.items():
            n = av / rv
            if n <= 0 or n != n.to_integral_value() or n > 1000:
                continue
            n = n.quantize(Decimal('1'))
            # A qty read wrong by a pen tick ("01", "55") is not a digit swap: it costs one change more.
            qc = qtys.get(n, 2 if q is not None else 1)
            opts.add((rc + ac + qc, n, rv, av))
    if not opts and q and a and q > 0 and (a / q) == (a / q).quantize(CENT):
        opts.add((2, q, (a / q).quantize(CENT), a))  # rate misread beyond the usual swaps
    return sorted(opts)[:8]


def repair(lines, total, discount, rounding, net, words=None):
    """
    The reading of the whole bill that makes its sums agree with the fewest digits changed: each line's
    qty × rate = amount, the lines add up to the total, and total − discount + rounding = net. A line's
    amount is never bent to fit a misread total: the total has to agree with the net as well.
    Returns (line choices, total, (discount, rounding, net) or None, score).
    """
    by_net = {}  # total -> (changes, discount, rounding, net) that agree with it
    if words is not None:
        net = words  # the amount in words decides the net total
    if discount is None and net is not None and total is not None:
        for t, tc in variants(total, 2).items():  # discount not read: any total gives one, but a small one
            d = (t + (rounding or Decimal(0)) - net).quantize(CENT)
            if 0 <= d <= t * Decimal('0.2') and (t not in by_net or tc + 2 < by_net[t][0]):
                by_net[t] = (tc + 2, d, rounding or Decimal(0), net)
    if discount is not None and net is not None:
        for dv, dc in variants(discount, 2).items():
            for rv, rc in variants(rounding or Decimal(0), 1).items():
                if abs(rv) > Decimal('0.5'):  # rounding to the nearest rupee is never more than 50 paisa
                    continue
                for nv, nc in (variants(net, 2) if words is None else {words: 0}).items():
                    t = (nv + dv - rv).quantize(CENT)
                    if t > 0 and (t not in by_net or dc + rc + nc < by_net[t][0]):
                        by_net[t] = (dc + rc + nc, dv, rv, nv)
    totals = variants(total, 2) if total is not None else {}

    def bill_cost(t):
        """Changes for the total and the net to agree with lines adding up to t; None if neither does."""
        a, b = totals.get(t), by_net.get(t, (None,))[0]
        if a is None and b is None:
            return None
        # The net total is checked by its own sum, so disagreeing with it weighs more than a misread total.
        return (a if a is not None else 3) + (b if b is not None else (6 if by_net else 0))

    options = [line_options(l['qty'], l['rate'], l['amount']) for l in lines]
    keep = 8  # fewer candidates per line on long bills, so the search stays quick
    while keep > 1 and keep ** max(1, len(options)) > 200000:
        keep -= 1
    options = [o[:keep] for o in options]
    best = None
    if options and all(options):
        for combo in itertools.product(*options):
            t = sum(c[3] for c in combo)
            extra = bill_cost(t)
            if extra is None:
                continue
            cost = sum(c[0] for c in combo) + extra
            if best is None or cost < best[0]:
                best = (cost, combo, t)
    if best is None:
        # Nothing agrees with the total: keep each line as read where it already adds up.
        return [o[0] if o and o[0][0] == 0 else None for o in options], None, None, None
    t = best[2]
    return list(best[1]), t, by_net.get(t, (None, None, None, None))[1:] if t in by_net else None, best[0]


NUMBER_WORDS = {w: i for i, w in enumerate(
    'zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen '
    'seventeen eighteen nineteen'.split())}
NUMBER_WORDS.update({w: 10 * (i + 2) for i, w in enumerate('twenty thirty forty fifty sixty seventy eighty ninety'.split())})
SCALE_WORDS = {'hundred': 100, 'thousand': 1000, 'lakh': 100000, 'lac': 100000, 'crore': 10000000}


def in_words(rows):
    """The net total from 'Inwords Rs.: Six Hundred Sixty-five Only.' (words read far better than faint digits)."""
    for r in rows:
        text = row_text(r)
        if not re.search(r'ONLY', text) or not re.search(r'WORD|RS', text):
            continue
        words = re.findall(r'[a-z]+', text.lower().split(':', 1)[-1])
        total, part, seen = 0, 0, False
        for w in words:
            w = re.sub(r'^ll', 'h', w)  # H read as ll: llundred
            if w in ('only', 'rs', 'and', 'inwords', 'words'):
                continue
            match = w if w in NUMBER_WORDS or w in SCALE_WORDS else next(
                (k for k in list(NUMBER_WORDS) + list(SCALE_WORDS) if len(k) >= 4 and one_letter_apart(w, k)), None)
            if match is None:
                if w.endswith('only'):
                    break
                return None  # a word it cannot place: do not guess
            seen = True
            if match in NUMBER_WORDS:
                part += NUMBER_WORDS[match]
            elif match == 'hundred':
                part *= 100
            else:
                total += part * SCALE_WORDS[match]
                part = 0
        return Decimal(total + part) if seen else None
    return None


def fix_year(exp, notes):
    """An expiry year outside the next few years is a misread digit (2023 for 2028, 7029 for 2029)."""
    if not exp:
        return exp
    month, year = exp.split('/')
    if not 1 <= int(month) <= 12:
        # Known digit swaps first, in their usual order (17 -> 12, as 7 is read for 2); a month from 13 to 19
        # with no such swap is taken as 12, the month whose 2 looks most like the misread digit.
        swaps = [month[:i] + alt + month[i + 1:] for i, ch in enumerate(month) for alt in CONFUSED.get(ch, '')]
        fit = next((m for m in swaps if 1 <= int(m) <= 12), None)
        guess = fit or ('12' if 13 <= int(month) <= 19 else None)
        if guess is None:
            notes.append(f'Expiry month "{month}" cannot be right: type the expiry from the bill.')
            return None
        notes.append(f'Expiry month read as "{month}"; {guess}/{year} is the nearest real month'
                     f'{"" if fit else " (a guess)"}. Check the bill.')
        month, exp = guess, f'{guess}/{year}'
    now = date.today().year
    if now <= int(year) <= now + 7:  # an expiry before this year cannot be on a bill being entered now
        return exp
    fits = sorted((c, y) for y, c in variants(Decimal(year), 2).items() if now <= y <= now + 7)
    if fits:
        notes.append(f'Expiry read as {exp}; {month}/{fits[0][1]} is the nearest that makes sense. Check the bill.')
        return f'{month}/{fits[0][1]}'
    notes.append(f'Expiry {exp} looks wrong; check the bill.')
    return exp


def parse(rows):
    head = next(i for i, r in enumerate(rows) if is_header(r))
    cols = fill_gaps(rows[head], columns(rows[head]))
    end = next((i for i in range(head + 1, len(rows)) if has_word(row_text(rows[i]), 'TOTAL', 'CUSTOMER')), len(rows))
    table = [b for r in rows[head + 1:end] for b in r]
    by_col = {}
    for b in table:
        by_col.setdefault(column_for(b, cols), []).append(b)
    order = [k for k, _ in sorted(cols, key=lambda c: c[1].cx) if k != 'name']
    anchors = sorted((b for b in by_col.get('name', []) if re.search(r'[A-Za-z]{3}', b.text) or is_ditto(b.text)),
                     key=lambda b: b.cy)
    used, lines, checks = set(), [], []
    for anchor in anchors:
        # Walk right from the medicine name one column at a time, so a curled page still keeps each line together.
        cells, cur = {}, anchor
        for k in order:
            nxt = beside(cur, [b for b in by_col.get(k, []) if id(b) not in used and b.cx > anchor.cx])
            if nxt is not None:
                cells[k] = nxt
                used.add(id(nxt))
                cur = nxt
        if not any(k in cells for k in ('qty', 'rate', 'amount')):
            continue  # a name with no numbers beside it is not a bill line
        get = lambda k: cells[k].text if k in cells else ''
        sure = lambda k: cells[k].score if k in cells else 0
        glued = re.fullmatch(r'[1LI]X(\d{1,3})\s*T[A4][B8]?\s*([A-Z0-9]{5,})', get('batch').strip().upper())
        if glued and 'batch' in cells:
            cells['batch'] = Box(cells['batch'].points, glued.group(2), cells['batch'].score)
        joined = re.fullmatch(r'(\d+\s*\S{1,5}(?:\s\S{1,3})?)\s+([A-Z0-9]{6,})', get('batch').strip().upper())
        if joined and 'pack' not in cells:
            cells['pack'] = Box(cells['batch'].points, joined.group(1), cells['batch'].score)
            cells['batch'] = Box(cells['batch'].points, joined.group(2), cells['batch'].score)
        size, word = pack(get('pack'))
        ditto = is_ditto(anchor.text) and lines
        if ditto:  # "- do -": the medicine and pack of the line above
            size, word = size or lines[-1]['pack_size'], word or lines[-1]['pack_word']
        line = {
            'bill_name': lines[-1]['bill_name'] if ditto else ' '.join(bill_words(anchor.text)),
            'name_sure': lines[-1]['name_sure'] if ditto else anchor.score,
            'pack_size': size, 'pack_word': word,
            'batch': re.sub(r'\s+', '', get('batch')).upper(),
            'expiry': expiry(get('expiry')),
            'qty': money(get('qty')),
            'rate': money(get('rate')),
            'amount': money(get('amount')),
            'mrp': money(get('mrp')),
            'disc': percent(get('disc')) if 'disc' in cells else None,
            'free': money(get('free')) if 'free' in cells else None,
            'read': {k: get(k) for k in ('qty', 'rate', 'amount', 'expiry')},
            'notes': [],
        }
        n = line['notes']
        if not line['batch']:
            n.append('Batch not read.')
        elif sure('batch') < 0.95:
            n.append('Batch read unsure.')
        if not line['expiry']:
            n.append(f'Expiry not read ("{get("expiry")}").')
        line['expiry'] = fix_year(line['expiry'], n)
        lines.append(line)

    below = [b for r in rows[end:] for b in r]
    total = label_value(below, ('TOTAL', 'SUB TOTAL', 'GROSS', 'TOTAI', 'TOIAL', 'T0TAL', 'TO1AI', 'T01AI', 'TOTAT'), avoid=('NET', 'GRAND'))
    discount = label_value(below, ('DISC', 'COUNT', 'OISC', 'DIGC'), anywhere=True)  # LESS DIGCOUNT, LESS OISCOUNT
    rounding = label_value(below, ('ROUN', 'GOUN', 'NDING'), anywhere=True)  # ROUNOING, GOUNDING
    net = label_value(below, ('NET',), anywhere=True)

    # How well this reading agrees with itself before any repair: used to choose between the two readings.
    as_read = sum(1 for l in lines if None not in (l['qty'], l['rate'], l['amount'])
                  and (l['qty'] * l['rate']).quantize(CENT) == l['amount'])
    adds_up = total is not None and sum((l['amount'] or Decimal(0) for l in lines), Decimal(0)) == total
    nets = None not in (total, discount, net) and (total - discount + (rounding or 0)).quantize(CENT) == net
    missing = sum((not l['batch']) + (not l['expiry']) for l in lines)
    words = in_words(rows[end:])
    if words is not None and net is not None and words != net:
        checks.append(f'Net total read as {net}; the amount in words says {words}.')
    chosen, fitted, tail, changes = repair(lines, total, discount, rounding, net, words)
    # The better reading needs fewer corrections to make the bill add up; one that cannot be made to add up is worst.
    # A figure not read is a check not made, so it counts against the reading.
    unread = sum(v is None for v in (total, discount, net)) * 4
    score = -(changes if changes is not None else 50) - sum(c is None for c in chosen) * 10 - missing - unread \
        + (2 * as_read + 3 * adds_up + 3 * nets) / 100 - (0 if lines else 100)
    for l, c in zip(lines, chosen):
        if c is None:
            l['notes'].insert(0, 'Qty × Rate is not the Amount, and no likely misreading fits. Check qty, rate and amount.')
            continue
        _, q, r, a = c
        for key_, label, new in (('qty', 'Qty', q), ('rate', 'Rate', r), ('amount', 'Amount', a)):
            if l[key_] != new:
                l['notes'].insert(0, f'{label} read as "{l["read"][key_] or "nothing"}"; {new} fits the bill\'s sums.')
                l[key_] = new
    fit = None
    if fitted is not None:
        if total is not None and fitted != total:
            checks.append(f'Total read as {total}; {fitted} fits the lines and the net total.')
        total = fitted
        if tail:
            fit = tail
            for label, old, new in (('Less discount', discount, tail[0]), ('Rounding', rounding or Decimal(0), tail[1]), ('Net total', net, tail[2])):
                if old is None:
                    checks.append(f'{label} not read; {new} worked out from the total and the net total.')
                elif old != new and not (label == 'Net total' and new == words):
                    checks.append(f'{label} read as {old}; {new} fits the bill\'s sums.')
            discount, rounding, net = tail
        elif None not in (discount, net):
            checks.append(f'Total − discount + rounding = {total - discount + (rounding or 0)}, bill net total is {net}.')
    else:
        added = sum((l['amount'] or Decimal(0) for l in lines), Decimal(0))
        checks.append(f'Lines add up to {added}, bill total is {total}. A line may be missing or misread: check this bill by hand.')
        for l in lines:
            l['notes'].append('The lines do not add up to the bill total: compare this bill with the paper.')
    if total is None:
        checks.append('Bill total not read.')
    if changes is not None and changes >= 4:
        # Many digits changed: the sums agree, but a wrong set of changes could too. Every line is marked.
        checks.append(f'{changes} digits had to be corrected to make this bill add up: compare every line with the paper.')
        for l in lines:
            if 'compare with the paper' not in ' '.join(l['notes']):
                l['notes'].append('This bill needed several corrections: compare with the paper.')

    if discount is not None and lines:
        by_lines = sum(((l['amount'] or Decimal(0)) * (l['disc'] or Decimal(0)) / 100 for l in lines), Decimal(0)).quantize(CENT)
        if by_lines != discount and fit:
            # One distributor takes its 3% on the amount ÷ 1.075, so the printed Disc % never gives its discount.
            # The overall discount box on New purchase takes the bill's own amount, to the paisa.
            for l in lines:
                l['disc'] = None
            checks.append(f'TYPE ON NEW PURCHASE: overall discount {discount} (amount, not %), rounding {rounding or 0}. '
                          f'The lines\' Disc % would give {by_lines}, not the bill\'s {discount}, so it is left blank.')
        elif by_lines != discount:
            checks.append(f'Line discounts come to {by_lines}, bill says less discount {discount}. '
                          'The app will use the line Disc %; compare totals before saving.')
    elif discount is None and lines:
        checks.append('Less discount not read: compare the app\'s total with the bill before saving.')

    for l in lines:
        l['sell'] = None
        if l['mrp'] and l['rate']:
            l['sell'] = (l['mrp'] * INR_TO_NPR).quantize(CENT) if l['mrp'] < l['rate'] else l['mrp']
    if any(l['mrp'] and l['rate'] and l['mrp'] < l['rate'] for l in lines):
        checks.append('M.R.P. is below the rate, so it is taken as Indian rupees: Sell price = MRP × 1.6.')
    return lines, checks, {'Total': total, 'Less discount': discount, 'Rounding': rounding, 'Net total': net, '_score': score}


# ---------------------------------------------------------------- names.xlsx

def cell(v):
    return '' if v is None else str(v).strip()


def read_sheet(path, wanted):
    """Rows of the first sheet as dicts by header; {} for a missing file."""
    if not path or not Path(path).exists():
        return []
    ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [cell(h) for h in rows[0]]
    missing = [w for w in wanted if w not in header]
    if missing:
        raise SystemExit(f'{path}: no column {", ".join(missing)}')
    return [{h: cell(r[i]) if i < len(r) else '' for i, h in enumerate(header)} for r in rows[1:] if any(r)]


def match_names(seen, names_path, known):
    """
    names: bill name -> row of names.xlsx. Rows already in the file are kept as the user left them.
    A new bill name takes the medicine of a row (or a shop medicine) with the same name and strength,
    or one spelled a letter apart (marked to check); otherwise it is a medicine of its own.
    """
    names = {r['Bill name']: r for r in read_sheet(names_path, NAME_COLUMNS[:5])}
    for r in names.values():
        r['_new'] = False
        r.setdefault('Note', '')
    medicines = [(r['Name'], r['Strength']) for r in names.values()] + list(known)

    for bill_name, info in sorted(seen.items(), key=lambda kv: -len(kv[1]['bills'])):
        if bill_name in names:
            continue
        name, strength, form = split_name(bill_name)
        unit = unit_for(info['pack_word'], form, bill_name)
        strip = info['pack_size'] if unit in ('tablet', 'capsule') and (info['pack_size'] or 0) > 1 else ''
        note = []
        exact = [m for m in medicines if key(*m) == key(name, strength)]
        close = [m for m in medicines if same_medicine(m, (name, strength))]
        # Rows made on this run can still change; rows from names.xlsx or the shop's list cannot.
        open_rows = lambda m: [r for r in names.values() if r['_new'] and (r['Name'], r['Strength']) == m]
        if exact and bare(exact[0][1]) and not bare(strength) and open_rows(exact[0]):
            # This bill prints the unit ("10MG") another left out ("10"): the medicine takes it.
            for r in open_rows(exact[0]):
                r['Strength'] = strength
                r['Note'] = re.sub(r'Strength "[^"]*" has no unit on the bill; add mg/ml if you know it\. ?', '', r['Note']).strip()
            medicines[medicines.index(exact[0])] = (exact[0][0], strength)
            name = exact[0][0]
        elif exact:
            name, strength = exact[0]
        elif close:
            other = close[0]
            note.append(f'Read as "{medicine_text(name, strength)}"; taken as {medicine_text(*other)}, a letter apart. '
                        'If they are different medicines, put this one\'s own name here.')
            for r in open_rows(other):
                if 'Also read as' not in r['Note']:
                    r['Note'] = (r['Note'] + ' Also' if r['Note'] else 'Also') + \
                        f' read as "{medicine_text(name, strength)}" on another bill; check which spelling is right.'
            name, strength = other
        else:
            # A first word a letter away from another medicine's (WINDLAP / WINOLAP DS) is kept apart, as the
            # words differ, but both are marked: it is either a misreading or a near-namesake.
            brand = name.split()[0].upper() if name else ''
            for m in [m for m in medicines if m[0] and one_letter_apart(brand, m[0].split()[0].upper())][:1]:
                note.append(f'Looks like {medicine_text(*m)} (first word a letter apart); kept as its own medicine. Check the spelling.')
                for r in open_rows(m):
                    if 'Looks like' not in r['Note']:
                        r['Note'] = (r['Note'] + ' ' if r['Note'] else '') + \
                            f'Looks like {medicine_text(name, strength)} on another bill (first word a letter apart). Check the spelling.'
            medicines.append((name, strength))
            if strength and bare(strength):
                note.append(f'Strength "{strength}" has no unit on the bill; add mg/ml if you know it.')
            if info['name_sure'] < 0.9:
                note.append('Name read unsure; compare with the bill.')
        names[bill_name] = {'Bill name': bill_name, 'Name': name, 'Strength': strength, 'Unit': unit,
                            'Strip size': strip, 'Note': ' '.join(note), '_new': True}

    for bill_name, r in names.items():
        info = seen.get(bill_name)
        if info:
            r['Bills'] = len(info['bills'])
            r['Distributor'] = ', '.join(sorted(info['suppliers']))
    return names


def write_names(path, names):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Names'
    ws.append(NAME_COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in sorted(names.values(), key=lambda r: (r['Note'] == '', r['Name'].upper(), r['Bill name'])):
        ws.append([r.get(c) if r.get(c) not in ('', None) else None for c in NAME_COLUMNS])
        if r['Note']:
            for c in ws[ws.max_row]:
                c.fill = YELLOW
    for col, width in zip('ABCDEFGH', [30, 24, 10, 9, 9, 6, 26, 70]):
        ws.column_dimensions[col].width = width
    h = wb.create_sheet('How to use')
    for line in [
        'One row per medicine name as printed on the bills.',
        'Name and Strength are the medicine the bills will use. Rows that share a Name and Strength are one medicine.',
        'Yellow rows have a Note: check them against the bill, correct Name, Strength, Unit or Strip size, then delete the Note.',
        'Unit must be one of: ' + ', '.join(UNITS) + '.',
        'Strip size: units in one strip (blank if not sold in strips).',
        'Save, close, and run the reader again. Your changes are kept and the bill files and new-medicines.xlsx follow them.',
    ]:
        h.append([line])
    h.column_dimensions['A'].width = 120
    wb.save(path)


# ---------------------------------------------------------------- output files

def medicine_text(name, strength):
    return f'{name} {strength}'.strip()


def write_bill(path, info, lines, checks, totals, names):
    wb = Workbook()
    ws = wb.active
    ws.title = BILL_SHEET
    ws.append(BILL_COLUMNS + ['Bill name', 'Check'])
    for c in ws[1]:
        c.font = Font(bold=True)
    # Blank cells are left empty (None): PharmaSathi's reader stops on an empty text cell.
    num = lambda v: None if v is None else (int(v) if v == v.to_integral_value() else float(v))
    for l in lines:
        n = names[l['bill_name']]
        strip = n['Strip size']
        unit = 'strip' if strip else 'unit'
        notes = list(l['notes']) + (['Name: see names.xlsx.'] if n['Note'] else [])
        row = [medicine_text(n['Name'], n['Strength']), l['batch'], l['expiry'], num(l['qty']), unit,
               int(strip) if strip else None, num(l['rate']), num(l['sell']), num(l['disc']), num(l['free']),
               None, None, None, l['bill_name'], ' '.join(notes)]
        ws.append([v if v != '' else None for v in row])
        if notes:
            for c in ws[ws.max_row]:
                c.fill = YELLOW
    for col, width in zip('ABCDEFGHIJKLMNO', [26, 12, 9, 6, 6, 9, 9, 9, 7, 9, 11, 11, 6, 26, 60]):
        ws.column_dimensions[col].width = width
    d = wb.create_sheet('Bill details')
    for k, v in info.items():
        d.append([k, v])
    for k, v in totals.items():
        d.append([k, None if v is None else float(v)])
    d.append([None])
    for c in checks:
        d.append(['Check', c])
    d.column_dimensions['A'].width = 14
    d.column_dimensions['B'].width = 90
    wb.save(path)


def write_new_medicines(path, names, known, latest, usual=None):
    """One row per Name + Strength the shop does not have, priced from the latest bill it was on."""
    known_keys = {key(*m) for m in known}
    wb = Workbook()
    ws = wb.active
    ws.title = 'Medicines'
    ws.append(MEDICINE_COLUMNS + ['Check'])
    for c in ws[1]:
        c.font = Font(bold=True)
    done, count = set(), 0
    for r in sorted(names.values(), key=lambda r: (r['Name'].upper(), r['Strength'])):
        k = key(r['Name'], r['Strength'])
        if k in known_keys or k in done or not r['Name']:
            continue
        done.add(k)
        group = [x for x in names.values() if key(x['Name'], x['Strength']) == k]
        line = max((latest[x['Bill name']] for x in group if x['Bill name'] in latest), key=lambda l: l['date'], default=None)
        strip = next((x['Strip size'] for x in group if x['Strip size']), '')
        unit = next((x['Unit'] for x in group if x['Unit']), 'piece')
        rate = line and line['rate']
        sell = line and line['sell']
        notes = [x['Note'] for x in group if x['Note']]
        if not sell and rate and usual:
            sell = (rate / usual).quantize(CENT)
            notes.append(f'No M.R.P. read: selling price {sell} estimated from the usual margin. Check it.')
        elif not sell:
            notes.append('No MRP on the bill: fill in the selling price.')
        elif rate and sell < rate:
            notes.append(f'Selling price {sell} is below the cost {rate}: the rate or M.R.P. was misread on the bill. '
                         'Fix it here, or PharmaSathi will not add this medicine.')
        num = lambda v: None if not v else (int(v) if v == v.to_integral_value() else float(v))
        row = {'Name': r['Name'], 'Strength': r['Strength'] or None, 'Unit': unit, 'Active': 'yes',
               'Strip size': int(strip) if strip else None}
        if strip:
            row.update({'Strip cost': num(rate), 'Strip price': num(sell)})
        else:
            row.update({'Cost price': num(rate), 'Selling price': num(sell)})
        ws.append([row.get(c) for c in MEDICINE_COLUMNS] + [' '.join(notes) or None])
        if notes:
            for c in ws[ws.max_row]:
                c.fill = YELLOW
        count += 1
    for col, width in zip('ABCDEFGHIJKLMNOPQ', [8, 10, 24, 10, 10, 10, 14, 9, 9, 10, 10, 10, 12, 10, 10, 7, 60]):
        ws.column_dimensions[col].width = width
    wb.save(path)
    return count


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('source', type=Path, help='a folder of bill photos, or one photo')
    ap.add_argument('--out', type=Path)
    ap.add_argument('--medicines', type=Path, help="a Medicines export from PharmaSathi: the shop's own medicines")
    ap.add_argument('--debug', action='store_true')
    a = ap.parse_args()
    photos = sorted(p for p in a.source.iterdir() if p.suffix.lower() in ('.jpg', '.jpeg', '.png', '.webp')) \
        if a.source.is_dir() else [a.source]
    out = a.out or (a.source if a.source.is_dir() else a.source.parent) / 'bills-read'
    (out / 'bills').mkdir(parents=True, exist_ok=True)
    (out / '.cache').mkdir(exist_ok=True)
    known = [(r['Name'], r['Strength']) for r in read_sheet(a.medicines, ['Name', 'Strength'])]

    engines, bills, summary, contents, scores = {}, [], [], {}, {}
    for photo in photos:
        try:
            parsed = []
            for name, rows in read_photo(photo, out / '.cache', engines).items():
                try:
                    parsed.append((name, rows) + parse(rows))
                except StopIteration:
                    continue
            if not parsed:
                raise ValueError('could not find the table header (QTY / RATE / AMOUNT) on the bill')
            # The reading whose numbers agree with each other best; on a tie the medium one, which reads names better.
            name, rows, lines, checks, totals = max(parsed, key=lambda p: (p[4]['_score'], p[0] != 'small'))
            score = totals.pop('_score')
            if a.debug:
                print('    reading:', name)
                for r in rows:
                    print('   ', ' | '.join(b.text for b in r))
            info = bill_info(rows)
            # The same bill photographed twice: the same number, or (number covered) the same lines and total.
            # Content alone is not enough: a repeat order can have the same lines and total, and a credit note
            # returns exactly the lines of a bill. So the date and the kind of document must match too.
            content = (info.get('Kind'), info.get('Bill date'), totals.get('Total'),
                       tuple(sorted((str(l['qty']), str(l['amount'])) for l in lines)))
            twin = next((b for b in bills if b[1].get('Kind') == info.get('Kind') and (
                (info.get('Bill no.') and b[1].get('Bill no.') == info.get('Bill no.'))
                or (lines and content[1] and content[2] is not None and contents[b[0].name] == content))), None)
            entry = (photo, info, lines, checks, totals)
            contents[photo.name], scores[photo.name] = content, score
            if twin and scores[twin[0].name] >= score:
                summary.append(f"SAME  {photo.name}: same bill as {twin[0].name} ({twin[1].get('Bill no.', 'same lines and total')}); "
                               'the clearer photo is used.')
                print(summary[-1], flush=True)
                continue
            if twin:
                bills[bills.index(twin)] = entry
                summary.append(f"SAME  {twin[0].name}: same bill as {photo.name} ({info.get('Bill no.') or twin[1].get('Bill no.', '')}); "
                               'the clearer photo is used.')
                print(summary[-1], flush=True)
                continue
            bills.append(entry)
        except Exception as e:  # one bad photo must not stop the rest
            summary.append(f'FAIL  {photo.name}: {e}')
            print(summary[-1], flush=True)

    # A supplier is known by its PAN (digits read reliably); its name is the spelling read most often.
    pans = Counter(info['PAN'] for _, info, _, _, _ in bills if info.get('PAN'))
    for _, info, _, _, _ in bills:  # a PAN a digit away from one read more often is that one, misread
        if info.get('PAN'):
            info['PAN'] = next((p for p, _ in pans.most_common() if sum(a != b for a, b in zip(p, info['PAN'])) <= 1), info['PAN'])
    spellings = {}
    for _, info, _, _, _ in bills:
        spellings.setdefault(info.get('PAN') or info.get('Supplier', ''), Counter())[info.get('Supplier', '')] += 1
    for _, info, _, _, _ in bills:
        name = spellings[info.get('PAN') or info.get('Supplier', '')].most_common(1)[0][0]
        info['Supplier'] = f"{name} (PAN {info['PAN']})" if info.get('PAN') else name

    # Each distributor prices at a steady share of the selling price (one tested distributor: rate ≈ 86% of MRP × 1.6).
    # A selling price far from that share, or below cost, means the M.R.P. was misread: the reading of it
    # that gives the usual share is taken, and marked.
    shares = sorted(l['rate'] / l['sell'] for _, _, lines, _, _ in bills for l in lines
                    if l['sell'] and l['rate'] and Decimal('0.5') < l['rate'] / l['sell'] < 1)
    usual = shares[len(shares) // 2] if len(shares) >= 5 else None
    for _, _, lines, _, _ in bills:
        for l in lines:
            if not (l['sell'] and l['rate'] and l['mrp']):
                continue
            share = l['rate'] / l['sell']
            if usual is None or abs(share - usual) <= usual * Decimal('0.03'):
                if share >= 1:
                    l['notes'].append(f'Sell price {l["sell"]} is below the rate: check the M.R.P.')
                continue
            fits = []
            for mv, mc in variants(l['mrp'], 2).items():
                sell = (mv * INR_TO_NPR).quantize(CENT) if mv < l['rate'] else mv
                if sell > 0:
                    fits.append((abs(l['rate'] / sell - usual), mc, mv, sell))
            gap, _, mv, sell = min(fits)
            if gap <= usual * Decimal('0.01'):
                l['notes'].append(f'M.R.P. read as {l["mrp"]}; {mv} gives this distributor\'s usual margin. Check it.')
                l['mrp'], l['sell'] = mv, sell
            else:  # no likely misreading fits: estimate from the usual margin, as for a missing M.R.P.
                est = (l['rate'] / usual).quantize(CENT)
                l['notes'].append(f'Sell price {l["sell"]} looks wrong for a rate of {l["rate"]} (M.R.P. misread); '
                                  f'{est} estimated from the usual margin. Check it.')
                l['sell'] = est

    # Batch numbers have no sums, but a distributor's batches share a few shapes (GTH1243C: 3 letters,
    # 4 digits, a letter). A batch one look-alike character away from a common shape (6TH0416D) is that.
    shape = lambda b: ''.join('L' if c.isalpha() else 'D' if c.isdigit() else c for c in b)
    shapes = Counter(shape(l['batch']) for _, _, lines, _, _ in bills for l in lines if l['batch'])
    common = {k for k, n in shapes.items() if n >= 3 and n >= sum(shapes.values()) * 0.1}
    alike = {'6': 'G', 'G': '6', '0': 'O', 'O': '0', '8': 'B', 'B': '8', '5': 'S', 'S': '5', '1': 'I', 'I': '1', '2': 'Z', 'Z': '2'}
    for _, _, lines, _, _ in bills:
        for l in lines:
            b = l['batch']
            if not b or shape(b) in common or sum(shapes.values()) < 10:  # too few batches to know the usual shapes
                continue
            fixes = {b[:i] + alike[c] + b[i + 1:] for i, c in enumerate(b) if c in alike}
            fixes = sorted(f for f in fixes if shape(f) in common)
            if len(fixes) == 1 and re.search(r'[A-Z]', b):  # an all-digit batch (11310019) is left as read
                l['notes'].append(f'Batch read as {b}; {fixes[0]} has the usual pattern. Check it.')
                l['batch'] = fixes[0]
            elif 'Batch' not in ' '.join(l['notes']):
                l['notes'].append(f'Batch {b} has an unusual pattern: check it.')

    seen, latest = {}, {}
    for photo, info, lines, _, _ in bills:
        for l in lines:
            s = seen.setdefault(l['bill_name'], {'bills': set(), 'suppliers': set(), 'pack_size': None, 'pack_word': '', 'name_sure': 1})
            s['bills'].add(photo.name)
            s['suppliers'].add(info.get('Supplier', ''))
            s['pack_size'] = s['pack_size'] or l['pack_size']
            s['pack_word'] = s['pack_word'] or l['pack_word']
            s['name_sure'] = min(s['name_sure'], l['name_sure'])
            l['date'] = info.get('Bill date', '')
            if l['bill_name'] not in latest or l['date'] >= latest[l['bill_name']]['date']:
                latest[l['bill_name']] = l

    names = match_names(seen, out / 'names.xlsx', known)
    write_names(out / 'names.xlsx', names)
    added = write_new_medicines(out / 'new-medicines.xlsx', names, known, latest, usual)

    for photo, info, lines, checks, totals in bills:
        name = re.sub(r'[^\w\-]+', '-', f"{photo.stem}-{info.get('Bill no.', '')}").strip('-')
        if info.get('Kind'):  # a credit note: goods returned to the supplier, never imported as a purchase
            (out / 'returns').mkdir(exist_ok=True)
            checks.insert(0, 'This is a CREDIT NOTE (goods returned to the supplier), not a purchase. Do not import it '
                             'as a purchase: enter it in PharmaSathi as a purchase return.')
            write_bill(out / 'returns' / f'{name}.xlsx', info, lines, checks, totals, names)
            summary.append(f'RETURN {photo.name}: credit note {info.get("Bill no.", "")}, {len(lines)} lines -> returns/{name}.xlsx')
            print(summary[-1], flush=True)
            continue
        write_bill(out / 'bills' / f'{name}.xlsx', info, lines, checks, totals, names)
        marked = sum(1 for l in lines if l['notes'] or names[l['bill_name']]['Note'])
        bad = any('add up' in c or 'not read' in c for c in checks)
        summary.append(f"{'CHECK' if marked or bad else 'OK':5} {photo.name}: {len(lines)} lines, {marked} marked, "
                       f"{len(checks)} bill notes -> bills/{name}.xlsx")
        print(summary[-1], flush=True)

    to_check = sum(1 for r in names.values() if r['Note'])
    tail = [f'{len(names)} names, {to_check} to check in names.xlsx; {added} new medicines in new-medicines.xlsx.']
    print(tail[0])
    (out / 'summary.txt').write_text('\n'.join(summary + [''] + tail) + '\n')


if __name__ == '__main__':
    main()
