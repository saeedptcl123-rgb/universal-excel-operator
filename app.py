import io

import openpyxl
import streamlit as st


# ==========================================
# 1. OMS اٹینڈنس اور ٹوٹل موڈ (موڈ 1)
# ==========================================
def process_oms_attendance_mode(raw_lines, template_file):
    try:
        wb = openpyxl.load_workbook(template_file)
        sheet = wb.active
    except Exception as e:
        return None, f"خامی: ٹیمپلیٹ فائل پڑھنے میں مسئلہ آیا: {str(e)}"

    records = []
    for line in raw_lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue

        lower_line = line.lower()
        if any(w in lower_line for w in ['please', 'thanks', 'inform', 'unable', 'today', 'late', 'issue', 'team']):
            if 'present' not in lower_line and 'leave' not in lower_line and 'fever' not in lower_line:
                continue

        try:
            if '-' in line and ':' in line:
                parts = line.split('-', 1)
                date_str = parts[0].strip().split(',')[0].strip()
                rest = parts[1].strip()

                try:
                    day_num = int(date_str.split('/')[0])
                    if not (1 <= day_num <= 31):
                        continue
                except ValueError:
                    continue

                if ':' in rest:
                    name_status = rest.split(':', 1)
                    raw_name = name_status[0].strip()
                    msg = name_status[1].strip().lower()
                else:
                    continue

                if len(raw_name) > 30 or 'please' in raw_name.lower():
                    continue

                name = raw_name

                if any(w in msg for w in ['fever', 'beemar', 'sick', 'leave', 'absent', 'unable']):
                    status = 'L'
                else:
                    status = 'P'

                records.append({'day': day_num, 'name': name, 'status': status})
        except Exception:
            continue

    if not records:
        return None, "انتباہ: واٹس ایپ اٹینڈنس ڈیٹا سے کوئی ریکارڈ نہیں ملا۔"

    start_row = 5
    sr_col = 1
    name_col = 2
    date_start_col = 4

    unique_names = sorted(list(set(r['name'] for r in records)))

    current_row = start_row
    for idx, emp_name in enumerate(unique_names, start=1):
        try:
            sheet.cell(row=current_row, column=sr_col, value=idx)
            sheet.cell(row=current_row, column=name_col, value=emp_name)
        except Exception:
            pass

        for rec in records:
            if rec['name'] == emp_name:
                day_val = rec['day']
                target_col = date_start_col + (day_val - 1)
                try:
                    sheet.cell(row=current_row, column=target_col, value=rec['status'])
                except Exception:
                    pass

        try:
            r_str = str(current_row)
            sheet.cell(row=current_row, column=34, value=f"=COUNTIF(D{r_str}:AG{r_str}, \"P\")")
            sheet.cell(row=current_row, column=35, value=f"=COUNTIF(D{r_str}:AG{r_str}, \"H\")")
            sheet.cell(row=current_row, column=36, value=f"=COUNTIF(D{r_str}:AG{r_str}, \"L\")")
            sheet.cell(row=current_row, column=37, value=f"=SUM(AH{r_str}:AJ{r_str})")
        except Exception:
            pass

        current_row += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, "کامیابی!"


# ==========================================
# 2. یونیورسل ڈیٹا پروسیسنگ (موڈ 2)
# ==========================================
import io
import re
import csv
from collections import Counter
from copy import copy
from itertools import permutations

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

REVIEW_FILL = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
NUM_RE = re.compile(r'^[-+]?\d[\d,]*(\.\d+)?$')
REAL_COMMA_RE = re.compile(r',(?!\d{3}(\D|$))')
FORMULA = 'FORMULA'


# ---------- بنیادی مددگار ----------
def _is_num(tok):
    return isinstance(tok, str) and bool(NUM_RE.match(tok.strip()))


def _to_num(tok):
    v = float(tok.replace(',', '').strip())
    return int(v) if v.is_integer() else v


def _header_kind(header):
    h = (header or '').strip().lower()
    if 'total' in h:
        return 'total'
    if re.search(r'\b(price|rate|cost|salary|pay|wage|amount|value|mrp|fee)\b', h):
        return 'price'
    if re.search(r'\b(qty|quantity|count|units?|pcs|nos|pieces)\b', h):
        return 'qty'
    if re.search(r'(^|\s)(sr|s\.?no|sno|serial|no)\b|#', h):
        return 'index'
    if re.search(r'\b(product|item|name|description|domain|category|region|city|'
                 r'status|country|type|dept|department|code|vendor|customer|title|location)\b', h):
        return 'text'
    return 'other'


# ---------- ڈیٹا کی ساخت ----------
def _sniff_mode(lines):
    votes = Counter()
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if '\t' in line:
            votes['tab'] += 1
        elif REAL_COMMA_RE.search(s):
            votes['comma'] += 1
        elif re.search(r'\s{2,}', s):
            votes['space'] += 1
        else:
            votes['single'] += 1
    return votes.most_common(1)[0][0] if votes else 'single'


def _split_tokens(text):
    cells, buf = [], []
    for tok in text.split():
        if _is_num(tok):
            if buf:
                cells.append(' '.join(buf))
                buf = []
            cells.append(tok)
        else:
            buf.append(tok)
    if buf:
        cells.append(' '.join(buf))
    return cells


def _parse_line(line, mode):
    s = line.strip()
    if mode == 'tab':
        cells = [c.strip() for c in line.rstrip('\r\n').split('\t')]
        while cells and not cells[-1]:
            cells.pop()
        return cells
    if mode == 'comma':
        return [c.strip() for c in next(csv.reader([s])) if c.strip()]
    if mode == 'space':
        cells = []
        for part in re.split(r'\s{2,}', s):
            cells.extend(_split_tokens(part))
        return [c for c in cells if c]
    return [c for c in _split_tokens(s) if c]


def _read_rows(raw_lines):
    mode = _sniff_mode(raw_lines)
    rows = []
    for line in raw_lines:
        s = line.strip()
        if not s or s.startswith('#'):
            continue
        cells = _parse_line(line, mode)
        if not any(c.strip() for c in cells):
            continue
        if any(not _is_num(c) and 'total' in c.lower() for c in cells):
            continue
        rows.append(cells)
    return mode, rows


def _frequent_last_words(rows):
    cnt = Counter()
    for r in rows:
        for c in r:
            if not _is_num(c):
                cnt[c.split()[-1].lower()] += 1
    return {w for w, n in cnt.items() if n >= 2}


def _messy_row(cells, frequent):
    texts, nums = [], []
    for c in cells:
        if _is_num(c):
            nums.append(_to_num(c))
            continue
        words = c.split()
        if len(words) >= 2 and words[-1].lower() in frequent:
            texts.extend([' '.join(words[:-1]), words[-1]])
        else:
            texts.append(c)
    return texts, nums


def _order_numbers(nums):
    """اگر 3 نمبرز ہوں تو Qty×Price=Total والی ترتیب لگاتا ہے؛ ورنہ جیسے ہیں ویسے۔"""
    if len(nums) == 3:
        for q, p, t in permutations(nums):
            if abs(q * p - t) < 0.01:
                return [q, p, t]
    return nums


# ---------- ٹیمپلیٹ ----------
def _read_template(sheet):
    header_row = 1
    for r in range(1, min(sheet.max_row, 15) + 1):
        vals = [sheet.cell(r, c).value for c in range(1, sheet.max_column + 1)]
        if sum(1 for v in vals if isinstance(v, str) and v.strip()) >= 2:
            header_row = r
            break
    cols = []
    for c in range(1, sheet.max_column + 1):
        v = sheet.cell(header_row, c).value
        header = str(v).strip() if v is not None else ''
        if header:
            cols.append({'col': c, 'header': header, 'kind': _header_kind(header)})
    total_row = None
    for r in range(header_row + 1, sheet.max_row + 1):
        for c in range(1, sheet.max_column + 1):
            v = sheet.cell(r, c).value
            if isinstance(v, str) and 'total' in v.lower():
                total_row = r
                break
        if total_row:
            break
    return header_row, cols, total_row


def _is_header_row(cells, tcols):
    headers = {c['header'].lower() for c in tcols}
    if any(_is_num(c) for c in cells):
        return False
    return any(c.lower() in headers for c in cells)


# ---------- UI کے لیے: قطاریں اور نقشہ ----------
def prepare_rows(raw_lines):
    """
    Returns: (mode, rows, text_count)
    rows: ہر قطار کے خانے (ٹیب والی فائل میں اصل کالم، سپیس والی میں: متنی خانے پھر عددی خانے)
    """
    mode, rows = _read_rows(raw_lines)
    if mode in ('tab', 'comma'):
        width = max(len(r) for r in rows) if rows else 0
        return mode, [r + [''] * (width - len(r)) for r in rows], None

    frequent = _frequent_last_words(rows)
    out, text_count = [], 0
    for r in rows:
        texts, nums = _messy_row(r, frequent)
        text_count = max(text_count, len(texts))
        out.append(texts + _order_numbers(nums))
    return mode, out, text_count


def slot_labels(rows):
    width = max((len(r) for r in rows), default=0)
    labels = []
    for i in range(width):
        example = next((str(r[i]) for r in rows if len(r) > i and str(r[i]).strip()), '')
        labels.append(f"کالم {i + 1} (مثال: {example})")
    return labels


def default_mapping(tcols, width, mode, text_count):
    """ٹیمپلیٹ کالم → ڈیٹا خانہ (index)، None، یا FORMULA۔ یہ صرف ابتدائی تجویز ہے، آپ بدل سکتے ہیں۔"""
    mapping = {}
    if mode in ('tab', 'comma'):
        for k, tc in enumerate(tcols):
            if tc['kind'] == 'total':
                mapping[tc['col']] = FORMULA
            else:
                mapping[tc['col']] = k if k < width else None
        return mapping

    t_iter = iter(range(text_count))
    for tc in tcols:
        if tc['kind'] == 'text':
            mapping[tc['col']] = next(t_iter, None)
        elif tc['kind'] == 'qty':
            mapping[tc['col']] = text_count if width > text_count else None
        elif tc['kind'] == 'price':
            mapping[tc['col']] = text_count + 1 if width > text_count + 1 else None
        elif tc['kind'] == 'total':
            mapping[tc['col']] = FORMULA
        else:
            mapping[tc['col']] = None
    return mapping


def build_items(rows, tcols, mapping):
    items = []
    for row in rows:
        slots, review = {}, []
        for tc in tcols:
            m = mapping.get(tc['col'])
            if m is None or m == FORMULA:
                continue
            if m >= len(row):
                review.append(f"{tc['header']} کا خانہ نہیں ملا")
                continue
            v = row[m]
            if v in ('', None):
                continue
            slots[tc['col']] = _to_num(v) if _is_num(v) else v
        items.append({'slots': slots, 'review': review})
    return items


# ---------- لکھنا ----------
def _shift_formula(formula, old_row, new_row):
    return re.sub(
        r'(\$?[A-Z]{1,3})(\$?)(\d+)',
        lambda m: m.group(1) + m.group(2) + (str(new_row) if int(m.group(3)) == old_row else m.group(3)),
        formula,
    )


def _write(sheet, tcols, header_row, total_row, items, need_total, formula_cols):
    data_start = header_row + 1
    n = len(items)
    max_col = sheet.max_column
    L = get_column_letter

    pattern = {}
    if total_row != data_start:
        for c in range(1, max_col + 1):
            v = sheet.cell(data_start, c).value
            if isinstance(v, str) and v.startswith('='):
                pattern[c] = v

    old_total = None
    if total_row:
        old_total = {c: sheet.cell(total_row, c).value for c in range(1, max_col + 1)}
        for c in range(1, max_col + 1):
            sheet.cell(total_row, c).value = None

    qty_col = next((c['col'] for c in tcols if c['kind'] == 'qty'), None)
    price_col = next((c['col'] for c in tcols if c['kind'] == 'price'), None)

    review_col = max_col + 1
    sheet.cell(header_row, review_col).value = "چیک نوٹ"

    for k, item in enumerate(items):
        r = data_start + k
        slots = item['slots']
        for col, val in slots.items():
            sheet.cell(r, col).value = val
        if qty_col and price_col and qty_col in slots and price_col in slots:
            for fc in formula_cols:
                sheet.cell(r, fc).value = f"={L(qty_col)}{r}*{L(price_col)}{r}"
        for col, f in pattern.items():
            if sheet.cell(r, col).value is None:
                sheet.cell(r, col).value = _shift_formula(f, data_start, r)
        if item['review']:
            for c in range(1, max_col + 1):
                sheet.cell(r, c).fill = REVIEW_FILL
            sheet.cell(r, review_col).value = "چیک کریں: " + "، ".join(item['review'])

    if old_total is not None or need_total:
        t = data_start + n
        end = t - 1
        if old_total is not None and total_row:
            for c in range(1, max_col + 1):
                sheet.cell(t, c)._style = copy(sheet.cell(total_row, c)._style)
        for c in range(1, max_col + 1):
            old = old_total.get(c) if old_total else None
            if isinstance(old, str) and old.startswith('=') and 'SUM' in old.upper():
                sheet.cell(t, c).value = f"=SUM({L(c)}{data_start}:{L(c)}{end})"
            elif old is not None and not (isinstance(old, str) and old.startswith('=')):
                sheet.cell(t, c).value = old
        if old_total is None:
            sheet.cell(t, tcols[0]['col']).value = "Grand Total"
        for c in tcols:
            if c['kind'] in ('qty', 'total') and sheet.cell(t, c['col']).value is None:
                sheet.cell(t, c['col']).value = f"=SUM({L(c['col'])}{data_start}:{L(c['col'])}{end})"

    output = io.BytesIO()
    sheet.parent.save(output)
    output.seek(0)
    return output


# ---------- مین فنکشن ----------
def process_with_mapping(template_bytes, rows, mapping, need_total=False):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
        sheet = wb.active
    except Exception as e:
        return None, f"خامی: ٹیمپلیٹ فائل پڑھنے میں مسئلہ آیا: {str(e)}"

    header_row, tcols, total_row = _read_template(sheet)
    if not tcols:
        return None, "خامی: ٹیمپلیٹ میں ہیڈرز نہیں ملے۔"

    # ہیڈر کے نیچے والے مرجڈ سیلز میں لکھا نہیں جا سکتا، اس لیے انہیں کھول دیں
    for rng in list(sheet.merged_cells.ranges):
        if rng.max_row > header_row:
            sheet.unmerge_cells(str(rng))

    if rows and _is_header_row([str(x) for x in rows[0]], tcols):
        rows = rows[1:]
    if not rows:
        return None, "انتباہ: ڈیٹا میں کوئی قطار نہیں ملی۔"

    items = build_items(rows, tcols, mapping)
    formula_cols = [c for c, m in mapping.items() if m == FORMULA]
    output = _write(sheet, tcols, header_row, total_row, items, need_total, formula_cols)

    review_count = sum(1 for it in items if it['review'])
    summary = f"کامیابی! {len(items)} قطاریں پروسیس ہوئیں۔"
    if review_count:
        summary += f" {review_count} قطاروں کو چیک کرنا ضروری ہے (پیلی ہائی لائٹ)۔"
    return output, summary


def load_template_info(template_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    return _read_template(wb.active)[1]
import datetime

from openpyxl.styles import Alignment, Font


# ==========================================
# 3. ایکسل ورک بینچ (موڈ 3)
# ==========================================
EXCEL_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
EXCEL_HEADER_FONT = Font(bold=True, color="FFFFFF")
EXCEL_NUM_FMT = '#,##0.00'


def _names_header(cells, headers):
    """کیا یہ قطار کالم کے ناموں والی ہیڈر لائن ہے؟"""
    return bool(cells) and not any(_is_num(c) for c in cells) and any(c in headers for c in cells)


def _cell_text(v):
    return '' if v is None else str(v).strip()


def _coerce(v):
    """نمبر کو نمبر ہی رکھتا ہے (ٹیکسٹ نہیں بنتا)۔ صفر سے شروع ہونے والے کوڈ (مثلاً 00123) ٹیکسٹ رہتے ہیں۔"""
    if v is None:
        return None
    if isinstance(v, (int, float, datetime.date)):
        return v
    s = str(v).strip()
    if s == '':
        return None
    if _is_num(s) and not (s.startswith('0') and len(s) > 1 and not s.startswith('0.')):
        return _to_num(s)
    return s


def _style_header(cell):
    cell.font = EXCEL_HEADER_FONT
    cell.fill = EXCEL_HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _split_header(rows, n_cols_hint=None):
    """پہلی قطار ہیڈر ہے یا نہیں، یہ دیکھ کر (headers, data_rows) دیتا ہے۔"""
    if rows and not any(_is_num(c) for c in rows[0]):
        headers = [c if c else f"کالم {i + 1}" for i, c in enumerate(rows[0])]
        return headers, rows[1:]
    width = max((len(r) for r in rows), default=0)
    return [f"کالم {i + 1}" for i in range(width)], rows


def create_workbook(headers, data_rows, sheet_name="Data"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = (sheet_name or "Data")[:31]
    for c, h in enumerate(headers, 1):
        _style_header(ws.cell(1, c, h))
    for r, row in enumerate(data_rows, 2):
        for c in range(1, len(headers) + 1):
            v = row[c - 1] if c - 1 < len(row) else None
            cell = ws.cell(r, c, _coerce(v))
            if isinstance(cell.value, float):
                cell.number_format = EXCEL_NUM_FMT
            elif isinstance(cell.value, int):
                cell.number_format = '#,##0'
    for c, h in enumerate(headers, 1):
        longest = max([len(str(h))] + [len(str(r[c - 1])) for r in data_rows if len(r) >= c])
        ws.column_dimensions[get_column_letter(c)].width = max(12, min(40, longest + 4))
    ws.freeze_panes = 'A2'
    if headers and data_rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(data_rows) + 1}"
    return wb


def _sheet_headers(ws):
    return [_cell_text(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]


def _table(ws):
    """(headers, list of dict rows) — خالی قطاریں چھوڑ دیتا ہے۔"""
    headers = _sheet_headers(ws)
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
        if any(v not in (None, '') for v in vals):
            rows.append(dict(zip(headers, vals)))
    return headers, rows


def _last_data_row(ws):
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(r, c).value not in (None, '') for c in range(1, ws.max_column + 1)):
            return r
    return 1


def append_rows(ws, dict_rows):
    """
    نئی قطاریں آخر میں جوڑتا ہے۔ ہر قطار کا رنگ/فارمیٹ آخری ڈیٹا قطار سے لیتا ہے،
    تاکہ موجودہ فارمیٹ خراب نہ ہو۔
    """
<<<<<<< HEAD
    from openpyxl.formula.translate import Translator
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    headers = _sheet_headers(ws)
    last = _last_data_row(ws)
    for i, d in enumerate(dict_rows):
        r = last + 1 + i
        for c, name in enumerate(headers, 1):
            cell = ws.cell(r, c)
<<<<<<< HEAD
            val = d.get(name)
            if val in (None, '') and last >= 2:
                srcv = ws.cell(last, c).value
                if isinstance(srcv, str) and srcv.startswith('='):
                    # خالی خانہ: آخری قطار کا فارمولا نئی قطار کے مطابق نقل ہوتا ہے
                    val = Translator(srcv, origin=ws.cell(last, c).coordinate).translate_formula(cell.coordinate)
            cell.value = _coerce(val)
=======
            cell.value = _coerce(d.get(name))
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            if last >= 2:
                cell._style = copy(ws.cell(last, c)._style)
            if isinstance(cell.value, float):
                cell.number_format = EXCEL_NUM_FMT
    return len(dict_rows)


def delete_rows_where(ws, column_name, target):
    headers = _sheet_headers(ws)
    if column_name not in headers:
        return 0
    col = headers.index(column_name) + 1
    t = _cell_text(target).lower()
    to_delete = [r for r in range(2, ws.max_row + 1)
                 if _cell_text(ws.cell(r, col).value).lower() == t]
    for r in reversed(to_delete):
        ws.delete_rows(r)
    return len(to_delete)


def _norm(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and _is_num(v):
        return float(_to_num(v))
    return _cell_text(v).lower()


def _same(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) < 0.01
    return a == b


def reconcile_rows(headers_a, rows_a, headers_b, rows_b, key):
    A = {_cell_text(r.get(key)): r for r in rows_a}
    B = {_cell_text(r.get(key)): r for r in rows_b}
    common = [h for h in headers_a if h in headers_b and h and h != key]
    out = []
    for k in dict.fromkeys(list(A) + list(B)):
        diffs = []
        if k in A and k in B:
            for h in common:
                a, b = _norm(A[k].get(h)), _norm(B[k].get(h))
                if not _same(a, b):
                    diffs.append(f"{h}: {A[k].get(h)} | {B[k].get(h)}")
            status = "Matched" if not diffs else "Mismatch"
        elif k in A:
            status = "Only in File 1"
        else:
            status = "Only in File 2"
        out.append([k, status, "; ".join(diffs)])
    return out


def add_consolidated_sheet(wb, ws_src, group_name, sum_names):
    """
    موجودہ ورک بک میں نئی شیٹ "Consolidated" جوڑتا ہے، جس میں SUMIFS فارمولے ہوتے ہیں
    (یعنی ڈیٹا بدلنے پر رپورٹ خود اپڈیٹ ہو جائے)۔
    """
    headers = _sheet_headers(ws_src)
    gi = headers.index(group_name) + 1
    gl = get_column_letter(gi)
    groups = list(dict.fromkeys(
        _cell_text(ws_src.cell(r, gi).value) for r in range(2, ws_src.max_row + 1)
        if _cell_text(ws_src.cell(r, gi).value)
    ))
    ws = wb.create_sheet("Consolidated")
    head = [group_name] + sum_names
    for c, h in enumerate(head, 1):
        _style_header(ws.cell(1, c, h))
    src = f"'{ws_src.title}'"
    for i, g in enumerate(groups, 2):
        ws.cell(i, 1, g)
        for j, sn in enumerate(sum_names, 2):
            sl = get_column_letter(headers.index(sn) + 1)
            cell = ws.cell(i, j, f"=SUMIFS({src}!${sl}:${sl},{src}!${gl}:${gl},$A{i})")
            cell.number_format = EXCEL_NUM_FMT
    if groups:
        tr = len(groups) + 2
        ws.cell(tr, 1, "Total").font = Font(bold=True)
        for j in range(2, len(head) + 1):
            L = get_column_letter(j)
            cell = ws.cell(tr, j, f"=SUM({L}2:{L}{tr - 1})")
            cell.font = Font(bold=True)
            cell.number_format = EXCEL_NUM_FMT
    ws.column_dimensions['A'].width = 24
    for j in range(2, len(head) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    ws.freeze_panes = 'A2'
    return ws


def save_bytes(wb):
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ==========================================
# 4. ایکسل ایجنٹ: کمانڈز پر مبنی آپریشنز (موڈ 3)
# ==========================================
AGENT_HELP = """کمانڈز (ہر لائن ایک کمانڈ):
insert row at N [: Col=Val, Col2=Val2]     ← N نمبر پر خالی/بھری قطار ڈالیں
add row [: Col=Val, ...]                    ← آخر میں قطار جوڑیں
insert column NAME at N                     ← N نمبر پر نیا کالم
add column NAME                             ← آخر میں نیا کالم
delete where COL OP VAL                     ← قطاریں حذف (OP: = != > < >= <= contains)
filter COL OP VAL                           ← فلٹر شدہ نئی شیٹ "Filter"
sort COL asc|desc                           ← ڈیٹا ترتیب دیں
fill COL by KEY                             ← دوسری فائل سے خالی خانے بھریں (KEY کے ذریعے)
reconcile KEY                               ← دوسری فائل سے موازنہ، شیٹ "Reconciliation"
consolidate GROUP sum COL1,COL2             ← شیٹ "Consolidated" (SUMIFS فارمولے)
merge                                       ← اپلوڈ کی گئی باقی فائلوں کی قطاریں جوڑیں
rename column OLD to NEW                    ← کالم کا نام بدلیں
<<<<<<< HEAD
format RANGE [fill=HEX] [font_color=HEX] [bold] [italic] [font=NAME] [size=N] [align=left|center|right] [number=FMT] [border=thin|medium|thick|none]
                                           ← RANGE: A1:E1 | row 3 | col C | header (پہلی قطار)
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
help                                        ← یہی مدد"""

_OP_TOKENS = r'(=|!=|>=|<=|>|<|contains)'


def _assign_pairs(text):
    out = {}
    if not text:
        return out
    for part in text.split(','):
        if '=' in part:
            k, v = part.split('=', 1)
            out[k.strip()] = _coerce(v.strip())
    return out


def _col_of(ws, name):
    headers = _sheet_headers(ws)
    for i, h in enumerate(headers, 1):
        if h.lower() == name.strip().lower():
            return i
    raise ValueError(f"کالم '{name}' نہیں ملا۔")


def _match(a, op, b):
    if op == 'contains':
        return _cell_text(b).lower() in _cell_text(a).lower()
    na, nb = _norm(a), _norm(b)
    if isinstance(na, float) and isinstance(nb, float):
        return {'=': abs(na - nb) < 0.01, '!=': abs(na - nb) >= 0.01, '>': na > nb,
                '<': na < nb, '>=': na >= nb, '<=': na <= nb}[op]
    if op == '=':
        return na == nb
    if op == '!=':
        return na != nb
    return False


def _copy_row_style(ws, src_row, dst_row, ncol):
    for c in range(1, ncol + 1):
        ws.cell(dst_row, c)._style = copy(ws.cell(src_row, c)._style)


<<<<<<< HEAD
_REF_RE = re.compile(r"(?<![A-Za-z0-9_\"'!])(\$?)([A-Z]{1,3})(\$?)([0-9]+)(?![0-9A-Za-z_(])")


def _remap_formulas(ws, row_map=None, col_map=None):
    """
    شیٹ کے تمام فارمولوں میں خانوں کے حوالہ جات نئی جگہ کے مطابق بدلتا ہے
    (قطار/کالم ڈالنے یا حذف کرنے کے بعد)۔ row_map/col_map کچھ لوٹائیں یا None (حذف شدہ)۔
    """
    from openpyxl.utils import column_index_from_string, get_column_letter

    def sub(m):
        cabs, col, rabs, row = m.group(1), m.group(2), m.group(3), int(m.group(4))
        ci = column_index_from_string(col)
        if col_map:
            ci = col_map(ci)
        if row_map:
            row = row_map(row)
        if ci is None or row is None:
            return '#REF!'
        return f"{cabs}{get_column_letter(ci)}{rabs}{row}"

    for row_cells in ws.iter_rows():
        for cell in row_cells:
            v = cell.value
            if isinstance(v, str) and v.startswith('='):
                cell.value = _REF_RE.sub(sub, v)


def _shift_rows_map(at, n=1):
    return lambda r: r + n if r >= at else r


def _deleted_rows_map(deleted):
    d = sorted(set(deleted))

    def f(r):
        if r in d:
            return None
        return r - sum(1 for x in d if x < r)
    return f


def _fill_formulas_from_neighbour(ws, dst_row, src_row, ncol):
    """نئی قطار کے خالی خانوں میں ساتھ والی قطار کا فارمولا نئی جگہ کے مطابق نقل ہوتا ہے۔"""
    from openpyxl.formula.translate import Translator
    if src_row < 1 or src_row == dst_row:
        return
    for c in range(1, ncol + 1):
        cell = ws.cell(dst_row, c)
        srcv = ws.cell(src_row, c).value
        if cell.value in (None, '') and isinstance(srcv, str) and srcv.startswith('='):
            cell.value = Translator(srcv, origin=ws.cell(src_row, c).coordinate).translate_formula(cell.coordinate)


def _strip_formulas(rows):
    """دوسری شیٹ کے فارمولے (ان کے حوالہ جات غلط ہو جاتے ہیں) ملاتے وقت خالی کر دیے جاتے ہیں۔"""
    out = []
    for r in rows:
        out.append({k: (None if isinstance(v, str) and v.startswith('=') else v) for k, v in r.items()})
    return out


def _insert_row(ws, n, values):
    ws.insert_rows(n)
    _remap_formulas(ws, row_map=_shift_rows_map(n))
=======
def _insert_row(ws, n, values):
    ws.insert_rows(n)
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    ncol = ws.max_column
    neighbour = n - 1 if n > 2 else n + 1
    _copy_row_style(ws, neighbour, n, ncol)
    headers = _sheet_headers(ws)
    for name, v in values.items():
        for i, h in enumerate(headers, 1):
            if h.lower() == name.lower():
                ws.cell(n, i).value = v
<<<<<<< HEAD
    _fill_formulas_from_neighbour(ws, n, neighbour, ncol)
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    return f"قطار {n} پر نئی قطار ڈالی گئی۔"


def _add_column(ws, name, at=None):
    ncol = ws.max_column
    if at is None or at > ncol:
        at = ncol + 1
    else:
        ws.insert_cols(at)
<<<<<<< HEAD
        _remap_formulas(ws, col_map=_shift_rows_map(at))
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    src = at - 1 if at > 1 else at + 1
    ws.cell(1, at).value = name
    ws.cell(1, at)._style = copy(ws.cell(1, src)._style)
    for r in range(2, _last_data_row(ws) + 1):
        ws.cell(r, at)._style = copy(ws.cell(r, src)._style)
    return f"کالم '{name}' شامل ہوا (کالم {at})۔"


def _filter_sheet(wb, ws, col_name, op, val):
    ci = _col_of(ws, col_name)
    headers = _sheet_headers(ws)
    out = wb.create_sheet(f"Filter {len(wb.sheetnames)}")
    for c in range(1, len(headers) + 1):
        out.cell(1, c).value = headers[c - 1]
        out.cell(1, c)._style = copy(ws.cell(1, c)._style)
    k = 2
    for r in range(2, _last_data_row(ws) + 1):
        if _match(ws.cell(r, ci).value, op, val):
            for c in range(1, len(headers) + 1):
                out.cell(k, c).value = ws.cell(r, c).value
                out.cell(k, c)._style = copy(ws.cell(r, c)._style)
            k += 1
    out.freeze_panes = 'A2'
    return f"فلٹر شیٹ '{out.title}' میں {k - 2} قطاریں آئیں۔"


def _delete_where(ws, col_name, op, val):
    ci = _col_of(ws, col_name)
    rows = [r for r in range(2, _last_data_row(ws) + 1) if _match(ws.cell(r, ci).value, op, val)]
    for r in reversed(rows):
        ws.delete_rows(r)
<<<<<<< HEAD
    _remap_formulas(ws, row_map=_deleted_rows_map(rows))
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    return f"{len(rows)} قطاریں حذف ہوئیں۔"


def _sort(ws, col_name, desc=False):
    ci = _col_of(ws, col_name)
    last = _last_data_row(ws)
    ncol = ws.max_column
    records = []
    for r in range(2, last + 1):
        vals = [ws.cell(r, c).value for c in range(1, ncol + 1)]
        styles = [copy(ws.cell(r, c)._style) for c in range(1, ncol + 1)]
<<<<<<< HEAD
        records.append((vals, styles, r))
=======
        records.append((vals, styles))
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7

    def key(rec):
        v = rec[0][ci - 1]
        n = _norm(v) if v not in (None, '') else None
        if n is None:
            return (2, 0.0, '')
        return (0, n, '') if isinstance(n, float) else (1, 0.0, n)

    records.sort(key=key, reverse=desc)
<<<<<<< HEAD
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    for i, (vals, styles, old_r) in enumerate(records, 2):
        for c in range(1, ncol + 1):
            v = vals[c - 1]
            if isinstance(v, str) and v.startswith('=') and old_r != i:
                L = get_column_letter(c)
                v = Translator(v, origin=f"{L}{old_r}").translate_formula(f"{L}{i}")
            ws.cell(i, c).value = v
=======
    for i, (vals, styles) in enumerate(records, 2):
        for c in range(1, ncol + 1):
            ws.cell(i, c).value = vals[c - 1]
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            ws.cell(i, c)._style = styles[c - 1]
    return f"{len(records)} قطاریں '{col_name}' کے مطابق ترتیب دی گئیں۔"


def _fill_blanks(ws, src_ws, col_name, key_name):
    ci = _col_of(ws, col_name)
    ki = _col_of(ws, key_name)
    src_headers, src_rows = _table(src_ws)
    if col_name not in src_headers or key_name not in src_headers:
        raise ValueError(f"دوسری فائل میں '{col_name}' یا '{key_name}' کالم نہیں ہے۔")
    lookup = {_cell_text(r[key_name]).lower(): r[col_name] for r in src_rows if _cell_text(r[key_name])}
    filled = 0
    for r in range(2, _last_data_row(ws) + 1):
        if ws.cell(r, ci).value in (None, ''):
            k = _cell_text(ws.cell(r, ki).value).lower()
            if k in lookup and lookup[k] not in (None, ''):
                ws.cell(r, ci).value = lookup[k]
                filled += 1
    return f"{filled} خالی خانے دوسری فائل سے بھر دیے گئے۔"


def _reconcile(wb, ws, src_ws, key_name):
    h1, r1 = _table(ws)
    h2, r2 = _table(src_ws)
    if key_name not in h1 or key_name not in h2:
        raise ValueError(f"دونوں فائلوں میں '{key_name}' کالم ہونا چاہیے۔")
    out_rows = reconcile_rows(h1, r1, h2, r2, key_name)
    rep = wb.create_sheet("Reconciliation")
    for c, h in enumerate(["Key", "Status", "Differences"], 1):
        _style_header(rep.cell(1, c, h))
    for i, row in enumerate(out_rows, 2):
        for c, v in enumerate(row, 1):
            rep.cell(i, c).value = v
    rep.column_dimensions['A'].width = 22
    rep.column_dimensions['B'].width = 18
    rep.column_dimensions['C'].width = 60
    rep.freeze_panes = 'A2'
    counts = Counter(r[1] for r in out_rows)
    return "ری کنسائل مکمل: " + "، ".join(f"{k}: {v}" for k, v in counts.items())


<<<<<<< HEAD
def _hex_color(v):
    if v is None or str(v).strip() == '':
        return None
    s = str(v).strip().lstrip('#').upper()
    if not re.fullmatch(r'[0-9A-F]{6}', s):
        raise ValueError(f"رنگ کا کوڈ غلط ہے: {v} (مثلاً FFFF00)")
    return s


_BORDER_STYLES = ('none', 'thin', 'medium', 'thick', 'dotted', 'dashed', 'double')


def apply_format(ws, rng, fill=None, font_color=None, bold=None, italic=None,
                 font_name=None, font_size=None, align=None, number_format=None, border=None):
    """
    خانوں کا بیک گراؤنڈ، فونٹ، رنگ، الائنمنٹ اور فارمیٹ بدلتا ہے۔
    rng: 'A1:E1' | 'row 3' | 'col C' | 'header' (پہلی قطار)۔ Returns: تبدیل شدہ خانوں کی تعداد۔
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils.cell import range_boundaries
    border_val = (str(border).strip().lower() if border is not None else None) or None
    if border_val and border_val not in _BORDER_STYLES:
        raise ValueError("border کو none، thin، medium، thick، dotted، dashed یا double ہونا چاہیے")
    s = str(rng).strip()
    if s.lower() == 'header':
        s = 'row 1'
    m = re.fullmatch(r'row\s+(\d+)', s, re.I)
    if m:
        r = int(m.group(1))
        min_col, min_row, max_col, max_row = 1, r, max(ws.max_column, 1), r
    else:
        m = re.fullmatch(r'col(?:umn)?\s+([A-Za-z]+)', s, re.I)
        if m:
            s = f"{m.group(1).upper()}:{m.group(1).upper()}"
        min_col, min_row, max_col, max_row = range_boundaries(s.upper())
        if min_row is None:
            min_row, max_row = 1, max(ws.max_row, 1)
        if min_col is None:
            min_col, max_col = 1, max(ws.max_column, 1)

    fill_hex = _hex_color(fill) if str(fill or '').strip().lower() != 'none' else 'NONE'
    fc_hex = _hex_color(font_color)
    align_val = (align or '').strip().lower() or None
    if align_val and align_val not in ('left', 'center', 'right'):
        raise ValueError("الائنمنٹ left، center یا right ہونی چاہیے")

    count = 0
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(r, c)
            if fill_hex == 'NONE':
                cell.fill = PatternFill(fill_type=None)
            elif fill_hex:
                cell.fill = PatternFill(fill_type='solid', start_color='FF' + fill_hex, end_color='FF' + fill_hex)
            f = cell.font
            changes = {}
            if font_name:
                changes['name'] = str(font_name)
            if font_size:
                changes['size'] = float(font_size)
            if bold is not None:
                changes['bold'] = bool(bold)
            if italic is not None:
                changes['italic'] = bool(italic)
            if fc_hex:
                changes['color'] = 'FF' + fc_hex
            if changes:
                cell.font = Font(name=changes.get('name', f.name), size=changes.get('size', f.size),
                                 bold=changes.get('bold', f.bold), italic=changes.get('italic', f.italic),
                                 underline=f.underline, strike=f.strike,
                                 color=changes.get('color', f.color))
            if align_val:
                a = cell.alignment
                cell.alignment = Alignment(horizontal=align_val, vertical=a.vertical or 'center',
                                           wrap_text=a.wrap_text)
            if number_format:
                cell.number_format = str(number_format)
            if border_val:
                if border_val == 'none':
                    cell.border = Border()
                else:
                    side = Side(style=border_val, color='FF000000')
                    cell.border = Border(left=side, right=side, top=side, bottom=side)
            count += 1
    return count


def _parse_format_command(text):
    """'format A1:E1 fill=FFFF00 bold' → (rng, kwargs)"""
    m = re.match(r'^format\s+(row\s+\d+|col(?:umn)?\s+[A-Za-z]+|header|[A-Za-z]+\d*:[A-Za-z]+\d*|[A-Za-z]+\d+)(.*)$',
                 text, re.I)
    if not m:
        raise ValueError("format کے بعد رینج لکھیں، مثلاً: format row 1 fill=FFFF00 bold")
    rng, rest = m.group(1), m.group(2)
    kw = {}
    names = {'fill': 'fill', 'font_color': 'font_color', 'color': 'font_color', 'font': 'font_name',
             'size': 'font_size', 'align': 'align', 'number': 'number_format', 'border': 'border'}
    for k, v in re.findall(r'(\w+)\s*=\s*("[^"]*"|\S+)', rest):
        key = names.get(k.lower())
        if not key:
            raise ValueError(f"نامعلوم خصوصیت: {k}")
        kw[key] = v.strip('"')
    words = [w.lower() for w in re.findall(r'\b(bold|italic|nobold|noitalic)\b', rest, re.I)]
    if 'bold' in words:
        kw['bold'] = True
    if 'nobold' in words:
        kw['bold'] = False
    if 'italic' in words:
        kw['italic'] = True
    if 'noitalic' in words:
        kw['italic'] = False
    if not kw:
        raise ValueError("کم از کم ایک خصوصیت لکھیں، مثلاً fill=FFFF00 یا bold")
    return rng, kw


=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
def run_command(line, wb, extras):
    """
    ایک کمانڈ چلاتا ہے۔ Returns (ok, message).
    wb: ورک بک (پہلی شیٹ پر کام ہوتا ہے)
    extras: اضافی اپلوڈ شدہ ورک بکس کی لسٹ
    """
    text = line.strip()
    if not text or text.startswith('#'):
        return True, ""
    ws = wb.worksheets[0]
    low = text.lower()
    try:
        if low == 'help':
            return True, AGENT_HELP

        m = re.match(r'^insert row at (\d+)(?:\s*:\s*(.*))?$', text, re.I)
        if m:
            return True, _insert_row(ws, int(m.group(1)), _assign_pairs(m.group(2)))

        m = re.match(r'^add row(?:\s*:\s*(.*))?$', text, re.I)
        if m:
            n = append_rows(ws, [_assign_pairs(m.group(1))])
            return True, f"{n} قطار جوڑی گئی۔"

        m = re.match(r'^insert column (.+?) at (\d+)$', text, re.I)
        if m:
            return True, _add_column(ws, m.group(1).strip(), int(m.group(2)))

        m = re.match(r'^add column (.+)$', text, re.I)
        if m:
            return True, _add_column(ws, m.group(1).strip())

        m = re.match(rf'^delete where (.+?) {_OP_TOKENS} (.+)$', text, re.I)
        if m:
            return True, _delete_where(ws, m.group(1).strip(), m.group(2).lower(), m.group(3).strip())

        m = re.match(rf'^filter (.+?) {_OP_TOKENS} (.+)$', text, re.I)
        if m:
            return True, _filter_sheet(wb, ws, m.group(1).strip(), m.group(2).lower(), m.group(3).strip())

        m = re.match(r'^sort (.+?) (asc|desc)$', text, re.I)
        if m:
            return True, _sort(ws, m.group(1).strip(), m.group(2).lower() == 'desc')

        m = re.match(r'^fill (.+?) by (.+)$', text, re.I)
        if m:
            if not extras:
                return False, "اس کمانڈ کے لیے دوسری فائل اپلوڈ کریں۔"
            return True, _fill_blanks(ws, extras[0].worksheets[0], m.group(1).strip(), m.group(2).strip())

        m = re.match(r'^reconcile (.+)$', text, re.I)
        if m:
            if not extras:
                return False, "ری کنسائل کے لیے دوسری فائل اپلوڈ کریں۔"
            return True, _reconcile(wb, ws, extras[0].worksheets[0], m.group(1).strip())

        m = re.match(r'^consolidate (.+?) sum (.+)$', text, re.I)
        if m:
            sums = [s.strip() for s in m.group(2).split(',') if s.strip()]
            add_consolidated_sheet(wb, ws, m.group(1).strip(), sums)
            return True, "شیٹ 'Consolidated' بن گئی۔"

        if low == 'merge':
            if not extras:
                return False, "مرج کے لیے فائل اپلوڈ کریں۔"
            total = 0
            for ex in extras:
                _, rows = _table(ex.worksheets[0])
<<<<<<< HEAD
                total += append_rows(ws, _strip_formulas(rows))
=======
                total += append_rows(ws, rows)
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            return True, f"{total} قطاریں ملا دی گئیں۔"

        m = re.match(r'^rename column (.+?) to (.+)$', text, re.I)
        if m:
            ci = _col_of(ws, m.group(1).strip())
            ws.cell(1, ci).value = m.group(2).strip()
            return True, "کالم کا نام بدل دیا گیا۔"

<<<<<<< HEAD
        if low.startswith('format '):
            rng, kw = _parse_format_command(text)
            n = apply_format(ws, rng, **kw)
            return True, f"{n} خانوں کی فارمیٹنگ بدل دی گئی۔"

=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
        return False, "یہ کمانڈ سمجھ نہیں آئی۔ مدد کے لیے 'help' لکھیں۔"
    except Exception as e:
        return False, f"خامی: {e}"


def run_commands(text, wb, extras):
    logs = []
    for line in text.splitlines():
        if not line.strip() or line.strip().startswith('#'):
            continue
        ok, msg = run_command(line, wb, extras)
        logs.append(("✅ " if ok else "❌ ") + f"{line.strip()} → {msg}")
    return logs


# ==========================================
# 5. ایکٹیویٹی پروگریس (کئی تاریخوں کے کالم میں فیصد)
# ==========================================
import difflib as _difflib
from datetime import datetime as _dt

_PROG_CODE_RE = re.compile(r'^\s*\*?\s*([A-I])\s?(\d+)\.(\d+)\s*')
_PROG_PCT_RE = re.compile(r'\(\s*(\d{1,3}(?:\.\d+)?)\s*%\s*Completed', re.I)
_PROG_SECTION_RE = re.compile(r'^\s*\*?\s*([A-I])[\.\s]+[A-Za-z]')
_PROG_BARE_SECTION_RE = re.compile(r'^\s*([A-I])\s*$')
_PROG_DATE_RE = re.compile(r'Date\s*:\s*\*?\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})', re.I)


def _read_text(raw):
    """فائل کا انکوڈنگ خود پہچانیں: UTF-16 (Notepad کا 'Unicode')، UTF-8 یا Windows-1252۔"""
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        text = raw.decode('utf-16', errors='ignore')
    elif raw.count(b'\x00') > len(raw) // 4:
        text = raw.decode('utf-16-le', errors='ignore')
    else:
        try:
            text = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = raw.decode('cp1252', errors='ignore')
    return text.lstrip('\ufeff')


def _norm_text(s):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9]+', ' ', str(s).lower())).strip()


def _sim(a, b):
    na, nb = _norm_text(a), _norm_text(b)
    if not na or not nb:
        return 0.0
    if min(len(na), len(nb)) >= 12 and (na in nb or nb in na):
        return 0.95
    return _difflib.SequenceMatcher(None, na, nb).ratio()


def _is_report_line(v):
    """رپورٹ کی اپنی لائن (کوڈ + فیصد) ٹیمپلیٹ کا خانہ نہیں ہے، اسے چھوڑ دیں۔"""
    return bool(_PROG_CODE_RE.match(v)) or '%' in v or 'completed' in v.lower()


def _clean_desc(s):
    """نام کے آخر/شروع سے فالتو ) ( اور خالی جگہ ہٹائیں، مگر جائز قوسین رہیں۔"""
    s = s.strip()
    while s:
        if s[-1] in '( ' or s[-1].isspace():
            s = s[:-1]
        elif s[-1] == ')' and s.count(')') > s.count('('):
            s = s[:-1]
        else:
            break
        s = s.rstrip()
    while s and (s[0] == ')' or s[0].isspace()):
        s = s[1:]
    return s.strip()


def _parse_date_match(s):
    m = _PROG_DATE_RE.search(s)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    y = y + 2000 if y < 100 else y
    try:
        return _dt(y, mo, d).date()
    except ValueError:
        return None


def split_activity_blocks(lines):
    """
    ہر 'Date:' لائن سے نیا بلاک شروع ہوتا ہے۔
    Returns: [(report_date_or_None, [lines_of_that_block])]
    پہلی 'Date:' سے پہلے کی لائنیں (مثلاً واٹس ایپ کا ہیڈر) چھوڑ دی جاتی ہیں۔
    """
    blocks = []
    cur = None
    for line in lines:
        s = line.strip()
        if _PROG_DATE_RE.search(s):
            cur = (_parse_date_match(s), [])
            blocks.append(cur)
            continue
        if cur is not None and s:
            cur[1].append(s)
    return blocks


def parse_activities(lines):
    """
    لائنوں سے ایکٹیویٹیاں نکالیں: کوڈ + نام + فیصد۔
    Returns: [{'letter', 'code', 'desc', 'pct'}]
    """
    acts = []
    for s in lines:
        cm = _PROG_CODE_RE.match(s)
        pm = _PROG_PCT_RE.search(s)
<<<<<<< HEAD
        if not pm:
            continue
        if cm:
=======
        if cm and pm:
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            acts.append({
                'letter': cm.group(1),
                'code': f"{cm.group(1)}{cm.group(2)}.{cm.group(3)}",
                'desc': _clean_desc(s[cm.end():pm.start()]),
                'pct': float(pm.group(1)),
            })
<<<<<<< HEAD
        else:
            # کوڈ کے بغیر رپورٹ (مثلاً "Backfilling (50% completed)")
            desc = _clean_desc(s[:pm.start()])
            if desc:
                acts.append({'letter': None, 'code': '', 'desc': desc, 'pct': float(pm.group(1))})
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    return acts


def _header_date(v):
    if isinstance(v, _dt):
        return v.date()
    if hasattr(v, 'year') and hasattr(v, 'month'):
        return v
    if isinstance(v, str):
        for fmt in ('%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y'):
            try:
                return _dt.strptime(v.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _header_map(ws, max_scan=20):
    """ٹیمپلیٹ میں ہر تاریخ کا خانہ: {date: (row, col)}"""
    out = {}
    for r in range(1, min(ws.max_row, max_scan) + 1):
        for c in range(1, ws.max_column + 1):
            d = _header_date(ws.cell(r, c).value)
            if d is not None and d not in out:
                out[d] = (r, c)
    return out


def _section_map(ws, last_row):
    """
    ہر قطار کے لیے موجودہ سیکشن کا حرف۔ سیکشن کی سرخی دو طرح کی ہو سکتی ہے:
    'A' اکیلا (ٹیمپلیٹ) یا 'A. Control House' (رپورٹ)۔
    Returns: (sections{row: letter}, header_rows set)
    """
    letter = None
    out = {}
    header_rows = set()
    for r in range(1, last_row + 1):
        found = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                m = _PROG_BARE_SECTION_RE.match(v) or _PROG_SECTION_RE.match(v)
                if m:
                    found = m.group(1)
                    break
        if found:
            letter = found
            header_rows.add(r)
        out[r] = letter
    return out, header_rows


def _pick_desc_col(ws, hr, last, skip_cols, acts):
    """وہ کالم جس میں سب سے زیادہ ایکٹیویٹی کے نام ملیں۔"""
    best_col, best_hits = None, 0
    for c in range(1, ws.max_column + 1):
        if c in skip_cols:
            continue
        hits = 0
        for a in acts:
            for r in range(hr + 1, last + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and not _is_report_line(v) and _sim(a['desc'], v) >= 0.8:
                    hits += 1
                    break
        if hits > best_hits:
            best_col, best_hits = c, hits
    return best_col


def _match_block(ws, hr, last, desc_col, sections, header_rows, acts):
    """
    ایک رپورٹ بلاک کی ایکٹیویٹیاں ٹیمپلیٹ کی قطاروں سے ملائیں۔
    پہلے اسی سیکشن میں، پھر باقی میں۔ Returns: {act_index: template_row}
    """
    cand = []
    for r in range(hr + 1, last + 1):
        if r in header_rows:
            continue
        v = ws.cell(r, desc_col).value
        if isinstance(v, str) and not _is_report_line(v):
            cand.append((r, v, sections.get(r)))

    used = set()
    assigned = {}

    def best_for(a, same_section_only):
        best_r, best_s = None, 0.0
        for r, v, sec in cand:
            if r in used:
                continue
<<<<<<< HEAD
            if same_section_only and a['letter'] and sec != a['letter']:
=======
            if same_section_only and sec != a['letter']:
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
                continue
            sc = _sim(a['desc'], v)
            if sc > best_s:
                best_r, best_s = r, sc
        return best_r, best_s

    for i, a in enumerate(acts):
        r, sc = best_for(a, True)
        if r is not None and sc >= 0.8:
            used.add(r)
            assigned[i] = r
    for i, a in enumerate(acts):
        if i in assigned:
            continue
        r, sc = best_for(a, False)
        if r is not None and sc >= 0.8:
            used.add(r)
            assigned[i] = r
    return assigned


def process_progress(activity_lines, template_bytes):
    """
    کئی تاریخوں والی رپورٹ کو ٹیمپلیٹ میں بھرتا ہے۔
    ہر 'Date:' بلاک کے فیصد اسی تاریخ کے کالم میں لکھے جاتے ہیں۔
    Returns: (BytesIO | None, [messages], [unmatched_lines])
    """
    raw_blocks = split_activity_blocks(activity_lines)
    if not raw_blocks:
        return None, ["خامی: رپورٹ میں 'Date: dd/mm/yy' والی کوئی لائن نہیں ملی۔"], []

    blocks = [(d, parse_activities(ls)) for d, ls in raw_blocks]

    try:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    except Exception as e:
        return None, [f"خامی: ٹیمپلیٹ پڑھنے میں مسئلہ: {e}"], []
    ws = wb.worksheets[0]

    hmap = _header_map(ws)
    if not hmap:
        return None, ["خامی: ٹیمپلیٹ میں کوئی تاریخ والا ہیڈر نہیں ملا (مثلاً '13-Nov-26')۔"], []
    hr_all = min(r for r, _ in hmap.values())
    date_cols = {c for _, c in hmap.values()}
    last = ws.max_row
    sections, header_rows = _section_map(ws, last)

    # ڈیٹا کالم کا تعین: تمام بلاکس کی ایکٹیویٹیوں کے نام (بغیر دہرائے) سے
    uniq = {}
    for _, acts in blocks:
        for a in acts:
            uniq.setdefault(_norm_text(a['desc']), a)
    desc_col = _pick_desc_col(ws, hr_all, last, date_cols, list(uniq.values()))
    if desc_col is None:
        return None, ["خامی: ٹیمپلیٹ میں ایکٹیویٹی کے نام والا کالم نہیں ملا۔"], []

    # ایک ہی تاریخ کے ایک سے زیادہ بلاک: وہ بلاک لیں جس کی سب سے زیادہ ایکٹیویٹیاں ٹیمپلیٹ سے ملیں
    groups, order = {}, []
    messages = []
    for d, acts in blocks:
        if d is None:
            messages.append(f"خامی: ایک بلاک کی تاریخ درست نہیں تھی ({len(acts)} ایکٹیویٹیاں چھوڑی گئیں)۔")
            continue
        if not acts:
            messages.append(f"خامی: تاریخ {d.strftime('%d-%b-%y')} کے بلاک میں کوئی ایکٹیویٹی با فیصد نہیں ملی۔")
            continue
        if d not in hmap:
            messages.append(f"خامی: تاریخ {d.strftime('%d-%b-%y')} کا کالم ٹیمپلیٹ میں نہیں ملا، یہ بلاک چھوڑ دیا گیا۔")
            continue
        if d not in groups:
            groups[d] = []
            order.append(d)
        groups[d].append(acts)

    unmatched = []
    total_written = 0
    for d in order:
        hr, date_col = hmap[d]
        results = []
        for acts in groups[d]:
            assigned = _match_block(ws, hr, last, desc_col, sections, header_rows, acts)
            results.append((len(assigned), acts, assigned))
        best_idx = 0
        for i, (score, _, _) in enumerate(results):
<<<<<<< HEAD
            # برابر میچ ہوں تو بعد والا بلاک (زیادہ تازہ رپورٹ) لیا جاتا ہے
            if score >= results[best_idx][0]:
=======
            if score > results[best_idx][0]:
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
                best_idx = i
        best_score, acts, assigned = results[best_idx]
        for i, (score, other, _) in enumerate(results):
            if i != best_idx:
                messages.append(
                    f"ℹ️ تاریخ {d.strftime('%d-%b-%y')} کا ایک اور بلاک ({len(other)} ایکٹیویٹیاں، ٹیمپلیٹ سے {score} ملیں) چھوڑ دیا گیا۔"
                )
        matched = 0
        for i, a in enumerate(acts):
            r = assigned.get(i)
            if r is None:
<<<<<<< HEAD
                unmatched.append(f"{d.strftime('%d-%b-%y')} | {a['code']} {a['desc']} ({a['pct']:g}%)".replace('  ', ' ').strip())
=======
                unmatched.append(f"{d.strftime('%d-%b-%y')} | {a['code']} {a['desc']} ({a['pct']:g}%)")
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
                continue
            cell = ws.cell(r, date_col)
            cell.value = a['pct'] / 100.0
            cell.number_format = '0%'
            matched += 1
        total_written += matched
        messages.append(
            f"✅ تاریخ {d.strftime('%d-%b-%y')}: {matched} فیصد لکھے گئے، کل ایکٹیویٹیاں: {len(acts)}۔"
        )

    if total_written == 0:
        return None, messages or ["خامی: کوئی بلاک نہیں لکھا جا سکا۔"], unmatched

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, messages, unmatched


def render_progress_ui():
<<<<<<< HEAD
    st.info("ایکٹیویٹی رپورٹ (ہر قسم کی فائلیں، ایک یا زیادہ) اور ٹیمپلیٹ (xlsx، PDF، Word یا تصویر) اپلوڈ کریں۔ رپورٹ میں جتنی بھی 'Date:' لائنیں ہوں گی، ہر تاریخ کے فیصد ٹیمپلیٹ میں اسی تاریخ کے کالم میں لکھے جائیں گے۔ نتیجہ ہمیشہ xlsx فائل ہوگی۔")
    up_rep = st.file_uploader("ایکٹیویٹی رپورٹ (txt, csv, xlsx, pdf, docx, تصویر، ایک یا زیادہ)", type=_RAW_TYPES,
                              accept_multiple_files=True, key="prog_txt")
    up_tpl = st.file_uploader("ٹیمپلیٹ (xlsx, pdf, docx یا تصویر)", type=_RAW_TYPES, key="prog_tpl")
    up_style = st.file_uploader("فارمیٹ ٹیمپلیٹ (اختیاری، xlsx: رنگ، ہیڈر، عنوان، Gantt کی شکل کے لیے، جیسے Askari کی فارمیٹ والی فائل)",
                                type=["xlsx"], key="prog_style")
    if up_rep and up_tpl:
        lines, notes, errs = [], [], []
        for f in up_rep:
            try:
                l, n = extract_raw_lines(f.name, f.getvalue())
                lines.extend(l)
                notes.append(f"{f.name}: {n}")
            except ValueError as e:
                errs.append(f"{f.name}: {e}")
        for n in notes:
            st.info(n)
        if lines:
            date_n = sum(1 for l in lines if _PROG_DATE_RE.search(l))
            pct_n = sum(1 for l in lines if _PROG_PCT_RE.search(l))
            st.info(f"کل لائنیں: {len(lines)} | 'Date:' والی لائنیں (تاریخیں): {date_n} | فیصد والی ایکٹیویٹی لائنیں: {pct_n}۔ "
                    "صرف یہی متعلقہ لائنیں استعمال ہوتی ہیں، باقی لائنیں نظر انداز ہوتی ہیں۔")
            if date_n == 0 or pct_n == 0:
                st.warning("رپورٹ میں 'Date:' یا فیصد والی لائنیں نہیں ملیں۔ رپورٹ کا فارمیٹ دیکھیں۔")
        for e in errs:
            st.error(e)
        if up_tpl.name.lower().endswith((".pdf", ".docx", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")):
            st.warning("ٹیمپلیٹ PDF/Word/تصویر ہے: اس سے خانے اور ڈھانچہ تو پڑھا جاتا ہے، مگر ہیڈر کے رنگ، گرین Gantt خانے اور merged عنوان نہیں آتے۔ "
                       "فارمیٹ والی xlsx ٹیمپلیٹ (جیسے b42a0bcf-Two_Week_Lookahead_-_Askari.xlsx) اپلوڈ کریں، تاکہ آؤٹ پٹ اور PDF دونوں اسی فارمیٹ میں آئیں۔")
        tpl_bytes = None
        try:
            tpl_bytes = template_xlsx_bytes(up_tpl.name, up_tpl.getvalue())
        except ValueError as e:
            st.error(str(e))
        if st.button("تمام تاریخوں کے فیصد لکھیں") and lines and tpl_bytes:
            res, messages, unmatched = process_progress(lines, tpl_bytes)
=======
    st.info("ایکٹیویٹی رپورٹ (.txt) اور ٹیمپلیٹ (.xlsx) اپلوڈ کریں۔ رپورٹ میں جتنی بھی 'Date:' لائنیں ہوں گی، ہر تاریخ کے فیصد ٹیمپلیٹ میں اسی تاریخ کے کالم میں لکھے جائیں گے۔")
    up_txt = st.file_uploader("ایکٹیویٹی رپورٹ (.txt)", type=["txt"], key="prog_txt")
    up_tpl = st.file_uploader("ٹیمپلیٹ (.xlsx)", type=["xlsx"], key="prog_tpl")
    if up_txt and up_tpl:
        lines = _read_text(up_txt.getvalue()).splitlines()
        if st.button("تمام تاریخوں کے فیصد لکھیں"):
            res, messages, unmatched = process_progress(lines, up_tpl.getvalue())
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            for m in messages:
                if m.startswith("✅"):
                    st.success(m)
                elif m.startswith("ℹ️"):
                    st.info(m)
                else:
                    st.error(m)
<<<<<<< HEAD
            if res is not None and up_style is not None:
                try:
                    styled = apply_style_template(res.getvalue(), up_style.getvalue())
                    res = io.BytesIO(styled)
                    st.info("فارمیٹ ٹیمپلیٹ کی شکل (رنگ، ہیڈر، عنوان، بارڈر) نتیجے پر لگا دی گئی ہے۔")
                except Exception as e:
                    st.error(f"فارمیٹ ٹیمپلیٹ لگانے میں خامی: {e}")
            if res is not None:
=======
            if res is None:
                pass
            else:
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
                if unmatched:
                    st.warning("یہ ایکٹیویٹیاں ٹیمپلیٹ میں نہیں ملیں:")
                    for u in unmatched:
                        st.write("• " + u)
                st.download_button("📥 پروگریس والی فائل ڈاؤن لوڈ کریں", res, "Progress_Updated.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
<<<<<<< HEAD
                _pdf_download_button(res.getvalue(), "Progress_Updated.pdf", "📄 پروگریس PDF میں ڈاؤن لوڈ کریں", "prog_pdf")
    else:
        st.warning("رپورٹ اور ٹیمپلیٹ دونوں اپلوڈ کریں۔")
=======
    else:
        st.warning("دونوں فائلیں اپلوڈ کریں۔")
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7


# ==========================================
# 6. AI ایجنٹ (Grok API، OpenAI جیسا فارمیٹ)
# ==========================================
import json as _json
import os as _os

<<<<<<< HEAD
GROQ_BASE_URL = "https://api.groq.com/openai/v1"
# ماڈل کا نام Groq کنسول کے Models صفحے سے دیکھیں؛ GROQ_MODEL ماحول کے متغیر سے بدلا جا سکتا ہے
GROQ_DEFAULT_MODEL = _os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")
AGENT_MAX_STEPS = 60
=======
XAI_BASE_URL = _os.environ.get("XAI_BASE_URL", "https://api.x.ai/v1")
XAI_DEFAULT_MODEL = _os.environ.get("XAI_MODEL", "grok-4")  # کنسول میں دستیاب ماڈل کا درست نام یہاں یا XAI_MODEL میں لکھیں
AGENT_MAX_STEPS = 25
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7

AGENT_SYSTEM = """آپ ایک ماہر ایکسل آپریٹر ہیں۔ آپ کو صارف کی ہدایت پر ایکسل فائل میں کام کرنا ہے۔
اصول:
1. ہمیشہ پہلے inspect_workbook کال کریں تاکہ شیٹیں، ہیڈر اور ڈھانچہ معلوم ہو۔
2. کالم کے لیے ہیڈر کا نام استعمال کریں جو inspect میں نظر آئے۔ کوئی نام طے شدہ نہیں، ہر فائل کی بناوٹ الگ ہو سکتی ہے۔
3. ڈیٹا بدلنے سے پہلے read_range سے اسے دیکھیں۔
4. جوڑ، ٹوٹل وغیرہ کے لیے فارمولے استعمال کریں، ٹیکسٹ نہیں۔ نمبر کو نمبر ہی لکھیں۔
5. موجودہ فارمیٹ نہ بگاڑیں۔ اصل شیٹ کو صرف ضرورت کے مطابق بدلیں۔
<<<<<<< HEAD
6. اگر صارف کہے کہ خود فیصلہ کرو، یا ہدایت میں تھوڑی کمی ہو، تو سوال نہ پوچھیں۔ فائل کے ڈھانچے اور موجودہ ڈیٹا کو دیکھ کر معقول فیصلہ خود کریں، اور summary میں بتائیں کہ کیا فیصلہ کیا۔
7. سوال صرف اس وقت پوچھیں جب فائل میں ضروری کالم ہی نہ ہو اور کام ممکن نہ ہو۔
9. کسی شیٹ میں دوسری فائل/شیٹ سے ڈیٹا جمع (compile) کرتے وقت وہی قطار دوبارہ نہ ڈالیں۔ جمع کرنے کے بعد remove_duplicates کے ساتھ کلیدی کالم (مثلاً ID) دے کر ڈپلیکیٹ ہٹائیں، اور summary میں لکھیں کہ کتنی قطاریں آئیں۔
8. کام ختم ہونے پر finish کال کریں، summary میں اردو میں بتائیں کہ کیا کیا بدلا۔ summary صرف سادہ اردو متن ہو، JSON یا کوڈ نہیں۔"""
=======
6. اگر ہدایت واضح نہ ہو یا کوئی کالم نہ ملے تو کام نہ کریں، finish کے summary میں سوال لکھیں۔
7. کام ختم ہونے پر finish کال کریں، summary میں اردو میں بتائیں کہ کیا کیا بدلا۔"""
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7


def _jsonable(v):
    if v is None or isinstance(v, (int, float, str, bool)):
        return v
    if hasattr(v, 'isoformat'):
        return v.isoformat()
    return str(v)


<<<<<<< HEAD
def _row_key_value(ws, r, c):
    """قطار کی پہچان کے لیے قیمت؛ فارمولا ہو تو اسے ایک مشترک قطار میں لا کر موازنہ ہوتا ہے۔"""
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    v = ws.cell(r, c).value
    if isinstance(v, str) and v.startswith('='):
        L = get_column_letter(c)
        return ('F', Translator(v, origin=f"{L}{r}").translate_formula(f"{L}2"))
    return _norm(v)


=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
class AgentWorkbook:
    """ایجنٹ کے ٹولز: ہر ٹول openpyxl پر ایک محفوظ عمل ہے۔"""

    def __init__(self, wb):
        self.wb = wb
        self.log = []

    # ---- مددگار ----
    def _ws(self, sheet):
        if not sheet:
            return self.wb.worksheets[0]
        if sheet not in self.wb.sheetnames:
            raise ValueError(f"شیٹ '{sheet}' نہیں ملی۔ موجود شیٹیں: {self.wb.sheetnames}")
        return self.wb[sheet]

    def _col(self, ws, col):
        if isinstance(col, int):
            return col
        return _col_of(ws, str(col))

    def _values_by_col(self, ws, values):
        out = {}
        for k, v in (values or {}).items():
            out[self._col(ws, k)] = v
        return out

    # ---- ٹولز ----
    def inspect_workbook(self):
        sheets = []
        for ws in self.wb.worksheets:
            hr = 1
            for r in range(1, min(ws.max_row, 15) + 1):
                if sum(1 for c in range(1, ws.max_column + 1)
                       if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip()) >= 2:
                    hr = r
                    break
            headers = [_cell_text(ws.cell(hr, c).value) for c in range(1, ws.max_column + 1)]
            preview = [[_jsonable(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 30) + 1)]
                       for r in range(1, min(ws.max_row, 12) + 1)]
            sheets.append({
                "name": ws.title, "max_row": ws.max_row, "max_col": ws.max_column,
                "header_row": hr, "headers": headers,
                "merged_ranges": [str(x) for x in list(ws.merged_cells.ranges)[:20]],
                "preview_first_12_rows": preview,
            })
        return {"sheets": sheets}

    def read_range(self, sheet=None, start_row=1, end_row=20, max_col=30):
        ws = self._ws(sheet)
        end_row = min(int(end_row), ws.max_row)
        max_col = min(int(max_col), ws.max_column)
        return [[_jsonable(ws.cell(r, c).value) for c in range(1, max_col + 1)]
                for r in range(int(start_row), end_row + 1)]

    def set_cell(self, sheet=None, row=1, col=1, value=None, number_format=None):
        ws = self._ws(sheet)
        if int(row) > ws.max_row + 200 or int(row) < 1:
            raise ValueError(f"قطار {row} حد سے باہر ہے (شیٹ میں {ws.max_row} قطاریں ہیں)")
        cell = ws.cell(int(row), self._col(ws, col))
        cell.value = value if (isinstance(value, str) and value.startswith('=')) else _coerce(value)
        if number_format:
            cell.number_format = number_format
        return f"{cell.coordinate} لکھا گیا"

    def append_row(self, sheet=None, values=None):
        ws = self._ws(sheet)
        vals = self._values_by_col(ws, values)
        headers = _sheet_headers(ws)
        dict_row = {headers[c - 1]: v for c, v in vals.items() if 1 <= c <= len(headers)}
        n = append_rows(ws, [dict_row])
        return f"{n} قطار آخر میں جوڑی گئی"

    def insert_row(self, sheet=None, at_row=2, values=None):
        ws = self._ws(sheet)
        vals = self._values_by_col(ws, values)
        return _insert_row(ws, int(at_row), {_sheet_headers(ws)[c - 1]: v for c, v in vals.items()})

    def insert_column(self, sheet=None, at_col=1, header=""):
        ws = self._ws(sheet)
        return _add_column(ws, str(header), int(at_col))

    def delete_rows_where(self, sheet=None, column="", op="=", value=""):
        ws = self._ws(sheet)
        return _delete_where(ws, str(column), op, str(value))

    def filter_to_new_sheet(self, sheet=None, column="", op="=", value="", new_sheet_name="Filtered"):
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        headers = _sheet_headers(ws)
        out = self.wb.create_sheet(str(new_sheet_name)[:31])
        for c in range(1, len(headers) + 1):
            out.cell(1, c).value = headers[c - 1]
            out.cell(1, c)._style = copy(ws.cell(1, c)._style)
        k = 2
        for r in range(2, _last_data_row(ws) + 1):
            if _match(ws.cell(r, ci).value, op, value):
                for c in range(1, len(headers) + 1):
                    out.cell(k, c).value = ws.cell(r, c).value
                    out.cell(k, c)._style = copy(ws.cell(r, c)._style)
                k += 1
        return f"شیٹ '{out.title}' میں {k - 2} قطاریں"

    def sort_sheet(self, sheet=None, column="", descending=False):
        ws = self._ws(sheet)
        return _sort(ws, str(column), bool(descending))

    def lookup_fill(self, sheet=None, target_column="", key_column="", source_sheet="", source_key="", source_value=""):
        ws = self._ws(sheet)
        src = self._ws(source_sheet)
        _, src_rows = _table(src)
        lookup = {_cell_text(r.get(source_key)).lower(): r.get(source_value)
                  for r in src_rows if _cell_text(r.get(source_key))}
        tci = self._col(ws, target_column)
        kci = self._col(ws, key_column)
        filled = 0
        for r in range(2, _last_data_row(ws) + 1):
            if ws.cell(r, tci).value in (None, ''):
                k = _cell_text(ws.cell(r, kci).value).lower()
                if lookup.get(k) not in (None, ''):
                    ws.cell(r, tci).value = lookup[k]
                    filled += 1
        return f"{filled} خانے بھرے گئے"

    def merge_sheet(self, sheet=None, source_sheet=""):
        ws = self._ws(sheet)
        _, rows = _table(self._ws(source_sheet))
<<<<<<< HEAD
        return f"{append_rows(ws, _strip_formulas(rows))} قطاریں ملائی گئیں"
=======
        return f"{append_rows(ws, rows)} قطاریں ملائی گئیں"
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7

    def reconcile(self, sheet=None, other_sheet="", key=""):
        return _reconcile(self.wb, self._ws(sheet), self._ws(other_sheet), str(key))

    def consolidate(self, sheet=None, group="", sum_columns=None):
        ws = self._ws(sheet)
        add_consolidated_sheet(self.wb, ws, str(group), list(sum_columns or []))
        return "شیٹ 'Consolidated' بن گئی"

    def create_sheet(self, name="NewSheet", headers=None, rows=None):
        ws = self.wb.create_sheet(str(name)[:31])
        headers = list(headers or [])
        for c, h in enumerate(headers, 1):
            _style_header(ws.cell(1, c, h))
        for r, row in enumerate(rows or [], 2):
            for c, v in enumerate(row, 1):
                ws.cell(r, c).value = _coerce(v)
        ws.freeze_panes = 'A2'
        return f"شیٹ '{ws.title}' بن گئی ({len(rows or [])} قطاریں)"

<<<<<<< HEAD
    def format_cells(self, sheet=None, cell_range="A1", fill=None, font_color=None, bold=None,
                     italic=None, font_name=None, font_size=None, align=None, number_format=None, border=None):
        ws = self._ws(sheet)
        n = apply_format(ws, str(cell_range), fill=fill, font_color=font_color, bold=bold, italic=italic,
                         font_name=font_name, font_size=font_size, align=align, number_format=number_format,
                         border=border)
        return f"{n} خانوں کی فارمیٹنگ بدل دی گئی"

    def delete_column(self, sheet=None, column=""):
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        ws.delete_cols(ci)
        _remap_formulas(ws, col_map=_deleted_rows_map([ci]))
        return f"کالم {ci} حذف ہوا"

    def rename_column(self, sheet=None, column="", new_name=""):
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        if not str(new_name).strip():
            raise ValueError("نیا نام خالی نہیں ہو سکتا")
        ws.cell(1, ci).value = str(new_name).strip()
        return f"کالم کا نام '{str(new_name).strip()}' رکھا گیا"

    def clean_text(self, sheet=None, column="", trim=True, case="none"):
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        if case not in ("none", "upper", "lower", "title"):
            raise ValueError("case کو none، upper، lower یا title ہونا چاہیے")
        n = 0
        for r in range(2, _last_data_row(ws) + 1):
            cell = ws.cell(r, ci)
            v = cell.value
            if isinstance(v, str) and not v.startswith('='):
                nv = re.sub(r'\s+', ' ', v).strip() if trim else v
                if case == "upper":
                    nv = nv.upper()
                elif case == "lower":
                    nv = nv.lower()
                elif case == "title":
                    nv = nv.title()
                if nv != v:
                    cell.value = nv
                    n += 1
        return f"{n} خانے صاف ہوئے"

    def remove_duplicates(self, sheet=None, columns=None):
        ws = self._ws(sheet)
        cols = [self._col(ws, c) for c in columns] if columns else list(range(1, ws.max_column + 1))
        seen, drop = set(), []
        for r in range(2, _last_data_row(ws) + 1):
            key = tuple(_row_key_value(ws, r, c) for c in cols)
            if all(k in ('', None) for k in key):
                continue
            if key in seen:
                drop.append(r)
            else:
                seen.add(key)
        for r in reversed(drop):
            ws.delete_rows(r)
        _remap_formulas(ws, row_map=_deleted_rows_map(drop))
        return f"{len(drop)} ڈپلیکیٹ قطاریں ہٹائی گئیں"

    def find_replace(self, sheet=None, find="", replace="", column=None, match_case=False):
        ws = self._ws(sheet)
        if not str(find):
            raise ValueError("کیا تلاش کرنا ہے یہ لکھیں")
        cols = [self._col(ws, column)] if column else list(range(1, ws.max_column + 1))
        pat = re.compile(re.escape(str(find)), 0 if match_case else re.IGNORECASE)
        n = 0
        for r in range(2, _last_data_row(ws) + 1):
            for c in cols:
                cell = ws.cell(r, c)
                v = cell.value
                if isinstance(v, str) and not v.startswith('=') and pat.search(v):
                    cell.value = pat.sub(lambda m: str(replace), v)
                    n += 1
        return f"{n} خانوں میں تبدیلی ہوئی"

    def filter_multi(self, sheet=None, conditions=None, logic="and", new_sheet_name="Filtered"):
        ws = self._ws(sheet)
        if not conditions:
            raise ValueError("کم از کم ایک شرط دیں")
        if logic not in ("and", "or"):
            raise ValueError("logic 'and' یا 'or' ہونا چاہیے")
        parsed = []
        for cond in conditions:
            op = str(cond.get("op", "=")).lower()
            if op not in ("=", "!=", ">", "<", ">=", "<=", "contains"):
                raise ValueError(f"نامعلوم آپریٹر: {op}")
            parsed.append((self._col(ws, cond.get("column", "")), op, cond.get("value", "")))
        headers = _sheet_headers(ws)
        out = self.wb.create_sheet(str(new_sheet_name)[:31])
        for c in range(1, len(headers) + 1):
            out.cell(1, c).value = headers[c - 1]
            out.cell(1, c)._style = copy(ws.cell(1, c)._style)
        k = 2
        for r in range(2, _last_data_row(ws) + 1):
            hits = [_match(ws.cell(r, ci).value, op, val) for ci, op, val in parsed]
            if (all(hits) if logic == "and" else any(hits)):
                for c in range(1, len(headers) + 1):
                    out.cell(k, c).value = ws.cell(r, c).value
                    out.cell(k, c)._style = copy(ws.cell(r, c)._style)
                k += 1
        return f"شیٹ '{out.title}' میں {k - 2} قطاریں"

    def set_sheet_options(self, sheet=None, freeze_header=None, autofilter=None, column=None, width=None):
        ws = self._ws(sheet)
        msgs = []
        if freeze_header is not None:
            ws.freeze_panes = 'A2' if freeze_header else None
            msgs.append("ہیڈر فریز" if freeze_header else "فریز ہٹایا")
        if autofilter is not None:
            ws.auto_filter.ref = (f"A1:{get_column_letter(max(ws.max_column, 1))}{max(ws.max_row, 1)}"
                                  if autofilter else None)
            msgs.append("فلٹر آن" if autofilter else "فلٹر بند")
        if column and width is not None:
            L = get_column_letter(self._col(ws, column))
            ws.column_dimensions[L].width = float(width)
            msgs.append(f"کالم {L} کی چوڑائی {width}")
        return "؛ ".join(msgs) or "کوئی تبدیلی نہیں مانگی گئی"

    def rename_sheet(self, sheet=None, new_name=""):
        ws = self._ws(sheet)
        new = str(new_name).strip()[:31]
        if not new or new in self.wb.sheetnames:
            raise ValueError("نیا نام خالی ہو یا پہلے سے موجود ہو")
        old = ws.title
        ws.title = new
        for other in self.wb.worksheets:
            for row_cells in other.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        cell.value = cell.value.replace(f"'{old}'!", f"'{new}'!").replace(f"{old}!", f"{new}!")
        return f"شیٹ '{old}' کا نام '{new}' رکھا گیا (اس نام کے فارمولے بھی بدل دیے گئے)"

    def delete_sheet(self, sheet=None):
        ws = self._ws(sheet)
        if len(self.wb.sheetnames) <= 1:
            raise ValueError("آخری شیٹ حذف نہیں ہو سکتی")
        title = ws.title
        self.wb.remove(ws)
        return f"شیٹ '{title}' حذف ہوئی (اگر کسی فارمولے نے اسے استعمال کیا تھا تو وہ خراب ہو سکتا ہے)"

    def split_text(self, sheet=None, column="", separator=",", parts=2):
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        parts = int(parts)
        if not str(separator) or parts < 2:
            raise ValueError("separator لکھیں اور parts کم از کم 2 ہوں")
        base = _cell_text(ws.cell(1, ci).value) or "کالم"
        start = ws.max_column + 1
        for k in range(parts):
            ws.cell(1, start + k).value = f"{base}_{k + 1}"
            _style_header(ws.cell(1, start + k))
        n = 0
        for r in range(2, _last_data_row(ws) + 1):
            v = ws.cell(r, ci).value
            if isinstance(v, str) and not v.startswith('='):
                pieces = [p.strip() for p in v.split(str(separator))]
                for k in range(parts):
                    ws.cell(r, start + k).value = _coerce(pieces[k]) if k < len(pieces) else None
                n += 1
        return f"{n} قطاریں تقسیم ہوئیں، نئے کالم آخر میں شامل ہوئے"

    def concat_columns(self, sheet=None, columns=None, separator=" ", new_name="Combined"):
        ws = self._ws(sheet)
        cols = [self._col(ws, c) for c in (columns or [])]
        if len(cols) < 2:
            raise ValueError("کم از کم دو کالم دیں")
        nc = ws.max_column + 1
        ws.cell(1, nc).value = str(new_name)
        _style_header(ws.cell(1, nc))
        for r in range(2, _last_data_row(ws) + 1):
            ws.cell(r, nc).value = str(separator).join(_cell_text(ws.cell(r, c).value) for c in cols).strip()
        return f"کالم '{new_name}' شامل ہوا"

    def delete_blank_rows(self, sheet=None):
        ws = self._ws(sheet)
        drop = [r for r in range(2, ws.max_row + 1)
                if all(ws.cell(r, c).value in (None, '') for c in range(1, ws.max_column + 1))]
        for r in reversed(drop):
            ws.delete_rows(r)
        _remap_formulas(ws, row_map=_deleted_rows_map(drop))
        return f"{len(drop)} خالی قطاریں ہٹائی گئیں"

    def sort_multi(self, sheet=None, keys=None):
        ws = self._ws(sheet)
        if not keys:
            raise ValueError("کم از کم ایک کالم دیں")
        for k in reversed(keys):
            _sort(ws, str(k.get("column", "")), bool(k.get("descending", False)))
        return "ایک سے زیادہ کالموں کے مطابق ترتیب دی گئی"

    def hide_column(self, sheet=None, column="", hidden=True):
        ws = self._ws(sheet)
        L = get_column_letter(self._col(ws, column))
        ws.column_dimensions[L].hidden = bool(hidden)
        return f"کالم {L} {'چھپا دیا' if hidden else 'دکھا دیا'} گیا"

    def merge_cells(self, sheet=None, cell_range="A1:B1", unmerge=False):
        ws = self._ws(sheet)
        rng = str(cell_range)
        if unmerge:
            ws.unmerge_cells(rng)
            return f"{rng} کا مرج ختم"
        from openpyxl.utils.cell import range_boundaries
        c1, r1, c2, r2 = range_boundaries(rng)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1) and ws.cell(r, c).value not in (None, ''):
                    raise ValueError("مرج سے پہلے دوسرے خانوں کا ڈیٹا ہٹائیں (مرج صرف پہلے خانے کی قیمت رکھتا ہے)")
        ws.merge_cells(rng)
        return f"{rng} مرج ہو گیا"

    def add_summary_row(self, sheet=None, columns=None, function="sum"):
        ws = self._ws(sheet)
        fn_map = {"sum": "SUM", "average": "AVERAGE", "count": "COUNT", "max": "MAX", "min": "MIN"}
        fn = str(function).lower()
        if fn not in fn_map:
            raise ValueError("function ایک ہونا چاہیے: sum، average، count، max یا min")
        cols = [self._col(ws, c) for c in (columns or [])]
        last = _last_data_row(ws)
        tr = last + 1
        ws.cell(tr, 1).value = "Total" if fn == "sum" else "Summary"
        ws.cell(tr, 1).font = Font(bold=True)
        for c in cols:
            L = get_column_letter(c)
            cell = ws.cell(tr, c)
            cell.value = f"={fn_map[fn]}({L}2:{L}{last})"
            cell.font = Font(bold=True)
            cell.number_format = '#,##0.00'
        return f"قطار {tr} پر {fn_map[fn]} کے فارمولے لگ گئے"

    def copy_sheet(self, sheet=None, new_name="Copy"):
        ws = self._ws(sheet)
        new = str(new_name).strip()[:31]
        if not new or new in self.wb.sheetnames:
            raise ValueError("نیا نام خالی ہو یا پہلے سے موجود ہو")
        dup = self.wb.copy_worksheet(ws)
        dup.title = new
        return f"شیٹ کی نقل '{new}' بن گئی"

    def add_lookup_column(self, sheet=None, new_name="", key_column="", source_sheet="",
                          source_key="", source_value=""):
        ws = self._ws(sheet)
        src = self._ws(source_sheet)
        _, src_rows = _table(src)
        lookup = {_cell_text(r.get(source_key)).lower(): r.get(source_value)
                  for r in src_rows if _cell_text(r.get(source_key))}
        kci = self._col(ws, key_column)
        nc = ws.max_column + 1
        ws.cell(1, nc).value = str(new_name)
        _style_header(ws.cell(1, nc))
        found = 0
        for r in range(2, _last_data_row(ws) + 1):
            k = _cell_text(ws.cell(r, kci).value).lower()
            if k in lookup:
                ws.cell(r, nc).value = lookup[k]
                found += 1
            else:
                ws.cell(r, nc).value = "نہیں ملا"
        return f"{found} قطاروں کے لیے قیمت ملی، باقی پر 'نہیں ملا' لکھا"

    def add_dropdown(self, sheet=None, column="", options=None):
        from openpyxl.worksheet.datavalidation import DataValidation
        ws = self._ws(sheet)
        ci = self._col(ws, column)
        opts = [str(o).replace('"', '').replace(',', ' ') for o in (options or []) if str(o).strip()]
        if not opts:
            raise ValueError("کم از کم ایک آپشن دیں")
        formula = '"' + ",".join(opts) + '"'
        if len(formula) > 255:
            raise ValueError("ڈراپ ڈاؤن کی فہرست بہت لمبی ہے")
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        L = get_column_letter(ci)
        dv.add(f"{L}2:{L}{max(ws.max_row, 2)}")
        return f"کالم {L} میں ڈراپ ڈاؤن لگ گیا"

    def highlight_rule(self, sheet=None, column="", op="<", value=0, fill="FFC7CE"):
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import PatternFill
        ops = {"=": "equal", "!=": "notEqual", ">": "greaterThan", "<": "lessThan",
               ">=": "greaterThanOrEqual", "<=": "lessThanOrEqual"}
        if str(op) not in ops:
            raise ValueError("op ایک ہونا چاہیے: =، !=، >، <، >=، <=")
        ws = self._ws(sheet)
        L = get_column_letter(self._col(ws, column))
        last = max(ws.max_row, 2)
        v = _coerce(value)
        formula = str(v) if isinstance(v, (int, float)) else f'"{v}"'
        hx = _hex_color(fill)
        rule = CellIsRule(operator=ops[str(op)], formula=[formula],
                          fill=PatternFill(start_color='FF' + hx, end_color='FF' + hx, fill_type='solid'))
        ws.conditional_formatting.add(f"{L}2:{L}{last}", rule)
        return f"کالم {L} میں شرط '{op} {value}' پر رنگ لگے گا"

    def add_chart(self, sheet=None, kind="bar", category_column="", value_columns=None,
                  title="", anchor="H2"):
        from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        ws = self._ws(sheet)
        kinds = {"bar": BarChart, "line": LineChart, "pie": PieChart}
        if str(kind).lower() not in kinds:
            raise ValueError("kind ایک ہونا چاہیے: bar، line یا pie")
        vals = [self._col(ws, c) for c in (value_columns or [])]
        if not vals:
            raise ValueError("کم از کم ایک ویلیو کالم دیں")
        cat = self._col(ws, category_column)
        last = _last_data_row(ws)
        chart = kinds[str(kind).lower()]()
        chart.title = str(title) or None
        for vc in vals:
            chart.add_data(Reference(ws, min_col=vc, min_row=1, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=cat, min_row=2, max_row=last))
        ws.add_chart(chart, str(anchor))
        return f"{kind} چارٹ {anchor} پر شامل ہوا"

=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    def call(self, name, args):
        method = getattr(self, name, None)
        if name.startswith('_') or method is None or not callable(method):
            raise ValueError(f"نامعلوم ٹول: {name}")
        return method(**args)


def _tool(name, description, properties=None, required=None):
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties or {}, "required": required or []},
    }}


_S = {"type": "string"}
_I = {"type": "integer"}
_B = {"type": "boolean"}
_V = {"description": "قیمت (نمبر، ٹیکسٹ یا فارمولا جو = سے شروع ہو)"}
_VALS = {"type": "object", "description": "کالم نام یا نمبر → قیمت"}
_SHEET = {"type": "string", "description": "شیٹ کا نام (خالی ہو تو پہلی شیٹ)"}
AGENT_TOOLS = [
    _tool("inspect_workbook", "فائل کی تمام شیٹیں، ہیڈر، مرج شدہ رینجز اور پہلی 12 قطاریں دیکھیں", {}),
    _tool("read_range", "ایک شیٹ کی قطاروں کی قیمتیں پڑھیں",
          {"sheet": _SHEET, "start_row": _I, "end_row": _I, "max_col": _I}, ["start_row", "end_row"]),
    _tool("set_cell", "ایک خانے میں قیمت یا فارمولا لکھیں",
          {"sheet": _SHEET, "row": _I, "col": {"description": "کالم نمبر یا ہیڈر نام"}, "value": _V,
           "number_format": {"type": "string", "description": "مثلاً 0% یا #,##0.00"}}, ["row", "col", "value"]),
    _tool("append_row", "آخر میں نئی قطار جوڑیں", {"sheet": _SHEET, "values": _VALS}, ["values"]),
    _tool("insert_row", "کسی قطار پر بیچ میں نئی قطار ڈالیں", {"sheet": _SHEET, "at_row": _I, "values": _VALS}, ["at_row"]),
    _tool("insert_column", "کسی کالم پر نیا کالم ڈالیں",
          {"sheet": _SHEET, "at_col": _I, "header": _S}, ["at_col", "header"]),
    _tool("delete_rows_where", "شرط پوری کرنے والی قطاریں حذف کریں (op: =, !=, >, <, >=, <=, contains)",
          {"sheet": _SHEET, "column": _S, "op": _S, "value": _S}, ["column", "op", "value"]),
    _tool("filter_to_new_sheet", "شرط پوری کرنے والی قطاریں نئی شیٹ میں نکالیں",
          {"sheet": _SHEET, "column": _S, "op": _S, "value": _S, "new_sheet_name": _S},
          ["column", "op", "value", "new_sheet_name"]),
    _tool("sort_sheet", "ایک کالم کے مطابق ڈیٹا ترتیب دیں",
          {"sheet": _SHEET, "column": _S, "descending": _B}, ["column"]),
    _tool("lookup_fill", "دوسری شیٹ سے کی (key) کے ذریعے خالی خانے بھریں",
          {"sheet": _SHEET, "target_column": _S, "key_column": _S, "source_sheet": _S,
           "source_key": _S, "source_value": _S},
          ["target_column", "key_column", "source_sheet", "source_key", "source_value"]),
    _tool("merge_sheet", "دوسری شیٹ کی قطاریں کالم نام سے ملا کر آخر میں جوڑیں",
          {"sheet": _SHEET, "source_sheet": _S}, ["source_sheet"]),
    _tool("reconcile", "دو شیٹوں کا کی کے ذریعے موازنہ، نتیجہ 'Reconciliation' شیٹ میں",
          {"sheet": _SHEET, "other_sheet": _S, "key": _S}, ["other_sheet", "key"]),
    _tool("consolidate", "گروپ کے مطابق جمع، نتیجہ 'Consolidated' شیٹ میں (SUMIFS فارمولے)",
          {"sheet": _SHEET, "group": _S, "sum_columns": {"type": "array", "items": _S}}, ["group", "sum_columns"]),
    _tool("create_sheet", "نئی شیٹ بنائیں جس میں ہیڈر اور قطاریں ہوں",
          {"name": _S, "headers": {"type": "array", "items": _S},
           "rows": {"type": "array", "items": {"type": "array"}}}, ["name", "headers"]),
<<<<<<< HEAD
    _tool("format_cells", "خانوں کا بیک گراؤنڈ رنگ (fill)، فونٹ کا رنگ، نام، سائز، بولڈ/ترچھا اور الائنمنٹ بدلیں۔ "
          "cell_range مثلاً 'A1:E1' یا 'row 1' (ہیڈر کے لیے 'row 1')؛ رنگ 6 ہندسوں کا ہیکس کوڈ جیسے FFFF00",
          {"sheet": _SHEET, "cell_range": {"type": "string"}, "fill": {"type": "string", "description": "بیک گراؤنڈ ہیکس رنگ یا none"},
           "font_color": {"type": "string", "description": "فونٹ کا ہیکس رنگ"}, "bold": _B, "italic": _B,
           "font_name": _S, "font_size": {"type": "number"},
           "align": {"type": "string", "description": "left، center یا right"},
           "number_format": {"type": "string"},
           "border": {"type": "string", "description": "none، thin (باریک)، medium، thick (موٹا)، dotted، dashed یا double"}}, ["cell_range"]),
    _tool("delete_column", "ایک کالم حذف کریں (فارمولوں کے حوالے درست رکھے جاتے ہیں)",
          {"sheet": _SHEET, "column": _S}, ["column"]),
    _tool("rename_column", "کالم کا ہیڈر نام بدلیں",
          {"sheet": _SHEET, "column": _S, "new_name": _S}, ["column", "new_name"]),
    _tool("clean_text", "ایک کالم کا متن صاف کریں: اضافی جگہیں ہٹائیں، اور case (none، upper، lower، title) بدلیں",
          {"sheet": _SHEET, "column": _S, "trim": _B,
           "case": {"type": "string", "description": "none، upper، lower یا title"}}, ["column"]),
    _tool("remove_duplicates", "ڈپلیکیٹ قطاریں ہٹائیں؛ columns دیں تو صرف انہی کالموں کے مطابق، ورنہ پوری قطار کے مطابق",
          {"sheet": _SHEET, "columns": {"type": "array", "items": _S}}, []),
    _tool("find_replace", "متن تلاش کر کے بدلیں؛ column دیں تو صرف اس کالم میں، match_case سے بڑے چھوٹے حروف کا فرق",
          {"sheet": _SHEET, "find": _S, "replace": _S, "column": _S, "match_case": _B}, ["find", "replace"]),
    _tool("filter_multi", "کئی شرطوں سے فلٹر کریں؛ logic 'and' (سب شرطیں) یا 'or' (کوئی ایک)، نتیجہ نئی شیٹ میں",
          {"sheet": _SHEET,
           "conditions": {"type": "array", "items": {"type": "object", "properties": {
               "column": _S, "op": {"type": "string", "description": "=، !=، >، <، >=، <=، contains"},
               "value": {"type": "string"}}, "required": ["column", "op", "value"]}},
           "logic": {"type": "string", "description": "and یا or"},
           "new_sheet_name": _S}, ["conditions"]),
    _tool("set_sheet_options", "ہیڈر فریز کریں، آٹو فلٹر آن/بند کریں، اور کسی کالم کی چوڑائی طے کریں",
          {"sheet": _SHEET, "freeze_header": _B, "autofilter": _B, "column": _S,
           "width": {"type": "number"}}, []),
    _tool("rename_sheet", "شیٹ کا نام بدلیں (اس نام کے فارمولے بھی بدل جاتے ہیں)",
          {"sheet": _SHEET, "new_name": _S}, ["new_name"]),
    _tool("delete_sheet", "ایک شیٹ مکمل حذف کریں (آخری شیٹ حذف نہیں ہوتی)",
          {"sheet": _SHEET}, []),
    _tool("split_text", "ایک کالم کا متن جدا کر کے نئے کالم آخر میں بنائیں (مثلاً 'نام، شہر')",
          {"sheet": _SHEET, "column": _S, "separator": _S, "parts": _I}, ["column", "separator"]),
    _tool("concat_columns", "کئی کالموں کا متن جوڑ کر نیا کالم بنائیں",
          {"sheet": _SHEET, "columns": {"type": "array", "items": _S}, "separator": _S, "new_name": _S},
          ["columns"]),
    _tool("delete_blank_rows", "پوری طرح خالی قطاریں ہٹائیں", {"sheet": _SHEET}, []),
    _tool("sort_multi", "ایک سے زیادہ کالموں کے مطابق ترتیب دیں (پہلا کالم سب سے اہم)",
          {"sheet": _SHEET, "keys": {"type": "array", "items": {"type": "object", "properties": {
              "column": _S, "descending": _B}, "required": ["column"]}}}, ["keys"]),
    _tool("hide_column", "کالم چھپائیں یا دکھائیں",
          {"sheet": _SHEET, "column": _S, "hidden": _B}, ["column"]),
    _tool("merge_cells", "خانوں کی رینج مرج کریں یا unmerge کریں (مثلاً 'A1:E1')",
          {"sheet": _SHEET, "cell_range": _S, "unmerge": _B}, ["cell_range"]),
    _tool("add_summary_row", "آخر میں خلاصہ قطار لگائیں: columns کا sum، average، count، max یا min (فارمولوں کے ساتھ)",
          {"sheet": _SHEET, "columns": {"type": "array", "items": _S},
           "function": {"type": "string", "description": "sum، average، count، max یا min"}}, ["columns"]),
    _tool("copy_sheet", "شیٹ کی مکمل نقل نئے نام سے بنائیں",
          {"sheet": _SHEET, "new_name": _S}, ["new_name"]),
    _tool("add_lookup_column", "دوسری شیٹ سے ملا کر نیا کالم بنائیں (مثلاً قیمت یا شہر)، نہ ملنے پر 'نہیں ملا'",
          {"sheet": _SHEET, "new_name": _S, "key_column": _S, "source_sheet": _S,
           "source_key": _S, "source_value": _S},
          ["new_name", "key_column", "source_sheet", "source_key", "source_value"]),
    _tool("add_dropdown", "کالم میں ڈراپ ڈاؤن فہرست لگائیں",
          {"sheet": _SHEET, "column": _S, "options": {"type": "array", "items": _S}}, ["column", "options"]),
    _tool("highlight_rule", "شرط پوری ہونے پر کالم کے خانے کا رنگ بدلنے کا قاعدہ لگائیں (مثلاً Qty < 3 پر سرخ)",
          {"sheet": _SHEET, "column": _S, "op": _S, "value": {"type": "string"},
           "fill": {"type": "string", "description": "ہیکس رنگ، مثلاً FFC7CE"}}, ["column", "op", "value"]),
    _tool("add_chart", "ڈیٹا سے چارٹ بنائیں: kind bar، line یا pie؛ category_column اور value_columns دیں",
          {"sheet": _SHEET, "kind": _S, "category_column": _S,
           "value_columns": {"type": "array", "items": _S}, "title": _S, "anchor": _S},
          ["category_column", "value_columns"]),
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
    _tool("finish", "کام ختم، summary میں اردو میں بتائیں کہ کیا ہوا (یا سوال)",
          {"summary": _S}, ["summary"]),
]


<<<<<<< HEAD
def _combine_workbooks(files):
    """
    ایک یا زیادہ ایکسل فائلیں ایک ورک بک میں۔ پہلی فائل بنیاد ہے؛ باقی فائلوں کی شیٹیں
    'فائل نام_شیٹ نام' کے ساتھ شامل ہوتی ہیں، تاکہ ایجنٹ دونوں کو ایک ساتھ استعمال کر سکے۔
    """
    from copy import copy as _cp
    base_wb = openpyxl.load_workbook(io.BytesIO(files[0].getvalue()))
    for f in files[1:]:
        src = openpyxl.load_workbook(io.BytesIO(f.getvalue()))
        stem = re.sub(r'\.xlsx$', '', str(f.name), flags=re.I)[:18]
        for ws in src.worksheets:
            base_title = f"{stem}_{ws.title}"[:31]
            title, k = base_title, 2
            while title in base_wb.sheetnames:
                title = f"{base_title[:27]}_{k}"
                k += 1
            dst = base_wb.create_sheet(title)
            for row in ws.iter_rows():
                for c in row:
                    d = dst.cell(row=c.row, column=c.column)
                    d.value = c.value
                    if c.has_style:
                        d.font = _cp(c.font)
                        d.fill = _cp(c.fill)
                        d.border = _cp(c.border)
                        d.alignment = _cp(c.alignment)
                        d.number_format = c.number_format
            for col_letter, dim in ws.column_dimensions.items():
                if dim.width:
                    dst.column_dimensions[col_letter].width = dim.width
    return base_wb


def _get_groq_key():
    """کلید پہلے st.secrets (.streamlit/secrets.toml یا Streamlit Cloud کی Secrets) سے، پھر ماحول کے متغیر سے۔"""
    try:
        v = str(st.secrets["GROQ_API_KEY"]).strip()
        if v:
            return v
    except Exception:
        pass
    return _os.environ.get("GROQ_API_KEY", "").strip()


def _make_client():
    key = _get_groq_key()
    if not key:
        raise ValueError("GROQ_API_KEY ماحول میں نہیں ملی۔ اسے سیٹ کریں، کوڈ میں نہ لکھیں۔")
    import openai
    # ہر API کال کے لیے 15 سیکنڈ کی حد؛ پوری ایجنٹ رن کی حد نہیں
    return openai.OpenAI(api_key=key, base_url=GROQ_BASE_URL, timeout=15.0)


def _salvage_summary(err):
    """
    Groq کبھی آخری جواب کو غلط ٹول نام ('json' وغیرہ) میں بھیج دیتا ہے۔
    اس صورت میں خلاصہ اس خرابی والے متن سے نکال لیتے ہیں۔
    """
    m = re.search(r'"summary"\s*:\s*"(.*?)"(?:\s|\\n)*\}', str(err), re.S)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return "ایجنٹ نے کام مکمل کیا (خلاصہ حاصل نہیں ہو سکا؛ نتیجہ فائل ڈاؤن لوڈ کر کے دیکھیں)۔"


def _clean_summary(text):
    """اگر ماڈل نے خلاصہ JSON یا کوڈ کی شکل میں لکھ دیا ہو تو صرف اس کا متن نکالیں۔"""
    t = (text or '').strip()
    try:
        obj = _json.loads(t)
        if isinstance(obj, dict):
            for key in ('summary', 'message', 'text'):
                if isinstance(obj.get(key), str) and obj[key].strip():
                    return obj[key].strip()
    except ValueError:
        pass
    m = re.search(r'"summary"\s*:\s*"(.*?)"\s*\}', t, re.S)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return t


def _inherit_styles(wb, keep, source_titles):
    """
    ایجنٹ کی بنائی ہوئی شیٹ میں ہیڈر اور ڈیٹا کی فارمیٹنگ (رنگ، بارڈر، فونٹ، سیدھ)
    اصل فائل سے لی جاتی ہے، ہیڈر کے نام کے ملاپ سے۔ ملاپ نہ ہو تو کچھ نہیں بدلتا۔
    """
    from copy import copy as _cp
    ws = wb[keep]
    kept_hdr = {}
    for c in ws[1]:
        if c.value is not None:
            kept_hdr[str(c.value).strip()] = c.column
    best, best_hits = None, 0
    for t in source_titles:
        if t == keep or t not in wb.sheetnames:
            continue
        src = wb[t]
        src_hdr = {str(c.value).strip(): c.column for c in src[1] if c.value is not None}
        hits = len(set(kept_hdr) & set(src_hdr))
        if hits > best_hits:
            best, best_hits = src, hits
    if best is None:
        return 0
    src_hdr = {str(c.value).strip(): c.column for c in best[1] if c.value is not None}
    body_row = 2 if best.max_row >= 2 else 1
    changed = 0
    for name, kcol in kept_hdr.items():
        if name not in src_hdr:
            continue
        scol = src_hdr[name]
        sh = best.cell(1, scol)
        sb = best.cell(body_row, scol)
        for r in range(1, ws.max_row + 1):
            src_cell = sh if r == 1 else sb
            d = ws.cell(row=r, column=kcol)
            d.font = _cp(src_cell.font)
            d.fill = _cp(src_cell.fill)
            d.border = _cp(src_cell.border)
            d.alignment = _cp(src_cell.alignment)
            d.number_format = src_cell.number_format
            changed += 1
    return changed


def _keep_one_output_sheet(wb, source_titles):
    """
    آؤٹ پٹ میں ایک ہی شیٹ رکھیں: وہ آخری شیٹ جو ایجنٹ نے نئی بنائی ہو، ورنہ پہلی شیٹ۔
    فیصلہ: اگر رکھی جانے والی شیٹ کے فارمولے کسی دوسری شیٹ کا حوالہ دیں، تو وہ شیٹ
    رکھی جاتی ہے تاکہ فارمولے نہ ٹوٹیں۔ باقی شیٹیں ہٹ جاتی ہیں۔
    واپسی: (ورک بک، رکھی گئی اضافی شیٹوں کے نام)
    """
    created = [t for t in wb.sheetnames if t not in source_titles]
    keep = created[-1] if created else wb.sheetnames[0]
    ws = wb[keep]
    referenced = set()
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                for t in wb.sheetnames:
                    if t != keep and re.search(r"(?:'" + re.escape(t) + r"'|(?<![A-Za-z0-9_])" + re.escape(t) + r")!", v):
                        referenced.add(t)
    if created:
        _inherit_styles(wb, keep, source_titles)
    kept_extra = sorted(referenced)
    for t in list(wb.sheetnames):
        if t != keep and t not in referenced:
            del wb[t]
    wb.active = wb.sheetnames.index(keep)
    return wb, kept_extra
=======
def _make_client():
    key = _os.environ.get("XAI_API_KEY", "").strip()
    if not key:
        raise ValueError("XAI_API_KEY ماحول میں نہیں ملی۔ اسے سیٹ کریں، کوڈ میں نہ لکھیں۔")
    import openai
    return openai.OpenAI(api_key=key, base_url=XAI_BASE_URL)
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7


def run_agent(instruction, wb, client=None, model=None, max_steps=AGENT_MAX_STEPS):
    """
    Returns: (summary, log). wb کو براہ راست بدلا جاتا ہے (UI میں کاپی بنا کر دیں)۔
    """
    client = client or _make_client()
<<<<<<< HEAD
    model = model or GROQ_DEFAULT_MODEL
    aw = AgentWorkbook(wb)
    messages = [{"role": "system", "content": AGENT_SYSTEM},
                {"role": "user", "content": instruction}]
    bad_replies = 0
    for _ in range(max_steps):
        try:
            resp = client.chat.completions.create(model=model, messages=messages, tools=AGENT_TOOLS)
        except Exception as e:
            text = str(e)
            if 'rate_limit' in text or '429' in text:
                aw.log.append("⛔ Groq کی روزانہ ٹوکن حد ختم ہو گئی ہے؛ کچھ دیر بعد یا اگلے دن دوبارہ کوشش کریں۔")
                return "ایجنٹ روک دیا گیا: Groq کی حد ختم ہو گئی ہے۔ کچھ دیر بعد دوبارہ چلائیں۔", aw.log
            if 'tool_use_failed' in text or 'output_parse_failed' in text:
                bad_replies += 1
                if bad_replies <= 2:
                    aw.log.append("ℹ️ ماڈل کا جواب غلط شکل میں تھا؛ اسے دوبارہ درست شکل میں کہا جا رہا ہے۔")
                    messages.append({"role": "user", "content":
                                     "پچھلا جواب غلط شکل میں تھا۔ اب صرف ایک درست ٹول کال کریں، "
                                     "یا کام مکمل ہو تو finish ٹول سے خلاصہ دیں۔ وضاحتی متن نہ لکھیں۔"})
                    continue
                aw.log.append("ℹ️ ماڈل بار بار غلط جواب دے رہا ہے؛ جو کام ہو چکا وہ محفوظ ہے۔")
                return _salvage_summary(e), aw.log
            raise
        msg = resp.choices[0].message
        tool_calls = getattr(msg, 'tool_calls', None) or []
        if not tool_calls:
            return (_clean_summary(msg.content) or "کوئی جواب نہیں ملا"), aw.log
=======
    model = model or XAI_DEFAULT_MODEL
    aw = AgentWorkbook(wb)
    messages = [{"role": "system", "content": AGENT_SYSTEM},
                {"role": "user", "content": instruction}]
    for _ in range(max_steps):
        resp = client.chat.completions.create(model=model, messages=messages, tools=AGENT_TOOLS)
        msg = resp.choices[0].message
        tool_calls = getattr(msg, 'tool_calls', None) or []
        if not tool_calls:
            return (msg.content or "کوئی جواب نہیں ملا"), aw.log
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
        messages.append({"role": "assistant", "content": msg.content or "",
                         "tool_calls": [{"id": tc.id, "type": "function",
                                         "function": {"name": tc.function.name,
                                                      "arguments": tc.function.arguments}}
                                        for tc in tool_calls]})
        for tc in tool_calls:
            name = tc.function.name
            try:
                args = _json.loads(tc.function.arguments or '{}')
            except ValueError:
                args = {}
            if name == 'finish':
<<<<<<< HEAD
                return _clean_summary(args.get('summary', '')), aw.log
=======
                return args.get('summary', ''), aw.log
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            try:
                result = aw.call(name, args)
                aw.log.append(f"✅ {name} {args} → {result if isinstance(result, str) else 'ٹھیک'}")
            except Exception as e:
                result = {"error": str(e)}
                aw.log.append(f"❌ {name} {args} → {e}")
            messages.append({"role": "tool", "tool_call_id": tc.id,
<<<<<<< HEAD
                             "content": _json.dumps(result, ensure_ascii=False, default=str)[:6000]})
    return "حد پوری ہو گئی، کام مکمل نہیں ہوا", aw.log
# ==========================================
# 6. را ڈیٹا کا ماخذ: txt/csv/xlsx/pdf/docx/تصویر → سطریں (موڈ 2 کے لیے)
# ==========================================
_RAW_TYPES = ["txt", "csv", "xlsx", "pdf", "docx", "png", "jpg", "jpeg", "bmp", "tif", "tiff"]
_IMG_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")


def _cells_to_line(cells):
    return "\t".join("" if c is None else str(c).replace("\n", " ").strip() for c in cells)


def _ocr_line_to_cells(line):
    # OCR میں کالموں کے درمیان 2 یا زیادہ خالی جگہیں یا ٹیب ہوں تو انہیں الگ خانے مانیں
    return [p.strip() for p in re.split(r"\t|\s{2,}", line.strip()) if p.strip()]


def extract_raw_lines(name, data):
    """
    Returns: (lines, note)
    lines: موجودہ ڈیٹا پارسر کے لیے سطریں (ٹیب یا سادہ متن)۔
    note: صارف کو دکھانے کے لیے ایک لائن کا خلاصہ۔
    خرابی پر ValueError اٹھتا ہے جس کا پیغام اردو میں ہوتا ہے۔
    """
    low = str(name).lower()
    if low.endswith((".txt", ".csv")):
        lines = _read_text(data).splitlines()
        return lines, f"متنی فائل: {len(lines)} لائنیں"

    if low.endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.worksheets[0]
        lines = [_cells_to_line(r) for r in ws.iter_rows(values_only=True) if any(v not in (None, "") for v in r)]
        return lines, f"ایکسل فائل (شیٹ '{ws.title}'): {len(lines)} قطاریں"

    if low.endswith(".pdf"):
        try:
            import pdfplumber
        except ImportError:
            raise ValueError("PDF پڑھنے کے لیے pdfplumber انسٹال ہونا چاہیے (requirements.txt میں لکھیں)۔")
        lines, tables = [], 0
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                found = page.extract_tables() or []
                if found:
                    for tbl in found:
                        tables += 1
                        for row in tbl:
                            if row and any(c not in (None, "") for c in row):
                                lines.append(_cells_to_line(row))
                else:
                    for t in (page.extract_text() or "").splitlines():
                        if t.strip():
                            lines.append(t.strip())
        if not lines:
            raise ValueError("PDF میں کوئی متن یا ٹیبل نہیں ملا۔ اگر یہ سکین شدہ تصویر ہے تو اسے تصویر (PNG/JPG) کے طور پر اپلوڈ کریں۔")
        return lines, f"PDF: {tables} ٹیبل اور {len(lines)} قطاریں"

    if low.endswith(".docx"):
        try:
            import docx as _docx
            from docx.table import Table as _Table
            from docx.text.paragraph import Paragraph as _Para
        except ImportError:
            raise ValueError("Word فائل پڑھنے کے لیے python-docx انسٹال ہونا چاہیے (requirements.txt میں لکھیں)۔")
        doc = _docx.Document(io.BytesIO(data))
        lines, tables = [], 0
        for child in doc.element.body.iterchildren():
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                t = _Para(child, doc).text.strip()
                if t:
                    lines.append(t)
            elif tag == "tbl":
                tables += 1
                for row in _Table(child, doc).rows:
                    cells = [c.text for c in row.cells]
                    if any(c.strip() for c in cells):
                        lines.append(_cells_to_line(cells))
        if not lines:
            raise ValueError("Word فائل میں کوئی ڈیٹا نہیں ملا۔")
        return lines, f"Word فائل: {tables} ٹیبل اور {len(lines)} قطاریں"

    if low.endswith(_IMG_EXTS):
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            raise ValueError("تصویر سے ڈیٹا کے لیے pytesseract اور Pillow انسٹال ہونے چاہئیں۔")
        try:
            img = Image.open(io.BytesIO(data))
            text = pytesseract.image_to_string(img, config="--psm 6")
        except pytesseract.TesseractNotFoundError:
            raise ValueError("تصویر پڑھنے کے لیے Tesseract OCR سسٹم پر انسٹال ہونا چاہیے (Streamlit Cloud پر packages.txt میں tesseract-ocr)۔")
        lines = []
        for t in text.splitlines():
            if t.strip():
                cells = _ocr_line_to_cells(t)
                lines.append("\t".join(cells) if len(cells) > 1 else t.strip())
        if not lines:
            raise ValueError("تصویر میں کوئی متن نہیں ملا؛ زیادہ واضح تصویر اپلوڈ کریں۔")
        return lines, f"تصویر (OCR): {len(lines)} قطاریں۔ نتیجہ ضرور دیکھیں، OCR غلطی کر سکتا ہے۔"

    raise ValueError("یہ فائل کی قسم سپورٹ نہیں ہے۔")


def template_xlsx_bytes(name, data):
    """
    ٹیمپلیٹ کو xlsx بائٹس میں بدلیں: xlsx جوں کی توں، ورنہ PDF/Word/تصویر کے ٹیبل سے نئی شیٹ۔
    """
    if str(name).lower().endswith(".xlsx"):
        return data
    lines, _ = extract_raw_lines(name, data)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"
    num_re = re.compile(r"^-?\d+(\.\d+)?$")
    for r, line in enumerate(lines, 1):
        cells = line.split("\t") if "\t" in line else re.split(r"\s{2,}", line.strip())
        for c, v in enumerate(cells, 1):
            v = v.strip()
            if not v:
                continue
            if num_re.match(v):
                ws.cell(r, c, float(v) if "." in v else int(v))
            else:
                ws.cell(r, c, v)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ==========================================
# 7. نتیجہ PDF میں (xlsx → PDF، اردو/عربی رسم الخط کے ساتھ)
# ==========================================
import os as _os

_PDF_FONT_CANDIDATES = [
    _os.path.join(_os.path.dirname(_os.path.abspath(globals().get("__file__", "app.py"))), "fonts", "NotoNaskhArabic-Regular.ttf"),
    "fonts/NotoNaskhArabic-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
_ARABIC_RE = re.compile(r"[؀-ۿ]")


def _pdf_font_name():
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return "Helvetica"
    if "UniFont" in pdfmetrics.getRegisteredFontNames():
        return "UniFont"
    for p in _PDF_FONT_CANDIDATES:
        if _os.path.exists(p):
            pdfmetrics.registerFont(TTFont("UniFont", p))
            return "UniFont"
    return "Helvetica"


def _pdf_shape(text):
    """عربی/اردو حروف کو جوڑ کر دائیں سے بائیں ترتیب دیں، اگر لائبریری دستیاب ہو۔"""
    text = str(text)
    if not _ARABIC_RE.search(text):
        return text
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except ImportError:
        return text


_LATIN_CANDIDATES = [
    _os.path.join(_os.path.dirname(_os.path.abspath(globals().get("__file__", "app.py"))), "fonts", "DejaVuSans.ttf"),
    "fonts/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]


def _latin_font_name():
    """انگریزی حروف اور ہندسوں کے لیے فونٹ (Noto Naskh میں لاطینی ہندسے نہیں ہوتے)۔"""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return "Helvetica"
    if "LatinFont" in pdfmetrics.getRegisteredFontNames():
        return "LatinFont"
    for p in _LATIN_CANDIDATES:
        if _os.path.exists(p):
            pdfmetrics.registerFont(TTFont("LatinFont", p))
            return "LatinFont"
    return "Helvetica"


def _latin_bold_name():
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return _latin_font_name()
    if "LatinBold" in pdfmetrics.getRegisteredFontNames():
        return "LatinBold"
    for p in [x.replace("DejaVuSans.ttf", "DejaVuSans-Bold.ttf") for x in _LATIN_CANDIDATES]:
        if _os.path.exists(p):
            pdfmetrics.registerFont(TTFont("LatinBold", p))
            return "LatinBold"
    return _latin_font_name()


def _pdf_markup(text, bold=False):
    """
    متن کو Paragraph کے لیے تیار کرتا ہے: اردو/عربی حصے شیپ کر کے اردو فونٹ میں،
    باقی حصے (ہندسے، انگریزی) لاطینی فونٹ میں (bold ہو تو بولڈ)۔
    """
    esc = lambda t: t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    latin = _latin_bold_name() if bold else _latin_font_name()
    out = []
    for tok in re.findall(r"[\u0600-\u06FF\u200c\u200d\s]+|[^\u0600-\u06FF\u200c\u200d\s]+", text):
        if _ARABIC_RE.search(tok):
            out.append(esc(_pdf_shape(tok)))
        elif tok.strip():
            out.append(f'<font name="{latin}">{esc(tok)}</font>')
        else:
            out.append(esc(tok))
    return "".join(out)


def _pdf_cell_text(c):
    v = c.value
    if v is None:
        return ""
    fmt = c.number_format or ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if "%" in fmt:
            pct = v * 100
            return f"{pct:.0f}%" if abs(pct - round(pct)) < 1e-9 else f"{pct:.1f}%"
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return f"{v:g}" if isinstance(v, float) else str(v)
    if hasattr(v, "strftime"):
        return v.strftime("%d-%m-%Y")
    return str(v).replace("\n", " ").strip()


def xlsx_to_pdf_bytes(xlsx_bytes, title="Result"):
    """
    xlsx کی پہلی شیٹ کو PDF میں بدلیں، فارمیٹنگ سمیت:
    خانوں کے رنگ (fill)، بولڈ، سیدھ، اصل کالم چوڑائی اور قطار کی اونچائی،
    merged خانے (مثلاً عنوان)، اور گرڈ لائنیں۔ ہیڈر قطاریں ہر صفحے پر دہرائی جاتی ہیں۔
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A3, A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.worksheets[0]
    ncol = max(ws.max_column, 1)
    nrow = max(ws.max_row, 1)

    # کالم چوڑائی: Excel کی اکائی → پوائنٹ (تقریباً 5.6 پوائنٹ فی اکائی)
    default_w = 8.43
    col_pts = []
    for j in range(1, ncol + 1):
        dim = ws.column_dimensions.get(openpyxl.utils.get_column_letter(j))
        w = dim.width if (dim is not None and dim.width) else default_w
        col_pts.append(max(w, 2) * 5.6 + 4)
    # ایک لفظ والے خانے (مثلاً تاریخ 22/Nov/2026) ایک ہی سطر میں آئیں: کالم اتنا چوڑا
    for r_ in range(1, nrow + 1):
        for c_ in range(1, ncol + 1):
            v_ = ws.cell(r_, c_).value
            if v_ in (None, "") or (r_, c_) == (None, None):
                continue
            t_ = _pdf_cell_text(ws.cell(r_, c_))
            if t_ and " " not in t_ and len(t_) > 3:
                col_pts[c_ - 1] = max(col_pts[c_ - 1], len(t_) * 3.3 + 6)
    total = sum(col_pts)
    a4_land = landscape(A4)
    a3_land = landscape(A3)
    if total > a3_land[0] - 30:
        pagesize = a3_land
    elif total > a4_land[0] - 30:
        pagesize = a3_land if ncol > 14 else a4_land
    else:
        pagesize = a4_land
    avail = pagesize[0] - 30
    fit = 1.0
    if total > avail:
        fit = avail / total
        col_pts = [w * fit for w in col_pts]

    # merged خانے
    merged_tl, merged_skip, spans = {}, set(), []
    for rng in ws.merged_cells.ranges:
        r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if r1 > nrow or c1 > ncol:
            continue
        r2, c2 = min(r2, nrow), min(c2, ncol)
        spans.append((c1 - 1, r1 - 1, c2 - 1, r2 - 1))
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (r1, c1):
                    merged_skip.add((rr, cc))

    # ہیڈر قطاریں: اوپر سے وہ قطاریں جن میں رنگ/عنوان ہو یا آدھے سے زیادہ خانے بھرے ہوں (زیادہ سے زیادہ 4)
    def _row_filled(r):
        cnt = 0
        for c in ws[r][:ncol]:
            if c.value not in (None, ""):
                cnt += 1
        return cnt

    header_rows = 0
    for r in range(1, nrow + 1):
        if header_rows >= 4:
            break
        has_fill = any(c.fill is not None and c.fill.fill_type == "solid" for c in ws[r][:ncol])
        if has_fill or _row_filled(r) * 2 >= ncol or any((r - 1, c) for c in []):
            header_rows += 1
        else:
            break

    base_style = ParagraphStyle("cell", fontName=_latin_font_name(), fontSize=7, leading=9)
    style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#808080")),
    ]
    data, row_heights = [], []
    for r in range(1, nrow + 1):
        row_cells = []
        h = ws.row_dimensions[r].height if r in ws.row_dimensions else None
        row_heights.append(h if h else None)
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            if (r, c) in merged_skip:
                row_cells.append("")
                continue
            txt = _pdf_cell_text(cell)
            bold = bool(cell.font and cell.font.b)
            size = 6
            try:
                if cell.font and cell.font.sz:
                    size = min(max(float(cell.font.sz) * 0.62 * min(1.0, fit * 1.15), 4.5), 9)
            except Exception:
                pass
            al = (cell.alignment.horizontal or "") if cell.alignment else ""
            align_code = {"center": 1, "centre": 1, "right": 2}.get(al, 0)
            sty = ParagraphStyle(f"c{r}_{c}", parent=base_style, fontSize=size, leading=size + 2,
                                 alignment=align_code)
            font_color = colors.black
            try:
                fc = cell.font.color.rgb if cell.font and cell.font.color else None
                if isinstance(fc, str) and len(fc) == 8 and fc[2:] != "000000":
                    font_color = colors.HexColor("#" + fc[2:])
            except Exception:
                pass
            sty.textColor = font_color
            if txt and " " not in txt:
                # ایک لفظ والا خانہ: فونٹ اتنا چھوٹا کہ ایک سطر میں آ جائے
                fit_size = (col_pts[c - 1] - 5) / (len(txt) * 0.66)
                if fit_size < size:
                    size = max(fit_size, 3.5)
                    sty = ParagraphStyle(f"c{r}_{c}", parent=sty, fontSize=size, leading=size + 1.5)
            row_cells.append(Paragraph(_pdf_markup(txt, bold=bold), sty) if txt else "")
            try:
                if cell.fill is not None and cell.fill.fill_type == "solid":
                    rgb = cell.fill.fgColor.rgb
                    if isinstance(rgb, str) and len(rgb) == 8:
                        style.append(("BACKGROUND", (c - 1, r - 1), (c - 1, r - 1), colors.HexColor("#" + rgb[2:])))
            except Exception:
                pass
        data.append(row_cells)
    for (c1, r1, c2, r2) in spans:
        style.append(("SPAN", (c1, r1), (c2, r2)))
    if not data:
        data, row_heights, col_pts = [[""]], [None], [avail]

    tbl = Table(data, colWidths=col_pts, rowHeights=row_heights, repeatRows=max(header_rows, 1))
    tbl.setStyle(TableStyle(style))

    out = io.BytesIO()
    logo_l = _asset_path("logo_left.png")
    logo_r = _asset_path("logo_right.png")
    top = 62 if (logo_l or logo_r) else 18
    doc = SimpleDocTemplate(out, pagesize=pagesize, leftMargin=15, rightMargin=15, topMargin=top, bottomMargin=18,
                            title=title)

    def _draw_logos(canvas, _doc):
        canvas.saveState()
        pw, ph = pagesize
        if logo_l:
            h = 40.0
            w = h * 584.0 / 182.0
            canvas.drawImage(logo_l, 15, ph - 12 - h, width=w, height=h, mask="auto")
        if logo_r:
            h = 30.0
            w = h * 130.0 / 63.0
            canvas.drawImage(logo_r, pw - 15 - w, ph - 10 - h, width=w, height=h, mask="auto")
        canvas.restoreState()

    doc.build([tbl], onFirstPage=_draw_logos, onLaterPages=_draw_logos)
    return out.getvalue()

_ASSET_DIRS = [
    _os.path.join(_os.path.dirname(_os.path.abspath(globals().get("__file__", "app.py"))), "assets"),
    "assets",
]


def _asset_path(name):
    """assets فولڈر میں فائل ہو تو اس کا راستہ، ورنہ None۔"""
    for d in _ASSET_DIRS:
        p = _os.path.join(d, name)
        if _os.path.exists(p):
            return p
    return None


def _pdf_download_button(xlsx_bytes, file_name, label, key):
    """xlsx نتیجے سے PDF بنا کر ڈاؤن لوڈ بٹن دکھائیں؛ خرابی پر صرف پیغام۔"""
    try:
        pdf_bytes = xlsx_to_pdf_bytes(xlsx_bytes, title=file_name.rsplit(".", 1)[0])
    except Exception as e:
        st.error(f"PDF بنانے میں خامی: {e}")
        return
    st.download_button(label, pdf_bytes, file_name, mime="application/pdf", key=key)


def apply_style_template(result_bytes, style_bytes):
    """
    نتیجے کی شیٹ پر فارمیٹ ٹیمپلیٹ کی فارمیٹنگ لگائیں: خانوں کے رنگ، فونٹ، بارڈر، سیدھ،
    merged خانے، کالم چوڑائی اور قطار کی اونچائی۔ خانوں کا ملاپ ایک ہی جگہ (قطار/کالم) سے ہوتا ہے۔
    ڈیٹا (اقدار) نہیں بدلتا۔
    """
    from copy import copy as _cp
    out_wb = openpyxl.load_workbook(io.BytesIO(result_bytes))
    style_wb = openpyxl.load_workbook(io.BytesIO(style_bytes))
    out = out_wb.worksheets[0]
    st_ws = style_wb.worksheets[0]
    max_r = max(out.max_row, st_ws.max_row)
    max_c = max(out.max_column, st_ws.max_column)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            sc = st_ws.cell(r, c)
            if not sc.has_style:
                continue
            d = out.cell(r, c)
            d.font = _cp(sc.font)
            d.fill = _cp(sc.fill)
            d.border = _cp(sc.border)
            d.alignment = _cp(sc.alignment)
            if sc.number_format and sc.number_format != "General":
                d.number_format = sc.number_format
    for rng in st_ws.merged_cells.ranges:
        try:
            out.merge_cells(str(rng))
        except Exception:
            pass
    for key, dim in st_ws.column_dimensions.items():
        if dim.width:
            out.column_dimensions[key].width = dim.width
    for key, dim in st_ws.row_dimensions.items():
        if dim.height:
            out.row_dimensions[key].height = dim.height
    buf = io.BytesIO()
    out_wb.save(buf)
    return buf.getvalue()
=======
                             "content": _json.dumps(result, ensure_ascii=False, default=str)[:20000]})
    return "حد پوری ہو گئی، کام مکمل نہیں ہوا", aw.log
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7


# ==========================================
# مین انٹرفیس (Streamlit)
# ==========================================
st.set_page_config(page_title="ڈیٹا پروسیسنگ ایجنٹ", layout="wide")
st.title("📊 یونیورسل ڈیٹا پروسیسنگ ایجنٹ")
st.write("اپنے کام کا موڈ منتخب کریں:")

app_mode = st.radio(
    "موڈ منتخب کریں:",
    ("🟢 موڈ 1: OMS ماہانہ اٹینڈنس اور ٹوٹل",
     "🔵 موڈ 2: یونیورسل را ڈیٹا (فول پروف)",
     "🟣 موڈ 3: ایکسل ورک بینچ",
     "🟠 موڈ 4: AI ایجنٹ (Grok API)")
)

st.markdown("---")

if app_mode == "🟢 موڈ 1: OMS ماہانہ اٹینڈنس اور ٹوٹل":
    st.subheader("🟢 OMS اٹینڈنس اور ٹوٹل فارمولے")

    uploaded_raw = st.file_uploader("واٹس ایپ را ڈیٹا (.txt)", type=["txt", "csv"], key="att_raw")
    uploaded_template = st.file_uploader("OMS اٹینڈنس ٹیمپلیٹ (.xlsx)", type=["xlsx"], key="att_temp")

    if st.button("اٹینڈنس اور ٹوٹل پراسیس کریں"):
        if uploaded_raw and uploaded_template:
            raw_lines = uploaded_raw.getvalue().decode("utf-8-sig", errors="ignore").splitlines()
            res_file, msg = process_oms_attendance_mode(raw_lines, uploaded_template)
            if res_file:
                st.success("اٹینڈنس شیٹ اور ٹوٹل کامیابی سے تیار ہو گئے ہیں!")
                st.download_button(
                    "📥 فائنل اٹینڈنس شیٹ ڈاؤن لوڈ کریں", res_file, "OMS_Final_Attendance.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(msg)
        else:
            st.warning("براہ کرم دونوں فائلیں اپ لوڈ کریں۔")

elif app_mode == "🔵 موڈ 2: یونیورسل را ڈیٹا (فول پروف)":
    st.subheader("🔵 یونیورسل ڈیٹا پروسیسنگ")
    proc_kind = st.radio(
        "پروسیسنگ کا طریقہ:",
        ["عام ڈیٹا (کالم خود چنیں)", "ایکٹیویٹی پروگریس (تاریخ کے کالم میں %)"],
        key="proc_kind",
    )
    if proc_kind.startswith("ایکٹیویٹی"):
        render_progress_ui()
        st.stop()

    st.info("ڈیٹا اور ٹیمپلیٹ اپلوڈ کریں، پھر ہر ٹیمپلیٹ کالم کے سامنے ڈیٹا کا کالم چنیں۔")

    enable_total = st.checkbox("ٹیمپلیٹ میں Grand Total نہ ہو تو بھی نیچے ٹوٹل لگائیں")

<<<<<<< HEAD
    uploaded_raw_gen = st.file_uploader("را ڈیٹا فائل (txt, csv, xlsx, pdf, docx, تصویر PNG/JPG)", type=_RAW_TYPES, key="gen_raw")
    uploaded_template_gen = st.file_uploader("ایکسل ٹیمپلیٹ (.xlsx)", type=["xlsx"], key="gen_temp")

    if uploaded_raw_gen and uploaded_template_gen:
        template_bytes = uploaded_template_gen.getvalue()
        try:
            raw_lines, raw_note = extract_raw_lines(uploaded_raw_gen.name, uploaded_raw_gen.getvalue())
            st.info(raw_note)
        except ValueError as e:
            st.error(str(e))
            raw_lines = []
=======
    uploaded_raw_gen = st.file_uploader("را ڈیٹا فائل (.txt یا .csv)", type=["txt", "csv"], key="gen_raw")
    uploaded_template_gen = st.file_uploader("ایکسل ٹیمپلیٹ (.xlsx)", type=["xlsx"], key="gen_temp")

    if uploaded_raw_gen and uploaded_template_gen:
        raw_lines = uploaded_raw_gen.getvalue().decode("utf-8-sig", errors="ignore").splitlines()
        template_bytes = uploaded_template_gen.getvalue()
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
        mode, rows, text_count = prepare_rows(raw_lines)
        tcols = load_template_info(template_bytes)

        if not rows or not tcols:
            st.error("ڈیٹا یا ٹیمپلیٹ میں کوئی قابلِ استعمال چیز نہیں ملی۔")
        else:
            width = max(len(r) for r in rows)
            labels = slot_labels(rows)

            st.markdown("### ڈیٹا کا پیش نظارہ (پہلی 5 قطاریں)")
            st.table([
                {f"کالم {k + 1}": (r[k] if k < len(r) else "") for k in range(width)}
                for r in rows[:5]
            ])

            st.markdown("### ٹیمپلیٹ کالم ← ڈیٹا کالم")
            defaults = default_mapping(tcols, width, mode, text_count)
            options = ["— خالی چھوڑیں —"] + labels + ["فارمولا: Qty × Price"]
            mapping = {}
            for tc in tcols:
                d = defaults.get(tc['col'])
                if d is None:
                    idx = 0
                elif d == FORMULA:
                    idx = len(options) - 1
                else:
                    idx = d + 1
                choice = st.selectbox(tc['header'], options, index=idx, key=f"map_{tc['col']}")
                if choice == options[0]:
                    mapping[tc['col']] = None
                elif choice == options[-1]:
                    mapping[tc['col']] = FORMULA
                else:
                    mapping[tc['col']] = options.index(choice) - 1

            if st.button("ڈیٹا پراسیس کریں"):
                res_file, msg = process_with_mapping(
                    template_bytes, rows, mapping, need_total=enable_total
                )
                if res_file:
                    st.success(msg)
                    st.download_button(
                        "📥 فائنل پروسیس شدہ فائل ڈاؤن لوڈ کریں", res_file, "Processed_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
<<<<<<< HEAD
                    _pdf_download_button(res_file.getvalue(), "Processed_Data.pdf", "📄 نتیجہ PDF میں ڈاؤن لوڈ کریں", "mode2_pdf")
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
                else:
                    st.error(msg)
    else:
        st.warning("براہ کرم دونوں فائلیں اپ لوڈ کریں۔")

elif app_mode == "🟣 موڈ 3: ایکسل ورک بینچ":
    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.subheader("🟣 ایکسل ایجنٹ")
    st.info("فائل بنائیں، بدلیں، فلٹر، سورٹ، خالی خانے بھریں، ری کنسائل، کنسالیڈیٹ اور مرج کریں۔ کمانڈز نیچے لکھیں، ہر لائن ایک کام ہے۔")

    main_up = st.file_uploader("موجودہ ایکسل فائل (اختیاری، نہ ہو تو نیچے ڈیٹا سے نئی فائل بنے گی)",
                               type=["xlsx"], key="ag_main")
    base_text = st.text_area("نئی فائل کے لیے ڈیٹا (صرف اس وقت جب اوپر فائل نہ ہو)", height=140, key="ag_base")
    sheet_name = st.text_input("نئی شیٹ کا نام", "Data", key="ag_sheet")
    extra_ups = st.file_uploader("دوسری فائلیں (fill، reconcile، merge کے لیے، ایک سے زیادہ ہو سکتی ہیں)",
                                 type=["xlsx"], accept_multiple_files=True, key="ag_extra")

    if main_up is not None:
        wb_work = openpyxl.load_workbook(io.BytesIO(main_up.getvalue()))
        st.caption(f"فائل لوڈ ہو گئی۔ پہلی شیٹ: {wb_work.worksheets[0].title}")
    elif base_text.strip():
        _, rows_base = _read_rows(base_text.splitlines())
        headers_base, data_base = _split_header(rows_base)
        wb_work = create_workbook(headers_base, data_base, sheet_name)
        st.caption(f"نئی فائل ڈیٹا سے بنی: {len(data_base)} قطاریں۔")
    else:
        wb_work = None

    if wb_work is not None:
        ws0 = wb_work.worksheets[0]
        h0, rows0 = _table(ws0)
        st.markdown("### موجودہ ڈیٹا (پہلی 10 قطاریں)")
        st.table(rows0[:10])

    cmds = st.text_area("کمانڈز (ہر لائن ایک کام)", height=200, key="ag_cmds",
                        placeholder="insert row at 3 : Item=Mug, Qty=3\nsort Price desc\nfilter Region = North")
    with st.expander("کمانڈز کی فہرست"):
        st.code(AGENT_HELP)

    if st.button("کمانڈز چلائیں"):
        if wb_work is None:
            st.warning("پہلے فائل اپلوڈ کریں یا نیچے ڈیٹا دیں۔")
        elif not cmds.strip():
            st.warning("کم از کم ایک کمانڈ لکھیں۔")
        else:
            extras = [openpyxl.load_workbook(io.BytesIO(f.getvalue())) for f in (extra_ups or [])]
            logs = run_commands(cmds, wb_work, extras)
            for line in logs:
                if line.startswith("✅"):
                    st.success(line)
                else:
                    st.error(line)
            st.download_button("📥 نتیجہ فائل ڈاؤن لوڈ کریں", save_bytes(wb_work),
                               "Agent_Output.xlsx", mime=XLSX_MIME)
<<<<<<< HEAD
            _pdf_download_button(save_bytes(wb_work).getvalue(), "Agent_Output.pdf", "📄 نتیجہ PDF میں ڈاؤن لوڈ کریں", "mode3_pdf")
=======
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7

else:
    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.subheader("🟠 AI ایجنٹ (Grok API)")
<<<<<<< HEAD
    st.info("ہدایت اردو یا انگریزی میں لکھیں۔ ایجنٹ پہلے فائل کی ساخت دیکھتا ہے، پھر کام کرتا ہے۔ Groq API key ماحول کے متغیر GROQ_API_KEY میں ہونی چاہیے۔")
    if not _get_groq_key():
        st.warning("GROQ_API_KEY سیٹ نہیں ہے (.streamlit/secrets.toml میں یا ماحول میں)۔ ایجنٹ شروع نہیں ہوگا۔")

    up_ai = st.file_uploader("ایکسل فائلیں (.xlsx) — ایک یا ایک سے زیادہ، مثلاً ڈیٹا اور دوسری فائل ری کنسائل کے لیے", type=["xlsx"], accept_multiple_files=True, key="ai_file")
    ai_files = list(up_ai or [])
    raw_ai = st.text_area("یا کچا ڈیٹا یہاں پیسٹ کریں (اختیاری؛ فائل کے ساتھ یا اکیلے)", height=140, key="ai_raw",
                          placeholder="ہیڈر کی پہلی قطار، پھر ڈیٹا؛ کالم ٹیب یا کومے سے الگ ہوں")
    raw_obj = None
    if raw_ai.strip():
        _, rows_raw = _read_rows(raw_ai.splitlines())
        h_raw, d_raw = _split_header(rows_raw)
        if d_raw or h_raw:
            _rb = io.BytesIO()
            create_workbook(h_raw, d_raw, "Data").save(_rb)
            class _RawUpload:
                name = "pasted.xlsx"
                def getvalue(self_inner):
                    return _rb.getvalue()
            raw_obj = _RawUpload()
    all_inputs = ai_files + ([raw_obj] if raw_obj is not None else [])
    instr = st.text_area("ہدایت", height=120, key="ai_instr",
                         placeholder="مثلاً: 20-Nov-26 کے کالم میں ہر ایکٹیویٹی کا فیصد لکھو")
    if st.button("ایجنٹ چلائیں") and all_inputs and instr.strip():
        try:
            wb_ai = _combine_workbooks(all_inputs)
            src_titles = set(wb_ai.sheetnames)
            with st.spinner("ایجنٹ کام کر رہا ہے..."):
                summary, log = run_agent(instr, wb_ai)
            wb_ai, kept_src = _keep_one_output_sheet(wb_ai, src_titles)
            if kept_src:
                st.info("فارمولے ان شیٹوں کا حوالہ دیتے ہیں اس لیے یہ شیٹیں بھی رکھی گئی ہیں: " + "، ".join(kept_src))
=======
    st.info("ہدایت اردو یا انگریزی میں لکھیں۔ ایجنٹ پہلے فائل کی ساخت دیکھتا ہے، پھر کام کرتا ہے۔ API key ماحول کے متغیر XAI_API_KEY میں ہونی چاہیے۔")
    if not _os.environ.get("XAI_API_KEY", "").strip():
        st.warning("XAI_API_KEY سیٹ نہیں ہے۔ ایجنٹ شروع نہیں ہوگا۔")

    up_ai = st.file_uploader("ایکسل فائل (.xlsx)", type=["xlsx"], key="ai_file")
    instr = st.text_area("ہدایت", height=120, key="ai_instr",
                         placeholder="مثلاً: 20-Nov-26 کے کالم میں ہر ایکٹیویٹی کا فیصد لکھو")
    if st.button("ایجنٹ چلائیں") and up_ai is not None and instr.strip():
        try:
            wb_ai = openpyxl.load_workbook(io.BytesIO(up_ai.getvalue()))
            with st.spinner("ایجنٹ کام کر رہا ہے..."):
                summary, log = run_agent(instr, wb_ai)
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
            st.markdown("### ایجنٹ کی کارروائی")
            for line in log:
                st.write(line)
            st.success(summary)
            st.download_button("📥 نتیجہ فائل ڈاؤن لوڈ کریں", save_bytes(wb_ai),
                               "AI_Output.xlsx", mime=XLSX_MIME)
<<<<<<< HEAD
            _pdf_download_button(save_bytes(wb_ai).getvalue(), "AI_Output.pdf", "📄 نتیجہ PDF میں ڈاؤن لوڈ کریں", "mode4_pdf")
        except Exception as e:
            st.error(f"خامی: {e}")
    elif not all_inputs:
        st.info("پہلے فائل اپلوڈ کریں یا کچا ڈیٹا پیسٹ کریں۔")
    elif not instr.strip():
        st.warning("ہدایت لکھیں، پھر ایجنٹ چلائیں۔")
=======
        except Exception as e:
            st.error(f"خامی: {e}")
    elif up_ai is None:
        st.info("پہلے فائل اپلوڈ کریں۔")
>>>>>>> 1a3888f059dc75f376856c8a9b2b260bc8ecf9d7
