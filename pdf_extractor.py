"""
Extract data from PXK (Phiếu Xuất Kho) PDF files.

Two known table layouts exist (columns differ between PDF variants):
  Layout A — 9 cols:  STT@0, Name@1, Code@3, DVT@4, SL@5, Giá@6-7, TT@8
  Layout B — 18 cols: STT@2, Name@4, Code@7, DVT@8, SL@10, Giá@12, TT@15

Column positions are detected dynamically from the header row so both
(and any future) variants are handled automatically.
"""

import re
import unicodedata
import pdfplumber
from pathlib import Path


# ── Number parsing ─────────────────────────────────────────────────────────────

def parse_number(s: str) -> float:
    """Parse Vietnamese number format.

    Vietnamese uses '.' as thousands separator and ',' as decimal separator.
    Examples:
        '928.590'   → 928590.0
        '30.953'    → 30953.0
        '232,59'    → 232.59
        '1.150.164' → 1150164.0
    """
    if not s:
        return 0.0
    s = str(s).strip().replace(' ', '')
    if not s or s in ('-', '—'):
        return 0.0

    if ',' in s:
        # Comma = decimal separator, dots = thousands
        s = s.replace('.', '').replace(',', '.')
    else:
        parts = s.split('.')
        if len(parts) > 2:
            # Multiple dots → all are thousands separators
            s = s.replace('.', '')
        elif len(parts) == 2 and len(parts[1]) == 3:
            # Single dot + 3 trailing digits → thousands separator
            s = s.replace('.', '')
        # else: genuine decimal point, keep as-is

    try:
        return float(s)
    except ValueError:
        return 0.0


# ── Code cell parsing ──────────────────────────────────────────────────────────

def split_codes(cell: str) -> list:
    """Split merged code cell: 'DC97-\n22471T\nDC97-\n22471T'
    → ['DC97-22471T', 'DC97-22471T'].
    Also handles cases where codes are not split with hyphens.
    """
    if not cell or not cell.strip():
        return []

    lines = [ln.strip() for ln in cell.split('\n') if ln.strip()]
    codes = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue
        # Check if line ends with hyphen and next line exists (split code pattern)
        if line.endswith('-') and i + 1 < len(lines) and lines[i + 1]:
            codes.append(line + lines[i + 1].strip())
            i += 2
        # Also handle patterns like "DC97-\n22471T" where parts might be in separate lines
        elif len(line) <= 10 and i + 1 < len(lines) and lines[i + 1] and not lines[i + 1].startswith('DC'):
            # This might be a partial code (e.g., "DC97-" followed by "22471T")
            combined = line + lines[i + 1].strip()
            # Validate it looks like a product code
            if re.match(r'^[A-Z0-9\-]+$', combined):
                codes.append(combined)
                i += 2
            else:
                codes.append(line)
                i += 1
        else:
            codes.append(line)
            i += 1

    return codes


def split_names(cell: str, n: int) -> list:
    """Split a multi-line product-name cell into exactly n names.
    Each name typically spans 2 lines in the cell.
    """
    if not cell or not cell.strip():
        return [''] * n

    lines = [ln for ln in cell.split('\n') if ln.strip()]
    names = []
    i = 0
    while i < len(lines) and len(names) < n:
        name = lines[i].strip()
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            # Check if next line should be appended (lowercase or starts with special chars)
            if next_line and (not next_line[0].isupper() or next_line[0] in '(-['):
                name = name + ' ' + next_line
                i += 2
            else:
                i += 1
        else:
            i += 1
        names.append(name)
    while len(names) < n:
        names.append('')
    return names[:n]


# ── Column-layout detection ────────────────────────────────────────────────────

def _detect_columns(table: list):
    """Scan table rows for the header row and return a column-index map.

    Returns None if no recognisable header is found.
    Returned dict keys: 'stt', 'name', 'code', 'dvt', 'sl', 'gia', 'tt'
    """
    for row in table:
        if not row:
            continue
        cells = [str(c or '') for c in row]
        full = ' '.join(cells)

        # Must contain both 'STT' and 'Đơn giá' (or 'Unit price')
        if 'STT' not in full or ('Đơn giá' not in full and 'Unit price' not in full):
            continue

        # Find column indices by keyword search
        def find_col(*keywords):
            for ki, keyword in enumerate(keywords):
                for ci, cell in enumerate(cells):
                    if keyword in cell:
                        return ci
            return -1

        stt_col  = find_col('STT')
        name_col = find_col('Tên hàng hóa', 'Description')
        code_col = find_col('Mã số', 'HHDV', 'Code')
        dvt_col  = find_col('Đơn vị\ntính', 'Unit)\n', '(Unit)')
        sl_col   = find_col('Số lượng', 'Quantity')
        gia_col  = find_col('Đơn giá', 'Unit price')
        tt_col   = find_col('Thành tiền', 'Amount')

        if stt_col >= 0 and code_col >= 0 and sl_col >= 0:
            return {
                'stt':  stt_col,
                'name': name_col if name_col >= 0 else stt_col + 1,
                'code': code_col,
                'dvt':  dvt_col  if dvt_col  >= 0 else code_col + 1,
                'sl':   sl_col,
                'gia':  gia_col,
                'tt':   tt_col,
            }
    return None


# ── Main extractor ─────────────────────────────────────────────────────────────

def extract_pxk(pdf_path: str) -> dict:
    """Extract all line items from a single PXK PDF.

    Returns:
        {
            'so_phieu': str,
            'ngay': str,
            'items': [{'ten_hang', 'ma_hang', 'dvt',
                        'so_luong', 'don_gia', 'thanh_tien'}],
            'file_name': str,
            'error': str | None,
        }
    """
    file_name = Path(pdf_path).name
    result = {
        'so_phieu': '',
        'ngay': '',
        'ma_cqt': '',
        'do_no': '',
        'ly_do': '',
        'phuong_tien': '',
        'items': [],
        'file_name': file_name,
        'error': None,
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = unicodedata.normalize('NFC', page.extract_text() or '')

                # PXK number: "Số (No) : 00001174"
                if not result['so_phieu']:
                    m = re.search(r'No\s*\)\s*:\s*0*(\d+)', text)
                    if m:
                        result['so_phieu'] = m.group(1)
                    else:
                        m2 = re.search(r'C26NAA_(\d+)_', file_name, re.IGNORECASE)
                        if m2:
                            result['so_phieu'] = m2.group(1)

                # Date: "Ngày (Date) 02 tháng (month) 03 năm (year) 2026"
                if not result['ngay']:
                    m = re.search(
                        r'Date\)\s+(\d{1,2})\s+tháng.*?(\d{1,2}).*?năm.*?(\d{4})',
                        text, re.IGNORECASE)
                    if m:
                        result['ngay'] = (f"{int(m.group(1)):02d}/"
                                          f"{int(m.group(2)):02d}/{m.group(3)}")
                    else:
                        m2 = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
                        if m2:
                            result['ngay'] = m2.group(1)

                # Mã CQT: "Mã của cơ quan thuế:\n00B641410114204FE09F4B18DC835FD467"
                if not result['ma_cqt']:
                    m = re.search(r'cơ quan thuế[:\s]*\n([A-F0-9]{20,})', text)
                    if m:
                        result['ma_cqt'] = m.group(1).strip()

                # D/O No: "Căn cứ lệnh điều động số (D/O No) : 0093/0094 Ngày..."
                if not result['do_no']:
                    m = re.search(r'D/O No\)\s*:\s*(.+?)\s+Ngày', text)
                    if m:
                        result['do_no'] = m.group(1).strip()

                # Lý do: "Lý do xuất kho (Reason) : Xuất kho bán hàng"
                if not result['ly_do']:
                    m = re.search(r'Reason\)\s*:\s*(.+)', text)
                    if m:
                        result['ly_do'] = m.group(1).strip()

                # Phương tiện: "Phương tiện vận chuyển (Transportation) : Xe tải 60H-30681"
                if not result['phuong_tien']:
                    m = re.search(r'Transportation\)\s*:\s*(.+)', text)
                    if m:
                        result['phuong_tien'] = m.group(1).strip()

                for table in page.extract_tables():
                    cols = _detect_columns(table)
                    if cols:
                        _parse_packed_table(table, cols, result['items'])

    except Exception as exc:
        result['error'] = str(exc)

    return result


# ── Row parser ─────────────────────────────────────────────────────────────────

def _extract_stt_count(cell: str) -> int:
    """Extract the count of STT entries from a cell."""
    if not cell:
        return 0
    stt_lines = []
    for s in cell.split('\n'):
        s_clean = s.strip()
        if s_clean and (s_clean.isdigit() or re.match(r'^\d+\.?$', s_clean)):
            stt_lines.append(s_clean.rstrip('.'))
    return len(stt_lines)

def _cell_value(row: list, col: int) -> str:
    """Safely get cell value."""
    if col < 0 or col >= len(row):
        return ''
    return str(row[col] or '').strip()

def _parse_packed_table(table: list, cols: dict, items: list):
    """Unpack all items from the packed data row using detected column positions.

    Handles multi-row tables where data spans across multiple PDF table rows
    for the same STT group (e.g., PXK 1994 format).
    """
    stt_col  = cols['stt']
    name_col = cols['name']
    code_col = cols['code']
    dvt_col  = cols['dvt']
    sl_col   = cols['sl']
    gia_col  = cols['gia']
    tt_col   = cols['tt']

    # First pass: collect all data rows and merge continuation rows
    raw_rows = []
    for row in table:
        if not row or len(row) <= stt_col:
            continue
        raw_rows.append(row)

    # Merge rows that belong to the same STT group
    merged_rows = []
    current_group = None

    for row in raw_rows:
        stt_cell = _cell_value(row, stt_col)
        stt_count = _extract_stt_count(stt_cell)

        if stt_count > 0:
            # This is a new STT group - save previous if exists
            if current_group:
                merged_rows.append(current_group)
            # Start new group
            current_group = {
                'stt_cell': stt_cell,
                'stt_count': stt_count,
                'name': _cell_value(row, name_col),
                'code': _cell_value(row, code_col),
                'dvt': _cell_value(row, dvt_col),
                'sl': _cell_value(row, sl_col),
                'gia': '',
                'tt': '',
            }
            # Find price and total cells
            if gia_col >= 0:
                current_group['gia'] = _cell_value(row, gia_col)
            if tt_col >= 0:
                current_group['tt'] = _cell_value(row, tt_col)

            # If price is empty, scan forward
            if not current_group['gia']:
                for ci in range(sl_col + 1, len(row)):
                    v = _cell_value(row, ci)
                    if v:
                        current_group['gia'] = v
                        break
            # If total is empty, scan forward
            if not current_group['tt']:
                price_start = gia_col if gia_col >= 0 else sl_col + 1
                for ci in range(price_start + 1, len(row)):
                    v = _cell_value(row, ci)
                    if v and v != current_group['gia']:
                        current_group['tt'] = v
                        break
        else:
            # This is a continuation row - append to current group
            if current_group:
                # Check if this is a footer/summary row
                code_val = _cell_value(row, code_col)
                if code_val and ('tổng cộng' in code_val.lower() or 'total' in code_val.lower()):
                    continue  # Skip footer rows

                # Append code values
                if code_val:
                    current_group['code'] += '\n' + code_val
                # Append DVT values
                dvt_val = _cell_value(row, dvt_col)
                if dvt_val:
                    current_group['dvt'] += '\n' + dvt_val
                # Append SL values
                sl_val = _cell_value(row, sl_col)
                if sl_val:
                    current_group['sl'] += '\n' + sl_val
                # Append price values
                price_val = ''
                if gia_col >= 0:
                    price_val = _cell_value(row, gia_col)
                if not price_val:
                    for ci in range(sl_col + 1, len(row)):
                        v = _cell_value(row, ci)
                        if v:
                            price_val = v
                            break
                if price_val:
                    current_group['gia'] += '\n' + price_val
                # Note: Don't append totals from continuation rows
                # The first row (with STT) already contains all totals

    # Don't forget the last group
    if current_group:
        merged_rows.append(current_group)

    # Second pass: process merged rows
    for group in merged_rows:
        n = group['stt_count']
        if n == 0:
            continue

        name_cell = group['name']
        code_cell = group['code']
        dvt_cell = group['dvt']
        sl_cell = group['sl']
        price_cell = group['gia']
        total_cell = group['tt']

        codes  = split_codes(code_cell)
        dvts   = [v.strip() for v in dvt_cell.split('\n')    if v.strip()]
        sls    = [v.strip() for v in sl_cell.split('\n')     if v.strip()]
        prices = [v.strip() for v in price_cell.split('\n')  if v.strip()]
        totals = [v.strip() for v in total_cell.split('\n')  if v.strip()]
        names  = split_names(name_cell, n)

        # Ensure all arrays have at least n items by padding with empty values
        while len(codes) < n:
            codes.append('')
        while len(dvts) < n:
            dvts.append('')
        while len(sls) < n:
            sls.append('')
        while len(prices) < n:
            prices.append('')
        while len(totals) < n:
            totals.append('')
        while len(names) < n:
            names.append('')

        for i in range(n):
            items.append({
                'ten_hang':   names[i]  if i < len(names)  else '',
                'ma_hang':    codes[i]  if i < len(codes)  else '',
                'dvt':        dvts[i]   if i < len(dvts)   else '',
                'so_luong':   parse_number(sls[i])    if i < len(sls)    else 0.0,
                'don_gia':    parse_number(prices[i]) if i < len(prices) else 0.0,
                'thanh_tien': parse_number(totals[i]) if i < len(totals) else 0.0,
            })
