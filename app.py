#!/usr/bin/env python3
"""
Univet Order Form Export API — v3 (XML-direct)

Instead of openpyxl (which destroys fonts, borders, images, etc.),
this version manipulates the xlsx XML directly so the output is
byte-for-byte identical to the company template except for filled cells.
"""

import io, os, re, zipfile, datetime, json
import xml.etree.ElementTree as ET
from flask import Flask, request, Response
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ── Locate template ──────────────────────────────────────────────────
_dir = os.path.dirname(os.path.abspath(__file__))
# Try both possible filenames
for _candidate in ['2026_ORDER_FORM__eng__rev0.xlsx',
                   '2026 ORDER FORM (eng) rev0.xlsx']:
    _p = os.path.join(_dir, _candidate)
    if os.path.isfile(_p):
        TEMPLATE_PATH = _p
        break
else:
    raise FileNotFoundError("Template xlsx not found in " + _dir)

# ── Spreadsheet XML namespace ────────────────────────────────────────
NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NSP = '{' + NS + '}'

# Register namespace so ET.tostring doesn't add ns0: prefixes everywhere
ET.register_namespace('', NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
ET.register_namespace('xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
ET.register_namespace('xr6', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision6')
ET.register_namespace('xr10', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision10')


# ── Helpers ──────────────────────────────────────────────────────────

def col_letter(n):
    """1-based column number → letter(s): 1→A, 4→D, 20→T."""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_ref(row, col):
    """(row, col) 1-based → 'D6'."""
    return f'{col_letter(col)}{row}'


def ref_to_row(ref):
    """'D6' → 6."""
    return int(re.sub(r'[A-Z]+', '', ref))


def ref_to_col_num(ref):
    """'D6' → 4."""
    letters = re.sub(r'[0-9]+', '', ref)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


# ── Core: build cell-update map from form data ──────────────────────

def build_cell_map(data):
    """
    Return dict  { 'D6': value, ... }
    where value is:
      str        → written as inlineStr
      int/float  → written as number
      ('formula', '=(2026-E9)')  → written as formula
    """
    cells = {}

    def put(addr, val):
        if val is None or val == '':
            return
        cells[addr] = val

    # ── Customer info ────────────────────────────────────────────────
    put('D6',  data.get('date', ''))
    put('D7',  data.get('fname', ''))
    put('D8',  data.get('lname', ''))

    yob = data.get('yob', '')
    if yob:
        try:
            put('E9', int(yob))
        except (ValueError, TypeError):
            put('E9', yob)
        current_year = datetime.date.today().year
        cells['G9'] = ('formula', f'({current_year}-E9)')

    put('D10', data.get('profession', ''))
    put('D11', data.get('specialty', ''))
    put('D12', data.get('address', ''))
    put('D13', data.get('town', ''))
    put('D14', data.get('country', ''))
    put('D15', data.get('tel', ''))
    put('D16', data.get('email', ''))

    # ── Agent / Sales Representative ─────────────────────────────────
    agent = data.get('agent', '')
    if agent:
        put('E87', agent)

    # ── Optic + WD checkmark ─────────────────────────────────────────
    wd_col = {'300': 6, '350': 7, '400': 8, '450': 9,
              '500': 10, '550': 11, '600': 12, '700': 13}
    optic_row = {
        'Galilean|2.0x': 20, 'Galilean|2.5x': 21, 'Galilean|3.0x': 22,
        'Galilean|3.5x': 23,
        'Prismatic|3.5x': 24, 'Prismatic|4.0x': 25, 'Prismatic|5.0x': 26,
        'Ergo|3.5x': 27, 'Ergo|4.5x': 28, 'Ergo|5.7x': 29,
    }
    optic_key = data.get('optic', '')
    wd_key    = str(data.get('wd', ''))
    row = optic_row.get(optic_key)
    col = wd_col.get(wd_key)
    if row and col:
        cells[cell_ref(row, col)] = '\u221A'

    # ── Frame + color checkmark ──────────────────────────────────────
    frame_row = {
        'Look': 31, 'Cool': 32, 'Techne 2025': 33, 'Techne [Old]': 34,
        'Techne RX [Old]': 35,
        'Techne ASIAN FITTING [OLD]': 36, 'Techne RX ASIAN FITTING [OLD]': 37,
        'Ash 55-17': 38, 'Ash 53-17': 39,
        'ITA': 40, 'ITA - Extended Fit': 41, 'ONE': 42,
    }
    frame_color_cols = {
        'Look':       {'Ruby': 5, 'Emerald': 8},
        'Cool':       {'Ruby': 5, 'Emerald': 8},
        'Techne 2025': {'Black/Green': 5, 'White/Green': 8,
                        'Midnight Blue': 12, 'White/Wisteria': 15},
        'Techne [Old]': {'White/Red': 5, 'Black/Green': 8,
                         'White/Pink': 12, 'Black Edition': 15},
        'Techne RX [Old]': {'White/Red': 5, 'Black/Green': 8, 'White/Pink': 12},
        'Techne ASIAN FITTING [OLD]': {'White/Red': 5, 'Black/Green': 8,
                                        'White/Pink': 12, 'Black Edition': 15},
        'Techne RX ASIAN FITTING [OLD]': {'White/Red': 5, 'Black/Green': 8,
                                           'White/Pink': 12},
        'Ash 55-17': {'Black/Red': 5, 'Grey/Grey': 8, 'Purple/Grey': 12},
        'Ash 53-17': {'Black/Red': 5, 'Grey/Grey': 8, 'Purple/Grey': 12},
        'ITA': {'Midnight Blue': 5, 'Brown Endura': 8,
                'Green Vespa': 12, 'Black Edition': 15},
        'ITA - Extended Fit': {'Black Edition': 15},
        'ONE': {'Desert Sage': 5, 'Black Edition / Limited Edition': 15,
                'Color Kit': 18},
    }

    frame_name = data.get('frame', '')
    fr = frame_row.get(frame_name)
    if fr:
        color = data.get('frame_color', '').strip()
        color_map = frame_color_cols.get(frame_name, {})
        matched = False
        for c_name, c_col in color_map.items():
            if c_name.strip().lower() == color.lower():
                cells[cell_ref(fr, c_col + 1)] = '\u221A'
                matched = True
                break
        if not matched and color:
            cells[cell_ref(fr, 4)] = '\u221A'
            cells[cell_ref(fr, 5)] = color

    # ── Customization ────────────────────────────────────────────────
    put('E46', data.get('custom_case', ''))
    put('E47', data.get('custom_frame', ''))
    put('N44', data.get('note', ''))

    # ── Lens type checkmark ──────────────────────────────────────────
    lens_row = {'neutral': 53, 'neutral_cl': 55, 'distance': 57,
                'intermediate': 59, 'reading': 61, 'bifocal': 63}
    lr = lens_row.get(data.get('lens_type', ''))
    if lr:
        cells[cell_ref(lr, 9)] = '\u221A'

    # ── RX values ────────────────────────────────────────────────────
    # OD: F(sph), G(cyl), H(axis)   OS: J(sph), K(cyl), M(axis)
    # Row 69=Distance, 70=Interm, 71=Reading, 72=Add
    rx = data.get('rx', {})
    for dist, rn in [('dist', 69), ('int', 70), ('read', 71)]:
        for eye, cols in [('od', (6, 7, 8)), ('os', (10, 11, 13))]:
            for field, cn in zip(('sph', 'cyl', 'axis'), cols):
                val = rx.get(f'{eye}_{dist}_{field}')
                if val not in (None, ''):
                    try:
                        val = float(val)
                    except (ValueError, TypeError):
                        pass
                    cells[cell_ref(rn, cn)] = val

    for eye, cn in [('od', 6), ('os', 10)]:
        val = rx.get(f'{eye}_add')
        if val not in (None, ''):
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass
            cells[cell_ref(72, cn)] = val

    # ── Declination ──────────────────────────────────────────────────
    decl_col = {'18': 4, '22': 7, 'MAX': 10}
    dc = decl_col.get(data.get('decl', 'MAX'))
    if dc:
        cells[cell_ref(75, dc)] = '\u221A'

    # ── IPD ──────────────────────────────────────────────────────────
    ipd = data.get('ipd', {})
    for key, addr in [('r1', 'F80'), ('r2', 'G80'), ('r3', 'H80'),
                      ('l1', 'I80'), ('l2', 'J80'), ('l3', 'L80')]:
        val = ipd.get(key)
        if val not in (None, ''):
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass
            cells[addr] = val

    # ── Headlight ────────────────────────────────────────────────────
    hl_row = {'LYNX': 53, 'LYNX PRO': 55, 'EOS Wireless': 57}
    hr = hl_row.get(data.get('headlight', ''))
    if hr:
        cells[cell_ref(hr, 20)] = '\u221A'

    # ── Accessories ──────────────────────────────────────────────────
    acc_map = {
        '701 Overloupes':          (67, 20),
        '710 Overloupes':          (69, 20),
        'Antifog cloth':           (71, 20),
        'Case Loupe&Headlight':    (73, 20),
        'Custom Magnetic Adapter': (63, 20),
    }
    for acc in data.get('accessories', []):
        pos = acc_map.get(acc)
        if pos:
            cells[cell_ref(pos[0], pos[1])] = '\u221A'

    return cells


# ── XML-level cell injection ─────────────────────────────────────────

def set_cell_value(c_elem, value):
    """
    Modify an existing <c> element in-place to hold `value`.
    Preserves the 's' (style) attribute.
    """
    # Remove existing children (<v>, <f>, <is>)
    for child in list(c_elem):
        c_elem.remove(child)

    if isinstance(value, tuple) and value[0] == 'formula':
        # Formula
        c_elem.attrib.pop('t', None)
        f_el = ET.SubElement(c_elem, f'{NSP}f')
        f_el.text = value[1]
    elif isinstance(value, (int, float)):
        # Number
        c_elem.attrib.pop('t', None)
        v_el = ET.SubElement(c_elem, f'{NSP}v')
        # Write integers without decimal point
        if isinstance(value, float) and value == int(value):
            v_el.text = str(int(value))
        else:
            v_el.text = str(value)
    else:
        # String → inline string (preserves style, no sharedStrings modification)
        c_elem.set('t', 'inlineStr')
        is_el = ET.SubElement(c_elem, f'{NSP}is')
        t_el  = ET.SubElement(is_el, f'{NSP}t')
        t_el.text = str(value)


def make_cell_elem(ref, value, style_idx=None):
    """Create a new <c> element for a cell that doesn't exist in the template."""
    c = ET.Element(f'{NSP}c')
    c.set('r', ref)
    if style_idx is not None:
        c.set('s', str(style_idx))
    set_cell_value(c, value)
    return c


def get_or_create_row(sheet_data, row_num):
    """Find or create a <row> element for the given row number."""
    for row_el in sheet_data.findall(f'{NSP}row'):
        if int(row_el.get('r')) == row_num:
            return row_el

    # Create new row in the right position
    new_row = ET.Element(f'{NSP}row')
    new_row.set('r', str(row_num))
    # Insert in order
    inserted = False
    for i, row_el in enumerate(sheet_data.findall(f'{NSP}row')):
        if int(row_el.get('r')) > row_num:
            # Insert before this row
            rows = list(sheet_data)
            idx = list(sheet_data).index(row_el)
            sheet_data.insert(idx, new_row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(new_row)
    return new_row


def insert_cell_in_row(row_el, c_elem):
    """Insert a <c> element into a <row> in correct column order."""
    new_col = ref_to_col_num(c_elem.get('r'))
    for i, existing in enumerate(row_el.findall(f'{NSP}c')):
        existing_col = ref_to_col_num(existing.get('r'))
        if existing_col > new_col:
            # Insert before this cell
            children = list(row_el)
            idx = children.index(existing)
            row_el.insert(idx, c_elem)
            return
    # Append at end
    row_el.append(c_elem)


def find_style_for_cell(sheet_data, row_num, col_num):
    """Try to find a style index from a nearby cell in the same row."""
    for row_el in sheet_data.findall(f'{NSP}row'):
        if int(row_el.get('r')) == row_num:
            for c_el in row_el.findall(f'{NSP}c'):
                s = c_el.get('s')
                if s:
                    return s
    return None


def fill_template(data):
    """
    Fill the template by directly editing the sheet XML inside the zip.
    This preserves ALL formatting, images, drawings, print settings, etc.
    """
    cell_map = build_cell_map(data)
    if not cell_map:
        # Nothing to fill — return template as-is
        with open(TEMPLATE_PATH, 'rb') as f:
            return io.BytesIO(f.read())

    # Read template zip into memory
    with open(TEMPLATE_PATH, 'rb') as f:
        template_bytes = f.read()

    # Parse sheet1.xml
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        sheet_xml_bytes = zin.read('xl/worksheets/sheet1.xml')

    # Parse with ElementTree
    tree = ET.ElementTree(ET.fromstring(sheet_xml_bytes))
    root = tree.getroot()

    # Find <sheetData>
    sheet_data = root.find(f'{NSP}sheetData')

    # Group cells by row
    cells_by_row = {}
    for ref, val in cell_map.items():
        rn = ref_to_row(ref)
        cells_by_row.setdefault(rn, []).append((ref, val))

    # Process each cell
    for row_num, cell_list in cells_by_row.items():
        row_el = None
        # Find existing row
        for r in sheet_data.findall(f'{NSP}row'):
            if int(r.get('r')) == row_num:
                row_el = r
                break

        for ref, val in cell_list:
            if row_el is not None:
                # Look for existing cell
                found = False
                for c_el in row_el.findall(f'{NSP}c'):
                    if c_el.get('r') == ref:
                        # Cell exists — update value, preserve style
                        set_cell_value(c_el, val)
                        found = True
                        break
                if not found:
                    # Cell doesn't exist in this row — create and insert
                    style = find_style_for_cell(sheet_data, row_num,
                                                 ref_to_col_num(ref))
                    c_new = make_cell_elem(ref, val, style)
                    insert_cell_in_row(row_el, c_new)
            else:
                # Row doesn't exist — create row and cell
                row_el = get_or_create_row(sheet_data, row_num)
                style = find_style_for_cell(sheet_data, row_num,
                                             ref_to_col_num(ref))
                c_new = make_cell_elem(ref, val, style)
                row_el.append(c_new)

    # Serialize modified XML
    new_sheet_xml = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    # Rebuild zip: copy everything from template, replace sheet1.xml
    result_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin, \
         zipfile.ZipFile(result_buf, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            if item.filename == 'xl/worksheets/sheet1.xml':
                zout.writestr(item, new_sheet_xml)
            else:
                # Copy byte-for-byte from original
                zout.writestr(item, zin.read(item.filename))

    result_buf.seek(0)
    return result_buf


# ── Flask routes ─────────────────────────────────────────────────────

@app.route('/ping')
def ping():
    return 'ok'


@app.route('/export', methods=['POST', 'OPTIONS'])
def export():
    if request.method == 'OPTIONS':
        return '', 204

    data = request.get_json()
    buf  = fill_template(data)

    fname    = data.get('lname') or data.get('fname') or 'Order'
    date_str = (data.get('date') or '').replace('/', '').replace('-', '') or 'nodate'
    filename = f"Univet_Order_{fname}_{date_str}.xlsx"

    xlsx_bytes = buf.read()
    return Response(
        xlsx_bytes,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': f'attachment; filename="{filename}"',
            'Content-Length': str(len(xlsx_bytes)),
            'X-Content-Type-Options': 'nosniff',
            'Cache-Control': 'no-transform, no-store',
        }
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 7788))
    app.run(host='0.0.0.0', port=port)
