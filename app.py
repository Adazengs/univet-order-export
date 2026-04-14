#!/usr/bin/env python3
"""Univet Order Form Export API — Fixed version"""

import json, io, os, zipfile, datetime
from flask import Flask, request, send_file
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             '2026_ORDER_FORM__eng__rev0.xlsx')

def fill_template(data):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    def w(addr, val):
        if val is None or val == '':
            return
        ws[addr] = val

    # Customer info
    w('D6',  data.get('date', ''))
    w('D7',  data.get('fname', ''))
    w('D8',  data.get('lname', ''))        # Fixed: was E8
    yob = data.get('yob', '')
    if yob:
        w('E9', yob)
        # Fix age formula to use current year
        current_year = datetime.date.today().year
        ws['G9'] = f'=({current_year}-E9)'

    w('D10', data.get('profession', ''))
    w('D11', data.get('specialty', ''))
    w('D12', data.get('address', ''))
    w('D13', data.get('town', ''))
    w('D14', data.get('country', ''))
    w('D15', data.get('tel', ''))
    w('D16', data.get('email', ''))

    # Agent / Sales Representative
    agent = data.get('agent', '')
    if agent:
        w('E87', agent)

    # Optic + WD checkmark
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
        ws.cell(row=row, column=col).value = '\u221A'   # Fixed: √ not ✓

    # Frame checkmark + color checkmark
    frame_row = {
        'Look': 31, 'Cool': 32, 'Techne 2025': 33, 'Techne [Old]': 34,
        'Techne RX [Old]': 35,
        'Techne ASIAN FITTING [OLD]': 36, 'Techne RX ASIAN FITTING [OLD]': 37,
        'Ash 55-17': 38, 'Ash 53-17': 39,
        'ITA': 40, 'ITA - Extended Fit': 41, 'ONE': 42,
    }

    # Color name -> column of the color LABEL; checkmark goes 1 column RIGHT
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
                # Fixed: checkmark goes 1 column RIGHT of the color label
                ws.cell(row=fr, column=c_col + 1).value = '\u221A'
                matched = True
                break
        if not matched:
            # Fallback: mark frame selected + write color name
            ws.cell(row=fr, column=4).value = '\u221A'
            ws.cell(row=fr, column=5).value = color

    # Customization
    w('E46', data.get('custom_case', ''))
    w('E47', data.get('custom_frame', ''))
    w('N44', data.get('note', ''))

    # Lens type checkmark
    lens_row = {'neutral': 53, 'neutral_cl': 55, 'distance': 57,
                'intermediate': 59, 'reading': 61, 'bifocal': 63}
    lr = lens_row.get(data.get('lens_type', ''))
    if lr:
        ws.cell(row=lr, column=9).value = '\u221A'   # Fixed: √ not ✓

    # RX values — OD: F(sph), G(cyl), H(axis)  OS: J(sph), K(cyl), M(axis)
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
                    ws.cell(row=rn, column=cn).value = val

    for eye, cn in [('od', 6), ('os', 10)]:
        val = rx.get(f'{eye}_add')
        if val not in (None, ''):
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass
            ws.cell(row=72, column=cn).value = val

    # Declination — D=18°, G=22°, J=MAX
    decl_col = {'18': 4, '22': 7, 'MAX': 10}
    dc = decl_col.get(data.get('decl', 'MAX'))
    if dc:
        ws.cell(row=75, column=dc).value = '\u221A'   # Fixed: √ not ✓

    # IPD — Right: F80,G80,H80  Left: I80,J80,L80
    ipd = data.get('ipd', {})
    for key, addr in [('r1', 'F80'), ('r2', 'G80'), ('r3', 'H80'),
                      ('l1', 'I80'), ('l2', 'J80'), ('l3', 'L80')]:
        val = ipd.get(key)
        if val not in (None, ''):
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass
            ws[addr] = val

    # Headlight — col T (column 20)
    hl_row = {'LYNX': 53, 'LYNX PRO': 55, 'EOS Wireless': 57}
    hr = hl_row.get(data.get('headlight', ''))
    if hr:
        ws.cell(row=hr, column=20).value = '\u221A'   # Fixed: √ not ✓

    # Accessories — col T (column 20)
    acc_map = {
        '701 Overloupes':         (67, 20),
        '710 Overloupes':         (69, 20),
        'Antifog cloth':          (71, 20),
        'Case Loupe&Headlight':   (73, 20),
        'Custom Magnetic Adapter': (63, 20),
    }
    for acc in data.get('accessories', []):
        pos = acc_map.get(acc)
        if pos:
            ws.cell(row=pos[0], column=pos[1]).value = '\u221A'  # Fixed: √ not ✓

    # ── Save & preserve images ──
    buf = io.BytesIO()
    wb.save(buf)
    filled_bytes = buf.getvalue()

    # Re-inject images/drawings from original template (openpyxl sometimes drops them)
    result_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(filled_bytes), 'r') as filled_zip, \
         zipfile.ZipFile(TEMPLATE_PATH, 'r') as orig_zip, \
         zipfile.ZipFile(result_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:

        filled_names = set(filled_zip.namelist())

        # Copy all files from filled version
        for name in filled_zip.namelist():
            out_zip.writestr(name, filled_zip.read(name))

        # Re-inject any media/drawing files that openpyxl dropped
        for name in orig_zip.namelist():
            if name not in filled_names and ('media' in name or 'drawing' in name):
                out_zip.writestr(name, orig_zip.read(name))
            # Also restore drawing rels if they exist in original but not in filled
            if name not in filled_names and 'drawings' in name:
                out_zip.writestr(name, orig_zip.read(name))

    result_buf.seek(0)
    return result_buf


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
    date     = (data.get('date') or '').replace('/', '') or 'nodate'
    filename = f"Univet_Order_{fname}_{date}.xlsx"

    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 7788))
    app.run(host='0.0.0.0', port=port)
