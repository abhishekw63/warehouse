"""GT Select parser — header/line mapping, join key, unit-price math, EAN clean.
DB-free (only parse_headers / parse_lines, which don't touch the DB)."""

import pandas as pd

from online_b2b.services import gt_select_import as gts


def _make(tmp_path):
    headers = pd.DataFrame({
        'No.': ['SO/06/26/0005001', 'SO/06/26/0005002'],
        'External Document No.': [224743253, 224721676],
        'Ship-to Name': ['Anand Trading Co.', 'Sri Kavyasree'],
        'Location Code': ['PICK', 'DS_BL_OFF1'],
        'Document Date': ['2026-06-23', '2026-06-23'],
        'Total Quantity': [1884, 744],
        'Total Amount Incl. GST': [558377.84, 245022.44],
        'Status': ['Open', 'Open'],
    })
    lines = pd.DataFrame({
        'Document No.': ['SO/06/26/0005001', 'SO/06/26/0005001', 'SO/06/26/0005002'],
        'Type': ['Item', 'Item', 'Item'],
        'GTIN': [8904473101214, 8904473102310, 8904473102488],
        'No.': [200652, 201055, 201146],
        'Description': ['COMPACT', 'SUNSCREEN', 'TINT'],
        'Quantity': [24, 36, 12],
        'Line Amount Excl. VAT': [8568.43, 6900.88, 1200.00],
        'Location Code': ['PICK', 'PICK', 'DS_BL_OFF1'],
    })
    hp, lp = tmp_path / 'h.xlsx', tmp_path / 'l.xlsx'
    headers.to_excel(hp, index=False)
    lines.to_excel(lp, index=False)
    return str(hp), str(lp)


def test_header_mapping(tmp_path):
    hp, _ = _make(tmp_path)
    r = gts.parse_headers(hp)
    assert r['ok']
    h = {x['so_no']: x for x in r['rows']}
    assert h['SO/06/26/0005001']['external_doc'] == '224743253'
    assert h['SO/06/26/0005001']['ship_to_name'] == 'Anand Trading Co.'
    assert h['SO/06/26/0005001']['warehouse'] == 'PICK'
    assert h['SO/06/26/0005001']['qty'] == 1884
    assert h['SO/06/26/0005001']['order_value'] == 558377.84
    assert h['SO/06/26/0005001']['order_type'] == 'SO'


def test_line_mapping_and_unit_price(tmp_path):
    _, lp = _make(tmp_path)
    r = gts.parse_lines(lp)
    assert r['ok']
    first = r['rows'][0]
    assert first['so_no'] == 'SO/06/26/0005001'      # join key (Document No.)
    assert first['item_no'] == '200652'
    assert first['ean'] == '8904473101214'           # GTIN cleaned (no .0)
    assert first['qty'] == 24
    # unit price = Line Amount Excl. VAT / Qty
    assert first['unit_price'] == round(8568.43 / 24, 2)
    assert len(r['rows']) == 3
