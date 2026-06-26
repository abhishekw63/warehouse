"""item_master_loader.build_rows() — effective-MRP selection + join, using tiny
temp Excel files (no engine; DB is optional — the Swiggy lookup degrades to {})."""

import datetime

import pandas as pd

from online_b2b.services import item_master_loader as iml


def _make_files(tmp_path):
    items = pd.DataFrame({
        'No.': ['200164', '200238', '999999'],          # 999999 has no MRP
        'GTIN': ['8904473100569', '8906121640588', 'SKU-X'],
        'Description': ['ITEM A', 'ITEM B', 'NO MRP'],
        'GST Group Code': ['G-5-S', 'G-18-S', 'G-18'],
        'HSN/SAC Code': ['33049110', '33041000', '33049990'],
    })
    mrp = pd.DataFrame({
        'Item No.': ['200164', '200164', '200164', '200238', '400111'],
        'M.R.P.': [249, 220, 249, 750, 99],            # 400111 not in Items file
        'Start Date': [datetime.date(2024, 4, 1), datetime.date(2025, 9, 22),
                       datetime.date(2026, 6, 12), datetime.date(2024, 4, 1),
                       datetime.date(2024, 4, 1)],
        'End Date': [datetime.date(2025, 9, 21), datetime.date(2026, 6, 11),
                     datetime.date(2030, 3, 31), datetime.date(2030, 3, 31),
                     datetime.date(2030, 3, 31)],
    })
    ip, mp = tmp_path / 'items.xlsx', tmp_path / 'mrp.xlsx'
    with pd.ExcelWriter(ip) as w:
        items.to_excel(w, sheet_name='Items', index=False)
    with pd.ExcelWriter(mp) as w:
        mrp.to_excel(w, sheet_name='Item M.R.P.', index=False)
    return str(ip), str(mp)


def test_effective_mrp_picks_todays_window(tmp_path):
    ip, mp = _make_files(tmp_path)
    rows, stats, warnings = iml.build_rows(ip, mp, as_of=datetime.date(2026, 6, 24))
    by = {r['item_no']: r for r in rows}
    # Multi-period item: today's window is the last row → 249, NOT the expired 220.
    assert by['200164']['mrp'] == 249.0
    assert by['200164']['mrp_start'] == datetime.date(2026, 6, 12)
    assert by['200238']['mrp'] == 750.0
    # The item carries its joined attributes (EAN / GST / HSN).
    assert by['200164']['ean'] == '8904473100569'
    assert by['200164']['gst_code'] == 'G-5-S'


def test_mrp_item_missing_from_items_is_warned_not_dropped_silently(tmp_path):
    ip, mp = _make_files(tmp_path)
    rows, stats, warnings = iml.build_rows(ip, mp, as_of=datetime.date(2026, 6, 24))
    item_nos = {r['item_no'] for r in rows}
    # 400111 is in the MRP file but not the Items file → excluded, but reported.
    assert '400111' not in item_nos
    assert any('not found in the Items file' in w for w in warnings)
    # 999999 is in Items but has no MRP → not in the (MRP-driven) master.
    assert '999999' not in item_nos
    assert stats['items'] == 2
