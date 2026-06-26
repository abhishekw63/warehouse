"""build_lines() unit test with a fake engine result (no engine, no DB)."""

from types import SimpleNamespace

from online_b2b.services import lines_store


def _so(**kw):
    base = dict(
        po_number='PO1', source_location='Loc A', item_no='IT1', ean='890123',
        description='Test SKU', qty=10, gst_code='GST18', forced_unit_price=None,
        cost_price_ref=100.0, fob_price=120.0, ref_fob_price=110.0,
        applied_margin_pct=None, mrp=200.0, vendor_mrp=200.0, diffn=0.0,
        validation_status='OK', exception_label='',
    )
    base.update(kw)
    return SimpleNamespace(**base)


def _result(rows):
    return SimpleNamespace(rows=rows, marketplace='Blink', margin_pct=0.70,
                           compare_basis='landing', output_type='so')


def test_build_lines_basic_shape():
    res = _result([_so()])
    lines = lines_store.build_lines(res, run_id=5, output_file='out.xlsx')
    assert len(lines) == 1
    ln = lines[0]
    assert ln['run_id'] == 5
    assert ln['marketplace'] == 'Blink'
    assert ln['po'] == 'PO1'
    assert ln['qty'] == 10
    assert ln['order_type'] == 'SO'
    assert ln['status'] == 'OK'
    assert ln['output_file'] == 'out.xlsx'
    # split model: build_lines carries both fact + validation keys incl received_ean
    assert 'received_ean' in ln and ln['received_ean'] is None
    assert set(lines_store.COLS).issubset(ln.keys())


def test_build_lines_ean_fix_swaps_and_audits():
    """A corrected EAN ships on the CORRECT one; the WRONG one is kept as
    received_ean (audit). order_lines never holds the wrong EAN."""
    res = _result([_so(ean='8904473103652', item_no='200238')])
    lines = lines_store.build_lines(
        res, run_id=1, ean_fixes={'8904473103652': '8906121640588'})
    ln = lines[0]
    assert ln['ean'] == '8906121640588'          # shipped on the correct EAN
    assert ln['received_ean'] == '8904473103652'  # wrong EAN remembered for audit


def test_build_lines_affected_is_status_subset():
    res = _result([
        _so(validation_status='OK'),
        _so(po_number='PO2', validation_status='MISMATCH', diffn=-29.07),
        _so(po_number='PO3', validation_status='NOT_IN_MASTER'),
    ])
    lines = lines_store.build_lines(res, run_id=1)
    affected = [l for l in lines if l['status'] in ('MISMATCH', 'NOT_IN_MASTER')]
    assert len(lines) == 3
    assert len(affected) == 2          # the 2-table model: affected = status filter


def test_build_lines_forced_unit_price_wins():
    res = _result([_so(forced_unit_price=4.0, cost_price_ref=100.0)])
    ln = lines_store.build_lines(res, run_id=1)[0]
    assert ln['unit_price'] == 4.0     # deal/override price, not the ref cost
