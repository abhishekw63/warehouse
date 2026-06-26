"""Pure helper tests for order_db (no DB connection)."""

import datetime

from online_b2b.services import order_db


def test_parse_date_formats():
    assert order_db._parse_date('2026-06-22') == datetime.date(2026, 6, 22)
    assert order_db._parse_date('22-06-2026') == datetime.date(2026, 6, 22)
    assert order_db._parse_date('') is None
    assert order_db._parse_date(None) is None


def test_days_to_expiry():
    today = datetime.date.today()
    assert order_db._days_to_expiry(today.isoformat()) == 0
    assert order_db._days_to_expiry((today + datetime.timedelta(days=5)).isoformat()) == 5
    assert order_db._days_to_expiry((today - datetime.timedelta(days=3)).isoformat()) == -3
    assert order_db._days_to_expiry(None) is None


def test_norm_filters_defaults():
    f = order_db._norm_filters({})
    assert f['marketplace'] == ''
    assert f['days'] == 0
    assert f['order_type'] == ''


def test_norm_filters_trims_and_casts():
    f = order_db._norm_filters({'marketplace': '  Blinkit ', 'days': '7',
                                'q': ' PO '})
    assert f['marketplace'] == 'Blinkit'
    assert f['days'] == 7
    assert f['q'] == 'PO'


def test_sort_allowlist():
    # only known sort keys map to columns
    assert order_db._SORT_COLS['value'] == 'order_value'
    assert 'date' in order_db._SORT_COLS
