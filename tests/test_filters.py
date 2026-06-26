"""Template filter unit tests (pure, deterministic — no DB)."""

from online_b2b.templatetags.b2b_extras import compact, inr_short, querystring


def test_inr_short_crore():
    assert inr_short(45_843_126) == "₹4.58 Cr"


def test_inr_short_lakh():
    assert inr_short(332_418) == "₹3.32 L"


def test_inr_short_thousands():
    assert inr_short(12_300) == "₹12,300"


def test_inr_short_small_and_zero():
    assert inr_short(540) == "₹540"
    assert inr_short(0) == "₹0"


def test_inr_short_negative():
    assert inr_short(-150_000) == "-₹1.50 L"


def test_inr_short_bad_input_passthrough():
    assert inr_short("n/a") == "n/a"


def test_compact():
    assert compact(180_127) == "180.1k"
    assert compact(8_508) == "8.5k"
    assert compact(92) == "92"
    assert compact(3_400_000) == "3.4M"


def test_querystring_active_filters_only():
    qs = querystring({
        'marketplace': 'Blinkit', 'days': 7, 'q': 'PO1',
        'warehouse': '', 'order_type': '', 'date_from': '', 'date_to': '',
        'sort': 'date', 'direction': 'desc',
    })
    assert 'marketplace=Blinkit' in qs
    assert 'days=7' in qs
    assert 'q=PO1' in qs
    # empty + default values excluded
    assert 'warehouse=' not in qs
    assert 'dir=' not in qs          # desc is default → omitted


def test_querystring_direction_asc_included():
    qs = querystring({'sort': 'value', 'direction': 'asc'})
    assert 'sort=value' in qs
    assert 'dir=asc' in qs


def test_querystring_non_dict():
    assert querystring(None) == ""
