from urllib.parse import urlencode

from django import template

register = template.Library()


@register.filter
def inr_short(value):
    """Indian-style short money: ₹4.58 Cr / ₹3.21 L / ₹12,300 / ₹540."""
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return value
    sign = '-' if v < 0 else ''
    v = abs(v)
    if v >= 1e7:
        return f"{sign}₹{v / 1e7:.2f} Cr"
    if v >= 1e5:
        return f"{sign}₹{v / 1e5:.2f} L"
    if v >= 1000:
        return f"{sign}₹{v:,.0f}"
    return f"{sign}₹{v:.0f}"


@register.filter
def indnum(value, dp=0):
    """Indian digit grouping — 22,47,616 (last 3 digits, then groups of 2)."""
    try:
        n = float(value or 0)
    except (TypeError, ValueError):
        return value
    neg = n < 0
    if dp:
        whole, frac = f'{abs(n):.{int(dp)}f}'.split('.')
    else:
        whole, frac = f'{int(round(abs(n)))}', ''
    if len(whole) > 3:
        head, last3 = whole[:-3], whole[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ','.join(parts) + ',' + last3
    s = whole + (('.' + frac) if frac else '')
    return ('-' + s) if neg else s


@register.filter
def inr(value):
    """₹ with Indian grouping + 2 dp — ₹22,47,616.24."""
    return f'₹{indnum(value, 2)}'


@register.filter
def compact(value):
    """Short count: 1.2k / 3.4M for big quantities."""
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return value
    if v >= 1e6:
        return f"{v / 1e6:.1f}M"
    if v >= 1000:
        return f"{v / 1000:.1f}k"
    return f"{int(v)}"


@register.filter
def dictget(d, key):
    """Look up ``d[key]`` with a VARIABLE key in a template (Django can't do
    ``row[c.key]`` natively). Used by the generic verification page to render an
    arbitrary, channel-supplied column list. Returns '' when absent/not a dict."""
    if isinstance(d, dict):
        return d.get(key, '')
    return getattr(d, str(key), '')


@register.filter
def querystring(filters):
    """Turn the dashboard filters dict into a URL query string carrying only
    the active filters (matches the names the views read). 'direction' → 'dir'."""
    if not isinstance(filters, dict):
        return ''
    out = {}
    for key in ('marketplace', 'q', 'warehouse', 'order_type',
                'date_from', 'date_to', 'sort'):
        v = filters.get(key)
        if v:
            out[key] = v
    days = filters.get('days')
    if days and int(days) > 0:
        out['days'] = days
    direction = filters.get('direction')
    if direction and direction != 'desc':
        out['dir'] = direction
    return urlencode(out)


# ── SKU Exceptions page: MP avatar helpers ──────────────────────────────────
_MP_COLORS = {
    'Blink': '#0f9d58', 'BlinkMP': '#0b8043', 'Flipkart': '#2874f0',
    'Flipkart-TO': '#1a5fb4', 'Myntra': '#ff3f6c', 'Nykaa': '#e5177d',
    'Zepto': '#5b2a9e', 'Swiggy': '#fc8019', 'RK': '#0e7490',
    'Dmart': '#0a7d5a', 'Purplle': '#7c3aed', 'Reliance': '#c1121f',
    'Meesho-TO': '#f43397', 'Bigbasket': '#84b100', 'Firstcry': '#f47b20',
}


@register.filter
def mp_color(name):
    """Stable brand-ish colour for a marketplace avatar; deterministic fallback."""
    n = (str(name or '')).strip()
    if n in _MP_COLORS:
        return _MP_COLORS[n]
    palette = ['#4f46e5', '#0e7490', '#b45309', '#7c3aed', '#0a7d5a', '#be123c', '#1d4ed8']
    return palette[sum(ord(c) for c in n) % len(palette)] if n else '#94a3b8'


@register.filter
def first_letters(name):
    """Up to two initials for the avatar (e.g. 'Flipkart-TO' → 'FT', 'RK' → 'RK')."""
    n = (str(name or '')).strip()
    if not n:
        return '?'
    parts = [p for p in n.replace('-', ' ').split() if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[1][0]).upper()
    return n[:2].upper()
