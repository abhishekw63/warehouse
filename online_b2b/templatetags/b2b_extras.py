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
