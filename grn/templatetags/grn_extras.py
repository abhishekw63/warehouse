from django import template

register = template.Library()


@register.filter
def dictkey(d, key):
    """Look up ``d[key]`` from a template with a variable key (dicts don't
    support subscripting by variable in the template language)."""
    try:
        return d.get(key)
    except AttributeError:
        return None
