"""
SELF-CONTAINED, REMOVABLE — registers every model in ``models_extra`` in the
Django admin, **read-only** (no add / change / delete), so staff can browse the
remaining ``renee_orders`` tables. See ``models_extra.py`` header for how to
remove the whole feature.

Read-only is deliberate: these tables carry money-path + operational data, and
the admin here is a *viewer*, not an editor. Heavy blob/JSON columns (draft file
bytes, master-workbook chunks, payloads) are deferred out of the list query so a
browse never pulls megabytes.
"""

from django.contrib import admin
from django.db import models as _m

from . import models_extra

# Columns that can be large (up to ~4 MB) — never load them into a browse list.
_HEAVY = {'content', 'payload', 'meta_json', 'order_nos', 'data'}


class ReadOnlyAdmin(admin.ModelAdmin):
    list_per_page = 50
    list_display_links = None            # browse-only → no row links / change page

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        heavy = [f.name for f in self.model._meta.concrete_fields
                 if f.name in _HEAVY]
        return qs.defer(*heavy) if heavy else qs


def _display_fields(model):
    """First ~15 non-heavy concrete columns (the composite-PK virtual field and
    blob columns are skipped)."""
    out = []
    for f in model._meta.concrete_fields:
        if f.name in _HEAVY or isinstance(f, _m.CompositePrimaryKey):
            continue
        out.append(f.name)
    return out[:15]


for _model in models_extra._EXTRA_MODELS:
    if admin.site.is_registered(_model):
        continue
    admin.site.register(_model, type(
        f'{_model.__name__}Admin', (ReadOnlyAdmin,),
        {'list_display': _display_fields(_model)}))
