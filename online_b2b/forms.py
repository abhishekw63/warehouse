from django import forms

from .services import engine_bridge


class _MultiFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True


class _MultiFileField(forms.FileField):
    """FileField that accepts multiple files (Django dropped built-in multi
    support; this is the documented re-implementation)."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault('widget', _MultiFileInput(attrs={'multiple': True}))
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        single = super().clean
        if isinstance(data, (list, tuple)):
            return [single(d, initial) for d in data]
        return [single(data, initial)]


class UploadForm(forms.Form):
    marketplace = forms.ChoiceField(
        choices=engine_bridge.pilot_choices(),
        initial='Blink',
    )
    warehouse = forms.ChoiceField(
        choices=[(w, w) for w in engine_bridge.warehouse_choices()],
        initial=engine_bridge.default_warehouse(),
    )
    margin_pct = forms.IntegerField(
        min_value=1, max_value=100,
        required=False,                       # blank → marketplace default (view)
        initial=engine_bridge.default_margin_pct('Blink'),
        label='Margin %',
    )
    po_files = _MultiFileField(label='PO file(s)')
