"""Static-file storage backends.

``ResilientManifestStaticFilesStorage`` gives every asset a content-hashed
filename (e.g. ``overview.abc123.js``) + a ``staticfiles.json`` manifest, so a
new deploy always serves a NEW url — browsers can never serve a stale JS/CSS
from cache after a change. ``manifest_strict = False`` means a template that
references a static file missing from the manifest falls back to the plain name
instead of raising ``Missing staticfiles manifest entry`` and 500-ing the page —
a safe rollout guard, not a licence to ship broken refs.
"""
from whitenoise.storage import CompressedManifestStaticFilesStorage


class ResilientManifestStaticFilesStorage(CompressedManifestStaticFilesStorage):
    manifest_strict = False
