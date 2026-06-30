"""
offline.flows
=============

Channel flow specs for the shared :mod:`online_b2b.services.po_flow` scaffold.
Each spec is the *only* per-channel glue — a new offline channel = a processor
adapter + one ``FlowSpec`` here (no new pages).
"""
from online_b2b.services.po_flow import FlowSpec

from .services.gt_mass_flow import GTMassProcessor

GT_MASS_SPEC = FlowSpec(
    key='gt_mass',
    title='GT Mass',
    segment='Offline',
    # Use the Online-B2B chrome (sidebar + b2b styling) so GT Mass looks/behaves
    # exactly like the online marketplace pages.
    base_template='online_b2b/base_b2b.html',
    upload_dirname='gt_mass_flow',
    processor=lambda meta: GTMassProcessor(meta),
    # GT Mass carries its own warehouse in the file (recorder defaults to AHD) —
    # no manual warehouse picker needed, so the 'warehouse' cap is intentionally off.
    caps=frozenset({'exclude', 'download'}),
    intro=('Upload GT Mass PO file(s) → review → record to the dashboard. '
           'The 7-sheet dump (SO Workbook) is downloadable any time.'),
    accept='.xlsx,.xls,.xlsm',
    urls={
        'upload': 'gtm_flow_upload', 'review': 'gtm_flow_review',
        'confirm': 'gtm_flow_confirm', 'decision': 'gtm_flow_decision',
        'discard': 'gtm_flow_discard', 'download': 'gtm_flow_download',
        'back': 'offline_dashboard', 'dashboard': 'offline_dashboard',
    },
)
