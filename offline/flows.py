"""
offline.flows
=============

Channel flow specs for the shared :mod:`online_b2b.services.po_flow` scaffold.
Each spec is the *only* per-channel glue — a new offline channel = a processor
adapter + one ``FlowSpec`` here (no new pages).
"""
from online_b2b.services.po_flow import FlowSpec

from . import services  # noqa: F401 — ensure package import order
from .services import mt_bridge
from .services.gt_mass_flow import GTMassProcessor
from .services.mt_flow import MTFlowProcessor


def _mt_channels() -> tuple:
    """MT child channels ``((code, display), …)`` for the upload picker. Falls
    back to a static list if the frozen automation can't be imported."""
    try:
        return tuple(mt_bridge.channel_choices())
    except Exception:  # noqa: BLE001
        return (('SS', 'Shoppers Stop'), ('HG', 'Health & Glow'),
                ('NT', 'Naturals'), ('BN', 'Nature Basket'), ('LL', 'Lifestyle'))


def _mt_warehouses() -> tuple:
    try:
        return tuple((w, w) for w in mt_bridge.warehouse_choices())
    except Exception:  # noqa: BLE001
        return (('AHD', 'AHD'), ('BLR', 'BLR'))


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
        'export': 'gtm_flow_export',
        'back': 'offline_dashboard', 'dashboard': 'offline_dashboard',
    },
)


MT_SPEC = FlowSpec(
    key='mt',
    title='Modern Trade (MT)',
    segment='Offline',
    base_template='online_b2b/base_b2b.html',
    upload_dirname='mt_flow',
    processor=MTFlowProcessor,
    # Channel (SS / HG / NT…) chosen at upload via the 'marketplace' capability;
    # warehouse picker on; Exclude decisions per line. NO 'download' cap — the SO
    # workbook must NOT be generated before confirm (assigning SO numbers burns
    # the shared sequence counter), so the Download link appears only AFTER
    # recording (via ``has_download``).
    caps=frozenset({'marketplace', 'warehouse', 'exclude'}),
    marketplaces=_mt_channels(),
    warehouses=_mt_warehouses(),
    intro=('Pick the MT channel → upload PO file(s) → review → record to the '
           'dashboard. Same 6-sheet SO workbook as the desktop app. '
           'Optionally drop a tester-requirement sheet (Store + SKU + Tester) to '
           'auto-generate tester SOs alongside (e.g. HG @ ₹0.54).'),
    # .xlsb (Lifestyle replenishment workbook) + .pdf (LS/RL optional PO PDFs for
    # the address cross-check) join the .xlsx/.csv MT inputs — all accepted by the
    # bridge (mt_bridge.ACCEPTED_EXTENSIONS).
    accept='.xlsx,.xls,.xlsb,.csv,.pdf',
    # Reusable "additional verification" modal — renders a trigger link + big
    # popup under the KPI cards ONLY when the channel produced a `verification`
    # dict (LS PDF address check is the first consumer; dark for every other
    # channel). Any MP reuses the same component by building the generic dict.
    slots={'after_kpis': 'po_flow/_verification_modal.html'},
    urls={
        'upload': 'mt_flow_upload', 'review': 'mt_flow_review',
        'confirm': 'mt_flow_confirm', 'decision': 'mt_flow_decision',
        'discard': 'mt_flow_discard', 'download': 'mt_flow_download',
        'export': 'mt_flow_export',
        'back': 'offline_dashboard', 'dashboard': 'offline_dashboard',
    },
)
