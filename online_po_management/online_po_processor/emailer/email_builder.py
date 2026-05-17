"""
emailer.email_builder
=====================

Pure function: ``ProcessingResult`` → (HTML body, subject line).

No network I/O — this module is fully deterministic and unit-testable
without SMTP credentials. The ``EmailSender`` class consumes what we
produce here.

Visual design
-------------
Matches the GT Mass Dump email shell (navy banner, colored-bar strip,
stats grid, data tables, navy footer) so both tools feel like they
come from the same suite. Content sections differ because the data
model differs:

* GT Mass    → SO Number / Distributor / City / State / Location
* Online PO  → Marketplace / PO / Ship-to / validation outcome

Structural sections (top → bottom):

    1. Header banner (marketplace name + run timestamp + timing)
    2. Colored bar strip (brand accent)
    3. Stats grid (POs, Items, Order Qty, Unmapped, Warnings)
    4. Per-PO table
    5. SKU Demand table (Item No → total qty)
    6. Validation summary panel (OK / MISMATCH / NOT_IN_MASTER counts)
    7. Footer (engineering credit + operational note)
"""

from __future__ import annotations
from collections import OrderedDict
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from online_po_processor.data.models import ProcessingResult, SORow


# ══════════════════════════════════════════════════════════════════════
#  COLOR PALETTE — re-used from GT Mass Dump for visual consistency
# ══════════════════════════════════════════════════════════════════════
#
# Hex values chosen for good contrast on both light and dark email
# clients. Change these in one place and the entire email restyles.

class _Colors:
    """Shared palette for email HTML."""
    NAVY   = '#1A237E'   # Header banners, stat totals
    GREEN  = '#2E7D32'   # SKU section, validation OK
    ORANGE = '#E65100'   # Order-qty accent, warnings
    PURPLE = '#6A1B9A'   # Unmapped accent
    GOLD   = '#FFD600'   # Footer branding highlight
    RED    = '#C62828'   # MISMATCH / errors
    GRAY   = '#666666'   # Subtle text
    LTGRAY = '#f5f5f5'   # Light backgrounds


def _format_indian(number: Any) -> str:
    """
    Format a number using the Indian numbering system (lakhs/crores).

    Examples:
        1643     → "1,643"
        123456   → "1,23,456"
        1234567  → "12,34,567"

    Falls back to ``str(number)`` if the value can't be coerced to a
    float (useful for empty-string placeholders).
    """
    try:
        number = float(number)
    except (TypeError, ValueError):
        return str(number)

    sign = '-' if number < 0 else ''
    number = abs(number)

    if number == int(number):
        int_part = str(int(number))
        dec_part = ''
    else:
        parts = f'{number:.2f}'.split('.')
        int_part = parts[0]
        dec_part = '.' + parts[1]

    if len(int_part) <= 3:
        return sign + int_part + dec_part

    # Indian grouping: last 3 digits, then groups of 2.
    result = int_part[-3:]
    remaining = int_part[:-3]
    while remaining:
        result = remaining[-2:] + ',' + result
        remaining = remaining[:-2]

    return sign + result + dec_part


# ══════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ══════════════════════════════════════════════════════════════════════


class EmailBuilder:
    """
    Compose the HTML report email for an Online PO run.

    Usage::

        html    = EmailBuilder.build_html(result)
        subject = EmailBuilder.build_subject(result)

    All methods are static — the class is a namespace, there is no
    meaningful instance state.
    """

    # ── Subject ────────────────────────────────────────────────────────

    @staticmethod
    def build_subject(result: ProcessingResult) -> str:
        """
        Build the email subject line.

        Format: ``📦 <Marketplace> SO Report: <N> PO(s), <M> Items —
        <dd-mm-YYYY HH:MM>``

        v2.0.0: Transfer Order results (output_type='to') use
        "TO Report" instead of "SO Report" — the recipient should
        be able to tell at a glance from the inbox which import
        shape this batch produced.

        Uses the marketplace name from ``result`` when available, else
        falls back to a generic label. Timestamp helps the recipient
        sort the inbox when multiple reports land on the same day.
        """
        marketplace = result.marketplace or 'Online PO'
        ts = datetime.now().strftime('%d-%m-%Y %H:%M')

        po_count = len({r.po_number for r in result.rows})
        item_count = len(result.rows)

        kind_label = (
            'TO Report' if getattr(result, 'output_type', 'so') == 'to'
            else 'SO Report'
        )
        return (
            f"📦 {marketplace} {kind_label}: {po_count} PO(s), "
            f"{item_count} Items — {ts}"
        )

    # ── Body ───────────────────────────────────────────────────────────

    @staticmethod
    def build_html(result: ProcessingResult) -> str:
        """
        Build the full HTML email body for a ``ProcessingResult``.

        The function is pure: no network I/O, no file I/O, no
        environmental dependencies. Two identical inputs always
        produce identical output (modulo the current timestamp in
        the header).

        Args:
            result: Populated ``ProcessingResult`` from the engine.

        Returns:
            Complete HTML document as a string — ready to pass to
            ``EmailMessage.add_alternative(html, subtype='html')``.
        """
        # Pre-aggregate everything the template needs; keeps the
        # f-string soup below focused on layout only.
        agg = EmailBuilder._aggregate(result)

        # Timestamp + elapsed time shown in the banner.
        ts = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
        elapsed_str = (
            f'{result.elapsed_seconds:.2f}s'
            if result.elapsed_seconds is not None
            else '—'
        )

        # v1.5.4: Layout reverted to show Sales Order Details (per-PO
        # with amount) instead of SKU Demand Summary. The per-PO view
        # tells the recipient at a glance which PO drives how much
        # revenue; the SKU rollup was operational noise that
        # duplicated data already in the Summary sheet of the SO
        # workbook. Divider added before the footer to keep the
        # TOTAL row from visually bleeding into the navy footer.
        html_parts: List[str] = []
        html_parts.append(EmailBuilder._render_shell_open())
        html_parts.append(EmailBuilder._render_banner(result, ts, elapsed_str))
        html_parts.append(EmailBuilder._render_color_strip())
        html_parts.append(EmailBuilder._render_stats_grid(agg))
        html_parts.append(EmailBuilder._render_section_divider())
        html_parts.append(EmailBuilder._render_po_table(agg))
        html_parts.append(EmailBuilder._render_section_divider())
        html_parts.append(EmailBuilder._render_footer(result, elapsed_str))
        html_parts.append(EmailBuilder._render_shell_close())

        return ''.join(html_parts)

    # ══════════════════════════════════════════════════════════════════
    #  Aggregation
    # ══════════════════════════════════════════════════════════════════

    @staticmethod
    def _aggregate(result: ProcessingResult) -> Dict[str, Any]:
        """
        Roll ``result.rows`` up into the summaries the template needs.

        Returns a dict with these keys::

            total_pos:      int
            total_items:    int
            total_qty:      int
            total_amount:   float       # v1.5.1 — sum of row.amount

            po_groups:      OrderedDict[po_number → {
                              'ship_to', 'mapped_location', 'location',
                              'qty', 'items', 'unmapped' (bool)
                            }]

            sku_groups:     list[(item_no, {
                              'description', 'qty', 'po_count'
                            })] — sorted by qty desc

        v1.5.1 changes:
            * Added ``total_amount`` — sum of ``SORow.amount`` across
              rows. Rows with ``amount=None`` (e.g. when the marketplace
              has no ``amount_col`` configured) contribute 0.
            * Removed ``unmapped_count``, ``warning_count``, and
              ``validation`` rollups — the trimmed email no longer
              shows the validation panel.
            * ``po_groups`` is still computed because the SKU table
              header line ("X PO(s)") and the unique-PO count for the
              stats grid both need it.
        """
        rows = result.rows

        # ── PO-level rollup ─────────────────────────────────────────────
        po_groups: "OrderedDict[str, Dict[str, Any]]" = OrderedDict()
        for row in rows:
            if row.po_number not in po_groups:
                po_groups[row.po_number] = {
                    'ship_to': row.ship_to,
                    'mapped_location': row.mapped_location or row.location,
                    'location': row.location,
                    'qty': 0,
                    'items': 0,
                    'amount': 0.0,
                    'unmapped': not row.mapped,
                }
            bucket = po_groups[row.po_number]
            bucket['qty'] += int(row.qty or 0)
            bucket['items'] += 1
            # v1.5.4: accumulate per-PO amount for the new Amount
            # column in the email's Sales Order Details table. Rows
            # whose marketplace has no ``amount_col`` configured
            # (Myntra today) contribute 0.0 silently.
            bucket['amount'] += float(row.amount or 0.0)
            # A PO is considered unmapped if ANY row in it is unmapped
            # (all rows in a PO share the same location, so this is
            # effectively the same condition as per-PO — defensive).
            if not row.mapped:
                bucket['unmapped'] = True

        # ── SKU-level rollup — sorted by total qty desc ────────────────
        sku_groups_dict: Dict[Any, Dict[str, Any]] = {}
        sku_po_sets: Dict[Any, set] = {}

        for row in rows:
            key = row.item_no
            if key not in sku_groups_dict:
                sku_groups_dict[key] = {
                    'description': row.description or '',
                    'qty': 0,
                    'po_count': 0,
                }
                sku_po_sets[key] = set()

            sku_groups_dict[key]['qty'] += int(row.qty or 0)
            sku_po_sets[key].add(row.po_number)

            # Fill description if it came in blank earlier.
            if not sku_groups_dict[key]['description'] and row.description:
                sku_groups_dict[key]['description'] = row.description

        # Finalize po_count from the tracked set of PO numbers.
        for key, bucket in sku_groups_dict.items():
            bucket['po_count'] = len(sku_po_sets[key])

        sku_groups = sorted(
            sku_groups_dict.items(),
            key=lambda kv: kv[1]['qty'],
            reverse=True,
        )

        # ── Top-line scalars ───────────────────────────────────────────
        # v1.5.1: total_amount sums the marketplace-native row amount
        # (when configured via ``amount_col`` per marketplace). For
        # marketplaces without an amount_col every ``row.amount`` is
        # None, so the sum is 0.0 — the email still renders the stat
        # cell, just as ₹0, keeping the headline grid visually
        # consistent across marketplaces.
        total_qty = sum(int(r.qty or 0) for r in rows)
        total_amount = sum(
            float(r.amount or 0.0) for r in rows
        )

        return {
            'total_pos':    len(po_groups),
            'total_items':  len(rows),
            'total_qty':    total_qty,
            'total_amount': total_amount,
            'po_groups':    po_groups,
            'sku_groups':   sku_groups,
            # v2.0.0: surface the result's output_type so the renderers
            # downstream (which only see ``agg``, not ``result``) can
            # decide between SO and TO copy without re-importing the
            # result. ``is_to=True`` flips the PO-table heading and any
            # other SO-specific copy in the body.
            'is_to':        getattr(result, 'output_type', 'so') == 'to',
        }

    # ══════════════════════════════════════════════════════════════════
    #  Shell — outer HTML wrapper
    # ══════════════════════════════════════════════════════════════════

    @staticmethod
    def _render_shell_open() -> str:
        """Open the outer table wrapper. Closed by ``_render_shell_close``."""
        return (
            '<html><body style="margin:0;padding:0;'
            'font-family:Arial,sans-serif;background:#f0f2f5;">'
            '<table width="100%" cellpadding="0" cellspacing="0" '
            'style="background:#f0f2f5;">'
            '<tr><td align="center" style="padding:20px 10px;">'
            '<table width="800" cellpadding="0" cellspacing="0" '
            'style="background:#fff;border-radius:8px;overflow:hidden;'
            'border:1px solid #ddd;">'
        )

    @staticmethod
    def _render_shell_close() -> str:
        """Close the outer table wrapper."""
        return '</table></td></tr></table></body></html>'

    # ══════════════════════════════════════════════════════════════════
    #  Sections
    # ══════════════════════════════════════════════════════════════════

    @staticmethod
    def _render_banner(
        result: ProcessingResult,
        ts: str,
        elapsed_str: str,
    ) -> str:
        """Top banner — marketplace name + timestamp + timing."""
        marketplace = result.marketplace or 'Online PO'
        file_name = result.input_file or ''
        # v1.7.0: when a multi-file batch was processed, show a
        # clearer "Source" line that reveals the batch size instead
        # of only showing the first filename. The full per-file list
        # lives on the Warnings sheet (if there are any warnings) and
        # in the Summary sheet's per-PO rows — the banner just needs
        # enough context that the recipient isn't surprised by
        # multi-PO stats later in the email.
        n_files = getattr(result, 'input_files_count', 1) or 1
        C = _Colors

        if file_name and n_files > 1:
            source_label = f"Source: {file_name} + {n_files - 1} more files"
        elif file_name:
            source_label = f"Source: {file_name}"
        else:
            source_label = ""

        file_line = (
            f'<p style="margin:6px 0 0;font-size:11px;color:#9fa8da;">'
            f'{source_label}</p>'
            if source_label else ''
        )

        # v1.9.0: show which warehouse fulfilled this batch. Helpful
        # for cross-team readers (operations, accounts) who need to
        # know which RENEE warehouse the ERP posted against.
        # Backwards-compatible: warehouse_display defaults to 'AHD'
        # on older results without the field, so legacy code paths
        # just keep showing AHD.
        wh_display = getattr(result, 'warehouse_display', '') or 'AHD'
        wh_code = getattr(result, 'warehouse_code', '') or 'PICK'
        warehouse_line = (
            f'<p style="margin:2px 0 0;font-size:11px;color:#9fa8da;">'
            f'Warehouse: {wh_display} ({wh_code})</p>'
        )

        # v2.0.0: banner kind label flips for TO results.
        kind_phrase = (
            'Transfer Order Report'
            if getattr(result, 'output_type', 'so') == 'to'
            else 'Sales Order Report'
        )

        return (
            f'<tr><td style="background:{C.NAVY};padding:25px 30px;'
            f'text-align:center;">'
            f'<p style="margin:0;font-size:22px;font-weight:bold;'
            f'color:white;">📦 {marketplace} — {kind_phrase}</p>'
            f'<p style="margin:8px 0 0;font-size:12px;color:#9fa8da;">'
            f'Generated: {ts} | Processing: {elapsed_str}</p>'
            f'{file_line}'
            f'{warehouse_line}'
            f'<table style="margin:10px auto 0;"><tr>'
            f'<td style="background:#283593;padding:5px 15px;'
            f'border-radius:15px;">'
            f'<span style="font-size:10px;color:#9fa8da;'
            f'letter-spacing:1px;">⚡ ONLINE PO PROCESSOR</span>'
            f'</td></tr></table>'
            f'</td></tr>'
        )

    @staticmethod
    def _render_color_strip() -> str:
        """Four-colored accent strip under the banner (brand touch)."""
        C = _Colors
        return (
            f'<tr><td style="height:4px;font-size:0;">'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td width="25%" style="background:{C.ORANGE};height:4px;">'
            f'</td>'
            f'<td width="25%" style="background:{C.GOLD};height:4px;">'
            f'</td>'
            f'<td width="25%" style="background:#00E676;height:4px;">'
            f'</td>'
            f'<td width="25%" style="background:#2979FF;height:4px;">'
            f'</td>'
            f'</tr></table></td></tr>'
        )

    @staticmethod
    def _render_stats_grid(agg: Dict[str, Any]) -> str:
        """
        4-cell stats grid under the banner:

            POs | Items | Order Qty | Amount

        Each cell shows the figure in large coloured type with a small
        uppercase label below. Cells are equal-width (25% each).

        v1.5.1 changes
        --------------
        * Removed: ``Unmapped`` cell (purple) and ``Warnings`` cell
          (red). The recipient already sees both pieces of information
          in the generated SO workbook's Warnings sheet — surfacing
          them in the email added noise without driving any action.
        * Added: ``Amount`` cell (purple) showing the total monetary
          value across all rows, formatted as Indian-grouped rupees
          (₹X,XX,XXX). Source is ``SORow.amount`` summed across rows
          — populated by the engine when the marketplace declares an
          ``amount_col`` in its config.

        Empty / unconfigured case
        -------------------------
        When the marketplace has no ``amount_col`` (currently Myntra),
        every row's ``amount`` is ``None`` and the sum is 0. The cell
        still renders, just as ``₹0`` — keeping the headline grid
        visually consistent across marketplaces. The recipient learns
        at a glance that no amount data was available, instead of
        wondering whether a stat was hidden because of a bug.
        """
        C = _Colors

        # Indian-grouped rupee total. Round to whole rupees for display
        # — the punch files store paisa-precision values but the
        # headline number is more legible without the .00 noise.
        amount_str = '\u20B9' + _format_indian(int(round(agg["total_amount"])))

        return (
            f'<tr><td style="padding:0;border-bottom:1px solid #eee;">'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'

            # Cell 1: POs
            f'  <td width="25%" style="text-align:center;'
            f'padding:20px 10px;border-right:1px solid #f0f0f0;">'
            f'    <p style="margin:0;font-size:32px;font-weight:bold;'
            f'color:{C.NAVY};">{agg["total_pos"]}</p>'
            f'    <p style="margin:5px 0 0;font-size:10px;color:#999;'
            f'text-transform:uppercase;letter-spacing:1px;">POs</p>'
            f'  </td>'

            # Cell 2: Items
            f'  <td width="25%" style="text-align:center;'
            f'padding:20px 10px;border-right:1px solid #f0f0f0;">'
            f'    <p style="margin:0;font-size:32px;font-weight:bold;'
            f'color:{C.GREEN};">{_format_indian(agg["total_items"])}</p>'
            f'    <p style="margin:5px 0 0;font-size:10px;color:#999;'
            f'text-transform:uppercase;letter-spacing:1px;">Items</p>'
            f'  </td>'

            # Cell 3: Order Qty
            f'  <td width="25%" style="text-align:center;'
            f'padding:20px 10px;border-right:1px solid #f0f0f0;">'
            f'    <p style="margin:0;font-size:32px;font-weight:bold;'
            f'color:{C.ORANGE};">{_format_indian(agg["total_qty"])}</p>'
            f'    <p style="margin:5px 0 0;font-size:10px;color:#999;'
            f'text-transform:uppercase;letter-spacing:1px;">Order Qty</p>'
            f'  </td>'

            # Cell 4: Amount (smaller font — strings like "₹14,29,26,574"
            # need more horizontal room than a 4-digit count, so we drop
            # to 28px so it fits comfortably in the 25% column).
            f'  <td width="25%" style="text-align:center;'
            f'padding:20px 10px;">'
            f'    <p style="margin:0;font-size:28px;font-weight:bold;'
            f'color:{C.PURPLE};">{amount_str}</p>'
            f'    <p style="margin:5px 0 0;font-size:10px;color:#999;'
            f'text-transform:uppercase;letter-spacing:1px;">Amount</p>'
            f'  </td>'

            f'</tr></table></td></tr>'
        )

    @staticmethod
    def _render_section_divider() -> str:
        """Thin three-colored bar used between major sections."""
        C = _Colors
        return (
            f'<tr><td style="padding:12px 20px;background:#f8f9fa;">'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td width="33%" style="height:2px;background:{C.NAVY};'
            f'font-size:0;">&nbsp;</td>'
            f'<td width="34%" style="height:2px;background:{C.GREEN};'
            f'font-size:0;">&nbsp;</td>'
            f'<td width="33%" style="height:2px;background:{C.ORANGE};'
            f'font-size:0;">&nbsp;</td>'
            f'</tr></table></td></tr>'
        )

    @staticmethod
    def _render_po_table(agg: Dict[str, Any]) -> str:
        """
        PO-level table: one row per PO with ship-to, quantity, and
        amount.

        Unmapped POs get a red-tinted row background so they're
        instantly spotted by the recipient.

        Columns (6 — v1.5.4):
            PO Number | Location | Ship-to | Items | Order Qty | Amount

        Amount is the sum of ``SORow.amount`` across all rows of this
        PO, displayed as an Indian-grouped rupee total (₹2,17,205
        for 2.17 lakh, ₹13,28,786 for 13.28 lakh, etc.). Rendered in
        the same purple as the headline Amount stat for visual link.
        Rounded to whole rupees for readability — the underlying
        precision isn't useful at PO-summary altitude.

        When the marketplace has no ``amount_col`` configured
        (Myntra today), every PO's amount is 0.0 and the cell
        displays ``₹0``. The column stays visible so the layout is
        consistent across marketplaces — the zero value itself tells
        the recipient amount data wasn't extracted.
        """
        C = _Colors
        rows_html: List[str] = []

        for i, (po_num, info) in enumerate(agg['po_groups'].items()):
            bg = '#fff5f5' if info['unmapped'] else (
                '#f9f9f9' if i % 2 == 1 else '#ffffff'
            )
            ship_to = info['ship_to'] or '—'
            loc = info['mapped_location'] or info['location'] or '—'
            flag = (
                ' <span style="color:#C62828;font-weight:bold;">'
                '⚠ unmapped</span>'
                if info['unmapped'] else ''
            )
            amount_val = info.get('amount', 0.0)
            amount_str = (
                '\u20B9' + _format_indian(int(round(amount_val)))
            )

            rows_html.append(
                f'<tr style="background:{bg};">'
                f'<td style="padding:9px 8px;text-align:center;'
                f'font-size:12px;border-bottom:1px solid #eee;'
                f'font-weight:bold;">{po_num}</td>'
                f'<td style="padding:9px 8px;text-align:left;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{loc}{flag}</td>'
                f'<td style="padding:9px 8px;text-align:center;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{ship_to}</td>'
                f'<td style="padding:9px 8px;text-align:center;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{_format_indian(info["items"])}</td>'
                f'<td style="padding:9px 8px;text-align:center;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{_format_indian(info["qty"])}</td>'
                f'<td style="padding:9px 8px;text-align:right;'
                f'font-size:12px;border-bottom:1px solid #eee;'
                f'font-weight:bold;color:{C.PURPLE};">'
                f'{amount_str}</td>'
                f'</tr>'
            )

        total_amount_val = agg.get('total_amount', 0.0)
        total_amount_str = (
            '\u20B9' + _format_indian(int(round(total_amount_val)))
        )

        # v2.0.0: For TO mode, the section heading is "Transfer Order
        # Details" rather than "Sales Order Details". The table layout
        # is unchanged — the Amount column simply shows ₹0 for TO
        # marketplaces (they have no ``amount_col`` configured,
        # following the same Myntra pattern in SO mode). The is_to
        # flag is plumbed in via _aggregate's output dict.
        section_label = (
            'Transfer Order Details'
            if agg.get('is_to')
            else 'Sales Order Details'
        )

        return (
            f'<tr><td style="padding:14px 20px;font-weight:bold;'
            f'font-size:14px;color:{C.NAVY};'
            f'border-left:5px solid {C.NAVY};'
            f'background:#E8EAF6;">📋 {section_label}</td></tr>'
            f'<tr><td style="padding:0;">'
            f'<table width="100%" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;">'
            f'<tr>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'PO Number</th>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'Location</th>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'Ship-to</th>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'Items</th>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'Order Qty</th>'
            f'<th style="background:{C.NAVY};color:white;'
            f'padding:10px 8px;font-size:11px;text-transform:uppercase;">'
            f'Amount</th>'
            f'</tr>'
            f'{"".join(rows_html)}'
            f'<tr style="background:#E8EAF6;font-weight:bold;">'
            f'<td style="padding:10px 8px;text-align:center;'
            f'font-size:12px;">TOTAL</td>'
            f'<td colspan="2" style="padding:10px 8px;text-align:left;'
            f'font-size:12px;">{agg["total_pos"]} PO(s)</td>'
            f'<td style="padding:10px 8px;text-align:center;'
            f'font-size:12px;">'
            f'{_format_indian(agg["total_items"])}</td>'
            f'<td style="padding:10px 8px;text-align:center;'
            f'font-size:12px;">'
            f'{_format_indian(agg["total_qty"])}</td>'
            f'<td style="padding:10px 8px;text-align:right;'
            f'font-size:12px;color:{C.PURPLE};">'
            f'{total_amount_str}</td>'
            f'</tr>'
            f'</table></td></tr>'
        )

    @staticmethod
    def _render_sku_table(agg: Dict[str, Any]) -> str:
        """SKU Demand table — one row per Item No, sorted by qty desc."""
        C = _Colors
        rows_html: List[str] = []
        grand_qty = 0

        for rank, (item_no, info) in enumerate(agg['sku_groups'], 1):
            bg = '#f1f8e9' if rank % 2 == 0 else '#ffffff'
            desc = info['description']
            if len(desc) > 50:
                desc = desc[:47] + '...'
            grand_qty += info['qty']

            rows_html.append(
                f'<tr style="background:{bg};">'
                f'<td style="padding:8px 6px;text-align:center;'
                f'font-size:12px;color:#999;border-bottom:1px solid #eee;">'
                f'{rank}</td>'
                f'<td style="padding:8px 6px;text-align:center;'
                f'font-size:12px;font-weight:bold;'
                f'border-bottom:1px solid #eee;">{item_no}</td>'
                f'<td style="padding:8px 6px;text-align:left;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{desc or "—"}</td>'
                f'<td style="padding:8px 6px;text-align:center;'
                f'font-size:12px;border-bottom:1px solid #eee;">'
                f'{info["po_count"]}</td>'
                f'<td style="padding:8px 6px;text-align:center;'
                f'font-size:12px;font-weight:bold;'
                f'border-bottom:1px solid #eee;">'
                f'{_format_indian(info["qty"])}</td>'
                f'</tr>'
            )

        return (
            f'<tr><td style="padding:14px 20px;font-weight:bold;'
            f'font-size:14px;color:{C.GREEN};'
            f'border-left:5px solid {C.GREEN};'
            f'background:#E8F5E9;">📦 SKU Demand Summary</td></tr>'
            f'<tr><td style="padding:0;">'
            f'<table width="100%" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;">'
            f'<tr>'
            f'<th style="background:{C.GREEN};color:white;'
            f'padding:10px 6px;font-size:11px;">#</th>'
            f'<th style="background:{C.GREEN};color:white;'
            f'padding:10px 6px;font-size:11px;">ITEM NO</th>'
            f'<th style="background:{C.GREEN};color:white;'
            f'padding:10px 6px;font-size:11px;">DESCRIPTION</th>'
            f'<th style="background:{C.GREEN};color:white;'
            f'padding:10px 6px;font-size:11px;">POs</th>'
            f'<th style="background:{C.GREEN};color:white;'
            f'padding:10px 6px;font-size:11px;">QTY</th>'
            f'</tr>'
            f'{"".join(rows_html)}'
            f'<tr style="background:#E8F5E9;font-weight:bold;">'
            f'<td></td>'
            f'<td style="padding:10px 6px;text-align:center;'
            f'font-size:12px;">TOTAL</td>'
            f'<td style="padding:10px 6px;text-align:left;'
            f'font-size:12px;">'
            f'{len(agg["sku_groups"])} unique SKU(s)</td>'
            f'<td style="padding:10px 6px;text-align:center;'
            f'font-size:12px;">{agg["total_pos"]}</td>'
            f'<td style="padding:10px 6px;text-align:center;'
            f'font-size:12px;">{_format_indian(grand_qty)}</td>'
            f'</tr>'
            f'</table></td></tr>'
        )

    @staticmethod
    def _render_validation_panel(agg: Dict[str, Any]) -> str:
        """
        Validation outcome counters — 4 colored cells in a row.

        Shows: OK, MISMATCH, NOT_IN_MASTER, NO_PRICE. Skips 'blank'
        because a blank status just means validation wasn't run for
        that row (e.g. master loader wasn't provided); not useful to
        show on-page.
        """
        v = agg['validation']
        C = _Colors

        return (
            f'<tr><td style="padding:14px 20px;font-weight:bold;'
            f'font-size:14px;color:{C.RED};'
            f'border-left:5px solid {C.RED};'
            f'background:#FFEBEE;">🔍 Validation Summary '
            f'({v["total"]} row(s) checked)</td></tr>'
            f'<tr><td style="padding:0;">'
            f'<table width="100%" cellpadding="0" cellspacing="0">'
            f'<tr>'
            f'  <td width="25%" style="text-align:center;'
            f'padding:18px 10px;border-right:1px solid #f0f0f0;'
            f'background:#E8F5E9;">'
            f'    <p style="margin:0;font-size:24px;font-weight:bold;'
            f'color:{C.GREEN};">{_format_indian(v["ok"])}</p>'
            f'    <p style="margin:3px 0 0;font-size:10px;color:#666;'
            f'text-transform:uppercase;letter-spacing:1px;">OK</p>'
            f'  </td>'
            f'  <td width="25%" style="text-align:center;'
            f'padding:18px 10px;border-right:1px solid #f0f0f0;'
            f'background:#FFEBEE;">'
            f'    <p style="margin:0;font-size:24px;font-weight:bold;'
            f'color:{C.RED};">{_format_indian(v["mismatch"])}</p>'
            f'    <p style="margin:3px 0 0;font-size:10px;color:#666;'
            f'text-transform:uppercase;letter-spacing:1px;">Mismatch</p>'
            f'  </td>'
            f'  <td width="25%" style="text-align:center;'
            f'padding:18px 10px;border-right:1px solid #f0f0f0;'
            f'background:#FFF3E0;">'
            f'    <p style="margin:0;font-size:24px;font-weight:bold;'
            f'color:{C.ORANGE};">'
            f'{_format_indian(v["not_in_master"])}</p>'
            f'    <p style="margin:3px 0 0;font-size:10px;color:#666;'
            f'text-transform:uppercase;letter-spacing:1px;">'
            f'Not In Master</p>'
            f'  </td>'
            f'  <td width="25%" style="text-align:center;'
            f'padding:18px 10px;background:#F3E5F5;">'
            f'    <p style="margin:0;font-size:24px;font-weight:bold;'
            f'color:{C.PURPLE};">{_format_indian(v["no_price"])}</p>'
            f'    <p style="margin:3px 0 0;font-size:10px;color:#666;'
            f'text-transform:uppercase;letter-spacing:1px;">'
            f'No Price</p>'
            f'  </td>'
            f'</tr></table></td></tr>'
        )

    @staticmethod
    def _render_footer(
        result: ProcessingResult,
        elapsed_str: str,
    ) -> str:
        """Footer — engineering credit + operational note."""
        C = _Colors
        marketplace = result.marketplace or 'Online PO'

        return (
            f'<tr><td style="background:{C.NAVY};padding:30px;'
            f'text-align:center;">'
            f'<p style="margin:0 0 5px;font-size:16px;font-weight:bold;'
            f'color:{C.GOLD};letter-spacing:1px;">'
            f'⚡ ONLINE PO PROCESSOR</p>'
            f'<p style="margin:0 0 18px;font-size:11px;color:#7986CB;">'
            f'Marketplace PO → ERP Sales Order Import Suite</p>'

            f'<table style="margin:0 auto;max-width:400px;"><tr>'
            f'<td style="background:rgba(255,255,255,0.08);'
            f'border:1px solid rgba(255,255,255,0.15);padding:18px;'
            f'border-radius:10px;text-align:center;">'
            f'<p style="margin:0 0 3px;font-size:10px;color:#7986CB;'
            f'text-transform:uppercase;letter-spacing:2px;">'
            f'🚀 Engineered by</p>'
            f'<p style="margin:0 0 5px;font-size:18px;'
            f'font-weight:bold;color:white;">Abhishek Wagh</p>'
            f'<p style="margin:0 0 3px;font-size:11px;color:#9FA8DA;">'
            f'Order Management and Automation</p>'
            f'<p style="margin:0;font-size:10px;color:#7986CB;">'
            f'📧 abhishek.wagh@reneecosmetics.in</p>'
            f'</td></tr></table>'

            f'<table style="margin:18px auto 0;max-width:550px;"><tr>'
            f'<td style="background:rgba(255,255,255,0.06);'
            f'padding:16px 20px;border-radius:8px;'
            f'border:1px solid rgba(255,255,255,0.1);text-align:left;">'
            f'<p style="margin:0 0 8px;font-size:11px;'
            f'font-weight:bold;color:#FFD600;">📌 Please Note</p>'
            f'<p style="margin:0 0 8px;font-size:10px;color:#B0BEC5;'
            f'line-height:1.6;">This is an auto-generated report from the '
            f'Online PO Processor. The {marketplace} Sales Order(s) above '
            f'have been generated and are ready for upload into '
            f'<span style="color:white;font-weight:bold;">Dynamics 365 '
            f'Business Central</span>.</p>'
            f'<p style="margin:0 0 4px;font-size:10px;color:#B0BEC5;'
            f'line-height:1.6;">Use this report to:</p>'
            f'<p style="margin:0 0 2px;font-size:10px;color:#9FA8DA;'
            f'line-height:1.5;">&nbsp;&nbsp;• Review SO-wise demand and '
            f'identify unmapped locations</p>'
            f'<p style="margin:0 0 2px;font-size:10px;color:#9FA8DA;'
            f'line-height:1.5;">&nbsp;&nbsp;• Cross-check marketplace '
            f'landing/cost prices against master</p>'
            f'<p style="margin:0 0 8px;font-size:10px;color:#9FA8DA;'
            f'line-height:1.5;">&nbsp;&nbsp;• Validate Ship-to codes '
            f'before dispatch</p>'
            f'<p style="margin:0;font-size:10px;color:#78909C;'
            f'line-height:1.5;">For any discrepancies, please contact the '
            f'<span style="color:#9FA8DA;">Order Management team</span>.'
            f'</p>'
            f'</td></tr></table>'

            f'<p style="margin:18px 0 0;font-size:9px;color:#5C6BC0;">'
            f'© 2026 RENEE Cosmetics Pvt. Ltd. | Warehouse Automation '
            f'Division | Confidential</p>'
            f'</td></tr>'
        )