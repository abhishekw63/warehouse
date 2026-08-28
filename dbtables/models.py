"""
SELF-CONTAINED, REMOVABLE app (``dbtables``) — read-only admin browse for the
remaining ``renee_orders`` tables (everything not already mapped in
``online_b2b/models.py``). They live in their OWN app so the admin groups them
under "Database — raw tables", not under Online B2B.

To REMOVE: delete the ``dbtables/`` folder, its ``'dbtables'`` line in
INSTALLED_APPS, and ``'dbtables'`` from ``_ORDER_APPS`` in
``online_b2b/db_router.py``.

Every model here is ``managed = False`` (Django NEVER creates/alters/drops these
tables — the engine + services own them) and is registered **read-only** in the
admin (see ``admin_extra.py``), so browsing can never mutate money-path data.
Routed to the ``orders`` DB by :class:`OrdersRouter` (which now routes every
``online_b2b`` ``managed=False`` model there). Field definitions were generated
by ``manage.py inspectdb --database=orders`` and lightly cleaned:
  * inspectdb FK/O2O primary keys → plain PK columns (browse-only, no relations)
  * composite-PK tables keep ``CompositePrimaryKey`` (Django 5.2+/6.0)
  * blob/large-text columns (``content``) are deferred in the admin queryset
"""

from django.db import models


class AuditLog(models.Model):
    id = models.BigAutoField(primary_key=True)
    ts = models.DateTimeField(blank=True, null=True)
    username = models.CharField(max_length=150, blank=True, null=True)
    method = models.CharField(max_length=10, blank=True, null=True)
    url_name = models.CharField(max_length=120, blank=True, null=True)
    path = models.CharField(max_length=300, blank=True, null=True)
    target = models.CharField(max_length=300, blank=True, null=True)
    detail = models.CharField(max_length=500, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'audit_log'
        verbose_name = 'Audit log'


class AvailabilityRun(models.Model):
    run_id = models.BigAutoField(primary_key=True)
    run_ts = models.DateTimeField(blank=True, null=True)
    actor = models.CharField(max_length=80, blank=True, null=True)
    n_orders = models.IntegerField(blank=True, null=True)
    n_skus = models.IntegerField(blank=True, null=True)
    order_nos = models.TextField(blank=True, null=True)
    wh_override = models.CharField(max_length=40, blank=True, null=True)
    inv_as_of = models.CharField(max_length=255, blank=True, null=True)
    fill_pct = models.DecimalField(max_digits=6, decimal_places=2, blank=True, null=True)
    fill_val_pct = models.DecimalField(max_digits=6, decimal_places=2, blank=True, null=True)
    ord_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    fillable_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    best_wh = models.CharField(max_length=20, blank=True, null=True)
    note = models.CharField(max_length=255, blank=True, null=True)
    payload = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'availability_run'
        verbose_name = 'Availability run'


class CustomTables(models.Model):
    name = models.CharField(max_length=255)
    slug = models.CharField(unique=True, max_length=255)
    columns = models.JSONField()
    color_rules = models.JSONField(blank=True, null=True)
    sort = models.IntegerField(blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'custom_tables'
        verbose_name = 'Custom table'


class CustomTableRows(models.Model):
    table_id = models.IntegerField()
    data = models.JSONField()
    sort = models.IntegerField(blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'custom_table_rows'
        verbose_name = 'Custom table row'


class D365PostingGroupMap(models.Model):
    pg_key = models.CharField(primary_key=True, max_length=120)
    posting_group = models.CharField(max_length=120, blank=True, null=True)
    segment = models.CharField(max_length=20, blank=True, null=True)
    marketplace = models.CharField(max_length=50, blank=True, null=True)
    marketplace_label = models.CharField(max_length=60, blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)
    created_by = models.CharField(max_length=150, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'd365_posting_group_map'
        verbose_name = 'D365 posting-group map'


class DailyAdhoc(models.Model):
    id = models.BigAutoField(primary_key=True)
    title = models.CharField(max_length=500)
    note = models.CharField(max_length=1000, blank=True, null=True)
    due = models.DateField(blank=True, null=True)
    done = models.IntegerField(blank=True, null=True)
    created_at = models.DateTimeField()
    created_by = models.CharField(max_length=80, blank=True, null=True)
    done_at = models.DateTimeField(blank=True, null=True)
    done_by = models.CharField(max_length=80, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'daily_adhoc'
        verbose_name = 'Daily ad-hoc task'


class DailyChecklist(models.Model):
    # Real PK is composite (day, channel, step); Django admin can't register a
    # composite PK, so we designate one column as a BROWSE pk (read-only, so its
    # non-uniqueness is cosmetic — the list still shows every row).
    day = models.DateField(primary_key=True)
    channel = models.CharField(max_length=40)
    step = models.CharField(max_length=20)
    checked = models.IntegerField(blank=True, null=True)
    checked_at = models.DateTimeField(blank=True, null=True)
    checked_by = models.CharField(max_length=80, blank=True, null=True)
    remark = models.CharField(max_length=500, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'daily_checklist'
        verbose_name = 'Daily checklist entry'


class DailyChecklistHoldLog(models.Model):
    id = models.BigAutoField(primary_key=True)
    day = models.DateField()
    channel = models.CharField(max_length=40)
    action = models.CharField(max_length=10)
    at = models.DateTimeField()
    by_user = models.CharField(max_length=80, blank=True, null=True)
    reason = models.CharField(max_length=500, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'daily_checklist_hold_log'
        verbose_name = 'Daily checklist hold log'


class EkaData(models.Model):
    id = models.BigAutoField(primary_key=True)
    descr = models.CharField(max_length=255, blank=True, null=True)
    bill_to = models.CharField(max_length=20, blank=True, null=True)
    ship_to = models.CharField(max_length=20, blank=True, null=True)
    location_code = models.CharField(max_length=60, blank=True, null=True)
    posting_group = models.CharField(max_length=40, blank=True, null=True)
    short_name = models.CharField(max_length=80, blank=True, null=True)
    prefix = models.CharField(max_length=10, blank=True, null=True)
    short_code = models.CharField(max_length=30, blank=True, null=True)
    transfer_code = models.CharField(max_length=60, blank=True, null=True)
    kind = models.CharField(max_length=20, blank=True, null=True)
    example_regular = models.CharField(max_length=60, blank=True, null=True)
    example_tester = models.CharField(max_length=60, blank=True, null=True)
    status = models.CharField(max_length=12, blank=True, null=True)
    margin_pct = models.DecimalField(max_digits=6, decimal_places=3, blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'eka_data'
        verbose_name = 'EKA store data'


class FlipkartWhMap(models.Model):
    origin_warehouse = models.CharField(primary_key=True, max_length=120)
    market_place = models.CharField(max_length=60)
    source = models.CharField(max_length=20, blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'flipkart_wh_map'
        verbose_name = 'Flipkart WH map'


class InventoryBinAudit(models.Model):
    id = models.BigAutoField(primary_key=True)
    snapshot_id = models.BigIntegerField(blank=True, null=True)
    warehouse = models.CharField(max_length=40, blank=True, null=True)
    bin_code = models.CharField(max_length=120, blank=True, null=True)
    zone_code = models.CharField(max_length=60, blank=True, null=True)
    decision = models.CharField(max_length=10, blank=True, null=True)
    n_lines = models.IntegerField(blank=True, null=True)
    qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'inventory_bin_audit'
        verbose_name = 'Inventory bin audit'


class InventoryBinLine(models.Model):
    id = models.BigAutoField(primary_key=True)
    snapshot_id = models.BigIntegerField(blank=True, null=True)
    warehouse = models.CharField(max_length=40, blank=True, null=True)
    item_no = models.CharField(max_length=60, blank=True, null=True)
    bin_code = models.CharField(max_length=120, blank=True, null=True)
    zone_code = models.CharField(max_length=60, blank=True, null=True)
    decision = models.CharField(max_length=10, blank=True, null=True)
    qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'inventory_bin_line'
        verbose_name = 'Inventory bin line'


class InventoryBinRule(models.Model):
    id = models.BigAutoField(primary_key=True)
    pattern = models.CharField(max_length=120, blank=True, null=True)
    match_type = models.CharField(max_length=10, blank=True, null=True)
    decision = models.CharField(max_length=10, blank=True, null=True)
    note = models.CharField(max_length=255, blank=True, null=True)
    updated_by = models.CharField(max_length=80, blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)
    warehouse = models.CharField(max_length=40, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'inventory_bin_rule'
        verbose_name = 'Inventory bin rule'


class InventorySnapshot(models.Model):
    snapshot_id = models.BigAutoField(primary_key=True)
    warehouse = models.CharField(max_length=40, blank=True, null=True)
    warehouse_name = models.CharField(max_length=120, blank=True, null=True)
    captured_at = models.DateTimeField(blank=True, null=True)
    source_file = models.CharField(max_length=255, blank=True, null=True)
    uploaded_by = models.CharField(max_length=80, blank=True, null=True)
    total_lines = models.IntegerField(blank=True, null=True)
    included_lines = models.IntegerField(blank=True, null=True)
    excluded_lines = models.IntegerField(blank=True, null=True)
    new_lines = models.IntegerField(blank=True, null=True)
    item_count = models.IntegerField(blank=True, null=True)
    included_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    excluded_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    new_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    is_current = models.IntegerField(blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'inventory_snapshot'
        verbose_name = 'Inventory snapshot'


class InventoryStock(models.Model):
    id = models.BigAutoField(primary_key=True)
    snapshot_id = models.BigIntegerField(blank=True, null=True)
    warehouse = models.CharField(max_length=40, blank=True, null=True)
    item_no = models.CharField(max_length=60, blank=True, null=True)
    ean = models.CharField(max_length=40, blank=True, null=True)
    description = models.CharField(max_length=255, blank=True, null=True)
    uom = models.CharField(max_length=20, blank=True, null=True)
    available_qty = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'inventory_stock'
        verbose_name = 'Inventory stock'


class OfflineMasterFile(models.Model):
    channel = models.CharField(primary_key=True, max_length=32)
    filename = models.CharField(max_length=255, blank=True, null=True)
    size_bytes = models.BigIntegerField(blank=True, null=True)
    n_chunks = models.IntegerField(blank=True, null=True)
    uploaded_at = models.DateTimeField(blank=True, null=True)
    uploaded_by = models.CharField(max_length=64, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'offline_master_file'
        verbose_name = 'Offline master file'


class OfflineMasterChunk(models.Model):
    # Real PK is composite (channel, seq); designate a browse pk (read-only).
    channel = models.CharField(primary_key=True, max_length=32)
    seq = models.IntegerField()
    content = models.TextField(blank=True, null=True)     # LONGBLOB — deferred in admin

    class Meta:
        managed = False
        db_table = 'offline_master_chunk'
        verbose_name = 'Offline master chunk'


class OfflineSeqState(models.Model):
    channel = models.CharField(primary_key=True, max_length=32)
    seq_date = models.CharField(max_length=16, blank=True, null=True)
    next_counter = models.BigIntegerField(blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'offline_seq_state'
        verbose_name = 'Offline SO-counter state'


class OrderIssueLines(models.Model):
    line_id = models.BigAutoField(primary_key=True)
    run_ts = models.DateTimeField(blank=True, null=True)
    marketplace = models.CharField(max_length=50, blank=True, null=True)
    po = models.CharField(max_length=100, blank=True, null=True)
    item_no = models.CharField(max_length=50, blank=True, null=True)
    ean = models.CharField(max_length=20, blank=True, null=True)
    description = models.CharField(max_length=255, blank=True, null=True)
    qty = models.IntegerField(blank=True, null=True)
    gst_code = models.CharField(max_length=20, blank=True, null=True)
    vendor_mrp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    our_mrp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    vendor_landing = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    our_landing = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    vendor_cp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    our_cp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    diff = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    margin_pct = models.DecimalField(max_digits=6, decimal_places=2, blank=True, null=True)
    status = models.CharField(max_length=20, blank=True, null=True)
    output_file = models.CharField(max_length=500, blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'order_issue_lines'
        verbose_name = 'Order issue line (legacy)'


class OrderLineValidation(models.Model):
    # inspectdb made this a OneToOne to OrderLines; browse-only → plain PK column.
    line_id = models.BigIntegerField(primary_key=True)
    our_mrp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    vendor_mrp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    our_landing = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    vendor_landing = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    our_cp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    vendor_cp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    diff = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    margin_pct = models.DecimalField(max_digits=6, decimal_places=2, blank=True, null=True)
    status = models.CharField(max_length=20, blank=True, null=True)
    exception_label = models.CharField(max_length=50, blank=True, null=True)
    received_ean = models.CharField(max_length=20, blank=True, null=True)
    action = models.CharField(max_length=20, blank=True, null=True)
    override_cp = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    remark = models.CharField(max_length=255, blank=True, null=True)
    decided_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'order_line_validation'
        verbose_name = 'Order-line validation'


class OrderLines(models.Model):
    line_id = models.BigAutoField(primary_key=True)
    run_id = models.BigIntegerField(blank=True, null=True)
    run_ts = models.DateTimeField(blank=True, null=True)
    marketplace = models.CharField(max_length=50, blank=True, null=True)
    po = models.CharField(max_length=100, blank=True, null=True)
    location = models.CharField(max_length=500, blank=True, null=True)
    item_no = models.CharField(max_length=50, blank=True, null=True)
    ean = models.CharField(max_length=20, blank=True, null=True)
    description = models.CharField(max_length=255, blank=True, null=True)
    qty = models.IntegerField(blank=True, null=True)
    order_type = models.CharField(max_length=10, blank=True, null=True)
    gst_code = models.CharField(max_length=20, blank=True, null=True)
    unit_price = models.DecimalField(max_digits=14, decimal_places=2, blank=True, null=True)
    output_file = models.CharField(max_length=500, blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'order_lines'
        verbose_name = 'Order line (facts)'


class OrderTat(models.Model):
    # inspectdb made this a OneToOne to order_headers; browse-only → plain PK column.
    order_id = models.BigIntegerField(primary_key=True)
    reason_code = models.CharField(max_length=40, blank=True, null=True)
    note = models.CharField(max_length=500, blank=True, null=True)
    reason_by = models.CharField(max_length=80, blank=True, null=True)
    reason_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'order_tat'
        verbose_name = 'Order TAT reason'


class OrderWhOverride(models.Model):
    po = models.CharField(primary_key=True, max_length=120)
    warehouse = models.CharField(max_length=40, blank=True, null=True)
    orig_warehouse = models.CharField(max_length=40, blank=True, null=True)
    note = models.CharField(max_length=255, blank=True, null=True)
    actor = models.CharField(max_length=80, blank=True, null=True)
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'order_wh_override'
        verbose_name = 'Order WH override'


class ParkedDraft(models.Model):
    token = models.CharField(primary_key=True, max_length=64)
    marketplace = models.CharField(max_length=64, blank=True, null=True)
    draft_at = models.CharField(max_length=24, blank=True, null=True)
    draft_note = models.CharField(max_length=320, blank=True, null=True)
    pos = models.IntegerField(blank=True, null=True)
    undecided = models.IntegerField(blank=True, null=True)
    files = models.IntegerField(blank=True, null=True)
    meta_json = models.TextField(blank=True, null=True)          # deferred in admin
    updated_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'parked_draft'
        verbose_name = 'Parked draft (Review Later)'


class ParkedDraftFile(models.Model):
    # Real PK is composite (token, filename, seq); designate a browse pk (read-only).
    token = models.CharField(primary_key=True, max_length=64)
    filename = models.CharField(max_length=255)
    seq = models.IntegerField()
    content = models.TextField(blank=True, null=True)           # LONGBLOB — deferred

    class Meta:
        managed = False
        db_table = 'parked_draft_file'
        verbose_name = 'Parked draft file'


class RecordVerificationLog(models.Model):
    id = models.BigAutoField(primary_key=True)
    po = models.CharField(unique=True, max_length=120)
    marketplace = models.CharField(max_length=80, blank=True, null=True)
    status = models.CharField(max_length=32, blank=True, null=True)
    our_qty = models.IntegerField(blank=True, null=True)
    d365_qty = models.IntegerField(blank=True, null=True)
    excluded_qty = models.IntegerField(blank=True, null=True)
    qty_delta = models.IntegerField(blank=True, null=True)
    our_val = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    d365_val = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    val_delta = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    our_pin = models.CharField(max_length=20, blank=True, null=True)
    d365_pin = models.CharField(max_length=20, blank=True, null=True)
    pin_ok = models.IntegerField(blank=True, null=True)
    mismatch_fields = models.CharField(max_length=255, blank=True, null=True)
    checked_by = models.CharField(max_length=150, blank=True, null=True)
    checked_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'record_verification_log'
        verbose_name = 'Record-verification log'


class ShipToField(models.Model):
    id = models.BigAutoField(primary_key=True)
    name = models.CharField(unique=True, max_length=60, blank=True, null=True)
    label = models.CharField(max_length=120, blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'ship_to_field'
        verbose_name = 'Ship-To custom field'


class TrackerManual(models.Model):
    id = models.BigAutoField(primary_key=True)
    dept = models.CharField(max_length=20, blank=True, null=True)
    warehouse = models.CharField(max_length=60, blank=True, null=True)
    marketplace = models.CharField(max_length=80, blank=True, null=True)
    po = models.CharField(max_length=120, blank=True, null=True)
    external_doc = models.CharField(max_length=120, blank=True, null=True)
    location = models.CharField(max_length=255, blank=True, null=True)
    pincode = models.CharField(max_length=12, blank=True, null=True)
    zone = models.CharField(max_length=20, blank=True, null=True)
    po_date = models.DateField(blank=True, null=True)
    exp_date = models.DateField(blank=True, null=True)
    order_value = models.DecimalField(max_digits=16, decimal_places=2, blank=True, null=True)
    qty = models.IntegerField(blank=True, null=True)
    omt = models.CharField(max_length=255, blank=True, null=True)
    created_by = models.CharField(max_length=80, blank=True, null=True)
    created_at = models.DateTimeField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'tracker_manual'
        verbose_name = 'Tracker manual row'


# Registry consumed by admin_extra.py (order = admin index order within the app).
_EXTRA_MODELS = [
    AuditLog, AvailabilityRun, CustomTables, CustomTableRows, D365PostingGroupMap,
    DailyAdhoc, DailyChecklist, DailyChecklistHoldLog, EkaData, FlipkartWhMap,
    InventoryBinAudit, InventoryBinLine, InventoryBinRule, InventorySnapshot,
    InventoryStock, OfflineMasterFile, OfflineMasterChunk, OfflineSeqState,
    OrderIssueLines, OrderLineValidation, OrderLines, OrderTat, OrderWhOverride,
    ParkedDraft, ParkedDraftFile, RecordVerificationLog, ShipToField, TrackerManual,
]
