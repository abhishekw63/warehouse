"""
Read/write ORM models mapped to the engine-owned MySQL ``renee_orders`` tables
(``managed = False`` — Django never creates/alters them; the engine owns the
schema). These exist ONLY to power admin CRUD; the dashboards still read via
raw pymysql in ``services/order_db.py``. Routed to the ``orders`` DB by
``OrdersRouter``.
"""

from django.db import models


class Run(models.Model):
    run_id = models.BigAutoField(primary_key=True)
    run_ts = models.DateTimeField(null=True, blank=True)
    mode = models.CharField(max_length=10, blank=True)
    source = models.CharField(max_length=500, blank=True, null=True)
    marketplaces = models.IntegerField(null=True, blank=True)
    total_pos = models.IntegerField(null=True, blank=True)
    total_items = models.IntegerField(null=True, blank=True)
    total_qty = models.IntegerField(null=True, blank=True)
    total_value = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    consolidated_path = models.CharField(max_length=500, blank=True, null=True)
    tracker_path = models.CharField(max_length=500, blank=True, null=True)
    created_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "runs"
        verbose_name = "Run"

    def __str__(self):
        return f"Run #{self.run_id} · {self.mode}"


class OrderHeader(models.Model):
    order_id = models.BigAutoField(primary_key=True)
    run_id = models.BigIntegerField(null=True, blank=True)
    run_ts = models.DateTimeField(null=True, blank=True)
    mode = models.CharField(max_length=10, blank=True)
    segment = models.CharField(max_length=20, blank=True)
    marketplace = models.CharField(max_length=50, blank=True)
    marketplace_label = models.CharField(max_length=50, blank=True)
    po = models.CharField(max_length=100, blank=True)
    location = models.CharField(max_length=500, blank=True)
    warehouse = models.CharField(max_length=20, blank=True)
    po_date = models.DateField(null=True, blank=True)
    exp_date = models.DateField(null=True, blank=True)
    order_type = models.CharField(max_length=10, blank=True)
    items = models.IntegerField(null=True, blank=True)
    qty = models.IntegerField(null=True, blank=True)
    order_value = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    output_file = models.CharField(max_length=500, blank=True)
    created_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "order_headers"
        verbose_name = "Order header"

    def __str__(self):
        return f"{self.marketplace_label} · {self.po}"


class OrderLine(models.Model):
    line_id = models.BigAutoField(primary_key=True)
    run_id = models.BigIntegerField(null=True, blank=True)
    run_ts = models.DateTimeField(null=True, blank=True)
    marketplace = models.CharField(max_length=50, blank=True)
    po = models.CharField(max_length=100, blank=True)
    location = models.CharField(max_length=500, blank=True)
    item_no = models.CharField(max_length=50, blank=True)
    ean = models.CharField(max_length=20, blank=True)
    description = models.CharField(max_length=255, blank=True)
    qty = models.IntegerField(null=True, blank=True)
    order_type = models.CharField(max_length=10, blank=True)
    gst_code = models.CharField(max_length=20, blank=True)
    unit_price = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    vendor_mrp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    our_mrp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    vendor_landing = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    our_landing = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    vendor_cp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    our_cp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    diff = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    margin_pct = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    status = models.CharField(max_length=20, blank=True)
    exception_label = models.CharField(max_length=50, blank=True)
    received_ean = models.CharField(max_length=20, blank=True, null=True)
    action = models.CharField(max_length=20, blank=True, null=True)
    override_cp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    remark = models.CharField(max_length=255, blank=True, null=True)
    decided_at = models.DateTimeField(null=True, blank=True)
    output_file = models.CharField(max_length=500, blank=True)
    created_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        # Reads the join VIEW (facts + validation) after the 2-table split, so
        # the admin browses the full line — the facts live in `order_lines` and
        # the comparison/decision cols in `order_line_validation`. Browse-only
        # (a view isn't updatable).
        db_table = "order_lines_full"
        verbose_name = "Order line"

    def __str__(self):
        return f"{self.po} · {self.item_no} ({self.status or 'OK'})"


# ── DB-sourced master data (Excels retired into these tables) ────────────────


class ItemMaster(models.Model):
    item_no = models.CharField(max_length=50, primary_key=True)
    ean = models.CharField(max_length=32, blank=True, null=True)
    description = models.CharField(max_length=512, blank=True, null=True)
    gst_code = models.CharField(max_length=20, blank=True, null=True)
    hsn = models.CharField(max_length=20, blank=True, null=True)
    mrp = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)
    mrp_start = models.DateField(null=True, blank=True)
    mrp_end = models.DateField(null=True, blank=True)
    # Per-channel SKU codes (Swiggy/HG/…) live in ChannelSkuMap, not here.
    base_uom = models.CharField(max_length=20, blank=True, null=True)
    brand = models.CharField(max_length=60, blank=True, null=True)
    category = models.CharField(max_length=100, blank=True, null=True)
    batch_id = models.CharField(max_length=40, blank=True, null=True)
    updated_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "item_master"
        verbose_name = "Item master"

    def __str__(self):
        return f"{self.item_no} · {self.description}"


class ChannelSkuMap(models.Model):
    """Per-channel SKU-code -> item/EAN map (Swiggy / Health & Glow / future
    code-only channels). Generalises the old item_swiggy_map."""

    id = models.BigAutoField(primary_key=True)
    channel = models.CharField(max_length=40, blank=True, null=True)
    sku_code = models.CharField(max_length=80, blank=True, null=True)
    ean = models.CharField(max_length=32, blank=True, null=True)
    item_no = models.CharField(max_length=50, blank=True, null=True)
    source = models.CharField(max_length=10, blank=True, null=True)
    updated_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "channel_sku_map"
        verbose_name = "Channel SKU map"

    def __str__(self):
        return f"[{self.channel}] {self.sku_code} → {self.item_no or self.ean}"


class ShipToMapping(models.Model):
    id = models.BigAutoField(primary_key=True)
    party = models.CharField(max_length=60, blank=True, null=True)
    del_location = models.CharField(max_length=500, blank=True, null=True)
    cust_no = models.CharField(max_length=40, blank=True, null=True)
    ship_to = models.CharField(max_length=60, blank=True, null=True)
    name = models.CharField(max_length=255, blank=True, null=True)
    address = models.CharField(max_length=500, blank=True, null=True)
    address2 = models.CharField(max_length=500, blank=True, null=True)
    postcode = models.CharField(max_length=20, blank=True, null=True)
    city = models.CharField(max_length=120, blank=True, null=True)
    source = models.CharField(max_length=10, blank=True, null=True)
    batch_id = models.CharField(max_length=40, blank=True, null=True)
    updated_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "ship_to_mapping"
        verbose_name = "Ship-To mapping"

    def __str__(self):
        return f"{self.party} · {self.del_location} → {self.ship_to}"


class ItemException(models.Model):
    """ALL per-code overrides in one table: EAN remap / CP override / vendor-CP
    (kind='exception') AND Swiggy deal SKUs (kind='swiggy_deal')."""

    id = models.BigAutoField(primary_key=True)
    kind = models.CharField(max_length=16, blank=True, null=True)
    source_code = models.CharField(max_length=80, blank=True, null=True)
    maps_to = models.CharField(max_length=80, blank=True, null=True)
    override_mrp = models.CharField(max_length=40, blank=True, null=True)
    override_margin = models.CharField(max_length=40, blank=True, null=True)
    use_vendor_cp = models.CharField(max_length=10, blank=True, null=True)
    marketplace = models.CharField(max_length=60, blank=True, null=True)
    note = models.CharField(max_length=500, blank=True, null=True)
    item_id = models.CharField(max_length=40, blank=True, null=True)
    correct_gst = models.CharField(max_length=40, blank=True, null=True)
    cost_with_gst = models.CharField(max_length=40, blank=True, null=True)
    cost_after_gst = models.CharField(max_length=40, blank=True, null=True)
    source = models.CharField(max_length=10, blank=True, null=True)
    updated_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        managed = False
        db_table = "item_exceptions"
        verbose_name = "Item exception"

    def __str__(self):
        return f"[{self.kind}] {self.source_code} ({self.marketplace or 'all'})"
