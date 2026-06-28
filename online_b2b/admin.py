from django.contrib import admin

from .models import (
    ChannelSkuMap,
    ItemException,
    ItemMaster,
    OrderHeader,
    OrderLine,
    Run,
    ShipToMapping,
)


@admin.register(Run)
class RunAdmin(admin.ModelAdmin):
    list_display = ('run_id', 'run_ts', 'mode', 'marketplaces', 'total_pos',
                    'total_items', 'total_qty', 'total_value')
    list_filter = ('mode',)
    search_fields = ('run_id', 'source')
    ordering = ('-run_id',)
    readonly_fields = ('created_at',)
    list_per_page = 50


@admin.register(OrderHeader)
class OrderHeaderAdmin(admin.ModelAdmin):
    list_display = ('order_id', 'run_id', 'marketplace_label', 'po', 'location',
                    'order_type', 'items', 'qty', 'order_value', 'po_date',
                    'exp_date')
    list_filter = ('segment', 'marketplace_label', 'order_type', 'warehouse')
    search_fields = ('po', 'location', 'marketplace', 'marketplace_label')
    list_editable = ('order_type', 'qty', 'order_value')
    ordering = ('-order_id',)
    readonly_fields = ('created_at',)
    list_per_page = 50


@admin.register(OrderLine)
class OrderLineAdmin(admin.ModelAdmin):
    # Backed by the `order_lines_full` view (facts + validation) → browse-only.
    list_display = ('line_id', 'run_id', 'marketplace', 'po', 'item_no', 'ean',
                    'received_ean', 'description', 'qty', 'unit_price', 'our_cp',
                    'vendor_cp', 'diff', 'status', 'exception_label', 'action')
    list_filter = ('status', 'marketplace', 'order_type', 'action')
    search_fields = ('po', 'item_no', 'ean', 'received_ean', 'description')
    ordering = ('-line_id',)
    list_per_page = 50

    # A SQL view is not updatable — keep the admin read-only.
    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


# ── DB-sourced master data (the retired Excels live here now) ────────────────

@admin.register(ItemMaster)
class ItemMasterAdmin(admin.ModelAdmin):
    list_display = ('item_no', 'ean', 'description', 'mrp', 'gst_code', 'hsn',
                    'brand', 'mrp_start', 'mrp_end', 'updated_at')
    list_filter = ('gst_code', 'brand', 'batch_id')
    search_fields = ('item_no', 'ean', 'description')
    ordering = ('item_no',)
    list_per_page = 50


@admin.register(ChannelSkuMap)
class ChannelSkuMapAdmin(admin.ModelAdmin):
    list_display = ('id', 'channel', 'sku_code', 'ean', 'item_no', 'source',
                    'updated_at')
    list_filter = ('channel', 'source')
    search_fields = ('sku_code', 'ean', 'item_no')
    ordering = ('channel', 'sku_code')
    list_per_page = 100


@admin.register(ShipToMapping)
class ShipToMappingAdmin(admin.ModelAdmin):
    list_display = ('id', 'party', 'del_location', 'cust_no', 'ship_to', 'city',
                    'postcode', 'source', 'updated_at')
    list_filter = ('party', 'source')
    search_fields = ('party', 'del_location', 'cust_no', 'ship_to', 'city')
    ordering = ('party', 'del_location')
    list_per_page = 100


@admin.register(ItemException)
class ItemExceptionAdmin(admin.ModelAdmin):
    # ALL overrides in one place: kind='exception' (remap / CP override /
    # vendor-CP) + kind='swiggy_deal' (deal SKUs).
    list_display = ('id', 'kind', 'source_code', 'maps_to', 'override_mrp',
                    'override_margin', 'use_vendor_cp', 'cost_after_gst',
                    'marketplace', 'note', 'source', 'updated_at')
    list_filter = ('kind', 'marketplace', 'source')
    search_fields = ('source_code', 'maps_to', 'marketplace', 'note', 'item_id')
    ordering = ('kind', 'id')
