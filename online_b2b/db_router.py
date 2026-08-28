"""
Database router — sends the engine-owned order models to the MySQL ``orders``
connection and keeps ALL migrations OFF it. The order tables are created and
owned by the engine (``online_po_processor.auto.history_db``); Django must
never CREATE/ALTER/DROP them. Admin CRUD only issues INSERT/UPDATE/DELETE on
rows, never DDL.
"""

ORDER_DB = 'orders'
# model_name (lowercase) → the ORIGINAL curated set (kept for allow_migrate's
# explicit belt-and-suspenders block). Read/write routing below is now by the
# general rule "every online_b2b managed=False model is a renee_orders table",
# so the read-only admin-browse models (models_extra.py) route here too without
# having to list all ~35 names.
ORDER_MODELS = {'run', 'orderheader', 'orderline',
                'itemmaster', 'channelskumap', 'shiptomapping',
                'itemexception'}


def _is_order(model) -> bool:
    # Every online_b2b model is managed=False and mapped to a renee_orders table
    # (the app has no Django-managed models); route them all to the orders DB.
    return (getattr(model._meta, 'app_label', '') == 'online_b2b'
            and model._meta.managed is False)


class OrdersRouter:
    def db_for_read(self, model, **hints):
        return ORDER_DB if _is_order(model) else None

    def db_for_write(self, model, **hints):
        return ORDER_DB if _is_order(model) else None

    def allow_relation(self, obj1, obj2, **hints):
        return None  # defer (no cross-db relations defined)

    def allow_migrate(self, db, app_label, model_name=None, **hints):
        # Never migrate anything onto the orders DB.
        if db == ORDER_DB:
            return False
        # Never migrate the order models (managed=False) on any DB.
        if model_name in ORDER_MODELS:
            return False
        return None
