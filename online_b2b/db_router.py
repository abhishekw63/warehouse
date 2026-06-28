"""
Database router — sends the engine-owned order models to the MySQL ``orders``
connection and keeps ALL migrations OFF it. The order tables are created and
owned by the engine (``online_po_processor.auto.history_db``); Django must
never CREATE/ALTER/DROP them. Admin CRUD only issues INSERT/UPDATE/DELETE on
rows, never DDL.
"""

ORDER_DB = 'orders'
# model_name (lowercase) → these live in MySQL renee_orders (order tables +
# the DB-sourced master data that retired the bundled Excels).
ORDER_MODELS = {'run', 'orderheader', 'orderline',
                'itemmaster', 'channelskumap', 'shiptomapping',
                'itemexception'}


def _is_order(model) -> bool:
    return (getattr(model._meta, 'app_label', '') == 'online_b2b'
            and model._meta.model_name in ORDER_MODELS)


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
