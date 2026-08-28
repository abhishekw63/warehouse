"""
Database router — sends the engine-owned order models to the MySQL ``orders``
connection and keeps ALL migrations OFF it. The order tables are created and
owned by the engine (``online_po_processor.auto.history_db``); Django must
never CREATE/ALTER/DROP them. Admin CRUD only issues INSERT/UPDATE/DELETE on
rows, never DDL.
"""

ORDER_DB = 'orders'
# Apps whose managed=False models map to renee_orders tables: online_b2b (the
# curated CRUD master-data models) + dbtables (the read-only admin browsers for
# every other table). Routing is by the general rule "managed=False model in one
# of these apps is a renee_orders table", so no per-model name list to maintain.
_ORDER_APPS = {'online_b2b', 'dbtables'}
# model_name (lowercase) → the ORIGINAL curated set (kept for allow_migrate's
# explicit belt-and-suspenders block only).
ORDER_MODELS = {'run', 'orderheader', 'orderline',
                'itemmaster', 'channelskumap', 'shiptomapping',
                'itemexception'}


def _is_order(model) -> bool:
    # These apps hold only managed=False models mapped to renee_orders tables;
    # route them all to the orders DB (never the default/SQLite connection).
    return (getattr(model._meta, 'app_label', '') in _ORDER_APPS
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
