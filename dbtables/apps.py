from django.apps import AppConfig


class DbTablesConfig(AppConfig):
    name = 'dbtables'
    # Shown as the admin section header (grouping) for the raw-table browsers,
    # so inventory / offline / daily-ops tables no longer sit under "Online B2B".
    verbose_name = 'Database - raw tables (read-only)'
