"""
Test settings — sqlite only, no MySQL.

Inherits everything from the real settings but drops the 'orders' MySQL
connection + router so the test runner never touches the production DB. Tests
that exercise the order DB do so via raw pymysql against the real (read-only)
data and are marked accordingly; Django ORM tests use the fast sqlite default.
"""

from .settings import *  # noqa: F401,F403

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': ':memory:',
    }
}
DATABASE_ROUTERS = []  # noqa: F405 — no orders router in tests
