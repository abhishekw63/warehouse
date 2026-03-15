import os
import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "renee_cosmetics.settings")
django.setup()

from django.contrib.auth.models import User
if not User.objects.filter(username='testuser').exists():
    User.objects.create_superuser('testuser', 'test@example.com', 'testpass')
