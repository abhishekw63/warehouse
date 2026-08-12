"""Create the 'Editors' group and seed all EXISTING users into it.

Rationale: turning on RBAC must not suddenly lock out anyone who already works
in the app. So every pre-existing user becomes an Editor (full write, exactly as
before); the admin then demotes the ones who should be Viewers on the Users &
Roles page. New users created afterwards default to Viewer. Superusers are always
Editors regardless of group, so this is belt-and-suspenders for them.
"""
from django.conf import settings
from django.db import migrations

EDITORS_GROUP = 'Editors'


def seed(apps, schema_editor):
    Group = apps.get_model('auth', 'Group')
    User = apps.get_model(settings.AUTH_USER_MODEL.split('.', 1)[0]
                          if '.' in settings.AUTH_USER_MODEL else 'auth',
                          settings.AUTH_USER_MODEL.split('.')[-1])
    grp, _ = Group.objects.get_or_create(name=EDITORS_GROUP)
    for u in User.objects.all():
        u.groups.add(grp)


def unseed(apps, schema_editor):
    Group = apps.get_model('auth', 'Group')
    Group.objects.filter(name=EDITORS_GROUP).delete()


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('auth', '0001_initial'),
    ]

    operations = [
        migrations.RunPython(seed, unseed),
    ]
