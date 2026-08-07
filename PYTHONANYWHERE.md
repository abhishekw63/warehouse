# Deploying the RENEE B2B web app on PythonAnywhere

This is the cloud-hosting guide (Linux). For the office-LAN Windows setup see
`HOSTING.md`. What you deploy is the Django `renee_cosmetics` project (the
`online_b2b` dashboards + Process-PO flow for Online **and** Offline: GT Mass ·
EKA · MT · GT Select). It reads the MySQL **`renee_orders`** database.

---

## ⚠️ 0. Do this FIRST — rotate the leaked secrets

The Gmail **App Password** was previously hardcoded in the repo, so it exists in
git history and is considered compromised. Before/So after go-live:

1. Google Account → Security → **App passwords** → **revoke** `bomn ktfx jhct xexy`.
2. Create a **new** App Password, put it in `.env` (`EMAIL_PASSWORD`) — never in code.
3. (Recommended) generate a **new** `DJANGO_SECRET_KEY` for production (below).

The MySQL password lives only in `db_config.json` (gitignored) — keep it that way.

---

## 1. Prerequisites (on PythonAnywhere)

- A PythonAnywhere account (Hacker plan or above — you need MySQL + a custom
  domain of `USERNAME.pythonanywhere.com`).
- **Create the MySQL DB**: Databases tab → create `renee_orders`. Its full name
  becomes `USERNAME$renee_orders`; host `USERNAME.mysql.pythonanywhere-services.com`.
- Python **3.13** (matches `requirements.txt`).

---

## 2. Get the code + virtualenv

```bash
git clone <your repo URL> ~/warehouse
cd ~/warehouse
python3.13 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt
```
No `mysqlclient` needed — the app uses **PyMySQL** (pure Python). `waitress` is
Windows-only and simply goes unused on PA.

---

## 3. Move the business data (MySQL) to PythonAnywhere

The app is data-driven — `order_headers`, `order_lines`, `item_master`,
`eka_data`, `ship_to_mapping`, `channel_sku_map`, `inventory_*`, etc. Export from
the office MySQL and import into PA's MySQL:

```bash
# On the office machine (or wherever the live renee_orders lives):
mysqldump -u <user> -p renee_orders > renee_orders.sql

# Upload renee_orders.sql to PA (Files tab), then on a PA Bash console:
mysql -u USERNAME -p -h USERNAME.mysql.pythonanywhere-services.com \
      'USERNAME$renee_orders' < renee_orders.sql
```
> The office MySQL isn't reachable from PA (private network), so PA runs on this
> **imported copy** — it won't stay live with the desktop engine unless you also
> point the desktop at the PA MySQL, or re-import periodically.

---

## 4. Create the two secret files (from the templates)

```bash
cp .env.example .env
cp db_config.example.json db_config.json    # or put it at ~/db_config.json
```
Edit **`.env`** — set `DJANGO_SECRET_KEY`, `DJANGO_DEBUG=0`,
`DJANGO_ALLOWED_HOSTS=USERNAME.pythonanywhere.com`,
`DJANGO_CSRF_TRUSTED_ORIGINS=https://USERNAME.pythonanywhere.com`,
`ONLINE_PO_DB_CONFIG=/home/USERNAME/db_config.json`, and the `EMAIL_*` values.

Generate a secret key:
```bash
.venv/bin/python -c "from django.core.management.utils import get_random_secret_key as g; print(g())"
```
Edit **`db_config.json`** — the PA MySQL creds (see the template).
Both files are gitignored.

---

## 5. Django one-time setup

```bash
cd ~/warehouse
.venv/bin/python manage.py migrate           # creates the default sqlite (auth/sessions only)
.venv/bin/python manage.py createsuperuser    # your login (db.sqlite3 is NOT in the repo)
.venv/bin/python manage.py collectstatic --noinput
```
> `default` DB = sqlite (`db.sqlite3`) — Django auth/sessions only. The business
> data is the MySQL `orders` connection (from `db_config.json`); a router keeps
> migrations OFF it, so it's never altered.

---

## 6. Web app config (PythonAnywhere → Web tab)

1. **Add a new web app** → **Manual configuration** → Python **3.13**.
2. **Virtualenv**: `/home/USERNAME/warehouse/.venv`
3. **Source code** / **Working directory**: `/home/USERNAME/warehouse`
4. **WSGI configuration file** (edit it) — replace contents with:
   ```python
   import os, sys
   from dotenv import load_dotenv
   path = '/home/USERNAME/warehouse'
   if path not in sys.path:
       sys.path.insert(0, path)
   load_dotenv(os.path.join(path, '.env'))          # load .env for the WSGI process
   os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'renee_cosmetics.settings')
   from django.core.wsgi import get_wsgi_application
   application = get_wsgi_application()
   ```
5. **Static files** mapping: URL `/static/` → Directory
   `/home/USERNAME/warehouse/staticfiles`
   (Optional) URL `/media/` → `/home/USERNAME/warehouse/media`
6. Click **Reload**.

Open `https://USERNAME.pythonanywhere.com/` and log in.

---

## 7. Gotchas / notes

- **CSRF 403 on POST** → `DJANGO_CSRF_TRUSTED_ORIGINS` must include
  `https://USERNAME.pythonanywhere.com` (step 4). Already wired in settings.
- **Uploads**: `media/` must be writable (it is, under `/home`). EKA/MT/GT
  uploads + generated workbooks land there.
- **OneDrive path**: one optional deal-seed lookup points at a Windows OneDrive
  path (`overrides_store.py`); on Linux it just returns None and deal-seeding is
  skipped — harmless.
- **Item master / EKA registry** are read from the DB (`item_master`,
  `eka_data`), so no Excel/OneDrive files are needed for those.
- **Timezone** is UTC (`settings.py`); dates display as stored.
- **Applying updates**: `git pull` → `collectstatic` if static changed → **Reload**
  the web app (PA loads code once per reload).

---

## 8. Pre-go-live checklist

- [ ] Old Gmail App Password **revoked**; new one in `.env` only.
- [ ] Fresh `DJANGO_SECRET_KEY` in `.env`.
- [ ] `DJANGO_DEBUG=0`, `DJANGO_ALLOWED_HOSTS` + `DJANGO_CSRF_TRUSTED_ORIGINS` set.
- [ ] `db_config.json` present, `ONLINE_PO_DB_CONFIG` points to it, MySQL reachable.
- [ ] `renee_orders` data imported into PA MySQL.
- [ ] `migrate` + `createsuperuser` + `collectstatic` done.
- [ ] `.env` and `db_config.json` are gitignored (they are) and **not** committed.
- [ ] Web app reloads clean; login + Tracker + a Process-PO preview all work.
