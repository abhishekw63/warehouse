# Hosting on Render + TiDB Cloud Serverless (free)

Modern free hosting for the RENEE B2B web app: **Render** runs the app (gunicorn),
**TiDB Cloud Serverless** is the free MySQL-compatible database. Email works
(unlike PythonAnywhere free). Only quirk: the free web service cold-starts after
~15 min idle. For the Windows-LAN setup see `HOSTING.md`; for PythonAnywhere see
`PYTHONANYWHERE.md`.

---

## 1. Create the TiDB database
1. Sign up at TiDB Cloud → create a **Serverless** cluster (free).
2. Create a database named `renee_orders`.
3. From **Connect**, note: host `gateway01.<region>.prod.aws.tidbcloud.com`,
   **port 4000**, user `xxxxxxxx.root`, password. TLS is **required**.

## 2. Migrate your current local data → TiDB (one-time)
On your machine (with the local MySQL running):
```bash
cp db_profiles/local.json.example db_profiles/local.json   # fill local creds
python db_switch.py save local                              # or snapshot the active one
cp db_profiles/tidb.json.example  db_profiles/tidb.json     # fill TiDB creds (ssl:true)
python db_push_to_tidb.py            # DRY RUN — lists tables + row counts
python db_push_to_tidb.py --push     # copies schema + data local → TiDB
```
Your local DB is untouched; TiDB is now an exact copy.

## 3. Deploy on Render
Push the repo to GitHub, then either use the **Blueprint** (`render.yaml`) or a
Manual web service:
- **Build**: `pip install -r requirements.txt && python manage.py collectstatic --noinput`
- **Start**: `gunicorn renee_cosmetics.wsgi --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
- **Env vars** (dashboard → Environment):
  `DJANGO_DEBUG=0`, `DJANGO_SECRET_KEY=<generate>`,
  `DJANGO_ALLOWED_HOSTS=<your>.onrender.com`,
  `DJANGO_CSRF_TRUSTED_ORIGINS=https://<your>.onrender.com`, `DJANGO_SECURE_SSL=1`,
  `DB_HOST`, `DB_NAME=renee_orders`, `DB_USER`, `DB_PASSWORD`, `DB_PORT=4000`, `DB_SSL=1`,
  `EMAIL_SENDER`, `EMAIL_PASSWORD`.
- First deploy runs `migrate` (creates Django's auth/session tables in TiDB).
  Then create your login via the Render **Shell**: `python manage.py createsuperuser`.

Open `https://<your>.onrender.com/` and log in.

## 4. Switching local ↔ server
- **In the app:** *Order Management → Setup* (staff only) has a toggle.
- **CLI:** `python db_switch.py local` / `python db_switch.py tidb` / `python db_switch.py status`.
Local remains the default and is always one click back.

## Notes
- No `mysqldump` needed — `db_push_to_tidb.py` is pure PyMySQL.
- TiDB requires TLS → `ssl:true` (or `DB_SSL=1`) is set for you; port is **4000**.
- Render free web sleeps after ~15 min idle (first hit ~30–60 s). Fine for a pilot.
- Same code runs local and on Render — only the config/env differs.
