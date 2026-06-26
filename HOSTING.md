# Hosting the RENEE B2B PO web app on your office LAN

This is the **internal pilot** setup: one Windows PC runs the app, and anyone on
the office network opens it in a browser. No nginx, no cloud, no Docker. We start
it **manually** (double-click `serve.bat`). When you later "move commercially",
the *Going commercial* section at the bottom lists what changes.

> **What this serves:** the Django `renee_cosmetics` project (the `online_b2b`
> dashboard + Process-PO flow). It reads the same MySQL `renee_orders` DB the
> Tkinter engine writes to. The Tkinter `online_po_processor` and the offline
> tools are **untouched backups** — hosting this does not affect them.

---

## ⭐ Quick reference — dev server vs. the real (hosted) server

You have **two** ways to run the site. Use the dev server while coding; use the
hosted server (waitress) for everyone on the office network.

| | **Dev server** (just you) | **Hosted / local server** (whole office) |
|---|---|---|
| Command | `.venv\Scripts\python manage.py runserver` | **double-click `serve.bat`** |
| URL | `http://127.0.0.1:8000/` (this PC only) | `http://<this-PC-LAN-IP>:8000/` (any PC) |
| Mode | DEBUG on (stack traces) | DEBUG off (production-ish) |
| Static files | Django auto-serves | WhiteNoise (collectstatic) |
| **How to STOP** | press **Ctrl + C** in that terminal | **close the `serve.bat` window** (or Ctrl + C in it) |

**Turn OFF the dev server and start the local/hosted server:**
1. Go to the terminal running `runserver` and press **Ctrl + C** (it stops).
2. **Double-click `serve.bat`** in `D:\PO tracking\Automation\warehouse\`.
3. Read the LAN URL it prints (`http://192.168.x.x:8000/`) and share that.
4. **Leave the `serve.bat` window open** — closing it stops the server for everyone.

> Don't run both at once — they both want port 8000. Stop one before starting the other.

---

## ⚙️ Applying updates (after code/template changes)

The hosted server loads code **once at startup** and (on prod) caches templates in
the process. So after pulling/making changes:

| Changed | What's needed |
|---|---|
| **Static files** (CSS/JS in `static/`) | `collectstatic` — **`serve.bat` runs it automatically** every start |
| **Templates** (`.html`) | Now **live** — uncached loaders are on, so just **refresh the browser** (no restart) |
| **Python** (views, `settings.py`, services) | **Restart `serve.bat`** (waitress loads code once) |

**Restarting cleanly:** `serve.bat` now **kills any stale process on port 8000**
before starting, so a restart always takes effect. If a change isn't showing:
1. Close the `serve.bat` window, double-click it again.
2. **Hard-refresh** the browser (Ctrl+F5).
3. Totally blank where a chart/section should be ⇒ the server is still stale —
   make sure only ONE `serve.bat` is running.

---

## 0. One-time: what's on the host machine

The serving PC needs:

1. **Python 3.13** (the same version the `.venv` was built with).
2. **MySQL** running with the `renee_orders` database (the engine's DB).
3. The **DB credentials file** at
   `C:\Users\<you>\AppData\Local\OnlinePOProcessor\db_config.json`
   (this is the *same* file the Tkinter app uses — it is **not** in the repo).
   Example:
   ```json
   { "backend": "mysql", "host": "127.0.0.1", "port": 3306,
     "user": "root", "password": "********", "database": "renee_orders" }
   ```
4. This repo checked out at `D:\PO tracking\Automation\warehouse`.

---

## 1. One-time: create the virtual environment

Open **PowerShell** in the project folder and run:

```powershell
cd "D:\PO tracking\Automation\warehouse"
py -3.13 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

This installs Django + **waitress** (the web server) + **whitenoise** (serves CSS/JS)
plus the engine's data/PDF stack.

---

## 2. One-time: create an admin login + a user

The app requires login. Create the first account:

```powershell
.venv\Scripts\python manage.py createsuperuser
```

Give it a username + password. Share that login (or make more users via
`http://localhost:8000/admin/`) with whoever needs access.

---

## 3. One-time: open the firewall for port 8000

So other PCs on the network can reach it. Run **PowerShell as Administrator**:

```powershell
netsh advfirewall firewall add rule name="RENEE B2B 8000" dir=in action=allow protocol=TCP localport=8000
```

(To remove it later: `netsh advfirewall firewall delete rule name="RENEE B2B 8000"`.)

---

## 4. Every time: start the server

**Double-click `serve.bat`** (in the project folder). A black window opens and:

- sets `DJANGO_DEBUG=0` (production mode — no debug pages),
- runs `collectstatic` (refreshes CSS/JS),
- prints the address to share, and
- starts **waitress** on port 8000.

**Leave that window open.** Closing it (or `Ctrl+C`) stops the server.

You'll see something like:
```
On THIS machine:        http://localhost:8000/
On the office network:  http://192.168.1.50:8000/   <-- share THIS one
```

Other people open `http://<that-IP>:8000/` in their browser and log in.

> **Find the IP manually** if needed: run `ipconfig` and read the **IPv4 Address**
> (e.g. `192.168.1.50`). It's stable as long as the PC keeps the same network.

---

## 5. Stopping / restarting

- **Stop:** close the `serve.bat` window (or `Ctrl+C` in it).
- **Restart:** double-click `serve.bat` again.
- After a **code update** (git pull): just restart — `serve.bat` re-runs
  `collectstatic` automatically. If you changed the DB schema, no migration is
  needed for the engine tables (they're `managed=False`); for Django's own auth
  tables run `.venv\Scripts\python manage.py migrate`.

---

## What each piece does

| Piece | Role |
|-------|------|
| **waitress** | Pure-Python WSGI server that works natively on Windows (gunicorn doesn't). Serves the Django app on `0.0.0.0:8000` so the LAN can reach it. |
| **whitenoise** | Serves the collected static files (CSS/JS/ApexCharts) directly from the app, gzip-compressed — no separate web server needed. |
| **collectstatic** | Copies every app's static files into `staticfiles/` where whitenoise serves them. Re-run on every start (cheap). |
| `DJANGO_DEBUG=0` | Turns OFF Django debug mode (no stack traces to users, proper static handling). Set in `serve.bat`. |
| `DJANGO_ALLOWED_HOSTS` | Which hostnames/IPs may serve the app. `*` = any (fine on a trusted LAN). Set in `serve.bat`. |
| `db_config.json` | MySQL credentials, shared with the Tkinter engine. **Never committed.** |

---

## Settings that make hosting work (already wired)

In `renee_cosmetics/settings.py`:

- `DEBUG = os.environ.get('DJANGO_DEBUG', '1') != '0'` — defaults to dev (on);
  `serve.bat` sets it to `0`.
- `SECRET_KEY` / `ALLOWED_HOSTS` read from env (`DJANGO_SECRET_KEY`,
  `DJANGO_ALLOWED_HOSTS`) with safe dev fallbacks.
- `whitenoise.middleware.WhiteNoiseMiddleware` sits right after `SecurityMiddleware`.
- `STORAGES["staticfiles"]` uses `whitenoise.storage.CompressedStaticFilesStorage`.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Others can't reach `http://<ip>:8000/` | Firewall rule (step 3) not added, or wrong IP. Confirm with `ipconfig`. Also check both PCs are on the same network/VLAN. |
| `.venv not found` on launch | Run step 1 (create venv + install requirements). |
| Page loads but **no styling** | `collectstatic` failed — run `.venv\Scripts\python manage.py collectstatic --noinput` and read the error. |
| `Bad Request (400)` | `ALLOWED_HOSTS` too narrow. Leave `DJANGO_ALLOWED_HOSTS=*` for LAN, or add the IP/name. |
| MySQL connection error | MySQL not running, or `db_config.json` missing/wrong. The dashboards need it. |
| Port 8000 already in use | Edit `serve.bat` and the firewall rule to another port (e.g. 8001). |

---

## Going commercial (later)

When you move off the single office PC, the upgrades are:

1. **Run as a Windows Service** (auto-start on boot, no open window) — wrap
   `waitress` with [NSSM](https://nssm.cc/) instead of `serve.bat`.
2. **Put a real server in front** — nginx or Caddy for HTTPS (TLS) + a real
   domain name, proxying to waitress on localhost.
3. **A real `SECRET_KEY`** via `DJANGO_SECRET_KEY` and a locked-down
   `DJANGO_ALLOWED_HOSTS` (no `*`).
4. **Backups** of MySQL `renee_orders` + the `media/` folder.
5. **A dedicated DB user** (not `root`) with only the needed grants.

None of that changes the app code — only how it's launched and fronted.
