# ⚡ Move Render to Singapore (the big page-speed fix)

**Why:** the app's data (TiDB Cloud) lives in **Singapore (`ap-southeast-1`)**, but
the Render web service is in the **US (Oregon)**. Every database round-trip crosses
the Pacific (~200 ms). A page makes several round-trips, so ~1–2 s of every page is
pure network wait. Putting the web service in **Singapore, next to the database**,
cuts each round-trip to **~10 ms** — the single biggest speed win, and it's **free**.

**What is NOT affected:** your **data is safe** — TiDB is untouched and stays in
Singapore. We only move the *web service* (the code), which holds no data. Uploads /
generated workbooks on Render's free disk are already ephemeral (they don't survive
restarts today), so nothing durable is lost.

Render **cannot change a service's region in place** — the region is fixed when the
service is created. So we create the service fresh in Singapore. ~15 minutes.

---

## Before you start — copy these from the current service
Render dashboard → your **renee-b2b** service → **Environment** tab. Copy the values
of (the `sync: false` secrets aren't in git, so you must re-enter them):

- `DB_HOST`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`, `DB_PORT` (4000), `DB_SSL` (1)
- `EMAIL_SENDER`, `EMAIL_PASSWORD`

Everything else is already in `render.yaml`.

---

## Steps

1. **Push the region change** (already staged in `render.yaml`: `region: singapore`).
   Make sure it's on your deployed branch (`main`).

2. **Create the new Singapore service** — pick ONE:
   - **Blueprint (recommended):** Render dashboard → **New +** → **Blueprint** →
     select this repo. Render reads `render.yaml` and proposes a service in
     **Singapore**. Add the 8 secrets from above when prompted → **Apply**.
   - **Manual:** **New +** → **Web Service** → this repo → set **Region = Singapore**,
     **Build** `pip install -r requirements.txt && python manage.py collectstatic --noinput`,
     **Start** `gunicorn renee_cosmetics.wsgi --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 300`,
     add all env vars (the 8 secrets **plus** `DJANGO_DEBUG=0`, `DJANGO_DEFAULT_DB=mysql`,
     `DJANGO_SECURE_SSL=1`, and `DJANGO_SECRET_KEY` = generate).

3. **Fix the hostname vars.** The new service gets a new `*.onrender.com` URL. Set,
   on the new service:
   - `DJANGO_ALLOWED_HOSTS` = the new host (e.g. `renee-b2b.onrender.com`)
   - `DJANGO_CSRF_TRUSTED_ORIGINS` = `https://<new-host>`
   *(If you want to keep the exact old URL, use a **custom domain** — stable across
   services — or delete the old service first and try to reuse its name.)*

4. **Deploy & verify.** Open a page — it should feel dramatically faster. Confirm
   data shows (it's the same TiDB). Everyone re-logs in once (a fresh
   `DJANGO_SECRET_KEY` invalidates old session cookies — expected).

5. **Delete the old Oregon service** once the new one is confirmed good (stops the
   slow one + frees the name).

---

## Optional, +₹600/mo — remove the last two slowdowns
Free tier is **0.1 shared CPU** and **spins down after 15 min idle** (first hit after
idle takes 30–60 s). To remove both, set on the service (or in `render.yaml`):

```yaml
plan: starter        # dedicated CPU, no spin-down (~$7/mo)
```

Region co-location (free) is the big win; Starter makes it consistently instant.

---

## What was already done in code (so this is the remaining lever)
- Tracker 11 → 3 queries/render; `_stable` reference-data cache; SQL aggregates + a
  composite index; persistent DB connections (`CONN_MAX_AGE`); IST-correct times.
- This pass: skip the per-request source-file scan in production; memoise the
  per-request role query.

These shaved the round-trip **count**; only co-location fixes the round-trip **cost**.
