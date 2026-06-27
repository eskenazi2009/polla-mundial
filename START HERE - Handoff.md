# Polla World Cup 2026 — Project handoff (continue on any PC)

This folder lives on your **OneDrive Desktop**, so it syncs to any PC where you sign into OneDrive.
It has everything to keep working on the project. Read this file first.

> **This folder IS the live site's Git repo now.** It's connected to GitHub
> `eskenazi2009/polla-mundial`, and `build.ps1` + `.github/workflows/` live right here at the
> root. Edit here, `git push` from here, and the live site rebuilds. There is no separate clone
> to keep in sync anymore — this is the single source of truth.

## What this project is
A live, auto-updating web dashboard of everyone's predictions in the **pollaworldcup.com** pool.
A GitHub Actions cron logs in, fetches the data, regenerates a no-JS HTML page, and publishes it to
GitHub Pages — free and hands-off.

- **Live site:** https://eskenazi2009.github.io/polla-mundial/
- **GitHub repo:** https://github.com/eskenazi2009/polla-mundial  (account: eskenazi2009)
- **Refresh schedule:** every 30 min, only 1:00pm–3:00am Panama time (UTC-5). Editable in the workflow.

## Key identifiers (you'll need these)
- **Supabase project ref:** `bqkkplrrlmuxylnibdho`  → host `https://bqkkplrrlmuxylnibdho.supabase.co`
- **SUPABASE_ANON_KEY** (public publishable key, valid to 2036):
  ```
  eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJxa2twbHJybG11eHlsbmliZGhvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNzAwNTgsImV4cCI6MjA4OTk0NjA1OH0.QeXUubRqnCrxE3stC1AE67eYprHNODC8tEZ6mKIPnlc
  ```
- **GitHub repo secrets already set (don't need re-adding unless you recreate the repo):**
  `POLLA_EMAIL`, `POLLA_PASSWORD` (your pool login — you know these), `SUPABASE_ANON_KEY` (above).

---

## Working on the live site (this folder is the repo)
This OneDrive folder is the Git repo, wired to GitHub `eskenazi2009/polla-mundial`. On a PC where
OneDrive has synced this folder, you can edit and push directly — no cloning needed.

1. **Install Git for Windows:** https://git-scm.com/download/win  (includes Git Credential Manager).
2. **Make changes** to `build.ps1` (the generator) or `.github/workflows/build.yml` (the schedule).
3. **Commit + push** from this folder — the `push` trigger auto-rebuilds and redeploys in ~1 minute:
   ```powershell
   git add -A
   git commit -m "your change"
   git push
   ```
   (First push pops a browser GitHub login — that's Git Credential Manager, complete it once.)
4. The new version is live at the URL above. **Hard-refresh (Ctrl+Shift+R)** or add `?v=N` to beat the cache.

### Brand-new PC where OneDrive hasn't synced yet
If you'd rather not wait for OneDrive, you can clone a fresh copy anywhere and work there instead:
```powershell
git clone https://github.com/eskenazi2009/polla-mundial.git
cd polla-mundial
git config user.name "Tony Eskenazi"
git config user.email "eskenazi2009@gmail.com"
```
(The Excel scripts and these docs are committed too, so a clone has everything this folder has.)

### Preview changes locally BEFORE pushing (no login needed)
`build.ps1` normally logs in and fetches live. To preview generation changes against saved data,
use the sample `model.json` in `excel-and-data-scripts\`:
```powershell
# from the cloned repo folder
$repo = (Get-Location).Path
$raw  = Get-Content .\build.ps1 -Raw
$tail = $raw.Substring($raw.IndexOf('# ---------- 4) Generate the static dashboard ----------'))
$head = "`$OutDir='.\preview'; `$model = (Get-Content '<path>\excel-and-data-scripts\model.json' -Raw) | ConvertFrom-Json`n"
Set-Content .\preview.local.ps1 ($head + $tail); .\preview.local.ps1
Start-Process .\preview\index.html
```

### Force an instant refresh anytime
GitHub → repo → **Actions → Build Polla dashboard → Run workflow**. Deploys in ~1 min, even outside the window.

### Recreating the repo from scratch (only if ever needed)
1. Create a NEW **public** GitHub repo.
2. Push these files (from this folder — `build.ps1` + `.github/` are at the root).
3. Add the 3 secrets (Settings → Secrets and variables → Actions): `POLLA_EMAIL`, `POLLA_PASSWORD`, `SUPABASE_ANON_KEY`.
4. Settings → Pages → Source = **GitHub Actions**.
5. Actions → Run workflow.

---

## The Excel / analysis side (`excel-and-data-scripts\`)
These run on **Windows PowerShell + Excel** (Excel COM). Pipeline:
- `parse.ps1` — extracts the data model from a saved page HTML into `model.json`.
- `build.ps1` — builds the multi-sheet Excel workbook (Predictions, Points, Score Summary, By Game, Leaderboard).
- `update_workbook.ps1`, `add_summary.ps1`, `add_cards.ps1` — incremental Excel sheet builders.
- `build_static.ps1` / `build_html.ps1` — generate the offline/standalone HTML dashboard.
- `model.json` — a saved snapshot of the pool data (input for the above + local previews).

To get FRESH data for these: in Chrome (logged into pollaworldcup.com), DevTools → Network →
Copy as cURL of `https://www.pollaworldcup.com/ranking/detailed`, run it to save the page HTML,
then `parse.ps1`. (Or reuse the live-site auth approach in `build.ps1`.)

---

## Also in this folder
- **`PLAYBOOK - Web data to live dashboard.md`** — the general recipe + every gotcha, for building a
  SIMILAR project on a different site from scratch.

## Quick reminders
- Data only changes when matches are played, so "refresh" mostly matters on game days.
- GitHub's free cron is best-effort (can run late / skip), but only fires inside the 1pm–3am window.
- Times on the dashboard are **Panama/EST (UTC-5)**.
