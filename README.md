# Polla World Cup 2026 — live dashboard

Auto-updating, always-on dashboard of everyone's predictions in the
[pollaworldcup.com](https://www.pollaworldcup.com) pool. A GitHub Actions cron
job logs in, fetches the latest data, regenerates a self-contained `index.html`
(no JavaScript, country flags embedded) and publishes it to GitHub Pages.

## One-time setup

### 1. Create the repo
Create a **public** repository on GitHub (Pages is free on public repos) and push
these files to it (see commands at the bottom).

### 2. Add your secrets
Repo **Settings → Secrets and variables → Actions → New repository secret**, add three:

| Name | Value |
|------|-------|
| `POLLA_EMAIL` | your pollaworldcup.com login email |
| `POLLA_PASSWORD` | your pollaworldcup.com password |
| `SUPABASE_ANON_KEY` | the site's public API key (see below) |

**Finding `SUPABASE_ANON_KEY`:** open pollaworldcup.com in Chrome while logged in →
DevTools (F12) → **Network** tab → reload → click any request to
`bqkkplrrlmuxylnibdho.supabase.co` and copy the **`apikey`** request header.
(If you don't see Supabase requests, open the **Console** tab and the long
`eyJ...` token printed by the app's config is the same key. It is a *publishable*
key — safe to store, but a secret keeps it tidy.)

### 3. Enable Pages
Repo **Settings → Pages → Build and deployment → Source: GitHub Actions**.

### 4. Run it
**Actions** tab → "Build Polla dashboard" → **Run workflow**. After it finishes,
your live URL appears under Settings → Pages (e.g. `https://<you>.github.io/<repo>/`).
Share that link — it refreshes itself every 4 hours.

## Adjusting the schedule
Edit the `cron` line in `.github/workflows/build.yml`. `0 */4 * * *` = every 4 hours.
Use `0 */2 * * *` for every 2 hours, etc. (times are UTC).

## Push commands
```bash
git remote add origin https://github.com/<you>/<repo>.git
git branch -M main
git push -u origin main
```
