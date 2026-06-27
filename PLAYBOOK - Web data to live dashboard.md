# PLAYBOOK — Turn a login-gated website's data into Excel + a live, free, auto-updating web dashboard

This is a reusable recipe based on the **Polla World Cup 2026** project. Hand this file to Claude
next time and say "do this for <new site>". It captures the pipeline AND the gotchas that cost time.

Reference implementation lives in:
- Data/Excel scripts: `C:\Users\esken\Downloads\...\rebundle\polla-worldcup\`
- Live site repo: `C:\Users\esken\polla-worldcup-live\`  →  GitHub `eskenazi2009/polla-mundial`  →  https://eskenazi2009.github.io/polla-mundial/

---

## The overall pattern (what we built)
1. **Access** the data behind a login.
2. **Extract** the data into a clean JSON model.
3. **Excel** workbook for analysis (Windows, via Excel COM).
4. **Shareable HTML dashboard** (no-JS, offline-capable) to send on WhatsApp.
5. **Live version**: GitHub Actions rebuilds it on a schedule and publishes to GitHub Pages for free.

---

## PHASE 0 — Access / reconnaissance
- First check if the page is public: a normal fetch. If you get a login form, it's gated.
- **Get in via the browser session, not by storing logins (at first):** open the site logged in →
  Chrome **DevTools → Network → Copy as cURL** of the data request. That cURL carries the auth cookie.
- **Identify the stack.** Ours was **Next.js + Supabase**. Tell-tale: cookie named
  `sb-<projectref>-auth-token=base64-<base64 JSON session>`. The project ref (`bqkkplrrlmuxylnibdho`)
  is also the Supabase host: `https://<ref>.supabase.co`.
- **KEY INSIGHT that saved us:** a server-rendered (SSR) page often **embeds the entire dataset inside
  the HTML** in Next.js RSC chunks (`self.__next_f.push([1,"..."])`). So there is **no separate API/XHR
  call** to capture ("no supabase requests in Network" is expected). Just fetch the page and parse it.
- The "prefetch" RSC response (`?_rsc=...` with `next-router-prefetch:1`) is only the shell — fetch the
  **full page** (normal document request) to get the data.

## PHASE 1 — Extract the data model (PowerShell)
The data is a JSON object embedded in escaped strings. Reconstruct + brace-match it:
```powershell
$rx = [regex]'self\.__next_f\.push\(\[1,"((?:[^"\\]|\\.)*)"\]\)'
$sb = [System.Text.StringBuilder]::new()
foreach ($m in $rx.Matches($pageHtml)) { [void]$sb.Append( ('"'+$m.Groups[1].Value+'"') | ConvertFrom-Json ) }
$stream = $sb.ToString()
$idx = $stream.IndexOf('"model":'); $start = $stream.IndexOf('{', $idx)
$depth=0;$inStr=$false;$esc=$false;$end=-1
for($i=$start;$i -lt $stream.Length;$i++){ $ch=$stream[$i]
  if($inStr){ if($esc){$esc=$false}elseif($ch -eq '\'){$esc=$true}elseif($ch -eq '"'){$inStr=$false} }
  else{ if($ch -eq '"'){$inStr=$true}elseif($ch -eq '{'){$depth++}elseif($ch -eq '}'){$depth--;if($depth -eq 0){$end=$i;break}} } }
$model = $stream.Substring($start, $end-$start+1) | ConvertFrom-Json
```
Then inspect `$model.PSObject.Properties.Name` and sample a row to learn the schema before coding.

## PHASE 2 — Excel workbook (Windows, Excel COM)
This machine has **no real Python/Node** (the `python.exe` on PATH is the Microsoft Store stub) and no
ImportExcel module — but **Excel COM works**. Build a 2D `object[,]` array and assign it to a range in
one shot (fast), then format.
**PowerShell 5.1 gotchas that bit us repeatedly:**
- **Comma binds tighter than `+`** → `$arr[$i+1, $c+1]` parses as `$arr[$i + (1,$c) + 1]`. ALWAYS
  precompute indices: `$ri=$i+1; $cc=$c+1; $arr[$ri,$cc]=...`.
- `$home` is the read-only automatic `$HOME` → rename your variable (`$hc`).
- `fl` / `FL` is an alias for `Format-List` → don't name a function `FL`.
- Strings like `"2-0"` get **auto-converted to dates** by Excel → set `Range.NumberFormat='@'` (text)
  **before** writing those cells.
- Excel COM throws transient `0x80080005` if a previous instance is still releasing → `Stop-Process
  -Name EXCEL`, brief wait, and retry `New-Object -ComObject Excel.Application` in a loop.
- Save as `.xlsx`: `$wb.SaveAs($path, 51)`; release COM objects + `[GC]::Collect()`.

## PHASE 3 — Shareable HTML dashboard (no JavaScript!)
- **WhatsApp / email / phone "Files" previews STRIP JavaScript.** A JS-driven page shows an empty shell
  ("template loads but no data"). So **pre-render everything as static HTML** — no `<script>`.
- Use `<details><summary>` cards for tap-to-expand "filtering" with zero JS.
- **Exclusive accordion (one open at a time), still no JS:** give all `<details>` the same `name="..."`
  attribute (native HTML behavior in modern browsers).
- **Offline:** embed images as base64 data URIs (we downloaded the flags and inlined them) so it works
  with no internet.
- **Encoding trap:** PowerShell 5.1 reads `.ps1` files as ANSI, so accented letters/symbols in a here-
  string get mangled in the output. Keep generated HTML **ASCII-only**: use HTML entities
  (`&aacute;`, `&mdash;`, `&#10003;`) for static text, and plain ASCII separators in JS `textContent`.
- **CSS bar-overlay trap:** a rule like `.row > * {position:relative}` will **override**
  `.bar{position:absolute}` (equal specificity, later wins) and knock the bar back into flow, shoving
  content sideways. Scope it: `.row > :not(.bar){position:relative;z-index:1}`.
- Write files without a BOM for web: `[System.IO.File]::WriteAllText($p,$html,(New-Object System.Text.UTF8Encoding($false)))`.
- Times: convert UTC to the user's zone (Panama = **UTC-5, no DST** → just `.AddHours(-5)`).

## PHASE 4 — Make it live & free (GitHub Actions + GitHub Pages)
- **Repo must be public** (Pages is free on public repos). Secrets are still encrypted and safe.
- The build script **self-authenticates** instead of using a captured cookie (which expires hourly):
  ```powershell
  # Supabase email/password grant
  $r = Invoke-WebRequest "https://$ref.supabase.co/auth/v1/token?grant_type=password" -Method Post `
       -Headers @{ apikey=$env:SUPABASE_ANON_KEY; 'Content-Type'='application/json' } `
       -Body (@{ email=$env:POLLA_EMAIL; password=$env:POLLA_PASSWORD } | ConvertTo-Json)
  $sessionJson = $r.Content
  # Rebuild the SSR cookie from the session JSON, then fetch the page:
  $cookie = "sb-$ref-auth-token=base64-" + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($sessionJson))
  $page = Invoke-WebRequest 'https://.../ranking/detailed' -Headers @{ Cookie=$cookie; 'User-Agent'='Mozilla/5.0...' }
  ```
- **Secrets** to add in repo Settings → Secrets and variables → Actions:
  `POLLA_EMAIL`, `POLLA_PASSWORD`, `SUPABASE_ANON_KEY`.
- **Finding the anon key:** it's a **public** publishable key (`role:anon` JWT, valid for years) embedded
  in the site's JS bundles. Grab it from DevTools → Network → any `*.supabase.co` request's `apikey`
  header during login, OR fetch the site's `/_next/static/chunks/*.js` and grep for the
  `eyJ...` JWT next to `createBrowserClient`.
- **Workflow** (`.github/workflows/build.yml`): triggers on `push` + `schedule` (cron) + `workflow_dispatch`;
  `runs-on: ubuntu-latest`; `shell: pwsh` runs `build.ps1`; then `actions/upload-pages-artifact` +
  `actions/deploy-pages`. Needs `permissions: pages:write, id-token:write` and Pages **Source = GitHub Actions**.
- Cron example `0 */2 * * *` = every 2 hours (UTC). GitHub pauses schedules after 60 days of no commits.
- **Cross-platform note:** the runner has PowerShell 7 (`pwsh`), which doesn't have the 5.1 ANSI/encoding
  or 2MB ConvertFrom-Json issues — but keep the script 5.1-compatible so you can test locally too.

## Testing locally without credentials
To preview generation changes without doing a live login, run the generation half against a saved
`model.json`: take everything from the `# ---- Generate ----` marker onward, prepend
`$OutDir=...; $model = (Get-Content model.json -Raw) | ConvertFrom-Json`, run, open in Chrome.
(To preview an expanded card, string-replace the first `<details ...>` to add ` open`.)

## Caching note
GitHub Pages + the browser cache the HTML. After a deploy, **hard-refresh (Ctrl+Shift+R)** or append
`?v=N` to the URL. Friends opening a fresh link always get the latest.

---

## Quick checklist for a NEW similar project
1. [ ] Confirm the site is gated; capture a `Copy as cURL` of the data page.
2. [ ] Identify stack + whether data is embedded in the page HTML (SSR) or fetched via API.
3. [ ] Extract the data model to `model.json`; learn the schema.
4. [ ] (Optional) Build the Excel via Excel COM — mind the 5.1 gotchas.
5. [ ] Build the no-JS, offline HTML dashboard (entities, base64 images, `<details name>` accordion).
6. [ ] Find the public anon/API key; create a PUBLIC GitHub repo; add secrets.
7. [ ] Add `build.ps1` (self-auth + fetch + generate) and the Pages workflow; enable Pages = Actions.
8. [ ] Run the workflow; share the `*.github.io` link.
