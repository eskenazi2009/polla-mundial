#!/usr/bin/env pwsh
# Polla World Cup 2026 - live dashboard builder.
# Runs on GitHub Actions (PowerShell 7 / Ubuntu). Logs into pollaworldcup.com,
# fetches /ranking/detailed, extracts the embedded data model, and writes a
# fully self-contained (no-JS, embedded-flags) index.html into -OutDir.
param([string]$OutDir = "./public")

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13 }
catch { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }

$ref   = 'bqkkplrrlmuxylnibdho'                 # public Supabase project ref for pollaworldcup.com
$anon  = $env:SUPABASE_ANON_KEY
$email = $env:POLLA_EMAIL
$pw    = $env:POLLA_PASSWORD
if (-not $anon -or -not $email -or -not $pw) {
    throw "Missing required secrets. Set SUPABASE_ANON_KEY, POLLA_EMAIL and POLLA_PASSWORD as GitHub Actions secrets."
}
$ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36'

# ---------- 1) Authenticate (Supabase email/password grant) ----------
Write-Host "Authenticating..."
$tokenUrl = "https://$ref.supabase.co/auth/v1/token?grant_type=password"
$loginBody = @{ email = $email; password = $pw } | ConvertTo-Json -Compress
try {
    $authResp = Invoke-WebRequest -Uri $tokenUrl -Method Post -UseBasicParsing `
        -Headers @{ apikey = $anon; 'Content-Type' = 'application/json' } -Body $loginBody
} catch {
    throw "Login failed (check POLLA_EMAIL / POLLA_PASSWORD / SUPABASE_ANON_KEY). $($_.Exception.Message)"
}
$sessionJson = $authResp.Content
$session = $sessionJson | ConvertFrom-Json
if (-not $session.access_token) { throw "Login returned no access_token." }
Write-Host "Logged in as $($session.user.email)"

# ---------- 2) Rebuild the SSR auth cookie and fetch the page ----------
$b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($sessionJson))
$cookie = "sb-$ref-auth-token=base64-$b64"
Write-Host "Fetching /ranking/detailed..."
$page = Invoke-WebRequest -Uri 'https://www.pollaworldcup.com/ranking/detailed' -UseBasicParsing `
    -Headers @{ Cookie = $cookie; 'User-Agent' = $ua; 'Accept' = 'text/html,application/xhtml+xml' }
$pageHtml = $page.Content

# ---------- 3) Extract the embedded data model ----------
$rx = [regex]'self\.__next_f\.push\(\[1,"((?:[^"\\]|\\.)*)"\]\)'
$sbS = [System.Text.StringBuilder]::new()
foreach ($m in $rx.Matches($pageHtml)) {
    $real = ('"' + $m.Groups[1].Value + '"') | ConvertFrom-Json
    [void]$sbS.Append($real)
}
$stream = $sbS.ToString()
$idx = $stream.IndexOf('"model":')
if ($idx -lt 0) { throw "Data model not found - the auth cookie was likely rejected (got the login page instead)." }
$start = $stream.IndexOf('{', $idx)
$depth = 0; $inStr = $false; $esc = $false; $end = -1
for ($i = $start; $i -lt $stream.Length; $i++) {
    $ch = $stream[$i]
    if ($inStr) {
        if ($esc) { $esc = $false } elseif ($ch -eq '\') { $esc = $true } elseif ($ch -eq '"') { $inStr = $false }
    } else {
        if ($ch -eq '"') { $inStr = $true } elseif ($ch -eq '{') { $depth++ } elseif ($ch -eq '}') { $depth--; if ($depth -eq 0) { $end = $i; break } }
    }
}
if ($end -lt 0) { throw "Could not parse the data model." }
$model = $stream.Substring($start, $end - $start + 1) | ConvertFrom-Json
Write-Host "Parsed model: $($model.totalPools) participants, $($model.groupCols.Count) games."

# ===== TEMP DIAGNOSTIC (remove after) =====
try {
    $tmap = @{}; foreach ($p in $model.teams.PSObject.Properties) { $tmap[$p.Name] = $p.Value.code }
    Write-Host "DIAG groupPlayed=$(@($model.groupCols | Where-Object { $_.played }).Count)/$($model.groupCols.Count)"
    foreach ($k in $model.koCols) { $w = @($k.slots | Where-Object { $null -ne $_.actualWinnerId -and "$($_.actualWinnerId)" -ne '' }).Count; Write-Host "DIAG ko $($k.key) resolved=$($k.resolved) winners=$w" }
    foreach ($sd in @('r32','r16','qf')) {
        $vals = @{}; $byVal = @{}
        foreach ($r in $model.rows) { $arr = $r.$sd; if ($arr) { foreach ($e in $arr) { if ($e) { $v = "$($e[1])"; if (-not $vals.ContainsKey($v)) { $vals[$v] = 0; $byVal[$v] = @{} }; $vals[$v]++; $byVal[$v]["$($e[0])"] = $true } } } }
        Write-Host "DIAG $sd value->count: $(@($vals.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ' ')"
        foreach ($v in ($byVal.Keys | Sort-Object)) { Write-Host "DIAG $sd val=$v distinctTeams=$($byVal[$v].Count): $(@($byVal[$v].Keys | ForEach-Object { $tmap[$_] } | Sort-Object) -join ',')" }
    }
    $mr = $model.rows | Where-Object { $_.mine } | Select-Object -First 1
    if ($mr) { Write-Host "DIAG mine=$($mr.name) score=$($mr.score) r32: $(@($mr.r32 | ForEach-Object { if ($_) { ($tmap["$($_[0])"] + ':' + $_[1]) } }) -join ' ')" }
} catch { Write-Host "DIAG err $($_.Exception.Message)" }
# ===== END TEMP DIAGNOSTIC =====

# ---------- 4) Generate the static dashboard ----------
$teams = @{}; $flagData = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code }
$wc = New-Object System.Net.WebClient
foreach ($p in $model.teams.PSObject.Properties) {
    $url = $p.Value.flagUrl
    if (-not $url) { continue }
    $cc = [System.IO.Path]::GetFileName($url)
    try { $bytes = $wc.DownloadData("https://flagcdn.com/w80/$cc"); $flagData[$p.Name] = "data:image/png;base64," + [Convert]::ToBase64String($bytes) }
    catch { $flagData[$p.Name] = $url }
}
function GetCode($id) { if ($teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }
function GetFlag($id) { if ($flagData.ContainsKey("$id")) { $flagData["$id"] } else { "" } }
function Esc($s) { if ($null -eq $s) { return '' } "$s".Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;') }

$cols = $model.groupCols; $rows = $model.rows; $nG = $cols.Count
$me = $rows | Where-Object { $_.mine -eq $true } | Select-Object -First 1
$months = 'Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'

# ---- Leaderboard: official pool standings (rank + score exactly as on the site) ----
$board = @($rows | Sort-Object @{Expression = 'rank'; Descending = $false}, @{Expression = 'name'; Descending = $false})
$lb = New-Object System.Text.StringBuilder
foreach ($b in $board) {
    $cls = if ($b.mine) { 'lbrow me' } else { 'lbrow' }
    [void]$lb.Append("<div class='$cls'><span class='lrk'>$($b.rank)</span><span class='lname'>$(Esc $b.name)</span><span class='lpts'>$($b.score)</span></div>")
}
$lbHtml = "<details class='lb'><summary>Tabla de posiciones &middot; $($board.Count) jugadores</summary><div class='lbhead'><span class='lrk'>#</span><span class='lname'>Jugador</span><span class='lpts'>Pts</span></div><div class='lbbody'>$($lb.ToString())</div></details>"

$sb = New-Object System.Text.StringBuilder
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $hc = GetCode $c.homeId; $ac = GetCode $c.awayId
    $hf = GetFlag $c.homeId; $af = GetFlag $c.awayId
    $mine = ''
    if ($me) { $mg = $me.g[$j]; if ($mg -and $mg[0] -ne $null) { $mine = "$($mg[0])-$($mg[1])" } }
    $counts = @{}; $total = 0
    foreach ($r in $rows) {
        $g = $r.g[$j]
        if ($g -and $g.Count -ge 2 -and $g[0] -ne $null) {
            $s = "$($g[0])-$($g[1])"
            if ($counts.ContainsKey($s)) { $counts[$s]++ } else { $counts[$s] = 1 }
            $total++
        }
    }
    $sorted = @($counts.GetEnumerator() | Sort-Object -Property Value -Descending)
    $maxp = if ($sorted.Count) { [double]$sorted[0].Value / [math]::Max($total,1) * 100 } else { 100 }
    $date = ''; $gtop = "G$($j+1)"
    if ($c.kickoff) { try { $pa=[datetimeoffset]::Parse($c.kickoff).UtcDateTime.AddHours(-5); $date = "{0} {1} {2}, {3:00}:{4:00} (Panam&aacute;)" -f $pa.Day,$months[$pa.Month-1],$pa.Year,$pa.Hour,$pa.Minute; $gtop = "{0} {1} {2:00}:{3:00}" -f $pa.Day,$months[$pa.Month-1],$pa.Hour,$pa.Minute } catch {} }
    $resBadge = if ($c.played) { "<span class='badge bres'>Resultado $(Esc $c.actual)</span>" } else { "<span class='badge bpend'>Por jugar</span>" }
    $youBadge = if ($mine) { "<span class='badge byou'>Tu pick $(Esc $mine)</span>" } else { "" }
    $resTxt = if ($c.played) { Esc $c.actual } else { 'Por jugar' }
    $resCls = if ($c.played) { '' } else { ' pend' }
    $pickCls = 'pend'
    if ($c.played -and $mine -ne '') { $pickCls = if ($mine -eq $c.actual) { 'win' } else { 'lose' } }
    $pickHtml = if ($mine) { "<div class='gpick $pickCls'>T&uacute;: $(Esc $mine)</div>" } else { "<div class='gpick pend'>T&uacute;: -</div>" }
    $bars = New-Object System.Text.StringBuilder
    foreach ($kv in $sorted) {
        $sc = $kv.Key; $n = $kv.Value; $pct = [math]::Round(100.0*$n/[math]::Max($total,1),1)
        $isHit = ($c.played -and $sc -eq $c.actual); $isMine = ($sc -eq $mine -and $mine -ne '')
        $cls = 'drow'; if ($isHit) { $cls += ' hit' }; if ($isMine) { $cls += ' mine' }
        $w = [math]::Max(4.0, $pct/$maxp*100)
        $tags = ''
        if ($isHit) { $tags += "<span class='tg h'>&#10003;</span>" }
        if ($isMine) { $tags += "<span class='tg m'>T&Uacute;</span>" }
        [void]$bars.Append("<div class='$cls'><div class='barfill' style='width:$([math]::Round($w,1))%'></div><span class='score'>$(Esc $sc)</span><span class='tags'>$tags</span><span class='pct'>$pct%</span><span class='cnt'>$n pers.</span></div>")
    }
    $imgH = if ($hf) { "<img src='$hf' alt=''>" } else { '' }
    $imgA = if ($af) { "<img src='$af' alt=''>" } else { '' }
    $sfH = if ($hf) { "<img class='sflag' src='$hf' alt=''>" } else { '' }
    $sfA = if ($af) { "<img class='sflag' src='$af' alt=''>" } else { '' }
    [void]$sb.Append("<details class='game' name='games'><summary><div class='ghead'><span class='gid'>$gtop</span></div><div class='gmatch'>$sfH<span class='gm'>$hc</span><span class='gvs'>-</span><span class='gm'>$ac</span>$sfA</div><div class='gres$resCls'>$resTxt</div>$pickHtml</summary><div class='body'><div class='match'>$imgH<span>$hc</span><span class='vs'>vs</span><span>$ac</span>$imgA</div><div class='meta'>$resBadge$youBadge<span class='mdate'>$date</span><span class='mtot'>$total predicciones</span></div><div class='dist'>$($bars.ToString())</div></div></details>")
}

# ---- Champion + knockout-bracket prediction distributions ----
function StagePicks($row, $key) {
    if ($key -eq 'champion') {
        if ($row.champion -and $row.champion[0] -ne $null) { return @($row.champion[0][0]) } else { return @() }
    }
    return @($row.$key | ForEach-Object { if ($_ -ne $null) { $_[0] } })
}
$stages = @(
    @{ key='champion'; label='Campe&oacute;n';        small=$true  },
    @{ key='final';    label='Finalistas';            small=$true  },
    @{ key='sf';       label='Semifinalistas';        small=$true  },
    @{ key='qf';       label='Cuartos de final';      small=$false },
    @{ key='r16';      label='Octavos de final';      small=$false },
    @{ key='r32';      label='Ronda de 32';           small=$false }
)
$nPlayersAll = $rows.Count

# ---- Actual advancers per stage ("made it through"), from the site's own scoring ----
# Every bracket pick is [teamId, pointsEarned]; pointsEarned > 0 means that team actually
# reached that round (the site awards R32=1, R16=2, QF=4, SF=8 pts per correct team).
# Aggregating pts>0 across all players gives the authoritative set of teams that advanced,
# and it updates live as games are played. (This is the only signal that covers Ronda de 32 -
# the koCols bracket only resolves from Octavos onward.)
$advanced = @{}      # stage.key -> @{ "teamId" = $true }
$advResolved = @{}   # stage.key -> $true once at least one team has advanced
foreach ($stk in @('r32', 'r16', 'qf', 'sf', 'final', 'champion')) {
    $set = @{}
    foreach ($r in $rows) {
        if ($stk -eq 'champion') {
            if ($r.champion -and $null -ne $r.champion[0]) {
                $pair = $r.champion[0]
                if ($null -ne $pair[1] -and [double]$pair[1] -gt 0) { $set["$($pair[0])"] = $true }
            }
        } else {
            $arr = $r.$stk
            if ($arr) { foreach ($e in $arr) { if ($null -ne $e -and $null -ne $e[1] -and [double]$e[1] -gt 0) { $set["$($e[0])"] = $true } } }
        }
    }
    $advanced[$stk] = $set
    $advResolved[$stk] = ($set.Count -gt 0)
}

# A stage is "decided" only when its round is fully resolved - so a team I picked is shown
# as failed (red) only once it can no longer advance, never while it's still pending.
# r32 is decided when the group stage is over; later rounds when the site marks them resolved.
$koResolved = @{}
foreach ($k in $model.koCols) { $koResolved[$k.key] = [bool]$k.resolved }
$groupDone = (@($model.groupCols | Where-Object { -not $_.played }).Count -eq 0)
$stageDecided = @{
    r32      = $groupDone
    r16      = [bool]$koResolved['round_of_16']
    qf       = [bool]$koResolved['quarter']
    sf       = [bool]$koResolved['semi']
    final    = [bool]$koResolved['final']
    champion = [bool]$koResolved['champion']
}

$koSb = New-Object System.Text.StringBuilder
foreach ($st in $stages) {
    $counts = @{}
    foreach ($r in $rows) { foreach ($tid in (StagePicks $r $st.key)) { if ($tid -ne $null) { $k = "$tid"; if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 } } } }
    if ($counts.Count -eq 0) { continue }
    $mineSet = @{}; if ($me) { foreach ($t in (StagePicks $me $st.key)) { $mineSet["$t"] = $true } }
    $sorted = @($counts.GetEnumerator() | Sort-Object -Property Value -Descending)
    $maxc = $sorted[0].Value
    $bars = New-Object System.Text.StringBuilder
    foreach ($kv in $sorted) {
        $tid = $kv.Key; $n = $kv.Value; $pct = [math]::Round(100.0 * $n / $nPlayersAll, 1)
        $code = GetCode $tid; $fl = GetFlag $tid
        $isMine = $mineSet.ContainsKey($tid)
        $advSet = $advanced[$st.key]
        $isAdv  = ($advSet -and $advSet.ContainsKey($tid))
        # Bar color (knockout): advanced+mine=green, advanced+not-mine=red, picked-but-out=red,
        # picked-and-still-pending=blue, not-picked-not-through=default.
        $bc = ''
        if ($isAdv -and $isMine) { $bc = ' bc-green' }
        elseif ($isAdv) { $bc = ' bc-red' }
        elseif ($isMine -and $stageDecided[$st.key]) { $bc = ' bc-red' }
        elseif ($isMine) { $bc = ' bc-blue' }
        $cls = 'drow' + $bc
        $w = [math]::Max(4.0, [double]$n / $maxc * 100)
        $img = if ($fl) { "<img class='kflag' src='$fl' alt=''>" } else { '' }
        $tag = ''
        if ($isAdv) { if ($isMine) { $tag += "<span class='tg a'>&#10003;</span>" } else { $tag += "<span class='tg miss'>&#10003;</span>" } }
        if ($isMine) { $tag += "<span class='tg m'>T&Uacute;</span>" }
        [void]$bars.Append("<div class='$cls'><div class='barfill' style='width:$([math]::Round($w,1))%'></div>$img<span class='score kteam'>$code</span><span class='tags'>$tag</span><span class='pct'>$pct%</span><span class='cnt'>$n pers.</span></div>")
    }
    $advSet = $advanced[$st.key]
    $nAdv = if ($advSet) { $advSet.Count } else { 0 }
    $isResolved = ($advResolved[$st.key] -and $nAdv -gt 0)
    $myHits = 0; if ($advSet) { foreach ($t in $mineSet.Keys) { if ($advSet.ContainsKey($t)) { $myHits++ } } }

    $chip = ''
    if ($isResolved) { $chip += "<span class='chip cres'>&#10003; $nAdv clasificados</span>" }
    if ($st.small -and $me) {
        $mp = @(StagePicks $me $st.key | ForEach-Object { GetCode $_ })
        $lbl = if ($mp.Count) { $mp -join ', ' } else { '-' }
        $chip += "<span class='chip cyou'>T&uacute;: $lbl</span>"
    }
    if ($isResolved -and $me) { $chip += "<span class='chip cyou'>Aciertos: $myHits</span>" }

    $legend = if ($isResolved) { "<span class='mtot'>verde: la elegiste y pas&oacute; &middot; rojo: pas&oacute; sin elegirla, o la elegiste y qued&oacute; fuera &middot; azul: tu pick a&uacute;n en juego</span>" } else { '' }
    [void]$koSb.Append("<details class='game' name='ko'><summary><span class='gid'>KO</span><span class='gm'>$($st.label)</span>$chip</summary><div class='body'><div class='meta'><span class='mtot'>equipos m&aacute;s elegidos para llegar a esta ronda &middot; $nPlayersAll jugadores</span>$legend</div><div class='dist'>$($bars.ToString())</div></div></details>")
}
$koHtml = "<div class='kohead'>Eliminatorias &mdash; pron&oacute;sticos del bracket</div>" + $koSb.ToString()

$uName = if ($me) { Esc $me.name } else { '-' }
$uRank = if ($me) { $me.rank } else { 0 }
$uScore = if ($me) { $me.score } else { 0 }
$nPlayers = $rows.Count
$genPa = (Get-Date).ToUniversalTime().AddHours(-5)   # Panama = UTC-5, no DST
$gen = "{0} {1} {2:00}:{3:00} (Panam&aacute;)" -f $genPa.Day, $months[$genPa.Month - 1], $genPa.Hour, $genPa.Minute

$html = @"
<!DOCTYPE html>
<html lang='es'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Polla World Cup 2026 &mdash; Predicciones</title>
<style>
  :root{--bg:#0f1623;--card:#182030;--card2:#1f2a3d;--line:#2c3a52;--txt:#e8eef7;--mut:#90a0b8;--bar:#2b3b57}
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--txt)}
  header{padding:14px 16px;border-bottom:1px solid var(--line);background:linear-gradient(180deg,#141d2e,#0f1623)}
  h1{font-size:17px;margin:0 0 2px}
  .sub{font-size:12px;color:var(--mut)}
  .me{margin-top:10px;display:flex;gap:9px;flex-wrap:wrap;font-size:14px}
  .pill{background:var(--card2);border:1px solid var(--line);border-radius:999px;padding:7px 13px}
  .pill b{color:#fff;font-size:18px}
  .hint{padding:10px 16px;color:var(--mut);font-size:12px;border-bottom:1px solid var(--line)}
  main{max-width:1200px;margin:0 auto;padding:10px 16px 30px}
  details.game{background:var(--card);border:1px solid var(--line);border-radius:12px;margin-top:8px;overflow:hidden}
  summary{list-style:none;cursor:pointer;padding:12px 14px;display:flex;align-items:center;gap:8px;font-size:14px}
  summary::-webkit-details-marker{display:none}
  .gid{font-weight:800;color:var(--mut);min-width:30px}
  .gm{font-weight:800;font-size:15px}
  .chip{margin-left:auto;font-size:11px;font-weight:700;padding:3px 8px;border-radius:999px}
  .chip+.chip{margin-left:6px}
  .cres{background:#13351f;color:#7ee2a0;border:1px solid #2f6b45}
  .cpend{background:#2a2333;color:#caa6e8;border:1px solid #4a3a5e}
  .cyou{background:#16263f;color:#8fb6ff;border:1px solid #2f4f86}
  .body{padding:6px 14px 16px;border-top:1px solid var(--line)}
  .match{display:flex;align-items:center;justify-content:center;gap:12px;font-size:22px;font-weight:800;padding:10px 0}
  .match img{width:32px;height:22px;border-radius:3px;object-fit:cover;box-shadow:0 0 0 1px #0006}
  .vs{color:var(--mut);font-size:13px}
  .meta{display:flex;justify-content:center;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--mut);margin-bottom:6px}
  .badge{padding:3px 9px;border-radius:999px;font-weight:700}
  .bres{background:#13351f;color:#7ee2a0;border:1px solid #2f6b45}
  .bpend{background:#2a2333;color:#caa6e8;border:1px solid #4a3a5e}
  .byou{background:#16263f;color:#8fb6ff;border:1px solid #2f4f86}
  .drow{position:relative;display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:9px;margin-top:6px;background:var(--card2);overflow:hidden}
  .barfill{position:absolute;left:0;top:0;bottom:0;background:var(--bar);z-index:0}
  .drow.hit .barfill{background:#1f5a35}
  .drow.mine .barfill{background:#23426e}
  .drow.bc-green .barfill{background:#1f6b3f}
  .drow.bc-red .barfill{background:#6e2233}
  .drow.bc-blue .barfill{background:#23426e}
  .drow>:not(.barfill){position:relative;z-index:1}
  .score{font-weight:800;font-size:16px;min-width:46px}
  .tags{display:flex;gap:5px}
  .tg{font-size:10px;font-weight:700;padding:2px 6px;border-radius:6px}
  .tg.h{background:#22c55e;color:#06250f}
  .tg.a{background:#22c55e;color:#06250f}
  .tg.miss{background:#ef4444;color:#2b0606}
  .tg.m{background:#3b82f6;color:#04122b}
  .pct{margin-left:auto;font-weight:800;font-size:16px}
  .cnt{color:var(--mut);font-size:12px;min-width:62px;text-align:right}
  .foot{color:var(--mut);font-size:11px;text-align:center;padding:18px}
  details.lb{background:var(--card);border:1px solid var(--line);border-radius:12px;margin:12px auto 0;max-width:580px;overflow:hidden}
  details.lb[open]{max-width:760px}
  details.lb>summary{list-style:none;cursor:pointer;padding:16px 18px;font-weight:700;font-size:18px;text-align:center}
  details.lb>summary::-webkit-details-marker{display:none}
  .lbhead,.lbrow{display:flex;align-items:center;gap:12px;padding:10px 18px;font-size:16px}
  .lbhead{color:var(--mut);font-weight:700;border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
  .lbbody{max-height:72vh;overflow:auto}
  .lbrow:nth-child(even){background:#1b2536}
  .lbrow.me{background:#16263f;font-weight:800;color:#cfe0ff}
  .lrk{min-width:52px;color:var(--mut);font-weight:700}
  .lbrow.me .lrk{color:#8fb6ff}
  .lname{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .lpts{font-weight:800;min-width:46px;text-align:right}
  .kohead{margin:22px 2px 2px;font-size:15px;font-weight:800;color:#cdd9ec;padding-top:10px;border-top:1px solid var(--line)}
  .kflag{width:24px;height:16px;border-radius:2px;object-fit:cover;box-shadow:0 0 0 1px #0006}
  .kteam{min-width:48px}
  .grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;align-items:start;margin-top:10px}
  .grid details.game{margin-top:0}
  .grid details.game[open]{grid-column:1 / -1}
  .grid summary{flex-direction:column;align-items:stretch;text-align:center;gap:9px;padding:20px 12px}
  .ghead{display:flex;justify-content:center;font-size:13px}
  .grid .gid{min-width:0}
  .grid .gm{font-size:19px}
  .gmatch{display:flex;align-items:center;justify-content:center;gap:7px;flex-wrap:wrap}
  .sflag{width:26px;height:17px;border-radius:2px;object-fit:cover;box-shadow:0 0 0 1px #0006}
  .gvs{color:var(--mut);font-size:14px;font-weight:700}
  .gres{font-size:26px;font-weight:800}
  .gres.pend{font-size:14px;font-weight:600;color:var(--mut)}
  .gpick{font-size:19px;font-weight:800;border-radius:8px;padding:6px 10px}
  .gpick.win{background:#13351f;color:#7ee2a0}
  .gpick.lose{background:#3a1620;color:#ff9aa6}
  .gpick.pend{background:#16263f;color:#8fb6ff}
</style>
</head>
<body>
<header>
  <h1>Polla World Cup 2026 &mdash; Predicciones</h1>
  <div class='sub'>Qui&eacute;n predijo qu&eacute;, en cada partido</div>
  <div class='me'>
    <span class='pill'>T&uacute;: <b>$uName</b></span>
    <span class='pill'>Rank <b>#$uRank</b> / $nPlayers</span>
    <span class='pill'>Puntos <b>$uScore</b></span>
  </div>
  $lbHtml
</header>
<div class='hint'>Toca un partido para ver todas las predicciones. En eliminatorias: Verde = la elegiste y clasific&oacute; &middot; Rojo = clasific&oacute; sin elegirla, o tu pick qued&oacute; fuera &middot; Azul = tu pick a&uacute;n en juego. Puntos y posiciones tomados en vivo del sitio oficial. Se actualiza autom&aacute;ticamente.</div>
<main>
<div class='grid'>
$($sb.ToString())
</div>
$koHtml
</main>
<div class='foot'>$nG partidos &middot; actualizado $gen</div>
</body>
</html>
"@

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$outFile = Join-Path $OutDir 'index.html'
[System.IO.File]::WriteAllText($outFile, $html, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Wrote $outFile ($([math]::Round((Get-Item $outFile).Length/1KB,1)) KB)"
