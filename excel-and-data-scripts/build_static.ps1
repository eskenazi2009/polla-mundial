$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$teams = @{}; $flags = @{}; $flagData = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code; $flags[$p.Name] = $p.Value.flagUrl }

# Download each flag once and embed as a base64 data URI (so the file works fully offline)
$wc = New-Object System.Net.WebClient
$dl = 0; $fail = 0
foreach ($p in $model.teams.PSObject.Properties) {
    $id = $p.Name; $url = $p.Value.flagUrl
    if (-not $url) { continue }
    $cc = [System.IO.Path]::GetFileName($url)            # e.g. mx.png
    $u2 = "https://flagcdn.com/w80/$cc"                   # normalize to width 80
    try {
        $bytes = $wc.DownloadData($u2)
        $flagData[$id] = "data:image/png;base64," + [Convert]::ToBase64String($bytes)
        $dl++
    } catch {
        $flagData[$id] = $url                            # fallback to remote URL
        $fail++
    }
}
Write-Host "flags embedded: $dl  (failed/remote-fallback: $fail)"

function GetCode($id) { if ($teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }
function GetFlag($id) { if ($flagData.ContainsKey("$id")) { $flagData["$id"] } else { "" } }
function Esc($s) { if ($null -eq $s) { return '' } $s = "$s"; $s = $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;'); $s }

$cols = $model.groupCols; $rows = $model.rows; $nG = $cols.Count
$me = $rows | Where-Object { $_.mine -eq $true } | Select-Object -First 1
$months = 'Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'

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

    $date = ''
    if ($c.kickoff) { try { $dt=[datetimeoffset]::Parse($c.kickoff).UtcDateTime; $date = "{0} {1} {2}, {3:00}:{4:00}" -f $dt.Day,$months[$dt.Month-1],$dt.Year,$dt.Hour,$dt.Minute } catch {} }

    $resBadge = if ($c.played) { "<span class='badge bres'>Resultado $(Esc $c.actual)</span>" } else { "<span class='badge bpend'>Por jugar</span>" }
    $youBadge = if ($mine) { "<span class='badge byou'>Tu pick $(Esc $mine)</span>" } else { "" }
    # summary chips
    $sumRes = if ($c.played) { "<span class='chip cres'>$(Esc $c.actual)</span>" } else { "<span class='chip cpend'>--</span>" }
    $sumYou = if ($mine) { "<span class='chip cyou'>T&uacute;: $(Esc $mine)</span>" } else { "" }

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

    $openAttr = if ($j -eq 0) { ' open' } else { '' }
    $imgH = if ($hf) { "<img src='$hf' alt=''>" } else { '' }
    $imgA = if ($af) { "<img src='$af' alt=''>" } else { '' }
    [void]$sb.Append(@"
<details class='game'$openAttr>
<summary><span class='gid'>G$($j+1)</span><span class='gm'>$hc-$ac</span>$sumRes$sumYou</summary>
<div class='body'>
  <div class='match'>$imgH<span>$hc</span><span class='vs'>vs</span><span>$ac</span>$imgA</div>
  <div class='meta'>$resBadge$youBadge<span class='mdate'>$date</span><span class='mtot'>$total predicciones</span></div>
  <div class='dist'>$($bars.ToString())</div>
</div>
</details>
"@)
}

$uName = if ($me) { Esc $me.name } else { '-' }
$uRank = if ($me) { $me.rank } else { 0 }
$uScore = if ($me) { $me.score } else { 0 }
$nPlayers = $rows.Count
$gen = Get-Date -Format 'yyyy-MM-dd HH:mm'

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
  .me{margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;font-size:12px}
  .pill{background:var(--card2);border:1px solid var(--line);border-radius:999px;padding:4px 10px}
  .pill b{color:#fff}
  .hint{padding:10px 16px;color:var(--mut);font-size:12px;border-bottom:1px solid var(--line)}
  main{max-width:680px;margin:0 auto;padding:10px 12px 30px}
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
  .drow>*{position:relative;z-index:1}
  .score{font-weight:800;font-size:16px;min-width:46px}
  .tags{display:flex;gap:5px}
  .tg{font-size:10px;font-weight:700;padding:2px 6px;border-radius:6px}
  .tg.h{background:#22c55e;color:#06250f}
  .tg.m{background:#3b82f6;color:#04122b}
  .pct{margin-left:auto;font-weight:800;font-size:16px}
  .cnt{color:var(--mut);font-size:12px;min-width:62px;text-align:right}
  .foot{color:var(--mut);font-size:11px;text-align:center;padding:18px}
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
</header>
<div class='hint'>Toca un partido para ver todas las predicciones. Verde = resultado real &#10003; &middot; Azul = tu pick.</div>
<main>
$($sb.ToString())
</main>
<div class='foot'>$nG partidos &middot; generado $gen</div>
</body>
</html>
"@

$out = Join-Path $dir 'polla_dashboard.html'
Set-Content -Path $out -Value $html -Encoding UTF8
Write-Host "WROTE static $out  ($([math]::Round((Get-Item $out).Length/1KB,1)) KB, $nG games, no JavaScript)"
