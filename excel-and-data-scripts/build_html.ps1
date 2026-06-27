$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

$teams = @{}; $flags = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code; $flags[$p.Name] = $p.Value.flagUrl }
function GetCode($id) { if ($teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }
function GetFlag($id) { if ($flags.ContainsKey("$id")) { $flags["$id"] } else { "" } }

$cols = $model.groupCols; $rows = $model.rows; $nG = $cols.Count
$me = $rows | Where-Object { $_.mine -eq $true } | Select-Object -First 1

$games = @()
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $counts = @{}; $total = 0
    foreach ($r in $rows) {
        $g = $r.g[$j]
        if ($g -and $g.Count -ge 2 -and $g[0] -ne $null) {
            $s = "$($g[0])-$($g[1])"
            if ($counts.ContainsKey($s)) { $counts[$s]++ } else { $counts[$s] = 1 }
            $total++
        }
    }
    $dist = @($counts.GetEnumerator() | Sort-Object -Property Value -Descending | ForEach-Object {
        [pscustomobject]@{ s = $_.Key; n = $_.Value; p = [math]::Round(100.0 * $_.Value / [math]::Max($total,1), 1) }
    })
    $mine = ''
    if ($me) { $mg = $me.g[$j]; if ($mg -and $mg[0] -ne $null) { $mine = "$($mg[0])-$($mg[1])" } }
    $games += [pscustomobject]@{
        id = "G$($j+1)"; n = $j+1
        home = (GetCode $c.homeId); away = (GetCode $c.awayId)
        homeFlag = (GetFlag $c.homeId); awayFlag = (GetFlag $c.awayId)
        kickoff = $c.kickoff; played = [bool]$c.played
        actual = if ($c.played) { $c.actual } else { '' }
        total = $total; mine = $mine; dist = $dist
    }
}

$payload = [pscustomobject]@{
    user = [pscustomobject]@{ name = $(if($me){$me.name}else{''}); rank = $(if($me){$me.rank}else{0}); score = $(if($me){$me.score}else{0}); totalPlayers = $rows.Count }
    generated = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    games = $games
}
$json = $payload | ConvertTo-Json -Depth 8 -Compress

$html = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>Polla World Cup 2026 &mdash; Predicciones</title>
<style>
  :root{ --bg:#0f1623; --card:#182030; --card2:#1f2a3d; --line:#2c3a52; --txt:#e8eef7; --mut:#90a0b8;
         --accent:#3b82f6; --green:#22c55e; --greenbg:#13351f; --bluebg:#16263f; --bar:#2b3b57; }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
       background:var(--bg);color:var(--txt);-webkit-text-size-adjust:100%}
  header{position:sticky;top:0;z-index:10;background:linear-gradient(180deg,#141d2e,#0f1623);
         border-bottom:1px solid var(--line);padding:12px 16px}
  h1{font-size:17px;margin:0 0 2px;font-weight:700}
  .sub{font-size:12px;color:var(--mut)}
  .me{margin-top:8px;display:flex;gap:10px;flex-wrap:wrap;font-size:12px}
  .pill{background:var(--card2);border:1px solid var(--line);border-radius:999px;padding:4px 10px}
  .pill b{color:#fff}
  .controls{padding:12px 16px;position:sticky;top:96px;background:var(--bg);z-index:9;border-bottom:1px solid var(--line)}
  .row{display:flex;gap:8px;align-items:center}
  select,input{font-size:16px;padding:11px 12px;border-radius:10px;border:1px solid var(--line);
               background:var(--card);color:var(--txt);width:100%}
  .nav{display:flex;gap:8px;margin-top:8px}
  .nav button{flex:1;font-size:15px;padding:11px;border-radius:10px;border:1px solid var(--line);
              background:var(--card2);color:var(--txt);font-weight:600}
  .nav button:active{background:var(--accent)}
  .filterToggle{margin-top:8px;font-size:12px;color:var(--mut);display:flex;align-items:center;gap:6px}
  main{padding:16px;max-width:680px;margin:0 auto}
  .gcard{background:var(--card);border:1px solid var(--line);border-radius:16px;overflow:hidden}
  .ghead{padding:16px;background:linear-gradient(180deg,#1d2740,#18223550)}
  .match{display:flex;align-items:center;justify-content:center;gap:14px;font-size:24px;font-weight:800}
  .match img{width:34px;height:24px;border-radius:3px;object-fit:cover;box-shadow:0 0 0 1px #0006}
  .vs{color:var(--mut);font-size:14px;font-weight:600}
  .meta{display:flex;justify-content:center;gap:14px;margin-top:10px;font-size:12px;color:var(--mut);flex-wrap:wrap}
  .badge{padding:3px 9px;border-radius:999px;font-weight:700;font-size:12px}
  .b-actual{background:var(--greenbg);color:#7ee2a0;border:1px solid #2f6b45}
  .b-you{background:var(--bluebg);color:#8fb6ff;border:1px solid #2f4f86}
  .b-pend{background:#2a2333;color:#caa6e8;border:1px solid #4a3a5e}
  .dist{padding:8px 12px 16px}
  .drow{position:relative;display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:9px;margin-top:6px;
        background:var(--card2);overflow:hidden}
  .barfill{position:absolute;left:0;top:0;bottom:0;background:var(--bar);z-index:0}
  .drow.hit .barfill{background:#1f5a35}
  .drow.mine .barfill{background:#23426e}
  .drow>*{position:relative;z-index:1}
  .score{font-weight:800;font-size:16px;min-width:46px}
  .tags{display:flex;gap:5px;margin-left:2px}
  .tg{font-size:10px;font-weight:700;padding:2px 6px;border-radius:6px}
  .tg.h{background:#22c55e;color:#06250f}
  .tg.m{background:#3b82f6;color:#04122b}
  .pct{margin-left:auto;font-weight:800;font-size:16px}
  .cnt{color:var(--mut);font-size:12px;min-width:64px;text-align:right}
  .empty{color:var(--mut);text-align:center;padding:20px}
  .foot{color:var(--mut);font-size:11px;text-align:center;padding:18px}
</style>
</head>
<body>
<header>
  <h1>Polla World Cup 2026 &mdash; Predicciones</h1>
  <div class="sub">Qui&eacute;n predijo qu&eacute;, en cada partido</div>
  <div class="me" id="me"></div>
</header>
<div class="controls">
  <div class="row">
    <select id="game"></select>
  </div>
  <div class="nav">
    <button id="prev">&lsaquo; Anterior</button>
    <button id="next">Siguiente &rsaquo;</button>
  </div>
  <label class="filterToggle"><input type="checkbox" id="onlyPlayed" style="width:auto"> Solo partidos jugados</label>
</div>
<main><div id="view"></div></main>
<div class="foot" id="foot"></div>
<script>
const DATA = __JSON__;
const $ = s => document.querySelector(s);
const sel = $('#game');
let idx = 0;

function me(){
  const u = DATA.user;
  $('#me').innerHTML =
    `<span class="pill">T&uacute;: <b>${u.name||'-'}</b></span>`+
    `<span class="pill">Rank <b>#${u.rank}</b> / ${u.totalPlayers}</span>`+
    `<span class="pill">Puntos <b>${u.score}</b></span>`;
  $('#foot').textContent = `${DATA.games.length} partidos | generado ${DATA.generated}`;
}

function options(){
  const only = $('#onlyPlayed').checked;
  sel.innerHTML='';
  DATA.games.forEach((g,i)=>{
    if(only && !g.played) return;
    const o=document.createElement('option');
    o.value=i;
    o.textContent=`${g.id}  |  ${g.home}-${g.away}` + (g.played?`  (${g.actual})`:'');
    sel.appendChild(o);
  });
  if(!sel.querySelector(`option[value="${idx}"]`)){ idx = sel.options.length?+sel.options[0].value:0; }
  sel.value = idx;
}

function render(){
  const g = DATA.games[idx];
  if(!g){ $('#view').innerHTML='<div class="empty">Sin partido</div>'; return; }
  const date = g.kickoff ? new Date(g.kickoff).toLocaleString('es', {dateStyle:'medium', timeStyle:'short'}) : '';
  const statusBadge = g.played
    ? `<span class="badge b-actual">Resultado ${g.actual}</span>`
    : `<span class="badge b-pend">Por jugar</span>`;
  const youBadge = g.mine ? `<span class="badge b-you">Tu pick ${g.mine}</span>` : '';
  const maxp = g.dist.length ? g.dist[0].p : 100;
  const rows = g.dist.map(d=>{
    const hit = g.played && d.s===g.actual;
    const mine = d.s===g.mine;
    const cls = ['drow', hit?'hit':'', mine?'mine':''].join(' ');
    const w = Math.max(4, (d.p/maxp)*100);
    const tags = (hit?'<span class="tg h">&#10003;</span>':'')+(mine?'<span class="tg m">T&Uacute;</span>':'');
    return `<div class="${cls}"><div class="barfill" style="width:${w}%"></div>`+
      `<span class="score">${d.s}</span><span class="tags">${tags}</span>`+
      `<span class="pct">${d.p}%</span><span class="cnt">${d.n} pers.</span></div>`;
  }).join('');
  $('#view').innerHTML =
    `<div class="gcard"><div class="ghead">
      <div class="match">
        ${g.homeFlag?`<img src="${g.homeFlag}" alt="">`:''}<span>${g.home}</span>
        <span class="vs">vs</span><span>${g.away}</span>${g.awayFlag?`<img src="${g.awayFlag}" alt="">`:''}
      </div>
      <div class="meta">${statusBadge}${youBadge}<span>${date}</span><span>${g.total} predicciones</span></div>
     </div>
     <div class="dist">${rows||'<div class="empty">Sin datos</div>'}</div></div>`;
  sel.value = idx;
}

sel.addEventListener('change', e=>{ idx=+e.target.value; render(); });
$('#prev').addEventListener('click', ()=>{ const o=[...sel.options].map(x=>+x.value); const p=o.indexOf(idx); if(p>0){idx=o[p-1];render();} });
$('#next').addEventListener('click', ()=>{ const o=[...sel.options].map(x=>+x.value); const p=o.indexOf(idx); if(p<o.length-1){idx=o[p+1];render();} });
$('#onlyPlayed').addEventListener('change', ()=>{ options(); render(); });

me(); options(); render();
</script>
</body>
</html>
'@

$html = $html.Replace('__JSON__', $json)
$out = Join-Path $dir 'polla_dashboard.html'
Set-Content -Path $out -Value $html -Encoding UTF8
Write-Host "WROTE $out  ($([math]::Round((Get-Item $out).Length/1KB,1)) KB, $nG games)"
