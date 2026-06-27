$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

# team id -> code
$teams = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code }
function TeamCode($id) { if ($id -ne $null -and $teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }

$cols = $model.groupCols
$rows = $model.rows
$nG = $cols.Count        # 72
$nR = $rows.Count        # 424

# ---------- Sheet 1: Predictions grid (predicted score per game) ----------
$fixed = 3               # Rank, Nickname, Score
$predCols = $fixed + $nG
$pred = New-Object 'object[,]' ($nR + 1), $predCols
$pred[0,0] = 'Rank'; $pred[0,1] = 'Nickname'; $pred[0,2] = 'Total Score'
# precompute game-column headers
$ghdr = New-Object 'string[]' $nG
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $ghdr[$j] = "G$($j+1) " + (TeamCode $c.homeId) + "-" + (TeamCode $c.awayId)
    $cc = $fixed + $j
    $pred[0, $cc] = $ghdr[$j]
}
for ($i = 0; $i -lt $nR; $i++) {
    $r = $rows[$i]
    $ri = $i + 1
    $pred[$ri,0] = $r.rank
    $pred[$ri,1] = $r.name
    $pred[$ri,2] = $r.score
    for ($j = 0; $j -lt $nG; $j++) {
        $g = $r.g[$j]
        $cc = $fixed + $j
        if ($g -ne $null -and $g.Count -ge 2 -and $g[0] -ne $null) {
            $pred[$ri, $cc] = "$($g[0])-$($g[1])"
        } else { $pred[$ri, $cc] = '' }
    }
}

# ---------- Sheet 2: Points earned per game ----------
$pts = New-Object 'object[,]' ($nR + 1), $predCols
$pts[0,0] = 'Rank'; $pts[0,1] = 'Nickname'; $pts[0,2] = 'Total Score'
for ($j = 0; $j -lt $nG; $j++) { $cc = $fixed + $j; $pts[0, $cc] = $ghdr[$j] }
for ($i = 0; $i -lt $nR; $i++) {
    $r = $rows[$i]
    $ri = $i + 1
    $pts[$ri,0] = $r.rank; $pts[$ri,1] = $r.name; $pts[$ri,2] = $r.score
    for ($j = 0; $j -lt $nG; $j++) {
        $g = $r.g[$j]
        $cc = $fixed + $j
        if ($g -ne $null -and $g.Count -ge 3) { $pts[$ri, $cc] = $g[2] } else { $pts[$ri, $cc] = '' }
    }
}

# ---------- Sheet 3: Games reference ----------
$gm = New-Object 'object[,]' ($nG + 1), 7
$gh = 'Col','MatchId','Home','Away','Actual Result','Kickoff (UTC)','# Correct (acertados)'
for ($k = 0; $k -lt 7; $k++) { $gm[0,$k] = $gh[$k] }
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $rj = $j + 1
    $gm[$rj,0] = "G$($j+1)"
    $gm[$rj,1] = $c.matchId
    $gm[$rj,2] = TeamCode $c.homeId
    $gm[$rj,3] = TeamCode $c.awayId
    $gm[$rj,4] = if ($c.played) { $c.actual } else { '(not played)' }
    $gm[$rj,5] = $c.kickoff
    $gm[$rj,6] = $c.acertados
}

# ---------- Sheet 4: Leaderboard ----------
$lb = New-Object 'object[,]' ($nR + 1), 4
$lb[0,0]='Rank'; $lb[0,1]='Nickname'; $lb[0,2]='Total Score'; $lb[0,3]='Champion Pick'
for ($i = 0; $i -lt $nR; $i++) {
    $r = $rows[$i]
    $ri = $i + 1
    $champId = $null
    if ($r.champion -ne $null -and $r.champion.Count -ge 1 -and $r.champion[0] -ne $null) { $champId = $r.champion[0][0] }
    $lb[$ri,0]=$r.rank; $lb[$ri,1]=$r.name; $lb[$ri,2]=$r.score; $lb[$ri,3]=TeamCode $champId
}

# ---------- Write workbook via Excel COM ----------
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add() }

function Fill($ws, $name, $arr, $textColFrom, $textColTo) {
    $ws.Name = $name
    $rCount = $arr.GetLength(0); $cCount = $arr.GetLength(1)
    # Force text format on score columns so "2-0" isn't coerced to a date
    if ($textColFrom -gt 0) {
        $tr = $ws.Range($ws.Cells.Item(1,$textColFrom), $ws.Cells.Item($rCount,$textColTo))
        $tr.NumberFormat = '@'
    }
    $rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($rCount,$cCount))
    $rng.Value2 = $arr
    $hdr = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$cCount))
    $hdr.Font.Bold = $true
    $hdr.Interior.Color = 15921906   # light blue-grey
    $ws.Application.ActiveWindow.SplitRow = 1
    $ws.Application.ActiveWindow.FreezePanes = $true
    [void]$ws.Columns.AutoFit()
}

$ws1 = $wb.Worksheets.Item(1); $ws1.Activate(); Fill $ws1 'Predictions' $pred ($fixed+1) $predCols
$ws2 = $wb.Worksheets.Item(2); $ws2.Activate(); Fill $ws2 'Points per Game' $pts 0 0
$ws3 = $wb.Worksheets.Item(3); $ws3.Activate(); Fill $ws3 'Games' $gm 5 5
$ws4 = $wb.Worksheets.Item(4); $ws4.Activate(); Fill $ws4 'Leaderboard' $lb 0 0
$ws1.Activate()

$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$out = Join-Path $dir "polla_worldcup_predictions_$stamp.xlsx"
$wb.SaveAs($out, 51)
$wb.Close($false)
$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws1)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Write-Host "SAVED: $out"
Write-Host "Participants: $nR  Games: $nG"
