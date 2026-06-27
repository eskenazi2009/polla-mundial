$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

$teams = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code }
function TeamCode($id) { if ($id -ne $null -and $teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }

$cols = $model.groupCols
$rows = $model.rows
$nG = $cols.Count
$nR = $rows.Count
$me = $rows | Where-Object { $_.mine -eq $true } | Select-Object -First 1

# my pick per game
$mineG = New-Object 'string[]' $nG
for ($j = 0; $j -lt $nG; $j++) {
    if ($me) { $g = $me.g[$j]; $mineG[$j] = if ($g -and $g[0] -ne $null) { "$($g[0])-$($g[1])" } else { '' } } else { $mineG[$j] = '' }
}

# per-game sorted distributions
$games = @()
$maxDistinct = 0
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $counts = @{}; $total = 0
    foreach ($r in $rows) {
        $g = $r.g[$j]
        if ($g -ne $null -and $g.Count -ge 2 -and $g[0] -ne $null) {
            $s = "$($g[0])-$($g[1])"
            if ($counts.ContainsKey($s)) { $counts[$s]++ } else { $counts[$s] = 1 }
            $total++
        }
    }
    $sorted = @($counts.GetEnumerator() | Sort-Object -Property Value -Descending)
    if ($sorted.Count -gt $maxDistinct) { $maxDistinct = $sorted.Count }
    $games += [pscustomobject]@{
        Label = "G$($j+1)"; Match = (TeamCode $c.homeId) + "-" + (TeamCode $c.awayId)
        Actual = if ($c.played) { $c.actual } else { '' }; Played = [bool]$c.played
        Mine = $mineG[$j]; Scores = $sorted
    }
}

# ---- points (real) matrix: hit=2, miss=0, pending=blank ----
$ptsData = New-Object 'object[,]' $nR, $nG
for ($i = 0; $i -lt $nR; $i++) {
    $r = $rows[$i]
    for ($j = 0; $j -lt $nG; $j++) {
        $st = $r.g[$j][2]
        $ptsData[$i, $j] = if ($st -eq 1) { 2 } elseif ($st -eq 2) { 0 } else { '' }
    }
}

# ================= open workbook =================
$f = Get-ChildItem (Join-Path $dir 'polla_worldcup_predictions_*.xlsx') | Sort-Object LastWriteTime | Select-Object -Last 1
$xl = $null
for ($t = 1; $t -le 6 -and -not $xl; $t++) { try { $xl = New-Object -ComObject Excel.Application } catch { Start-Sleep -Milliseconds 1500 } }
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($f.FullName)

# ---- (A) fix Points per Game sheet ----
$wp = $wb.Worksheets.Item('Points per Game')
$fixed = 3
$wp.Range($wp.Cells.Item(2,$fixed+1), $wp.Cells.Item($nR+1,$fixed+$nG)).Value2 = $ptsData
$wp.Cells.Item(1,1).Value2 = 'Rank'  # keep
# add note in a far header cell
$wp.Range($wp.Cells.Item(1,1), $wp.Cells.Item(1,$fixed+$nG)).Font.Bold = $true

# ---- (B) rebuild By Game card sheet with "You" row ----
foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq 'By Game') { $s.Delete() } }
$blockW = 3; $gap = 1; $perGame = $blockW + $gap
$headerRows = 5
$totalRows = $headerRows + $maxDistinct
$totalCols = $nG * $perGame - $gap
$arr = New-Object 'object[,]' $totalRows, $totalCols
for ($j = 0; $j -lt $nG; $j++) {
    $g = $games[$j]; $b = $j * $perGame; $b1 = $b + 1; $b2 = $b + 2
    $arr[0,$b] = $g.Label
    $arr[1,$b] = $g.Match
    $arr[2,$b] = if ($g.Played) { "Actual: " + $g.Actual } else { "Actual: -" }
    $arr[3,$b] = if ($g.Mine) { "You: " + $g.Mine } else { "You: -" }
    $arr[4,$b] = 'Score'; $arr[4,$b1] = '#'; $arr[4,$b2] = '%'
    for ($k = 0; $k -lt $g.Scores.Count; $k++) {
        $row = $headerRows + $k; $kv = $g.Scores[$k]
        $tot = ($g.Scores | Measure-Object -Property Value -Sum).Sum
        $arr[$row,$b]  = $kv.Key
        $arr[$row,$b1] = $kv.Value
        $arr[$row,$b2] = if ($tot -gt 0) { [double]$kv.Value / $tot } else { 0 }
    }
}
$ws = $wb.Worksheets.Add($wb.Worksheets.Item(1), [Type]::Missing, 1)
$ws.Name = 'By Game'
for ($j = 0; $j -lt $nG; $j++) {
    $bc = $j * $perGame + 1; $bc2 = $bc + 2
    $ws.Range($ws.Cells.Item(6,$bc),  $ws.Cells.Item($totalRows,$bc)).NumberFormat = '@'
    $ws.Range($ws.Cells.Item(6,$bc2), $ws.Cells.Item($totalRows,$bc2)).NumberFormat = '0.0%'
}
$ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($totalRows,$totalCols)).Value2 = $arr

$GREEN = 11854022; $BLUE = 15652797; $YOUFILL = 15123099; $DARK = 6710886; $MED = 13417386; $LT = 15921906
for ($j = 0; $j -lt $nG; $j++) {
    $g = $games[$j]; $bc = $j * $perGame + 1; $bc2 = $bc + 2
    $t1 = $ws.Range($ws.Cells.Item(1,$bc), $ws.Cells.Item(1,$bc2)); $t1.Merge(); $t1.HorizontalAlignment = -4108; $t1.Font.Bold = $true; $t1.Font.Color = 16777215; $t1.Interior.Color = $DARK
    $t2 = $ws.Range($ws.Cells.Item(2,$bc), $ws.Cells.Item(2,$bc2)); $t2.Merge(); $t2.HorizontalAlignment = -4108; $t2.Font.Bold = $true; $t2.Interior.Color = $MED
    $t3 = $ws.Range($ws.Cells.Item(3,$bc), $ws.Cells.Item(3,$bc2)); $t3.Merge(); $t3.HorizontalAlignment = -4108; $t3.Font.Italic = $true; $t3.Interior.Color = $LT
    $t4 = $ws.Range($ws.Cells.Item(4,$bc), $ws.Cells.Item(4,$bc2)); $t4.Merge(); $t4.HorizontalAlignment = -4108; $t4.Font.Bold = $true; $t4.Interior.Color = $YOUFILL
    $h  = $ws.Range($ws.Cells.Item(5,$bc), $ws.Cells.Item(5,$bc2)); $h.Font.Bold = $true; $h.Interior.Color = $LT
    for ($k = 0; $k -lt $g.Scores.Count; $k++) {
        $sc = $g.Scores[$k].Key; $er = $headerRows + $k + 1
        $isHit = ($g.Played -and $sc -eq $g.Actual); $isMine = ($sc -eq $g.Mine -and $g.Mine -ne '')
        if ($isHit -or $isMine) {
            $rr = $ws.Range($ws.Cells.Item($er,$bc), $ws.Cells.Item($er,$bc2))
            if ($isHit) { $rr.Interior.Color = $GREEN } else { $rr.Interior.Color = $BLUE }
            $rr.Font.Bold = $true
            if ($isMine -and $isHit) { $rr.Font.Color = 12611584 }  # blue text on green = your hit
        }
    }
    $ws.Columns.Item($bc).ColumnWidth = 6; $ws.Columns.Item($bc+1).ColumnWidth = 5; $ws.Columns.Item($bc+2).ColumnWidth = 7
    if ($j -lt $nG - 1) { $ws.Columns.Item($bc+3).ColumnWidth = 2 }
}
$ws.Activate(); $xl.ActiveWindow.SplitRow = 5; $xl.ActiveWindow.FreezePanes = $true

$wb.Save(); $wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wp)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
Write-Host "Updated: By Game (with your picks) + fixed Points sheet. You=$($me.name) rank=$($me.rank)"
