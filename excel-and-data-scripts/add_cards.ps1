$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

$teams = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code }
function TeamCode($id) { if ($id -ne $null -and $teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }

$cols = $model.groupCols
$rows = $model.rows
$nG = $cols.Count

# Build per-game sorted score distributions
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
        Label = "G$($j+1)"
        Match = (TeamCode $c.homeId) + "-" + (TeamCode $c.awayId)
        Actual = if ($c.played) { $c.actual } else { '' }
        Played = [bool]$c.played
        Total = $total
        Scores = $sorted   # array of DictionaryEntry (Key=score, Value=count)
    }
}

# Layout: each game = 3 columns (Score, #, %) + 1 spacer.  Rows:
# 1 Game label | 2 Match | 3 Actual | 4 header(Score/#/%) | 5.. data
$blockW = 3; $gap = 1; $perGame = $blockW + $gap
$headerRows = 4
$totalRows = $headerRows + $maxDistinct
$totalCols = $nG * $perGame - $gap   # no trailing spacer

$arr = New-Object 'object[,]' $totalRows, $totalCols
for ($j = 0; $j -lt $nG; $j++) {
    $g = $games[$j]
    $b = $j * $perGame    # 0-based start col of this block
    $b1 = $b + 1; $b2 = $b + 2
    $arr[0, $b]  = $g.Label
    $arr[1, $b]  = $g.Match
    $arr[2, $b]  = if ($g.Played) { "Actual: " + $g.Actual } else { "Actual: -" }
    $arr[3, $b]  = 'Score'
    $arr[3, $b1] = '#'
    $arr[3, $b2] = '%'
    for ($k = 0; $k -lt $g.Scores.Count; $k++) {
        $row = $headerRows + $k
        $kv = $g.Scores[$k]
        $arr[$row, $b]  = $kv.Key
        $arr[$row, $b1] = $kv.Value
        $arr[$row, $b2] = if ($g.Total -gt 0) { [double]$kv.Value / $g.Total } else { 0 }
    }
}

# ---- Open workbook and write the sheet ----
$f = Get-ChildItem (Join-Path $dir 'polla_worldcup_predictions_*.xlsx') | Sort-Object LastWriteTime | Select-Object -Last 1
$xl = $null
for ($t = 1; $t -le 5 -and -not $xl; $t++) { try { $xl = New-Object -ComObject Excel.Application } catch { Start-Sleep -Milliseconds 1500 } }
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($f.FullName)
foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq 'By Game') { $s.Delete() } }
$ws = $wb.Worksheets.Add($wb.Worksheets.Item(1), [Type]::Missing, 1)  # place first
$ws.Name = 'By Game'

# Pre-set number formats per block (Score col = text, % col = percent)
for ($j = 0; $j -lt $nG; $j++) {
    $bc = $j * $perGame + 1   # 1-based
    $ws.Range($ws.Cells.Item(4,$bc),     $ws.Cells.Item($totalRows,$bc)).NumberFormat = '@'
    $ws.Range($ws.Cells.Item(4,$bc + 2), $ws.Cells.Item($totalRows,$bc + 2)).NumberFormat = '0.0%'
}

$ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($totalRows,$totalCols)).Value2 = $arr

# ---- Styling per block ----
for ($j = 0; $j -lt $nG; $j++) {
    $g = $games[$j]
    $bc = $j * $perGame + 1
    # title (row1) merged, dark fill, white bold
    $t1 = $ws.Range($ws.Cells.Item(1,$bc), $ws.Cells.Item(1,$bc + 2)); $t1.Merge()
    $t1.HorizontalAlignment = -4108; $t1.Font.Bold = $true; $t1.Font.Color = 16777215; $t1.Interior.Color = 6710886
    # match (row2) merged, medium fill bold
    $t2 = $ws.Range($ws.Cells.Item(2,$bc), $ws.Cells.Item(2,$bc + 2)); $t2.Merge()
    $t2.HorizontalAlignment = -4108; $t2.Font.Bold = $true; $t2.Interior.Color = 13417386
    # actual (row3) merged
    $t3 = $ws.Range($ws.Cells.Item(3,$bc), $ws.Cells.Item(3,$bc + 2)); $t3.Merge()
    $t3.HorizontalAlignment = -4108; $t3.Font.Italic = $true; $t3.Interior.Color = 15921906
    # header (row4) bold
    $h = $ws.Range($ws.Cells.Item(4,$bc), $ws.Cells.Item(4,$bc + 2)); $h.Font.Bold = $true; $h.Interior.Color = 15921906
    # highlight the row whose score == actual (the hit)
    if ($g.Played) {
        for ($k = 0; $k -lt $g.Scores.Count; $k++) {
            if ($g.Scores[$k].Key -eq $g.Actual) {
                $hr = $ws.Range($ws.Cells.Item($headerRows + $k + 1,$bc), $ws.Cells.Item($headerRows + $k + 1,$bc + 2))
                $hr.Interior.Color = 11854022; $hr.Font.Bold = $true   # green
            }
        }
    }
    # column widths
    $ws.Columns.Item($bc).ColumnWidth = 6
    $ws.Columns.Item($bc + 1).ColumnWidth = 5
    $ws.Columns.Item($bc + 2).ColumnWidth = 7
    if ($j -lt $nG - 1) { $ws.Columns.Item($bc + 3).ColumnWidth = 2 }
}

$ws.Activate()
$xl.ActiveWindow.SplitRow = 4
$xl.ActiveWindow.FreezePanes = $true

$wb.Save(); $wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
Write-Host "Added 'By Game': $nG game cards, up to $maxDistinct scores each ($totalRows x $totalCols) -> $($f.Name)"
