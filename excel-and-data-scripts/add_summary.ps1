$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$model = (Get-Content (Join-Path $dir 'model.json') -Raw) | ConvertFrom-Json

$teams = @{}
foreach ($p in $model.teams.PSObject.Properties) { $teams[$p.Name] = $p.Value.code }
function TeamCode($id) { if ($id -ne $null -and $teams.ContainsKey("$id")) { $teams["$id"] } else { "?" } }

$cols = $model.groupCols
$rows = $model.rows
$nG = $cols.Count

# Aggregate predicted scores per game
$records = New-Object System.Collections.ArrayList
for ($j = 0; $j -lt $nG; $j++) {
    $c = $cols[$j]
    $match = (TeamCode $c.homeId) + "-" + (TeamCode $c.awayId)
    $actual = if ($c.played) { $c.actual } else { '' }
    $counts = @{}
    $total = 0
    foreach ($r in $rows) {
        $g = $r.g[$j]
        if ($g -ne $null -and $g.Count -ge 2 -and $g[0] -ne $null) {
            $s = "$($g[0])-$($g[1])"
            if ($counts.ContainsKey($s)) { $counts[$s] = $counts[$s] + 1 } else { $counts[$s] = 1 }
            $total++
        }
    }
    $sorted = $counts.GetEnumerator() | Sort-Object -Property Value -Descending
    foreach ($kv in $sorted) {
        $pct = if ($total -gt 0) { [double]$kv.Value / $total } else { 0 }
        $hit = if ($c.played -and $kv.Key -eq $actual) { 'YES' } else { '' }
        [void]$records.Add([pscustomobject]@{
            Game = "G$($j+1)"; Match = $match; Actual = $actual
            Predicted = $kv.Key; Count = $kv.Value; Pct = $pct; Hit = $hit
        })
    }
}

# Build 2D array
$nRec = $records.Count
$arr = New-Object 'object[,]' ($nRec + 1), 7
$hdr = 'Game','Match','Actual Result','Predicted Score','# Predicting','% of Predictions','Hit (=actual)'
for ($k = 0; $k -lt 7; $k++) { $arr[0,$k] = $hdr[$k] }
for ($i = 0; $i -lt $nRec; $i++) {
    $rec = $records[$i]; $ri = $i + 1
    $arr[$ri,0] = $rec.Game
    $arr[$ri,1] = $rec.Match
    $arr[$ri,2] = $rec.Actual
    $arr[$ri,3] = $rec.Predicted
    $arr[$ri,4] = $rec.Count
    $arr[$ri,5] = $rec.Pct
    $arr[$ri,6] = $rec.Hit
}

# Open latest workbook and append sheet
$f = Get-ChildItem (Join-Path $dir 'polla_worldcup_predictions_*.xlsx') | Sort-Object LastWriteTime | Select-Object -Last 1
$xl = $null
for ($t = 1; $t -le 5 -and -not $xl; $t++) { try { $xl = New-Object -ComObject Excel.Application } catch { Start-Sleep -Milliseconds 1500 } }
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($f.FullName)

# remove existing summary sheet if rerun
foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq 'Score Summary') { $s.Delete() } }

$ws = $wb.Worksheets.Add($wb.Worksheets.Item($wb.Worksheets.Count), [Type]::Missing, 1)
$ws.Name = 'Score Summary'
$rCount = $nRec + 1
# text format for Actual + Predicted, percent format for Pct
$ws.Range($ws.Cells.Item(1,3), $ws.Cells.Item($rCount,4)).NumberFormat = '@'
$rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($rCount,7))
$rng.Value2 = $arr
$ws.Range($ws.Cells.Item(2,6), $ws.Cells.Item($rCount,6)).NumberFormat = '0.0%'
$hrng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,7))
$hrng.Font.Bold = $true
$hrng.Interior.Color = 15921906
$ws.Activate()
$xl.ActiveWindow.SplitRow = 1
$xl.ActiveWindow.FreezePanes = $true
[void]$ws.Columns.AutoFit()

$wb.Save()
$wb.Close($false)
$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
Write-Host "Added 'Score Summary': $nRec rows across $nG games -> $($f.Name)"
