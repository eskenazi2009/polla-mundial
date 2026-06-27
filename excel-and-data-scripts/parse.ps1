$ErrorActionPreference = 'Stop'
$dir = 'C:\Users\esken\Downloads\rebundle-RB-KDSZ-M9JC\rebundle\polla-worldcup'
$html = Get-Content (Join-Path $dir 'detailed_full.html') -Raw

# 1) Reconstruct the RSC stream: concatenate every self.__next_f.push([1,"<literal>"]) body, unescaped.
$rx = [regex]'self\.__next_f\.push\(\[1,"((?:[^"\\]|\\.)*)"\]\)'
$sb = New-Object System.Text.StringBuilder
foreach ($m in $rx.Matches($html)) {
    $lit = $m.Groups[1].Value
    # turn the JS/JSON string literal body back into the real string
    $real = ('"' + $lit + '"') | ConvertFrom-Json
    [void]$sb.Append($real)
}
$stream = $sb.ToString()
Write-Host "stream length: $($stream.Length)"

# 2) Locate "model":{ and brace-match (tracking string state) to extract the object.
$idx = $stream.IndexOf('"model":')
if ($idx -lt 0) { throw 'model not found' }
$start = $stream.IndexOf('{', $idx)
$depth = 0; $inStr = $false; $esc = $false; $end = -1
for ($i = $start; $i -lt $stream.Length; $i++) {
    $ch = $stream[$i]
    if ($inStr) {
        if ($esc) { $esc = $false }
        elseif ($ch -eq '\') { $esc = $true }
        elseif ($ch -eq '"') { $inStr = $false }
    } else {
        if ($ch -eq '"') { $inStr = $true }
        elseif ($ch -eq '{') { $depth++ }
        elseif ($ch -eq '}') { $depth--; if ($depth -eq 0) { $end = $i; break } }
    }
}
if ($end -lt 0) { throw 'unbalanced model object' }
$modelJson = $stream.Substring($start, $end - $start + 1)
Write-Host "model json length: $($modelJson.Length)"
Set-Content -Path (Join-Path $dir 'model.json') -Value $modelJson -Encoding UTF8

# 3) Parse and report shape
$model = $modelJson | ConvertFrom-Json
Write-Host "totalPools: $($model.totalPools)"
Write-Host "groupCols (games): $($model.groupCols.Count)"
Write-Host "rows (participants): $($model.rows.Count)"
Write-Host "teams: $(@($model.teams.PSObject.Properties).Count)"
$r0 = $model.rows[0]
Write-Host "row0: rank=$($r0.rank) name=$($r0.name) score=$($r0.score) g.len=$($r0.g.Count) champion=$($r0.champion) final=$($r0.final)"
$c0 = $model.groupCols[0]
Write-Host "col0: matchId=$($c0.matchId) home=$($c0.homeId) away=$($c0.awayId) actual=$($c0.actual) acertados=$($c0.acertados)"
Write-Host "row props: $((@($r0.PSObject.Properties.Name)) -join ', ')"
Write-Host "col props: $((@($c0.PSObject.Properties.Name)) -join ', ')"
