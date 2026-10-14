# 과거 현품서(2601~)에서 "품목명 -> 매입처" 대응표(vendor_map.csv)를 만든다.
# 사용: pwsh -File build_vendor_map.ps1 [-Root '\\nas2dual\2026\4.공급유통\01 학교급식\01 친환경급식']
param(
  [string]$Root = '\\nas2dual\2026\4.공급유통\01 학교급식\01 친환경급식',
  [string]$Out  = (Join-Path $PSScriptRoot 'vendor_map.csv')
)
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-Sheet1($path) {
  $z = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $ss = @()
    $e = $z.GetEntry('xl/sharedStrings.xml')
    if ($e) { $x = [xml]([IO.StreamReader]::new($e.Open()).ReadToEnd()); foreach ($si in $x.sst.si) { $ss += $si.InnerText } }
    $e = $z.GetEntry('xl/worksheets/sheet1.xml')
    $x = [xml]([IO.StreamReader]::new($e.Open()).ReadToEnd())
    foreach ($r in $x.worksheet.sheetData.row) {
      $h = @{}
      foreach ($c in $r.c) {
        $col = $c.r -replace '\d', ''
        $v = $c.v
        if ($c.t -eq 's' -and $null -ne $v) { $v = $ss[[int]$v] }
        elseif ($c.t -eq 'inlineStr') { $v = $c.is.InnerText }
        if ($null -ne $v -and "$v" -ne '') { $h[$col] = "$v" }
      }
      [pscustomobject]@{ Row = [int]$r.r; Cells = $h }
    }
  } finally { $z.Dispose() }
}

$tally = @{}
Get-ChildItem -LiteralPath $Root -Directory | Where-Object Name -match '^\d{4}$' | ForEach-Object {
  $d = Join-Path $_.FullName '두탑용\현품서'
  if (-not (Test-Path -LiteralPath $d)) { return }
  Get-ChildItem -LiteralPath $d -Filter '현품서_*.xlsx' -File | ForEach-Object {
    try { $rows = @(Read-Sheet1 $_.FullName) } catch { Write-Warning "skip $($_.Name): $_"; return }
    foreach ($row in $rows) {
      if ($row.Row -le 1) { continue }
      $n = $row.Cells['B']; $v = $row.Cells['G']
      if (-not $n -or -not $v) { continue }
      if ($v -match '^[\d,.\s]+$' -or $v.StartsWith('※') -or $v.Length -gt 20) { continue }  # 숫자/주석은 매입처가 아님
      if (-not $tally.ContainsKey($n)) { $tally[$n] = @{} }
      $tally[$n][$v] = 1 + [int]$tally[$n][$v]
    }
  }
}

$rows = foreach ($k in ($tally.Keys | Sort-Object)) {
  $c = $tally[$k].GetEnumerator() | Sort-Object Value -Descending
  [pscustomobject]@{
    item         = $k
    vendor       = $c[0].Name
    count        = $c[0].Value
    alternatives = (($c | Select-Object -Skip 1 | ForEach-Object { "$($_.Name)x$($_.Value)" }) -join ';')
  }
}
$rows | Export-Csv -LiteralPath $Out -NoTypeInformation -Encoding utf8
"items: $(@($rows).Count)  with alternatives: $(@($rows | Where-Object alternatives).Count)  -> $Out"
