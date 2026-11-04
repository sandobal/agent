param([string]$Sp)
# 장애인복지관 발주서_01의 빈 매입처/단가를 규칙으로 채운다 (이미 값이 있는 칸은 건드리지 않음).
# 규칙: 단가파일 후보 중 (1) 10월 발주서와 같은 값이 있으면 그것, (2) 아니면 가장 많이 나온 값, (3) 동률이면 파일에서 가장 아래(최근) 행.
#       단가파일에 없으면 이전 발주서(05->02월 순) 값. 둘 다 없으면 비워 둔다. 단가 0원은 넣지 않는다.
$ErrorActionPreference = 'Stop'
$d = "\\nas2dual\2026\4.공급유통\02 공공급식\07. 괴산군장애인복지관\발주서"
$cur = "$d\2611\발주서_장애인복지관_02.xlsx"
function Norm([string]$s) { ($s -replace '\s', '' -replace '[,()/·.]', '').ToUpper() }
function ToNum($t) { $s = ("$t" -replace '[^\d.]', ''); if ($s) { [double]$s } else { [double]0 } }
$lock = Join-Path (Split-Path $cur) ('~$' + (Split-Path $cur -Leaf))
if (Test-Path -LiteralPath $lock) { "건너뜀: 발주서가 Excel에서 열려 있습니다."; return }
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$price = @(); $prev = @()
try {
  $tmp = Join-Path ([IO.Path]::GetTempPath()) 'jb2_price.xlsx'; Copy-Item -LiteralPath "$d\2611\장애인복지관_단가.xlsx" $tmp -Force
  $wb = $xl.Workbooks.Open($tmp, 0, $true); $ws = $wb.Worksheets.Item(1)
  for ($r = 2; $r -le $ws.UsedRange.Rows.Count; $r++) { $n = $ws.Cells.Item($r, 2).Text.Trim(); if ($n) { $price += [pscustomobject]@{ row = $r; vendor = $ws.Cells.Item($r, 1).Text.Trim(); name = $n; price = (ToNum $ws.Cells.Item($r, 3).Text) } } }
  $wb.Close($false)
  foreach ($f in '11-01','05', '04', '03', '02') {
    $tmp = Join-Path ([IO.Path]::GetTempPath()) 'jb2_prev.xlsx'; $pp = if ($f -eq '11-01') { "$d\2611\발주서_장애인복지관_01.xlsx" } else { "$d\2610\발주서_장애인복지관_$f.xlsx" }; Copy-Item -LiteralPath $pp $tmp -Force
    $wb = $xl.Workbooks.Open($tmp, 0, $true); $ws = $wb.Worksheets.Item(1)
    for ($r = 2; $r -le $ws.UsedRange.Rows.Count; $r++) { $n = $ws.Cells.Item($r, 2).Text.Trim(); if ($n -and $ws.Cells.Item($r, 5).Text.Trim()) { $prev += [pscustomobject]@{ order = $(if ($f -eq '11-01') { 99 } else { [int]$f }); name = $n; vendor = $ws.Cells.Item($r, 5).Text.Trim(); price = (ToNum $ws.Cells.Item($r, 6).Text) } } }
    $wb.Close($false)
  }
  New-Item -ItemType Directory -Force "$Sp\backup" | Out-Null
  Copy-Item -LiteralPath $cur "$Sp\backup\발주서_장애인복지관_02_before.xlsx" -Force
  $wb = $xl.Workbooks.Open($cur); if ($wb.ReadOnly) { throw '읽기 전용' }
  $ws = $wb.Worksheets.Item(1); $ur = $ws.UsedRange
  $cV = $null; $cP = $null
  for ($c = 1; $c -le $ur.Columns.Count; $c++) { $t = $ws.Cells.Item(1, $c).Text.Trim(); if ($t -eq '매입처') { $cV = $c }; if ($t -eq '단가') { $cP = $c } }
  $report = @()
  for ($r = 2; $r -le $ur.Rows.Count; $r++) {
    $name = $ws.Cells.Item($r, 2).Text.Trim(); if (-not $name) { continue }
    $hasV = [bool]$ws.Cells.Item($r, $cV).Text.Trim(); $hasP = [bool]$ws.Cells.Item($r, $cP).Text.Trim()
    if ($hasV -and $hasP) { continue }
    $nn = Norm $name
    $cands = @($price | Where-Object { (Norm $_.name) -eq $nn })
    $pv = @($prev | Where-Object { (Norm $_.name) -eq $nn } | Sort-Object order -Descending)
    $pick = $null; $why = ''
    if ($cands.Count -gt 0) {
      if ($hasV) { $cands = @($cands | Where-Object { $_.vendor -eq $ws.Cells.Item($r, $cV).Text.Trim() }) }
      if ($cands.Count -gt 0) {
        $groups = @($cands | Group-Object { "$($_.vendor)|$($_.price)" })
        $hit = $null
        foreach ($p in $pv) { $k = "$($p.vendor)|$($p.price)"; $g = $groups | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($g) { $hit = $g; $why = '10월과 같은 값'; break } }
        if (-not $hit) {
          $max = ($groups | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
          $top = @($groups | Where-Object { $_.Count -eq $max })
          if ($top.Count -eq 1) { $hit = $top[0]; $why = '가장 많이 나온 값' } else { $hit = $top | Sort-Object { ($_.Group | ForEach-Object { $_.row } | Measure-Object -Maximum).Maximum } -Descending | Select-Object -First 1; $why = '동률 중 최근 행' }
        }
        $s = $hit.Name -split '\|'; $pick = [pscustomobject]@{ vendor = $s[0]; price = [double]$s[1] }
        if ($groups.Count -eq 1) { $why = '단가파일(단일)' }
      }
    }
    if (-not $pick -and $pv.Count -gt 0) { $pick = [pscustomobject]@{ vendor = $pv[0].vendor; price = $pv[0].price }; $why = "이전 발주서($($pv[0].order)차)" }
    if (-not $pick) { $report += "비움: [$r] $name (자료 없음)"; continue }
    $w = @()
    if (-not $hasV) { $ws.Cells.Item($r, $cV).Formula = [string]$pick.vendor; $w += "매입처 $($pick.vendor)" }
    if (-not $hasP -and $pick.vendor -eq '중앙상회') { $w += '단가 중앙상회라 비움' }
    elseif (-not $hasP -and $pick.price -gt 0) { $ws.Cells.Item($r, $cP).Value2 = [double]$pick.price; $w += "단가 $($pick.price)" }
    elseif (-not $hasP) { $w += '단가 0원이라 비움' }
    $report += "입력: [$r] $name -> " + ($w -join ', ') + " ($why)"
  }
  $wb.Save(); $wb.Close($false)
  $report
} finally { $xl.Quit() }


