# 학교 부식 발주서(발주서_○○(부식|조식)_YYMM.xlsx; 한살림 등 상품 단위, 헤더 NO/식품명/규격/매입처/판매단가/일자/총량/총액)
#   -> 두탑용 통계표 + 현품서 (친환경 양식 B와 같은 표준 템플릿 서식 사용)
#   통계표: 품목행을 식품명 유니코드순으로 정렬, NO 재부여. 평균단가 = 판매단가, 총금액 = 총액. 식품속성 설명은 원본에 없어 빈칸.
#   현품서: NO/식품명/규격/식품속성 + 총량 + 평균단가(판매단가) + 매입처(원본 매입처 열)
# 사용: pwsh -File convert_busik.ps1 -InputPath <발주서.xlsx> -OutDir <출력폴더> [-Force]
#   출력: <OutDir>\통계표\통계표_<접미>.xlsx , <OutDir>\현품서\현품서_<접미>.xlsx  (기존 파일은 -Force 없이는 덮어쓰지 않음)
param(
  [Parameter(Mandatory)][string]$InputPath,
  [Parameter(Mandatory)][string]$OutDir,
  [switch]$Force,
  [switch]$StatsOnly   # 통계표만 만든다(현품서 생략)
)
$ErrorActionPreference = 'Stop'
$name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
if ($name -notmatch '^발주서_(.+)$') { throw "파일명이 '발주서_...' 형식이 아닙니다: $name" }
$suffix = $Matches[1]
if ($suffix -notmatch '(\d{2})(\d{2})$') { throw "파일명 끝에서 YYMM을 찾지 못했습니다: $suffix" }
$year = 2000 + [int]$Matches[1]; $mon = [int]$Matches[2]
$days = [DateTime]::DaysInMonth($year, $mon)
$tongPath = Join-Path $OutDir "통계표\통계표_$suffix.xlsx"
$hpPath   = Join-Path $OutDir "현품서\현품서_$suffix.xlsx"
$checkPaths = if ($StatsOnly) { @($tongPath) } else { @($tongPath, $hpPath) }
foreach ($p in $checkPaths) {
  if ((Test-Path -LiteralPath $p) -and -not $Force) { throw "이미 존재: $p (덮어쓰려면 -Force)" }
  New-Item -ItemType Directory -Force -Path (Split-Path $p) | Out-Null
}
$tplT = Join-Path $PSScriptRoot 'template_통계표.xlsx'
$tplH = Join-Path $PSScriptRoot 'template_현품서.xlsx'
foreach ($p in $tplT, $tplH) { if (-not (Test-Path -LiteralPath $p)) { throw "템플릿 파일이 없습니다: $p" } }

# 일(1~31) -> 표준 통계표 열 번호 (1~6: E~J, 7~18: L~W, 19~31: Y~AK; K·X는 여백 열)
function Get-DayCol([int]$d) { if ($d -le 6) { 4 + $d } elseif ($d -le 18) { 5 + $d } else { 6 + $d } }
# 셀에 쓰고 다시 읽어 확인한다(Excel COM에서 값 대입이 조용히 무시되는 경우가 있었음).
function Set-Cell($ws, [int]$r, [int]$c, $v) {
  $cell = $ws.Cells.Item($r, $c)
  if ($null -eq $v -or "$v" -eq '') { $cell.ClearContents() | Out-Null; return }
  if ($v -is [double]) { $cell.Value2 = $v; if ([math]::Abs([double]$cell.Value2 - $v) -gt 1e-9) { throw "셀 쓰기 확인 실패 r$r c$c : $v" } }
  else { $cell.NumberFormat = '@'; $cell.Formula = [string]$v; if ($cell.Text.Trim() -ne ([string]$v).Trim()) { throw "셀 쓰기 확인 실패 r$r c$c : $v / $($cell.Text)" } }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$report = [System.Collections.Generic.List[string]]::new()
try {
  $src = $xl.Workbooks.Open($InputPath, 0, $true)
  # '판매단가' 헤더가 있는 보이는 시트를 찾는다 (예: '11월', '통계표'; '원본' 같은 메뉴표 시트는 건너뜀)
  $ws = $null; $hrow = $null
  foreach ($cand in $src.Worksheets) {
    if ($cand.Visible -ne -1) { continue }
    for ($r = 1; $r -le 12; $r++) {
      if ($cand.Cells.Item($r, 1).Text.Trim() -eq 'NO') {
        for ($c = 1; $c -le $cand.UsedRange.Columns.Count; $c++) { if ($cand.Cells.Item($r, $c).Text.Trim() -eq '판매단가') { $ws = $cand; $hrow = $r; break } }
        break
      }
    }
    if ($ws) { break }
  }
  if (-not $ws) { throw "'판매단가' 헤더가 있는 시트를 찾지 못했습니다(부식 양식이 아님)." }
  # 열 위치: 헤더 글자로 찾는다 (양식마다 식품속성/제조사 열 유무가 다름)
  $colName = 2; $colSpec = $null; $colAttr = $null; $colVendor = $null; $colPrice = $null; $colQty = $null; $colAmt = $null; $colDay1 = $null
  for ($c = 1; $c -le $ws.UsedRange.Columns.Count; $c++) {
    $t = ($ws.Cells.Item($hrow, $c).Text -replace '\s+', '')
    switch -Regex ($t) { '^식품명' { $colName = $c } '^규격$' { $colSpec = $c } '^식품속성' { $colAttr = $c } '^매입처$' { $colVendor = $c } '^판매단가$' { $colPrice = $c } '^총량$' { $colQty = $c } '^총액$|^총금액$|^금액$' { $colAmt = $c } '^1$' { if (-not $colDay1) { $colDay1 = $c } } }
  }
  if (-not $colSpec) { throw "'규격' 열을 찾지 못했습니다." }
  foreach ($pair in @(@('판매단가', $colPrice), @('총량', $colQty), @('총액', $colAmt), @('1일 열', $colDay1))) { if (-not $pair[1]) { throw "부식 양식이 아닙니다: '$($pair[0])' 열을 찾지 못했습니다." } }

  $dayMax = $colQty - $colDay1   # 일자 열 개수(총량 열 앞까지)
  # 품목행 수집
  $items = [System.Collections.Generic.List[object]]::new()
  for ($r = $hrow + 1; $r -le $ws.UsedRange.Rows.Count; $r++) {
    if ($ws.Cells.Item($r, 1).Text.Trim() -match '^\d+$' -and $ws.Cells.Item($r, $colName).Text.Trim()) {
      $items.Add([pscustomobject]@{ row = $r; name = $ws.Cells.Item($r, $colName).Text.Trim() })
    }
  }
  if ($items.Count -eq 0) { throw '품목행을 찾지 못했습니다.' }
  # 식품명 유니코드(서수)순 정렬
  $keys = [string[]]@($items | ForEach-Object { $_.name })
  $idx = [int[]](0..($items.Count - 1)); [Array]::Sort($keys, $idx, [StringComparer]::Ordinal)
  $sorted = @($idx | ForEach-Object { $items[$_] })

  # ---- 통계표 ----
  $tp = $xl.Workbooks.Open($tplT, 0, $true); $tpw = $tp.Worksheets.Item(1)
  $t = $xl.Workbooks.Add(); $tw = $t.Worksheets.Item(1); $tw.Name = 'sheet1'
  $lastCol = 41; $nameRow = 2
  [void]$tpw.Range($tpw.Cells.Item(1, 1), $tpw.Cells.Item(1, $lastCol)).Copy($tw.Cells.Item(1, 1))
  $tw.Rows.Item(1).RowHeight = $tpw.Rows.Item(1).RowHeight
  if ($days -ge 31) { [void]$tpw.Cells.Item(1, 36).Copy($tw.Cells.Item(1, 37)); Set-Cell $tw 1 37 ([double]31) } else { Set-Cell $tw 1 37 $null }
  $i = 1
  foreach ($it in $sorted) {
    $i++; $r = $it.row
    [void]$tpw.Range($tpw.Cells.Item($nameRow, 1), $tpw.Cells.Item($nameRow, $lastCol)).Copy($tw.Cells.Item($i, 1))
    $tw.Rows.Item($i).RowHeight = $tpw.Rows.Item($nameRow).RowHeight
    $tw.Range($tw.Cells.Item($i, 5), $tw.Cells.Item($i, 37)).ClearContents() | Out-Null
    Set-Cell $tw $i 1 ([double]($i - 1))
    Set-Cell $tw $i 2 ([string]$ws.Cells.Item($r, $colName).Text.Trim())
    Set-Cell $tw $i 3 ([string]$ws.Cells.Item($r, $colSpec).Text.Trim())
    if ($colAttr) { Set-Cell $tw $i 4 ([string]$ws.Cells.Item($r, $colAttr).Text.Trim()) } else { Set-Cell $tw $i 4 $null }
    for ($d = 1; $d -le $dayMax; $d++) {
      $v = $ws.Cells.Item($r, $colDay1 - 1 + $d).Value2
      if ($null -ne $v -and "$v" -ne '' -and $v -is [double]) {
        if ($d -gt $days) { $report.Add("경고: $($it.name) 의 ${d}일 수량($v)이 ${mon}월 일수를 넘어 누락되었습니다.") }
        else { Set-Cell $tw $i (Get-DayCol $d) ([double]$v) }
      }
    }
    Set-Cell $tw $i 38 ([double]$ws.Cells.Item($r, $colQty).Value2)
    Set-Cell $tw $i 39 ([double]$ws.Cells.Item($r, $colAmt).Value2)
    Set-Cell $tw $i 40 ([double]$ws.Cells.Item($r, $colPrice).Value2)
    # 검증: 일자 합 = 총량, 총량 x 단가 = 총액
    $sum = 0.0; for ($d = 1; $d -le $dayMax; $d++) { $v = $ws.Cells.Item($r, $colDay1 - 1 + $d).Value2; if ($v -is [double]) { $sum += $v } }
    $q = [double]$ws.Cells.Item($r, $colQty).Value2
    if ([math]::Abs($sum - $q) -gt 1e-6) { $report.Add("확인: $($it.name) 일자 합($sum)과 총량($q)이 다릅니다.") }
    $amt = [double]$ws.Cells.Item($r, $colAmt).Value2; $pr = [double]$ws.Cells.Item($r, $colPrice).Value2
    if ([math]::Abs($q * $pr - $amt) -gt 0.5) { $report.Add("확인: $($it.name) 총량x단가($($q * $pr))와 총액($amt)이 다릅니다.") }
  }
  for ($c = 1; $c -le $lastCol; $c++) { $tw.Columns.Item($c).ColumnWidth = $tpw.Columns.Item($c).ColumnWidth }
  $tp.Close($false)
  $t.SaveAs($tongPath, 51)
  if ($StatsOnly) {
    $t.Close($false); $src.Close($false)
    "OK [부식/통계표만] $suffix : 품목 $($items.Count)개 -> $tongPath"
    $report | ForEach-Object { "  ! $_" }
    return
  }

  # ---- 현품서 ----
  $hp = $xl.Workbooks.Open($tplH, 0, $true); $hpw = $hp.Worksheets.Item(1)
  $h = $xl.Workbooks.Add(); $hw = $h.Worksheets.Item(1); $hw.Name = 'sheet1'
  [void]$hpw.Range($hpw.Cells.Item(1, 1), $hpw.Cells.Item(1, 7)).Copy($hw.Cells.Item(1, 1))
  $hw.Rows.Item(1).RowHeight = $hpw.Rows.Item(1).RowHeight
  $k = 1
  foreach ($it in $sorted) {
    $k++; $r = $it.row
    [void]$hpw.Range($hpw.Cells.Item($nameRow, 1), $hpw.Cells.Item($nameRow, 7)).Copy($hw.Cells.Item($k, 1))
    $hw.Rows.Item($k).RowHeight = $hpw.Rows.Item($nameRow).RowHeight
    Set-Cell $hw $k 1 ([double]($k - 1))
    Set-Cell $hw $k 2 ([string]$tw.Cells.Item($k, 2).Value2)
    Set-Cell $hw $k 3 ([string]$tw.Cells.Item($k, 3).Value2)
    Set-Cell $hw $k 4 ([string]$tw.Cells.Item($k, 4).Value2)
    Set-Cell $hw $k 5 ([double]$tw.Cells.Item($k, 38).Value2)
    Set-Cell $hw $k 6 ([double]$tw.Cells.Item($k, 40).Value2)
    $vd = if ($colVendor) { $ws.Cells.Item($r, $colVendor).Text.Trim() } else { '' }
    if ($vd) { Set-Cell $hw $k 7 ([string]$vd) } else { Set-Cell $hw $k 7 $null; $report.Add("매입처 빈칸: $($it.name)") }
  }
  for ($c = 1; $c -le 7; $c++) { $hw.Columns.Item($c).ColumnWidth = $hpw.Columns.Item($c).ColumnWidth }
  $hp.Close($false)
  $t.Close($false)
  $h.SaveAs($hpPath, 51); $h.Close($false)
  $src.Close($false)
  "OK [부식] $suffix : 품목 $($items.Count)개 -> $tongPath / $hpPath"
  $report | ForEach-Object { "  ! $_" }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

