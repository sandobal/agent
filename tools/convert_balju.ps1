# 학교 친환경 발주서(발주서_○○(친환경)_YYMM.xlsx) -> 두탑용 통계표 + 현품서 변환
#   통계표: 발주서의 품목행을 합쳐 식품명 가나다순(유니코드순)으로 정렬, NO 재부여, 헤더 1줄
#   현품서: NO/식품명/규격/식품속성 + 총량/평균단가/매입처(vendor_map.csv 조회)
#
# 지원하는 발주서 양식
#   A) 급식품 구매예정량 통계표 (헤더에 '평균단가') : 시트 전체(농산물/축산물 ...)를 합침
#   B) 물품 구매 요구서        (헤더에 '결정단가') : '(현물제외)' 시트만 사용(없으면 보이는 시트 전체).
#      결정단가를 평균단가로, 일자 열을 표준 통계표 배치로 옮김
#
# 사용: pwsh -File convert_balju.ps1 -InputPath <발주서.xlsx> -OutDir <출력폴더> [-Force]
#   출력: <OutDir>\통계표\통계표_*.xlsx , <OutDir>\현품서\현품서_*.xlsx  (기존 파일은 -Force 없이는 덮어쓰지 않음)
param(
  [Parameter(Mandatory)][string]$InputPath,
  [Parameter(Mandatory)][string]$OutDir,
  [string]$VendorMap = (Join-Path $PSScriptRoot 'vendor_map.csv'),
  [string]$SupportProducers = (Join-Path $PSScriptRoot 'support_producers.csv'),   # -SupportOnly 시 품목 패턴 -> 생산자(매입처)
  [switch]$Force,
  [switch]$SupportOnly   # 현물지원 시트(...지원, (현물))의 품목만으로 통계표/현품서 생성 (파일명 (친환경)->(현물지원))
)
$ErrorActionPreference = 'Stop'

$vendor = @{}
if (Test-Path -LiteralPath $VendorMap) {
  foreach ($r in Import-Csv -LiteralPath $VendorMap -Encoding utf8) { $vendor[$r.item] = $r }
}

$supportRules = @()
if ($SupportOnly -and (Test-Path -LiteralPath $SupportProducers)) { $supportRules = @(Import-Csv -LiteralPath $SupportProducers -Encoding utf8) }

$name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
if ($name -notmatch '^발주서_(.+)$') { throw "파일명이 '발주서_...' 형식이 아닙니다: $name" }
$suffix = $Matches[1]
if ($SupportOnly) { $suffix = $suffix -replace '\(친환경\)', '(현물지원)' }
$tongPath = Join-Path $OutDir "통계표\통계표_$suffix.xlsx"
$hpPath   = Join-Path $OutDir "현품서\현품서_$suffix.xlsx"
foreach ($p in $tongPath, $hpPath) {
  if ((Test-Path -LiteralPath $p) -and -not $Force) { throw "이미 존재: $p (덮어쓰려면 -Force)" }
  New-Item -ItemType Directory -Force -Path (Split-Path $p) | Out-Null
}

# 일(1~31) -> 표준 통계표 열 번호 (1~6: E~J, 7~18: L~W, 19~31: Y~AK; K·X는 여백 열)
function Get-DayCol([int]$d) { if ($d -le 6) { 4 + $d } elseif ($d -le 18) { 5 + $d } else { 6 + $d } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$report = [System.Collections.Generic.List[string]]::new()
try {
  $src = $xl.Workbooks.Open($InputPath, 0, $true)

  # 1) 헤더/품목행 수집. 품목행 = A열이 숫자(NO)이고 B열에 식품명이 있는 행.
  $layout = $null
  $sheets = @($src.Worksheets | Where-Object { $_.Visible -eq -1 })
  if ($SupportOnly) {
    $sheets = @($sheets | Where-Object { $_.Name -match '지원' -or $_.Name -match '\(현물\)\s*$' })
    if ($sheets.Count -eq 0) { $src.Close($false); "SKIP $suffix : 현물지원 시트 없음"; return }
  }
  $first = $null
  foreach ($ws in $sheets) {
    $ur = $ws.UsedRange; $last = $ur.Row + $ur.Rows.Count - 1
    for ($r = 1; $r -le [Math]::Min($last, 40); $r++) {
      if ($ws.Cells.Item($r, 1).Text.Trim() -ne 'NO') { continue }
      for ($c = 5; $c -le 60; $c++) {
        $t = $ws.Cells.Item($r, $c).Text.Trim()
        if ($t -eq '평균단가') { $layout = 'A'; break } elseif ($t -eq '결정단가') { $layout = 'B'; break }
      }
      if ($layout) { $first = $ws; break }
    }
    if ($layout) { break }
  }
  if (-not $layout) { throw '지원하지 않는 양식입니다(헤더에 평균단가/결정단가 없음).' }

  # 2) 사용할 시트 결정 (과거 수작업본 기준 규칙)
  #   - 현물지원 시트('...지원' 또는 '(현물)'로 끝나는 시트)는 통계표에 넣지 않고, 그 품목명은 다른 시트에서도 제외
  #   - '(현물포함)' 시트가 있으면 그 시트만, 없고 '(현물제외)'가 있으면 그 시트만, 둘 다 없으면 현물지원 시트를 뺀 전체
  function Get-SheetItems($ws) {
    $ur = $ws.UsedRange; $last = $ur.Row + $ur.Rows.Count - 1
    for ($r = 1; $r -le $last; $r++) {
      $a = $ws.Cells.Item($r, 1).Text.Trim()
      $b = $ws.Cells.Item($r, 2).Text.Trim()
      if ($a -match '^\d+$' -and $b) { [pscustomobject]@{ ws = $ws; row = $r; name = $b } }
    }
  }
  $support = @($sheets | Where-Object { $_.Name -match '지원' -or $_.Name -match '\(현물\)\s*$' })
  $incl = @($sheets | Where-Object { $_.Name -like '*현물포함*' })
  $excl = @($sheets | Where-Object { $_.Name -like '*현물제외*' })
  $rest = @($sheets | Where-Object { $support.Name -notcontains $_.Name })
  $use = if ($SupportOnly) { $sheets } elseif ($incl.Count -gt 0) { $incl } elseif ($excl.Count -gt 0) { $excl } else { $rest }
  $skipNames = @{}
  if (-not $SupportOnly) { foreach ($ws in $support) { foreach ($it in (Get-SheetItems $ws)) { $skipNames[$it.name] = $ws.Name } } }
  if (-not $SupportOnly -and $support.Count -gt 0) { $report.Add("현물지원 시트 제외: " + (($support | ForEach-Object { $_.Name }) -join ', ')) }

  $items = [System.Collections.Generic.List[object]]::new()
  $hdr = $null
  foreach ($ws in $use) {
    $ur = $ws.UsedRange; $last = $ur.Row + $ur.Rows.Count - 1
    for ($r = 1; $r -le $last; $r++) {
      $a = $ws.Cells.Item($r, 1).Text.Trim()
      $b = $ws.Cells.Item($r, 2).Text.Trim()
      if ($a -eq 'NO') { if (-not $hdr) { $hdr = @{ ws = $ws; row = $r } }; continue }
      if ($a -match '^\d+$' -and $b) {
        if (-not $SupportOnly -and ($skipNames.ContainsKey($b) -or $b -match '현물')) { $report.Add("현물지원 품목 제외: $b"); continue }
        $items.Add([pscustomobject]@{ ws = $ws; row = $r; name = $b; seq = $items.Count })
      }
    }
  }
  if (-not $hdr) { throw '헤더행(NO)을 찾지 못했습니다.' }
  if ($items.Count -eq 0) { throw '품목행이 없습니다.' }
  $hws = $hdr.ws; $hrow = $hdr.row

  # 열 위치 (총량 / 단가 / 총금액)
  $colQty = $null; $colAvg = $null; $colAmt = $null
  $ncol = $hws.UsedRange.Columns.Count
  for ($c = 1; $c -le $ncol; $c++) {
    $t = $hws.Cells.Item($hrow, $c).Text.Trim()
    if ($t -eq '총량') { $colQty = $c } elseif ($t -in '평균단가', '결정단가') { $colAvg = $c } elseif ($t -eq '총금액') { $colAmt = $c }
  }
  if (-not $colQty -or -not $colAvg) { throw '헤더에서 총량/단가 열을 찾지 못했습니다.' }

  # 식품명 ordinal(유니코드) 순 = 가나다순. 동명이면 원래 순서 유지.
  $sorted = [System.Collections.Generic.List[object]]::new($items)
  $sorted.Sort([Comparison[object]]{
    param($x, $y)
    $c = [string]::CompareOrdinal($x.name, $y.name)
    if ($c -ne 0) { $c } else { $x.seq - $y.seq }
  })

  # 셀 쓰기: 이 환경에서 .Value2/.Value 대입이 조용히 무시되는 경우가 있어, 읽어서 확인하고 다른 방식으로 재시도한다.
  function Set-Cell($ws, [int]$r, [int]$c, $v) {
    $cell = $ws.Cells.Item($r, $c)
    if ($null -eq $v -or "$v" -eq '') { [void]$cell.ClearContents(); return }
    $methods = if ($v -is [double]) { 'Value2', 'Value' } else { 'Formula', 'Value2', 'Value' }
    foreach ($m in $methods) {
      try { $cell.$m = $v } catch { continue }
      $got = $cell.Value2
      if ($v -is [double]) { if ($got -is [double] -and [math]::Abs($got - $v) -lt 1e-9) { return } }
      elseif ("$got".Trim() -eq "$v".Trim()) { return }
    }
    throw "셀 쓰기 실패: $($ws.Name) R${r}C${c} 값='$v'"
  }
  function Write-Vendor($hw, $i, $itemName) {
    if ($null -eq $itemName -or "$itemName" -eq '') { throw "품목명이 비어 있습니다(행 $i)" }
    foreach ($rule in $supportRules) {
      if ($itemName -match $rule.pattern) { Set-Cell $hw $i 7 $rule.producer; return }
    }
    $v = $vendor[$itemName]
    if ($v) {
      Set-Cell $hw $i 7 $v.vendor
      if ($v.alternatives) { $report.Add("매입처 후보 여러 개: $itemName -> $($v.vendor) (대안: $($v.alternatives))") }
    } else {
      Set-Cell $hw $i 7 $null
      $report.Add("매입처 미확인(빈칸): $itemName")
    }
  }

  if ($layout -eq 'A') {
    # ---- 양식 A: 원본 서식을 그대로 복사 ----
    $lastCol = $colAvg
    $t = $xl.Workbooks.Add(); $tw = $t.Worksheets.Item(1); $tw.Name = 'sheet1'
    [void]$hws.Range($hws.Cells.Item($hrow, 1), $hws.Cells.Item($hrow, $lastCol)).Copy($tw.Cells.Item(1, 1))
    $tw.Rows.Item(1).RowHeight = $hws.Rows.Item($hrow).RowHeight
    $i = 1
    foreach ($it in $sorted) {
      $i++; $s = $it.ws
      [void]$s.Range($s.Cells.Item($it.row, 1), $s.Cells.Item($it.row, $lastCol)).Copy($tw.Cells.Item($i, 1))
      $tw.Rows.Item($i).RowHeight = $s.Rows.Item($it.row).RowHeight
      Set-Cell $tw $i 1 ([double]($i - 1))
    }
    for ($c = 1; $c -le $lastCol; $c++) { $tw.Columns.Item($c).ColumnWidth = $hws.Columns.Item($c).ColumnWidth }
    $ur = $tw.UsedRange; [void]$ur.Copy(); [void]$ur.PasteSpecial(-4163); $xl.CutCopyMode = 0   # 수식 -> 값
    $t.SaveAs($tongPath, 51); $t.Close($false)

    $h = $xl.Workbooks.Add(); $hw = $h.Worksheets.Item(1); $hw.Name = 'sheet1'
    [void]$hws.Range($hws.Cells.Item($hrow, 1), $hws.Cells.Item($hrow, 4)).Copy($hw.Cells.Item(1, 1))
    foreach ($pair in @(@($colQty, 5), @($colAvg, 6), @($colAvg, 7))) { [void]$hws.Cells.Item($hrow, $pair[0]).Copy($hw.Cells.Item(1, $pair[1])) }
    Set-Cell $hw 1 5 '총량'; Set-Cell $hw 1 6 '평균단가'; Set-Cell $hw 1 7 '매입처'
    $hw.Rows.Item(1).RowHeight = $hws.Rows.Item($hrow).RowHeight
    $i = 1
    foreach ($it in $sorted) {
      $i++; $s = $it.ws; $r = $it.row
      [void]$s.Range($s.Cells.Item($r, 1), $s.Cells.Item($r, 4)).Copy($hw.Cells.Item($i, 1))
      foreach ($pair in @(@($colQty, 5), @($colAvg, 6), @($colAvg, 7))) { [void]$s.Cells.Item($r, $pair[0]).Copy($hw.Cells.Item($i, $pair[1])) }
      Set-Cell $hw $i 1 ([double]($i - 1))
      Write-Vendor $hw $i $it.name
      $hw.Rows.Item($i).RowHeight = $s.Rows.Item($r).RowHeight
    }
    $w = [double[]](1.7, 16, 3, 7.7, 4.3, 4.4, 7.6)
    for ($c = 1; $c -le 7; $c++) { $hw.Columns.Item($c).ColumnWidth = $w[$c - 1] }
    $ur = $hw.UsedRange; [void]$ur.Copy(); [void]$ur.PasteSpecial(-4163); $xl.CutCopyMode = 0
    $h.SaveAs($hpPath, 51); $h.Close($false)
  }
  else {
    # ---- 양식 B: 표준 통계표 템플릿 서식 위에 값을 채움 ----
    $tplT = Join-Path $PSScriptRoot 'template_통계표.xlsx'
    $tplH = Join-Path $PSScriptRoot 'template_현품서.xlsx'
    if (-not (Test-Path -LiteralPath $tplT) -or -not (Test-Path -LiteralPath $tplH)) { throw "템플릿 파일이 없습니다: $tplT / $tplH" }
    if ($suffix -notmatch '\d{4}$') { throw "파일명 끝에서 YYMM을 찾지 못했습니다: $suffix" }
    $year = 2000 + [int]($suffix.Substring($suffix.Length - 4, 2)); $mon = [int]$suffix.Substring($suffix.Length - 2, 2)
    $days = [DateTime]::DaysInMonth($year, $mon)
    $nameRow = 2   # 템플릿의 서식 행

    $tp = $xl.Workbooks.Open($tplT, 0, $true); $tpw = $tp.Worksheets.Item(1)
    $t = $xl.Workbooks.Add(); $tw = $t.Worksheets.Item(1); $tw.Name = 'sheet1'
    $lastCol = 41
    [void]$tpw.Range($tpw.Cells.Item(1, 1), $tpw.Cells.Item(1, $lastCol)).Copy($tw.Cells.Item(1, 1))
    $tw.Rows.Item(1).RowHeight = $tpw.Rows.Item(1).RowHeight
    # 일자 헤더: 템플릿은 30일 기준이므로 31일이 있는 달만 AK1에 31 추가
    if ($days -ge 31) { [void]$tpw.Cells.Item(1, 36).Copy($tw.Cells.Item(1, 37)); Set-Cell $tw 1 37 ([double]31) }
    else { Set-Cell $tw 1 37 $null }

    $i = 1
    foreach ($it in $sorted) {
      $i++; $s = $it.ws; $r = $it.row
      [void]$tpw.Range($tpw.Cells.Item($nameRow, 1), $tpw.Cells.Item($nameRow, $lastCol)).Copy($tw.Cells.Item($i, 1))
      $tw.Rows.Item($i).RowHeight = $tpw.Rows.Item($nameRow).RowHeight
      $tw.Range($tw.Cells.Item($i, 5), $tw.Cells.Item($i, 37)).ClearContents() | Out-Null
      Set-Cell $tw $i 1 ([double]($i - 1))
      $nm = ([string]$s.Cells.Item($r, 2).Text).Trim()
      $cv = ([string]$s.Cells.Item($r, 3).Text).Trim()
      $at = ([string]$s.Cells.Item($r, 4).Text).Trim()
      Set-Cell $tw $i 2 $nm
      Set-Cell $tw $i 3 $cv
      Set-Cell $tw $i 4 $at
      for ($d = 1; $d -le 31; $d++) {
        $v = $s.Cells.Item($r, 4 + $d).Value2
        if ($null -ne $v -and "$v" -ne '') {
          if ($d -gt $days) { $report.Add("경고: $($it.name) 의 ${d}일 수량($v)이 ${mon}월 일수를 넘어 누락되었습니다.") }
          else { Set-Cell $tw $i (Get-DayCol $d) ([double]$v) }
        }
      }
      Set-Cell $tw $i 38 ([double]$s.Cells.Item($r, $colQty).Value2)
      if ($colAmt) { Set-Cell $tw $i 39 ([double]$s.Cells.Item($r, $colAmt).Value2) }
      Set-Cell $tw $i 40 ([double]$s.Cells.Item($r, $colAvg).Value2)
    }
    for ($c = 1; $c -le $lastCol; $c++) { $tw.Columns.Item($c).ColumnWidth = $tpw.Columns.Item($c).ColumnWidth }
    $tp.Close($false)
    $t.SaveAs($tongPath, 51)

    $hp = $xl.Workbooks.Open($tplH, 0, $true); $hpw = $hp.Worksheets.Item(1)
    $h = $xl.Workbooks.Add(); $hw = $h.Worksheets.Item(1); $hw.Name = 'sheet1'
    [void]$hpw.Range($hpw.Cells.Item(1, 1), $hpw.Cells.Item(1, 7)).Copy($hw.Cells.Item(1, 1))
    $hw.Rows.Item(1).RowHeight = $hpw.Rows.Item(1).RowHeight
    for ($k = 2; $k -le $i; $k++) {
      [void]$hpw.Range($hpw.Cells.Item($nameRow, 1), $hpw.Cells.Item($nameRow, 7)).Copy($hw.Cells.Item($k, 1))
      $hw.Rows.Item($k).RowHeight = $hpw.Rows.Item($nameRow).RowHeight
      Set-Cell $hw $k 1 ([double]$tw.Cells.Item($k, 1).Value2)
      Set-Cell $hw $k 2 ([string]$tw.Cells.Item($k, 2).Value2)
      Set-Cell $hw $k 3 ([string]$tw.Cells.Item($k, 3).Value2)
      Set-Cell $hw $k 4 ([string]$tw.Cells.Item($k, 4).Value2)
      Set-Cell $hw $k 5 ([double]$tw.Cells.Item($k, 38).Value2)
      Set-Cell $hw $k 6 ([double]$tw.Cells.Item($k, 40).Value2)
      Write-Vendor $hw $k ($tw.Cells.Item($k, 2).Value2)
    }
    for ($c = 1; $c -le 7; $c++) { $hw.Columns.Item($c).ColumnWidth = $hpw.Columns.Item($c).ColumnWidth }
    $hp.Close($false)
    $t.Close($false)
    $h.SaveAs($hpPath, 51); $h.Close($false)
  }

  $src.Close($false)
  "OK [양식 $layout] $suffix : 품목 $($items.Count)개 -> $tongPath / $hpPath"
  $report | ForEach-Object { "  ! $_" }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
