param([string]$Sp, [string]$Target, [string]$SchoolList = '')
# 학교별 발주서에서 육류(돼지/닭/한우/오리/소사골/소잡뼈)만 골라 통계표 형식 시트로 만들어 Target 통합문서에 추가한다.
$ErrorActionPreference = 'Stop'
$base = "\\nas2dual\2026\4.공급유통\01 학교급식\01 친환경급식\2611\"
$schools = '오성중', '장연초', '청안초', '청천초', '칠성초', '감물초', '괴산북중', '괴산중', '명덕초', '목도나루', '목도초', '문광초', '백봉초', '보광초', '송면초', '연풍초'
if ($SchoolList) { $schools = @($SchoolList -split ',' | Where-Object { $_ }) }
$meat = '^(돼지|닭|한우|오리|소사골|소잡뼈|소고기)'
New-Item -ItemType Directory -Force "$Sp\w" | Out-Null

# 1) 학교/식사별 임시 발주서 만들기
$jobs = @()
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  foreach ($s in $schools) {
    $p = "${base}발주서_${s}(친환경)_2611.xlsx"
    $wb = $xl.Workbooks.Open($p, 0, $true)
    $groups = @()
    if ($s -eq '오성중') { $groups = @('조식', '중식', '석식') }
    else {
      $label = ''
      foreach ($ws in $wb.Worksheets) {
        if ($ws.Visible -ne -1 -or $ws.Name -match '지원|\(현물\)') { continue }
        $t = $ws.Cells.Item(5, 1).Text.Trim()
        if ($t -match '\(([^)]*)\)\s*$') { $label = $Matches[1]; break }
      }
      if (-not $label) { $label = '중식' }
      $groups = @($label)
    }
    $wb.Close($false)
    foreach ($g in $groups) {
      $d = "$Sp\w\${s}_$g"; New-Item -ItemType Directory -Force $d | Out-Null
      $tmp = "$d\발주서_${s}(친환경)_2611.xlsx"
      Copy-Item -LiteralPath $p $tmp -Force
      if ($s -eq '오성중') {
        $a = $xl.Workbooks.Open($tmp)
        foreach ($ws in @($a.Worksheets)) { if ($ws.Name -ne $g) { $ws.Delete() } }
        $a.Save(); $a.Close($false)
      }
      $jobs += [pscustomobject]@{ school = $s; group = $g; dir = $d; tmp = $tmp }
    }
  }
} finally { $xl.Quit() }

# 2) 변환
foreach ($j in $jobs) {
  $o = & pwsh -NoProfile -File C:\Users\PC\agent\tools\convert_balju.ps1 -InputPath $j.tmp -OutDir "$($j.dir)\out" 2>&1 | Select-Object -First 1
  "$($j.school)($($j.group)) : " + ($o -replace ' -> .*$', '')
}

# 3) 육류만 남겨 Target에 시트로 추가
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $tw = $xl.Workbooks.Open($Target)
  foreach ($j in $jobs) {
    $f = "$($j.dir)\out\통계표\통계표_$($j.school)(친환경)_2611.xlsx"
    if (-not (Test-Path -LiteralPath $f)) { "건너뜀(변환 결과 없음): $($j.school)($($j.group))"; continue }
    $uniq = "$($j.dir)\m_$($j.school)_$($j.group).xlsx"
    Copy-Item -LiteralPath $f $uniq -Force
    $wb = $xl.Workbooks.Open($uniq); $ws = $wb.Worksheets.Item(1)
    $last = $ws.UsedRange.Rows.Count
    for ($r = $last; $r -ge 2; $r--) {
      $nm = $ws.Cells.Item($r, 2).Text.Trim()
      if ($nm -and $nm -notmatch $meat) { $ws.Rows.Item($r).Delete() | Out-Null }
    }
    $n = 0
    $last = $ws.UsedRange.Rows.Count
    for ($r = 2; $r -le $last; $r++) { if ($ws.Cells.Item($r, 2).Text.Trim()) { $n++; $ws.Cells.Item($r, 1).Value2 = [double]$n } }
    if ($n -eq 0) { "육류 없음 -> 시트 만들지 않음: $($j.school)($($j.group))"; $wb.Close($false); continue }
    $name = "$($j.school)($($j.group))"
    $ws.Copy([Type]::Missing, $tw.Worksheets.Item($tw.Worksheets.Count))
    $new = $tw.Worksheets.Item($tw.Worksheets.Count); $new.Name = $name
    $wb.Close($false)
    "시트 추가: $name (육류 ${n}품목)"
  }
  $tw.Save();
  "최종 시트: " + (($tw.Worksheets | ForEach-Object { $_.Name }) -join ', ')
  $tw.Close($false)
} finally { $xl.Quit() }
