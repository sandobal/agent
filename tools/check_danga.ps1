# 단가심의자료(품목/공급단가/원물공급자)와 변환된 현품서의 단가·매입처를 대조한다.
# 사용: pwsh -File check_danga.ps1 -DangaXlsx <단가심의자료.xlsx> -Folder <output\2611_친환경> [-Month 11] [-Report <csv경로>]
#   Folder\현품서\현품서_*(친환경)_*.xlsx 의 각 행(식품명/규격/총량/평균단가/매입처)을 단가자료에 대응시켜
#   단가 불일치 / 매입처 불일치 / 대응 품목 없음 을 구분해 출력한다. (읽기 전용: 원본 파일은 수정하지 않음)
param(
  [Parameter(Mandatory)][string]$DangaXlsx,
  [Parameter(Mandatory)][string]$Folder,
  [int]$Month = 11,
  [string]$Report
)
$ErrorActionPreference = 'Stop'

# 발주서 식품명(정규식) -> 단가자료 지원품목명. 위에서부터 먼저 맞는 규칙을 사용한다.
$rules = @(
  @('^연근','깐연근채'), @('^한우\(채끝\)','한우 채끝'), @('^찹쌀\(흑미\)','찰흑미'), @('^우엉','깐우엉채'), @('^혼합잡곡','혼합5곡(2025년산)'), @('^보리\(쌀보리\)','쌀보리'), @('^오리고기','오리 정육'),
  @('^닭고기\(윙\)|^닭고기\(닭날개\)','계육 윙'), @('^돼지고기\(등갈비\)','돈육 등갈비'), @('^배\(kg\)','배(신고/25내)'),
  @('^멥쌀\(백미\)','멥쌀백미(2025년산)'), @('^찹쌀\(백미\)','찹쌀백미(2025년산)'), @('^찹쌀\(현미\)','찰현미'),
  @('^멥쌀\(발아현미\)','발아현미'), @('^멥쌀\(현미\)','현미'), @('^멥쌀\(흑미\)','찰흑미'),
  @('^귀리','귀리'), @('^기장','찰기장'), @('^녹두','깐녹두'), @('^보리\(찰보리\)','찰보리'), @('^보리\(할맥\)','압맥'),
  @('^수수','찰수수'), @('^차조','청차조'), @('^통밀','통밀쌀'), @('^혼합10곡','혼합15곡'), @('^완두콩','깐완두콩(냉동)'),
  @('^깨\(참깨\)','볶은참깨'), @('^된장','된장(과세)'), @('^표고버섯가루','표고버섯가루'), @('^찹쌀가루','찹쌀가루'), @('^콩가루','생콩가루'),
  @('^고춧가루','고춧가루(2025년산)'),
  @('^감자/껍질제거','깐감자'), @('^감자','감자'), @('^양파','양파'), @('^마늘','깐마늘'), @('^파\(','대파'), @('^무\(조선무\)','무우'),
  @('^양배추\(적채','적양배추'), @('^양배추','깐양배추'), @('^당근','당근'),
  @('^깻잎','깻잎'), @('^근대','근대'), @('^시금치','시금치'), @('^쑥갓','쑥갓'), @('^아욱','아욱'), @('^미나리','미나리'),
  @('^청경채','청경채'), @('^부추','부추'), @('^배추\(얼갈이','얼갈이배추'), @('^상추','적상추'), @('^브로콜리','브로콜리'),
  @('^호박\(애호박\)','애호박'), @('^오이','백오이'), @('^토마토\(대추방울','대추방울토마토'), @('^토마토\(방울토마토','대추방울토마토'), @('^토마토','완숙토마토'),
  @('^큰느타리','새송이버섯'), @('^애느타리|^느타리','느타리버섯'), @('^양송이','갈색양송이버섯'), @('^표고버섯','생표고버섯(향고)'), @('^팽이버섯','팽이버섯'),
  @('^숙주','숙주나물'), @('^콩나물','콩나물'), @('^사과','사과(부사/40내)'), @('^배\(신고\)|^배/','배(신고/25내)'),
  @('^고사리','데친고사리'), @('^곤드레','데친곤드레'), @('^도라지','깐도라지채'), @('^달걀','유정란(동물복지)'),
  @('^닭고기\(가슴살, 안심살\)/덩어리','계육 가슴살(스킨무)'), @('^닭고기\(가슴살, 안심살\)','계육 가슴살'),
  @('^닭고기\(넓적다리\)/냉장','계육 정육'), @('^닭고기\(넓적다리\)','계육 정육(스킨무)'),
  @('^닭고기\(다리, 북채\)','계육 북채'), @('^닭고기\(다리살\)','계육 정육(스킨무)'), @('^닭고기\(봉\)','계육 봉'),
  @('^닭고기\(살코기\)','계육 정육'), @('^닭고기\(성계\)/껍질제거','계육 통마리 1/2(스킨무)'),
  @('^닭고기\(성계\)/친환경','계육 통마리(스킨무)'), @('^닭고기\(성계\)','계육 통마리'),
  @('^돼지고기\(갈비\)','돈육 갈비(냉동)'), @('^돼지고기\(뒷다리','돈육 뒷다리살'), @('^돼지고기\(등심\)','돈육 등심살'),
  @('^돼지고기\(목살\)','돈육 목살'), @('^돼지고기\(사태\)','돈육 사태살'), @('^돼지고기\(삼겹살\)','돈육 삼겹살'),
  @('^돼지고기\(안심\)','돈육 안심'), @('^돼지고기\(앞다리','돈육 앞다리살'), @('^돼지등뼈','돈육 등뼈(냉동)'),
  @('^한우\(갈비\)','한우 갈비(냉동)'), @('^한우\(목심','한우 목심'), @('^한우\(사태\)','한우 사태'), @('^한우\(설도\)','한우 설도'),
  @('^한우\(안심\)','한우 안심'), @('^한우\(앞다리','한우 앞다리'), @('^한우\(양지\)','한우 양지'), @('^한우\(우둔\)','한우 우둔'),
  @('^소사골','한우 사골(냉동)'), @('^소잡뼈','한우 잡뼈(냉동)')
)

# 발주서 규격 -> 단가자료 규격
function Get-DangaSpec([string]$spec, [string]$name) {
  switch -Regex ($spec) {
    '^0\.50kg$' { return '500g' } '^1\.00kg$' { return '1kg' } '^10\.00kg$' { return '10kg' } '^4\.00kg$' { return '4kg' }
    '^3\.00kg$' { return '3kg' } '^0\.10kg$' { return '100g' } '^0\.30kg$' { return '300g' }
    '^1\.[58]0kg$' { return '30개/판' }
  }
  return ''
}
function ToNum($t) { $s = ("$t" -replace '[^\d.]', ''); if ($s) { [double]$s } else { $null } }
function Split-Sup($s) { @(((("$s" -replace '한살림축산식품|한축식품|한축','한살림축산')) -split '[/,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$danga = [System.Collections.Generic.List[object]]::new()
$rows  = [System.Collections.Generic.List[object]]::new()
try {
  $wb = $xl.Workbooks.Open($DangaXlsx, 0, $true); $ws = $wb.Worksheets.Item(1)
  $hdr = $null
  for ($c = 1; $c -le 20; $c++) { if ($ws.Cells.Item(2, $c).Text -match "^\s*${Month}월") { $hdr = $c } }
  if (-not $hdr) { throw "$Month월 단가 열을 찾지 못했습니다." }
  $supCol = $null
  for ($c = 1; $c -le 20; $c++) { if ($ws.Cells.Item(2, $c).Text.Trim() -eq '원물공급자') { $supCol = $c } }
  if (-not $supCol) { throw '원물공급자 열을 찾지 못했습니다.' }
  for ($r = 3; $r -le $ws.UsedRange.Rows.Count; $r++) {
    $n = $ws.Cells.Item($r, 4).Text.Trim(); if (-not $n) { continue }
    $danga.Add([pscustomobject]@{ name = $n; spec = $ws.Cells.Item($r, 5).Text.Trim(); price = (ToNum ($ws.Cells.Item($r, $hdr).Text)); supplier = $ws.Cells.Item($r, $supCol).Text.Trim() })
  }
  $wb.Close($false)

  foreach ($f in Get-ChildItem -LiteralPath (Join-Path $Folder '현품서') -Filter '현품서_*(친환경*)_*.xlsx' -File) {
    $school = (($f.BaseName -replace '^현품서_', '') -replace '\(친환경_육류\).*$', '(육류)') -replace '\(친환경\).*$', ''
    $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
    foreach ($ws in $wb.Worksheets) {   # 중식/석식처럼 시트가 나뉜 현품서도 모두 확인
      for ($r = 2; $r -le $ws.UsedRange.Rows.Count; $r++) {
        $n = $ws.Cells.Item($r, 2).Text.Trim(); if (-not $n) { continue }
        if ($n -match '^닭뼈') { continue }   # 사용자 지시(2026-10-16): 닭뼈는 단가 확인에서 제외
        $rows.Add([pscustomobject]@{ school = $school; name = $n; spec = $ws.Cells.Item($r, 3).Text.Trim(); qty = $ws.Cells.Item($r, 5).Value2
                                    price = [double]$ws.Cells.Item($r, 6).Value2; vendor = $ws.Cells.Item($r, 7).Text.Trim() })
      }
    }
    $wb.Close($false)
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

$out = foreach ($it in $rows) {
  $target = $null
  foreach ($rule in $rules) { if ($it.name -match $rule[0]) { $target = $rule[1]; break } }
  $cands = @($danga | Where-Object { $_.name -eq $target })
  $res = [ordered]@{ school = $it.school; item = $it.name; spec = $it.spec; 발주단가 = $it.price; 현품서매입처 = $it.vendor
                     단가자료품목 = ''; 단가자료규격 = ''; 단가자료단가 = $null; 단가자료공급자 = ''; 단가 = ''; 매입처 = ''; 비고 = '' }
  if (-not $target -or $cands.Count -eq 0) {
    $res.단가 = '대응없음'; $res.매입처 = '대응없음'
    # 단가가 같은 다른 품목이 있으면 힌트로 표시
    $same = @($danga | Where-Object { $_.price -eq $it.price } | Select-Object -First 3 | ForEach-Object { $_.name })
    if ($same.Count) { $res.비고 = '단가 동일 품목: ' + ($same -join ', ') }
    [pscustomobject]$res; continue
  }
  # 규격 선택
  $want = Get-DangaSpec $it.spec $it.name
  $row = $null
  if ($want) { $row = $cands | Where-Object { $_.spec -eq $want } | Select-Object -First 1 }
  if (-not $row) { $row = $cands | Where-Object { $_.spec -in 'kg', '1kg' } | Select-Object -First 1 }
  if (-not $row) { $row = $cands | Select-Object -First 1 }
  $expected = $row.price
  # 쌀류: 발주서 단위가 kg 이고 단가자료가 10kg/4kg 단위면 kg당 단가로 환산
  if ($it.name -match '\(kg\)$' -and $row.spec -match '^(\d+)kg$' -and [int]$Matches[1] -gt 1) { $expected = $row.price / [int]$Matches[1] }
  # 단가자료가 500g 단가이고 발주서 규격이 1.00kg 이면 1kg 단가(x2)로 환산 (예: 완두콩 냉동)
  if ($it.spec -eq '1.00kg' -and $row.spec -eq '500g') { $expected = $row.price * 2 }
  $res.단가자료품목 = $row.name; $res.단가자료규격 = $row.spec; $res.단가자료단가 = $expected; $res.단가자료공급자 = $row.supplier
  $priceOk = ($null -ne $expected) -and ([math]::Abs($expected - $it.price) -lt 0.5)
  $res.단가 = if ($priceOk) { '일치' } else { '불일치' }
  if (-not $priceOk) {
    $alt = @($danga | Where-Object { $_.price -eq $it.price } | Select-Object -First 3 | ForEach-Object { "$($_.name) $($_.spec)" })
    if ($alt.Count) { $res.비고 = '발주단가와 같은 품목: ' + ($alt -join ', ') }
  }
  if (-not $it.vendor) { $res.매입처 = '현품서 빈칸' }
  else {
    $a = Split-Sup $it.vendor; $b = Split-Sup $row.supplier
    $res.매입처 = if (($a -join '|') -eq ($b -join '|')) { '일치' } else { '불일치' }
  }
  [pscustomobject]$res
}

if ($Report) { $out | Export-Csv -LiteralPath $Report -NoTypeInformation -Encoding utf8 }
$out
