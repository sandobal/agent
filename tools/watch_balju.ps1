# 감시 폴더에 놓인 발주서_*.xlsx 를 찾아 통계표/현품서(+현물지원)를 자동 생성한다.
# 작업 스케줄러가 주기적으로 실행한다(run_watch_hidden.vbs 경유). 직접 실행해도 된다.
#
#   <Folder>\발주서_○○(친환경)_YYMM.xlsx   (입력: 여기에 넣는다)
#   <Folder>\통계표\통계표_○○(친환경)_YYMM.xlsx
#   <Folder>\현품서\현품서_○○(친환경)_YYMM.xlsx
#   <Folder>\통계표\통계표_○○(현물지원)_YYMM.xlsx, 현품서\... (현물지원 시트가 있는 학교만)
#   <Folder>\처리결과\발주서_○○...txt        (매입처 빈칸/후보 여러 개 같은 확인 사항)
#
# 규칙: 결과가 없거나 입력보다 오래됐을 때만 만든다. 방금 복사 중인 파일은 건드리지 않는다.
#       변환에 실패한 파일은 파일이 바뀔 때까지 다시 시도하지 않는다. 로그: tools\watch_balju.log
param(
  [string]$Folder = 'C:\Users\PC\agent\output\2611_친환경',   # 발주서를 넣는 폴더
  [string]$OutRoot = '',                                     # 결과(통계표\, 현품서\, 처리결과\)를 둘 폴더. 비우면 $Folder 와 같음 (NAS는 '<월폴더>\두탑용')
  [int]$SettleSeconds = 15
)
$ErrorActionPreference = 'Continue'
$tools = $PSScriptRoot
$log   = Join-Path $tools 'watch_balju.log'
$state = Join-Path $tools 'watch_state.json'
function Log([string]$m) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $m" | Add-Content -LiteralPath $log -Encoding utf8 }

if (-not (Test-Path -LiteralPath $Folder)) { exit 0 }
if (-not $OutRoot) { $OutRoot = $Folder }

# 동시 실행 방지
$mutex = New-Object System.Threading.Mutex($false, 'Local\agent_watch_balju')
if (-not $mutex.WaitOne(0)) { exit 0 }
try {
  $failed = @{}
  if (Test-Path -LiteralPath $state) {
    try { (Get-Content -LiteralPath $state -Raw -Encoding utf8 | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $failed[$_.Name] = [long]$_.Value } } catch {}
  }
  $pwsh = (Get-Command pwsh -ErrorAction Stop).Source
  $conv = Join-Path $tools 'convert_balju.ps1'

  $inputs = Get-ChildItem -LiteralPath $Folder -Filter '발주서_*.xlsx' -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' }
  foreach ($f in $inputs) {
    if (((Get-Date) - $f.LastWriteTime).TotalSeconds -lt $SettleSeconds) { continue }   # 아직 복사/저장 중일 수 있음
    try { $fs = [IO.File]::Open($f.FullName, 'Open', 'Read', 'None'); $fs.Dispose() } catch { continue }   # 다른 프로그램이 열고 있음

    $suffix = [IO.Path]::GetFileNameWithoutExtension($f.Name) -replace '^발주서_', ''
    $t = Join-Path $OutRoot "통계표\통계표_$suffix.xlsx"
    $h = Join-Path $OutRoot "현품서\현품서_$suffix.xlsx"
    $have = (Test-Path -LiteralPath $t) -and (Test-Path -LiteralPath $h)
    if ($have -and (Get-Item -LiteralPath $t).LastWriteTime -ge $f.LastWriteTime -and (Get-Item -LiteralPath $h).LastWriteTime -ge $f.LastWriteTime) { continue }
    if ($failed.ContainsKey($f.FullName) -and $failed[$f.FullName] -eq $f.LastWriteTimeUtc.Ticks) { continue }

    Log "변환 시작: $($f.Name)"
    $lines = [System.Collections.Generic.List[string]]::new()
    $force = if ((Test-Path -LiteralPath $t) -or (Test-Path -LiteralPath $h)) { @('-Force') } else { @() }

    $out = & $pwsh -NoProfile -File $conv -InputPath $f.FullName -OutDir $OutRoot @force 2>&1 | ForEach-Object { "$_" }
    $ok = [bool]($out | Where-Object { $_ -like 'OK [[]*' })
    $out | ForEach-Object { $lines.Add($_) }

    if ($ok) {
      $out2 = & $pwsh -NoProfile -File $conv -InputPath $f.FullName -OutDir $OutRoot -SupportOnly -Force 2>&1 | ForEach-Object { "$_" }
      $out2 | ForEach-Object { $lines.Add($_) }
      Log "완료: $($f.Name)"
      $failed.Remove($f.FullName)
    } else {
      Log "실패: $($f.Name) :: $(($out | Select-Object -First 3) -join ' / ')"
      $failed[$f.FullName] = $f.LastWriteTimeUtc.Ticks
    }

    $rd = Join-Path $OutRoot '처리결과'
    New-Item -ItemType Directory -Force -Path $rd | Out-Null
    $lines | Set-Content -LiteralPath (Join-Path $rd "발주서_$suffix.txt") -Encoding utf8
  }

  ($failed.GetEnumerator() | ForEach-Object -Begin { $o = [ordered]@{} } -Process { $o[$_.Key] = $_.Value } -End { $o }) | ConvertTo-Json | Set-Content -LiteralPath $state -Encoding utf8
} finally {
  $mutex.ReleaseMutex()
}
