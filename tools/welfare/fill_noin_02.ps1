param([string]$Sp)
$ErrorActionPreference='Stop'
$d='\\nas2dual\2026\4.공급유통\02 공공급식\03. 괴산군노인복지관'
$cur="$d\2611\발주서_괴산노인복지관_02.xlsx"
function Norm([string]$s){($s -replace '\s','' -replace '[,()/·.]','').ToUpper()}
function ToNum($t){$x=("$t" -replace '[^\d.]','');if($x){[double]$x}else{[double]0}}
if(Test-Path -LiteralPath (Join-Path (Split-Path $cur) ('~$'+(Split-Path $cur -Leaf)))){'건너뜀: 열려 있음';return}
$xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;$xl.DisplayAlerts=$false
$price=@();$prev=@()
try{
 $tmp=Join-Path ([IO.Path]::GetTempPath()) 'nh_price.xlsx';Copy-Item -LiteralPath "$d\2611\복지관 매입처.xlsx" $tmp -Force
 $wb=$xl.Workbooks.Open($tmp,0,$true);$ws=$wb.Worksheets.Item(1)
 $v=$ws.UsedRange.Value2;$n=$ws.UsedRange.Rows.Count
 for($r=2;$r -le $n;$r++){$nm=("$($v[$r,2])").Trim();if($nm){$price+=[pscustomobject]@{row=$r;vendor=("$($v[$r,1])").Trim();name=$nm;price=(ToNum $v[$r,3])}}}
 $wb.Close($false)
 $srcs=@(,@(99,"$d\2611\발주서_괴산노인복지관_01.xlsx"));foreach($f in '04','03','02','01'){$srcs+=,@([int]$f,"$d\2610\발주서_괴산노인복지관_$f.xlsx")}
 foreach($sx in $srcs){ if(-not(Test-Path -LiteralPath $sx[1])){continue}
  $tmp=Join-Path ([IO.Path]::GetTempPath()) 'nh_prev.xlsx';Copy-Item -LiteralPath $sx[1] $tmp -Force
  $wb=$xl.Workbooks.Open($tmp,0,$true);$ws=$wb.Worksheets.Item(1)
  for($r=2;$r -le $ws.UsedRange.Rows.Count;$r++){$nm=$ws.Cells.Item($r,2).Text.Trim();if($nm -and $ws.Cells.Item($r,5).Text.Trim()){$prev+=[pscustomobject]@{order=$sx[0];name=$nm;vendor=$ws.Cells.Item($r,5).Text.Trim();price=(ToNum $ws.Cells.Item($r,6).Text)}}}
  $wb.Close($false)}
 Copy-Item -LiteralPath $cur "$Sp\backup\발주서_괴산노인복지관_02_before2.xlsx" -Force
 $wb=$xl.Workbooks.Open($cur);$ws=$wb.Worksheets.Item(1);$report=@()
 for($r=2;$r -le 60;$r++){
  $name=$ws.Cells.Item($r,2).Text.Trim();if(-not $name){continue}
  $hasV=[bool]$ws.Cells.Item($r,5).Text.Trim();$hasP=[bool]$ws.Cells.Item($r,6).Text.Trim();if($hasV -and $hasP){continue}
  $nn=Norm $name
  $cands=@($price|?{(Norm $_.name) -eq $nn});$how='정확'
  if($cands.Count -eq 0 -and $nn.Length -ge 2){$cands=@($price|?{$pn=Norm $_.name;$pn.Length -ge 2 -and ($pn.Contains($nn) -or $nn.Contains($pn))});$how='포함'}
  $pv=@($prev|?{(Norm $_.name) -eq $nn}|Sort-Object order -Descending)
  $pick=$null;$why=''
  if($cands.Count -gt 0){
   $groups=@($cands|Group-Object{"$($_.vendor)|$($_.price)"});$hit=$null
   foreach($p in $pv){$k="$($p.vendor)|$($p.price)";$g=$groups|?{$_.Name -eq $k}|Select-Object -First 1;if($g){$hit=$g;$why='이전 발주서와 같은 값';break}}
   if(-not $hit){$max=($groups|%{$_.Count}|Measure-Object -Maximum).Maximum;$top=@($groups|?{$_.Count -eq $max})
    if($top.Count -eq 1){$hit=$top[0];$why='가장 많이 나온 값'}else{$hit=$top|Sort-Object{($_.Group|%{$_.row}|Measure-Object -Maximum).Maximum} -Descending|Select-Object -First 1;$why='동률 중 최근 행'}}
   $kv=$hit.Name -split '\|';$pick=[pscustomobject]@{vendor=$kv[0];price=[double]$kv[1]};if($groups.Count -eq 1){$why='단일'};$why+=" [$how]"}
  if(-not $pick -and $pv.Count -gt 0){$pick=[pscustomobject]@{vendor=$pv[0].vendor;price=$pv[0].price};$why="이전 발주서($($pv[0].order))"}
  if(-not $pick){$report+="비움: [$r] $name";continue}
  $w=@()
  if(-not $hasV){$ws.Cells.Item($r,5).Formula=[string]$pick.vendor;$w+="매입처 $($pick.vendor)"}
  if(-not $hasP -and $pick.vendor -eq '중앙상회'){$w+='단가 중앙상회라 비움'}
  elseif(-not $hasP -and $pick.price -gt 0){$ws.Cells.Item($r,6).Value2=[double]$pick.price;$w+="단가 $($pick.price)"}
  elseif(-not $hasP){$w+='단가 0원이라 비움'}
  $report+="입력: [$r] $name -> "+($w -join ', ')+" ($why)"}
 $wb.Save();$wb.Close($false);$report
}finally{$xl.Quit()}


