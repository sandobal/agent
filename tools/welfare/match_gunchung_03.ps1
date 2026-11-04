$ErrorActionPreference='Stop'
$sp='C:\Users\PC\AppData\Local\Temp\claude\C--Users-PC-agent\1ae278cc-bd52-48c7-ad0d-7cbb295648ac\scratchpad\gc'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
function Norm($s){ ([string]$s).ToUpper() -replace '\s','' }
function ReadOrder($f){
  $wb=$xl.Workbooks.Open((Join-Path $sp $f),0,$true); $ws=$wb.Worksheets.Item(1); $ur=$ws.UsedRange
  $h=0; for($r=1;$r -le 6;$r++){ if(([string]$ws.Cells.Item($r,1).Text).Trim().ToUpper() -eq 'NO'){$h=$r;break} }
  $cm=@{}; for($c=1;$c -le 14;$c++){ $cm[([string]$ws.Cells.Item($h,$c).Text).Trim()]=$c }
  $out=@(); for($r=$h+1;$r -le $ur.Rows.Count;$r++){
    $nm=([string]$ws.Cells.Item($r,2).Text).Trim(); if(-not $nm){continue}
    $out+=[pscustomobject]@{row=$r;name=$nm;spec=([string]$ws.Cells.Item($r,$cm['규격']).Text).Trim();vendor=([string]$ws.Cells.Item($r,$cm['매입처']).Text).Trim();price=[double]($ws.Cells.Item($r,$cm['단가']).Value2+0)}
  }
  $wb.Close($false); ,$out }
$cur=ReadOrder '괴산군청 구내식당_3.xlsx'
$p1=ReadOrder '괴산군청 구내식당_01.xlsx'
$p2=ReadOrder '괴산군청 구내식당_02.xlsx'
$dg=Import-Csv (Join-Path $sp 'gdanga.csv') | %{ [pscustomobject]@{vendor=$_.v;name=$_.n;spec=$_.u;price=[double]$_.p} } | ?{ $_.price -gt 0 -and $_.vendor }
$req=Import-Csv (Join-Path $sp 'req.csv')
$res=@()
foreach($it in $cur){
  $nn=Norm $it.name; $ns=Norm $it.spec; $src=''; $v=''; $pr=0; $note=''
  foreach($pp in @(@('02',$p2),@('01',$p1))){
    if($src){break}
    $h=@($pp[1] | ?{ (Norm $_.name) -eq $nn -and (Norm $_.spec) -eq $ns -and $_.price -gt 0 -and $_.vendor })
    if($h.Count){ $v=$h[-1].vendor; $pr=$h[-1].price; $src="전월발주_$($pp[0])" }
  }
  if(-not $src){
    $h=@($dg | ?{ (Norm $_.name) -eq $nn -and (Norm $_.spec) -eq $ns })
    if(-not $h.Count){ $h=@($dg | ?{ (Norm $_.name) -eq $nn }); if($h.Count){$note='규격불일치'} }
    if($h.Count){
      $g=@($h | Group-Object { "$($_.vendor)|$($_.price)" } | Sort-Object Count -Descending)
      $top=$g[0]; $kv=$top.Name -split '\|'; $v=$kv[0]; $pr=[double]$kv[1]; $src="군청단가($($top.Count)/$($h.Count))"
      if($g.Count -gt 1){ $note+=' 후보:'+(($g|%{"$($_.Name)x$($_.Count)"}) -join ',') }
    }
  }
  if(-not $src){
    $h=@($req | ?{ (Norm $_.name) -eq $nn -and $_.price })
    if(-not $h.Count){ $h=@($req | ?{ (Norm $_.name) -like "$nn*" -and $_.price }) }
    if($h.Count){ $pr=[double]$h[0].price; $src="계약리스트:$($h[0].sheet)/$($h[0].name)/$($h[0].unit)"; $note='매입처없음' }
  }
  $res+=[pscustomobject]@{row=$it.row;name=$it.name;spec=$it.spec;vendor=$v;price=$pr;src=$src;note=$note}
}
$res | %{ "{0,2}| {1} | {2} | {3} | {4:N0} | {5} {6}" -f $_.row,$_.name,$_.spec,$_.vendor,$_.price,$_.src,$_.note }
$res | Export-Csv (Join-Path $sp 'proposal.csv') -NoTypeInformation -Encoding UTF8
$xl.Quit()
