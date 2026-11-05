$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $MyInvocation.MyCommand.Path
$Path=Join-Path $Root 'build_daily.ps1'
if(!(Test-Path -LiteralPath $Path)){throw "Khong tim thay build_daily.ps1 cung folder voi file patch."}
$s=Get-Content -LiteralPath $Path -Raw -Encoding UTF8
if($s -match 'DISPLAY_RESULT' -and $s -match 'Read-DisplayResultSheet'){Write-Host 'build_daily.ps1 da co fix Ket qua trung bay.';exit 0}
$backup=Join-Path $Root ('build_daily_before_result_fix_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.ps1')
Copy-Item -LiteralPath $Path -Destination $backup -Force

$helpers=@'
function HeaderKey($v){
  $x=(Text $v).ToUpperInvariant().Normalize([Text.NormalizationForm]::FormD)
  $x=-join($x.ToCharArray()|Where-Object{[Globalization.CharUnicodeInfo]::GetUnicodeCategory($_)-ne[Globalization.UnicodeCategory]::NonSpacingMark})
  return ([regex]::Replace($x,'[^A-Z0-9]+',' ').Trim())
}
function SetKey($v){
  $x=HeaderKey $v
  $x=$x -replace '^FOODSET ','FOOD SET ' -replace '^BPCSET ','BPC SET ' -replace '^HCSET ','HC SET '
  $x=$x -replace '^SACHET 1 3$','SACHET 1-3'
  return $x
}
function ResultValue($v){
  $raw=Text $v;if(!$raw){return ''};$x=HeaderKey $raw
  if($x -match '(^| )KTG($| )'){return 'KTG'}
  if($x -match '(^| )KTC($| )'){return 'KTC'}
  if($x -match 'ROT'){return 'Rớt'}
  if($x -match 'DAT'){return 'Đạt'}
  return $raw
}
function FindHeaderRow($sheet){
  $max=[Math]::Min(15,$sheet.UsedRange.Rows.Count)
  for($r=1;$r-le$max;$r++){
    $m=@{};for($c=1;$c-le$sheet.UsedRange.Columns.Count;$c++){ $k=HeaderKey $sheet.Cells.Item($r,$c).Value2;if($k){$m[$k]=$c} }
    $hasCode=@('MASTER OUTLET CODE','OUTLET CODE','MA CH','CODE CH','MA OUTLET')|Where-Object{$m.ContainsKey($_)}
    $hasSet=@('SET','TEN SET','MUC TRUNG BAY','CHUONG TRINH','PROGRAM')|Where-Object{$m.ContainsKey($_)}
    $hasResult=@('KET QUA TRUNG BAY','KET QUA','RESULT','TRANG THAI')|Where-Object{$m.ContainsKey($_)}
    if($hasCode-and$hasSet-and$hasResult){return [pscustomobject]@{Row=$r;Map=$m}}
  }
  return $null
}
function FirstCol($map,$names){foreach($n in $names){if($map.ContainsKey($n)){return $map[$n]}};return 0}
function Read-DisplayResultSheet($workbook){
  $out=@{};$sheet=$null
  foreach($ws in @($workbook.Worksheets)){if((HeaderKey $ws.Name)-eq'KET QUA TRUNG BAY'){$sheet=$ws;break}}
  if($null-eq$sheet){Write-Warning 'Khong tim thay sheet Ket qua trung bay.';return $out}
  try{
    $h=FindHeaderRow $sheet;if($null-eq$h){Write-Warning 'Khong nhan dien duoc cot Ma CH / Set / Ket qua trong sheet Ket qua trung bay.';return $out}
    $cm=FirstCol $h.Map @('MASTER OUTLET CODE','OUTLET CODE','MA CH','CODE CH','MA OUTLET')
    $sm=FirstCol $h.Map @('SET','TEN SET','MUC TRUNG BAY','CHUONG TRINH','PROGRAM')
    $rm=FirstCol $h.Map @('KET QUA TRUNG BAY','KET QUA','RESULT','TRANG THAI')
    $u=$sheet.UsedRange;$a=$u.Value2;$rows=$u.Rows.Count
    for($r=$h.Row+1;$r-le$rows;$r++){
      $code=Code(GetValue $a $r $cm);$set=SetKey(GetValue $a $r $sm);$result=ResultValue(GetValue $a $r $rm)
      if(!$code-or!$set-or!$result){continue}
      if(!$out.ContainsKey($code)){$out[$code]=@{}}
      $out[$code][$set]=$result
    }
    Release-Com $u
    Write-Host ("  Ket qua trung bay: {0} CH" -f $out.Count)
    return $out
  } finally {Release-Com $sheet}
}
'@
$s=$s.Replace('function GetValue($arr,$r,$c) { if($null -eq $arr){return $null}; return $arr[$r,$c] }',"function GetValue(`$arr,`$r,`$c) { if(`$null -eq `$arr){return `$null}; return `$arr[`$r,`$c] }`r`n$helpers")

$s=$s.Replace("`$wb.RefreshAll();`$excel.CalculateUntilAsyncQueriesDone()","`$wb.RefreshAll();`$excel.CalculateUntilAsyncQueriesDone()")
$s=$s.Replace('    $stores=@{}',"    `$displayResult=Read-DisplayResultSheet `$wb`r`n    `$stores=@{}")

# Remove banker-derived display result injection completely.
$s=[regex]::Replace($s,'(?s)  # Dat/Rot cung bat buoc dung N-1 \+ N\..*?  # CSS cho badge Dat/Rot; chi chen mot lan\..*?  }\r?\n','',1)

# Extend Add-CurrentSalesSupport and inject DISPLAY_RESULT.
$s=$s.Replace('function Add-CurrentSalesSupport($html,$sales){','function Add-CurrentSalesSupport($html,$sales,$displayResult){')
$needle='  $json=Json $sales'
$inject=@'
  $json=Json $sales
  $resultJson=Json $displayResult
  if($html -match 'const CURRENT_SALES='){$html=$html.Replace('const CURRENT_SALES=',('const DISPLAY_RESULT='+$resultJson+';const CURRENT_SALES='))}
  elseif($html -match 'const RULES='){$html=$html.Replace('const RULES=',('const DISPLAY_RESULT='+$resultJson+';const RULES='))}
'@
$s=$s.Replace($needle,$inject)

# Insert file-based badge helper before currentBankerHTML and rename banker wording.
$marker='  # Chi thay metric trong currentBankerHTML, khong thay metric cua de nghi.'
$displayHelper=@'
  # Ket qua trung bay lay tu sheet Ket qua trung bay, theo Ma CH + Set.
  if($html -notmatch 'function currentSetResultHTML\('){
    $helper=@"
function displayResultFor(o,label){let byStore=DISPLAY_RESULT[String(o.c)]||{};return byStore[norm(label)]||''}
function currentSetResultHTML(o,label){let x=displayResultFor(o,label);if(!x)return '<span class="tb-result tb-na">Chưa có kết quả</span>';let n=norm(x),cls=n==='DAT'?'tb-pass':n==='ROT'?'tb-fail':'tb-special';return '<span class="tb-result '+cls+'">Kết quả trưng bày: '+esc(x)+'</span>'}
"@
    $html=$html.Replace('function currentBankerHTML',$helper+'function currentBankerHTML')
    $html=$html.Replace('${esc(x.replaceAll(''_'','' ''))}${currentBankerHTML(o,x)}','${esc(x.replaceAll(''_'','' ''))} ${currentSetResultHTML(o,x)}${currentBankerHTML(o,x)}')
  }
  if($html -notmatch '\.tb-special\{'){
    $css='.tb-result{display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:900;vertical-align:middle}.tb-pass{background:#D9FBEA;color:#087A4F}.tb-fail{background:#FFE1E1;color:#B42318}.tb-special{background:#FFF2CC;color:#8A5A00}.tb-na{background:#EDF2F4;color:#60717B}'
    $html=$html.Replace('</style>',$css+'</style>')
  }
'@
$s=$s.Replace($marker,$displayHelper+"`r`n"+$marker)
$s=$s.Replace("'Đã đạt':'Gap: '+money(gap)","'Đủ DS':'Thiếu DS: '+money(gap)")
$s=$s.Replace("'Đã đạt':'Gap: '+money(gap)","'Đủ DS':'Thiếu DS: '+money(gap)")

# Render receives and passes result map.
$s=$s.Replace('function Render($rows,$cur,$sales,$ex,$am,$h100){','function Render($rows,$cur,$sales,$results,$ex,$am,$h100){')
$s=$s.Replace('$h=Add-CurrentSalesSupport $h $sales','$h=Add-CurrentSalesSupport $h $sales $results')
$s=$s.Replace('(Render $data $current $currentSales $extra $asm $hc100)','(Render $data $current $currentSales $displayResult $extra $asm $hc100)')
$s=$s.Replace('$cur=@{};$sales=@{};$ex=@{};$am=@{};$h100=@{}','$cur=@{};$sales=@{};$results=@{};$ex=@{};$am=@{};$h100=@{}')
$s=$s.Replace('if($currentSales.ContainsKey($id)){$sales[$id]=$currentSales[$id]}',"if(`$currentSales.ContainsKey(`$id)){`$sales[`$id]=`$currentSales[`$id]}`r`n    if(`$displayResult.ContainsKey(`$id)){`$results[`$id]=`$displayResult[`$id]}")
$s=$s.Replace('(Render $rows $cur $sales $ex $am $h100)','(Render $rows $cur $sales $results $ex $am $h100)')

[IO.File]::WriteAllText($Path,$s,[Text.UTF8Encoding]::new($true))
Write-Host "Da patch build_daily.ps1. Backup: $backup"
Write-Host 'Chay lai RUN_UPDATE_DAILY.bat de build dashboard.'
