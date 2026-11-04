$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Source = Join-Path $Root 'source'
$Docs = Join-Path $Root 'docs'
$TemplatePath = Join-Path $Source 'template.html'
$SalesPath = Join-Path $Source 'New Microsoft Excel Worksheet.xlsx'
$DisplayPath = Join-Path $Source 'Storelist.xlsx'
$ConfigPath = Join-Path $Root 'config.json'

function Text($v) { if ($null -eq $v) { return '' }; return ([string]$v).Trim() }
function Code($v) { if ($null -eq $v) { return '' }; if ($v -is [double] -and $v -eq [math]::Floor($v)) { return ([int64]$v).ToString() }; return (Text $v) }
function Num($v) { if ($null -eq $v -or $v -eq '') { return 0.0 }; try { return [double]$v } catch { return 0.0 } }
function Slug($s) { $x=(Text $s).ToLowerInvariant(); $x=[regex]::Replace($x,'[^a-z0-9]+','-').Trim('-'); if(!$x){$x='npp'}; return $x }
function Json($o) { return ($o | ConvertTo-Json -Depth 12 -Compress).Replace('</','<\/') }
function ColMap($sheet,$row=1) { $used=$sheet.UsedRange; $cols=$used.Columns.Count; $m=@{}; for($c=1;$c -le $cols;$c++){ $h=Text $sheet.Cells.Item($row,$c).Value2; if($h){$m[$h]=$c} }; return $m }
function GetValue($arr,$r,$c) { if($null -eq $arr){return $null}; return $arr[$r,$c] }
function ReplaceBlock($html,$name,$obj,$next) { $json=Json $obj; $pattern='(?s)const '+[regex]::Escape($name)+'=.*?;const '+[regex]::Escape($next)+'='; $replacement='const '+$name+'='+$json+';const '+$next+'='; $out=[regex]::Replace($html,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($m)$replacement},1); if($out -eq $html){throw "Khong tim thay block $name trong template"}; return $out }
function Release-Com($o){ if($null-ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)} }
function MonthKey($v) {
  if($null -eq $v -or $v -eq ''){ return $null }
  if($v -is [double] -or $v -is [int]){ try { return [int]([datetime]::FromOADate([double]$v).ToString('yyyyMM')) } catch {} }
  $s=Text $v; $dt=[datetime]::MinValue
  if([datetime]::TryParse($s,[ref]$dt)){ return [int]$dt.ToString('yyyyMM') }
  if($s -match '^N-2$'){return -2}; if($s -match '^N-1$'){return -1}; if($s -match '^N$'){return 0}
  return $null
}
function AddSales($bucket,$cat,$seg,$pack,$v){
  $bucket.t += $v
  if($cat -eq 'Hair'){$bucket.hair += $v}
  if($cat -in @('Food','Foods')){$bucket.food += $v}
  if($cat -eq 'Skincare'){$bucket.skin += $v}
  if($cat -eq 'Oral'){$bucket.oral += $v}
  if($seg -eq 'Toothpaste'){$bucket.tp += $v}
  if($seg -eq 'Toothbrush'){$bucket.tb += $v}
  if($pack -match 'Sachet'){$bucket.sachet += $v}
  if($cat -eq 'Fabclean' -and $seg -eq 'Liquid'){$bucket.liquid += $v; $bucket.hc100 += $v}
  if($cat -eq 'Fabenc'){$bucket.hc100 += $v}
  if($cat -eq 'HNH' -and $seg -eq 'Dishwash'){$bucket.hc100 += $v}
  if($cat -eq 'SCL'){$bucket.scl += $v}
  if($cat -eq 'Deo'){$bucket.deo += $v}
}
function NewBucket(){ return [ordered]@{t=0.0;hair=0.0;food=0.0;skin=0.0;oral=0.0;tp=0.0;tb=0.0;sachet=0.0;liquid=0.0;hc100=0.0;scl=0.0;deo=0.0;herbal=0.0;men=0.0} }

$required=[ordered]@{'template.html'=$TemplatePath;'New Microsoft Excel Worksheet.xlsx'=$SalesPath;'Storelist.xlsx'=$DisplayPath;'config.json'=$ConfigPath}
foreach($name in $required.Keys){$p=$required[$name];if([string]::IsNullOrWhiteSpace($p)-or !(Test-Path -LiteralPath $p)){throw "Thieu file: $name | $p"}}
$Config=Get-Content $ConfigPath -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $Docs | Out-Null
$excel=$null;$wb=$null;$tb=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
  Write-Host '[1/5] Refresh Query T9-T10-T11 va Clear...'
  $wb=$excel.Workbooks.Open($SalesPath,0,$false)
  $wb.RefreshAll();$excel.CalculateUntilAsyncQueriesDone()
  while($excel.CalculationState -ne 0){Start-Sleep -Milliseconds 500}

  $stores=@{}
  function EnsureStore($code,$site,$name,$channel){
    if(!$stores.ContainsKey($code)){$stores[$code]=[ordered]@{s=$site;c=$code;n=$name;ch=$channel;proposed=(NewBucket);current=(NewBucket)}}
    $d=$stores[$code];if($site){$d.s=$site};if($name){$d.n=$name};if($channel){$d.ch=$channel};return $d
  }

  # Scan ALL sales output sheets after Power Query refresh. Do not assume worksheet 1 contains T9/T10/T11.
  $req=@('Site Name','Master Outlet Code','Outlet Name','Outlet Channel','Category','Segment','Pack Group','Year Month','KPI')
  $salesSheets=New-Object System.Collections.Generic.List[object]
  $monthSet=New-Object 'System.Collections.Generic.HashSet[int]'

  foreach($ws in @($wb.Worksheets)){
    try {
      if((Text $ws.Name) -eq 'Clear'){ continue }
      $map=ColMap $ws 1
      $ok=$true
      foreach($x in $req){ if(!$map.ContainsKey($x)){ $ok=$false; break } }
      if(!$ok){ continue }
      $u=$ws.UsedRange; $a=$u.Value2; $rows=$u.Rows.Count
      $sheetMonths=New-Object 'System.Collections.Generic.HashSet[int]'
      for($r=2;$r -le $rows;$r++){
        $mk=MonthKey(GetValue $a $r $map['Year Month'])
        if($null-ne$mk -and $mk-gt 0){ [void]$monthSet.Add($mk); [void]$sheetMonths.Add($mk) }
      }
      $salesSheets.Add([pscustomobject]@{Name=(Text $ws.Name); Map=$map; Data=$a; Rows=$rows; Months=@($sheetMonths)})
      Write-Host ("  Sales sheet: {0} | months: {1}" -f (Text $ws.Name), (@($sheetMonths|Sort-Object) -join ','))
      Release-Com $u
    } finally { Release-Com $ws }
  }

  $months=@($monthSet|Sort-Object)
  if($months.Count -lt 3){
    throw "Can it nhat 3 thang tren cac sheet sales. Tim thay: $($months -join ', '). Kiem tra Query T9/T10/T11 da Load To worksheet/table."
  }
  # Use the newest 3 months only: N-2, N-1, N.
  $N2=$months[-3]; $N1=$months[-2]; $N=$months[-1]
  Write-Host "Mapping: N-2=$N2 | N-1=$N1 | N=$N"

  foreach($ss in $salesSheets){
    $map=$ss.Map; $a=$ss.Data; $rows=$ss.Rows
    for($r=2;$r -le $rows;$r++){
      $code=Code(GetValue $a $r $map['Master Outlet Code']); if(!$code){continue}
      $mk=MonthKey(GetValue $a $r $map['Year Month']); if($mk-notin@($N2,$N1,$N)){continue}
      $d=EnsureStore $code (Text(GetValue $a $r $map['Site Name'])) (Text(GetValue $a $r $map['Outlet Name'])) (Text(GetValue $a $r $map['Outlet Channel']))
      $cat=Text(GetValue $a $r $map['Category']); $seg=Text(GetValue $a $r $map['Segment']); $pack=Text(GetValue $a $r $map['Pack Group'])
      # KPI is sales in million VND; dashboard uses VND.
      $v=(Num(GetValue $a $r $map['KPI']))*1000000
      if($mk-in@($N2,$N1)){ AddSales $d.proposed $cat $seg $pack $v }
      if($mk-in@($N1,$N)){ AddSales $d.current $cat $seg $pack $v }
    }
  }
  try {
    $clear=$wb.Worksheets.Item('Clear');$cm=ColMap $clear 1;$cu=$clear.UsedRange;$ca=$cu.Value2;$cr=$cu.Rows.Count
    if($cm.ContainsKey('Master Outlet Code')-and$cm.ContainsKey('KPI')){
      $clearMonthCol=if($cm.ContainsKey('Year Month')){$cm['Year Month']}else{0}
      for($r=2;$r -le $cr;$r++){
        $mk=if($clearMonthCol){MonthKey(GetValue $ca $r $clearMonthCol)}else{$N}
        if($mk-notin@($N2,$N1,$N)){continue}
        $code=Code(GetValue $ca $r $cm['Master Outlet Code']);if(!$code){continue}
        $d=EnsureStore $code '' '' '';$v=(Num(GetValue $ca $r $cm['KPI']))*1000000
        $sub=if($cm.ContainsKey('Sub Brand')){(Text(GetValue $ca $r $cm['Sub Brand'])).ToUpperInvariant()}else{''}
        $view=if($cm.ContainsKey('CD View')){(Text(GetValue $ca $r $cm['CD View'])).ToUpperInvariant()}else{''}
        if($mk-in@($N2,$N1)){if($sub-match'HERBAL'){$d.proposed.herbal+=$v};if($sub-match'MEN'-and$view-notmatch'SACHET'){$d.proposed.men+=$v}}
        if($mk-in@($N1,$N)){if($sub-match'HERBAL'){$d.current.herbal+=$v};if($sub-match'MEN'-and$view-notmatch'SACHET'){$d.current.men+=$v}}
      }
    }
    Release-Com $cu;Release-Com $clear
  } catch { Write-Host "Khong doc duoc sheet Clear: $($_.Exception.Message)" }

  $wb.Save();$wb.Close($false);Release-Com $wb;$wb=$null
  Write-Host '[2/5] Doc Storelist / Trung bay hien tai...'
  $tb=$excel.Workbooks.Open($DisplayPath,0,$true);$tw=$tb.Worksheets.Item(1);$tu=$tw.UsedRange;$ta=$tu.Value2;$tr=$tu.Rows.Count;$tc=$tu.Columns.Count
  $current=@{};$asm=@{}
  for($r=3;$r -le $tr;$r++){$code=Code(GetValue $ta $r 6);if(!$code){continue};$av=Text(GetValue $ta $r 1);if($av){$asm[$code]=$av};$vals=New-Object System.Collections.Generic.List[string];for($c=10;$c -le [Math]::Min(28,$tc);$c++){$v=GetValue $ta $r $c;if($null-eq$v-or$v-eq''-or$v-eq0){continue};$head=Text(GetValue $ta 2 $c);if($head-eq'Super FG'){$vals.Add('Super FG')}else{$vals.Add((Text $v))}};if($vals.Count){$current[$code]=@($vals)}}
  $tb.Close($false);Release-Com $tu;Release-Com $tw;Release-Com $tb;$tb=$null
} finally {if($wb){$wb.Close($false);Release-Com $wb};if($tb){$tb.Close($false);Release-Com $tb};if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()}

Write-Host '[3/5] Tao HTML tong va tung NPP...'
$data=New-Object System.Collections.Generic.List[object];$extra=@{};$hc100=@{}
foreach($k in $stores.Keys){$d=$stores[$k];$p=$d.proposed;$c=$d.current;$data.Add([ordered]@{s=$d.s;c=$d.c;n=$d.n;ch=$d.ch;t=[math]::Round($p.t);hair=[math]::Round($p.hair);food=[math]::Round($p.food);skin=[math]::Round($p.skin);oral=[math]::Round($p.oral);tp=[math]::Round($p.tp);tb=[math]::Round($p.tb);sachet=[math]::Round($p.sachet)});$extra[$k]=[ordered]@{liquid=[math]::Round($c.liquid);scl=[math]::Round($c.scl);deo=[math]::Round($c.deo);herbal=[math]::Round($c.herbal);men=[math]::Round($c.men)};$hc100[$k]=[math]::Round($c.hc100)}
$data=@($data|Sort-Object t -Descending);$template=Get-Content $TemplatePath -Raw -Encoding UTF8
function Render($rows,$cur,$ex,$am,$h100){$h=$template;$h=ReplaceBlock $h 'HC100' $h100 'ASM';$h=ReplaceBlock $h 'ASM' $am 'EXTRA';$h=ReplaceBlock $h 'EXTRA' $ex 'CURRENT';$h=ReplaceBlock $h 'CURRENT' $cur 'DATA';$h=ReplaceBlock $h 'DATA' $rows 'RULES';return $h}
[IO.File]::WriteAllText((Join-Path $Docs 'index.html'),(Render $data $current $extra $asm $hc100),[Text.UTF8Encoding]::new($false))
$nppRoot=Join-Path $Docs 'npp';if(Test-Path $nppRoot){Remove-Item $nppRoot -Recurse -Force};New-Item -ItemType Directory -Force $nppRoot|Out-Null
$base="https://$($Config.github_username).github.io/$($Config.repository_name)"
$links=New-Object System.Collections.Generic.List[object];$links.Add([pscustomobject]@{Loai='ASM';NPP='TONG';SoCH=$data.Count;Link="$base/"})
$sites=@($data|ForEach-Object{$_.s}|Where-Object{$_}|Sort-Object -Unique)
foreach($site in $sites){$rows=@($data|Where-Object{$_.s-eq$site});$ids=@{};$rows|ForEach-Object{$ids[$_.c]=$true};$cur=@{};$ex=@{};$am=@{};$h100=@{};foreach($id in $ids.Keys){if($current.ContainsKey($id)){$cur[$id]=$current[$id]};if($extra.ContainsKey($id)){$ex[$id]=$extra[$id]};if($asm.ContainsKey($id)){$am[$id]=$asm[$id]};if($hc100.ContainsKey($id)){$h100[$id]=$hc100[$id]}};$folder=Join-Path $nppRoot(Slug $site);New-Item -ItemType Directory -Force $folder|Out-Null;[IO.File]::WriteAllText((Join-Path $folder 'index.html'),(Render $rows $cur $ex $am $h100),[Text.UTF8Encoding]::new($false));$links.Add([pscustomobject]@{Loai='NPP';NPP=$site;SoCH=$rows.Count;Link="$base/npp/$(Slug $site)/"})}
$links|Export-Csv(Join-Path $Root 'LINK_NPP.csv') -NoTypeInformation -Encoding UTF8
Set-Content(Join-Path $Docs '.nojekyll') '' -Encoding ascii
Write-Host "Da tao $($sites.Count) link NPP va 1 link ASM."
Write-Host '[4/5] Kiem tra output...'
if(!(Test-Path(Join-Path $Docs 'index.html'))){throw 'Khong tao duoc docs\index.html'}
if((Get-ChildItem $nppRoot -Directory).Count-ne$sites.Count){throw 'So folder NPP khong khop'}
Write-Host '[5/5] Hoan tat build.'
