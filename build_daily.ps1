$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Source = Join-Path $Root 'source'
$Docs = Join-Path $Root 'docs'
$TemplatePath = Join-Path $Source 'template.html'
$SalesPath = Join-Path $Source 'New Microsoft Excel Worksheet.xlsx'
$DisplayPath = Join-Path $Source 'Trung bay.xlsx'
$Config = Get-Content (Join-Path $Root 'config.json') -Raw | ConvertFrom-Json

function Text($v) { if ($null -eq $v) { return '' }; return ([string]$v).Trim() }
function Code($v) { if ($null -eq $v) { return '' }; if ($v -is [double] -and $v -eq [math]::Floor($v)) { return ([int64]$v).ToString() }; return (Text $v) }
function Num($v) { if ($null -eq $v -or $v -eq '') { return 0.0 }; try { return [double]$v } catch { return 0.0 } }
function Slug($s) { $x=(Text $s).ToLowerInvariant(); $x=[regex]::Replace($x,'[^a-z0-9]+','-').Trim('-'); if(!$x){$x='npp'}; return $x }
function Json($o) { return ($o | ConvertTo-Json -Depth 12 -Compress).Replace('</','<\/') }
function ColMap($sheet,$row=1) {
  $used=$sheet.UsedRange; $cols=$used.Columns.Count; $m=@{}
  for($c=1;$c -le $cols;$c++){ $h=Text $sheet.Cells.Item($row,$c).Value2; if($h){$m[$h]=$c} }
  return $m
}
function GetValue($arr,$r,$c) { if($null -eq $arr){return $null}; return $arr[$r,$c] }
function ReplaceBlock($html,$name,$obj,$next) {
  $json=Json $obj
  $pattern='(?s)const '+[regex]::Escape($name)+'=.*?;const '+[regex]::Escape($next)+'='
  $replacement='const '+$name+'='+$json+';const '+$next+'='
  $out=[regex]::Replace($html,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($m)$replacement},1)
  if($out -eq $html){throw "Khong tim thay block $name trong template"}
  return $out
}
function Release-Com($o){ if($null-ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)} }

if(!(Test-Path $TemplatePath)){throw 'Thieu source\template.html'}
if(!(Test-Path $SalesPath)){throw 'Thieu source\New Microsoft Excel Worksheet.xlsx'}
if(!(Test-Path $DisplayPath)){throw 'Thieu source\Trung bay.xlsx'}
New-Item -ItemType Directory -Force -Path $Docs | Out-Null

$excel=$null;$wb=$null;$tb=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false
  Write-Host '[1/5] Refresh workbook doanh so...'
  $wb=$excel.Workbooks.Open($SalesPath,0,$false)
  try { $wb.RefreshAll(); $excel.CalculateUntilAsyncQueriesDone(); Start-Sleep -Seconds 2 } catch {}

  $stores=@{}
  function EnsureStore($code,$site,$name,$channel){
    if(!$stores.ContainsKey($code)){
      $stores[$code]=[ordered]@{s=$site;c=$code;n=$name;ch=$channel;t=0.0;hair=0.0;food=0.0;skin=0.0;oral=0.0;tp=0.0;tb=0.0;sachet=0.0;liquid=0.0;hc100=0.0;scl=0.0;deo=0.0;herbal=0.0;men=0.0}
    }
    $d=$stores[$code]; if($site){$d.s=$site};if($name){$d.n=$name};if($channel){$d.ch=$channel};return $d
  }

  $ws=$wb.Worksheets.Item(1); $map=ColMap $ws 1; $u=$ws.UsedRange; $a=$u.Value2; $rows=$u.Rows.Count
  $req=@('Site Name','Master Outlet Code','Outlet Name','Outlet Channel','Category','Segment','Pack Group','KPI')
  foreach($x in $req){if(!$map.ContainsKey($x)){throw ("Sales sheet missing column: " + $x)}}
  for($r=2;$r -le $rows;$r++){
    $code=Code (GetValue $a $r $map['Master Outlet Code']); if(!$code){continue}
    $d=EnsureStore $code (Text (GetValue $a $r $map['Site Name'])) (Text (GetValue $a $r $map['Outlet Name'])) (Text (GetValue $a $r $map['Outlet Channel']))
    $cat=Text (GetValue $a $r $map['Category']);$seg=Text (GetValue $a $r $map['Segment']);$pack=Text (GetValue $a $r $map['Pack Group']);$v=(Num (GetValue $a $r $map['KPI']))*1000000
    $d.t+=$v
    if($cat -eq 'Hair'){$d.hair+=$v};if($cat -in @('Food','Foods')){$d.food+=$v};if($cat -eq 'Skincare'){$d.skin+=$v};if($cat -eq 'Oral'){$d.oral+=$v};if($seg -eq 'Toothpaste'){$d.tp+=$v};if($seg -eq 'Toothbrush'){$d.tb+=$v};if($pack -match 'Sachet'){$d.sachet+=$v}
    if($cat -eq 'Fabclean' -and $seg -eq 'Liquid'){$d.liquid+=$v;$d.hc100+=$v}
    if($cat -eq 'Fabenc'){$d.hc100+=$v}
    if($cat -eq 'HNH' -and $seg -eq 'Dishwash'){$d.hc100+=$v}
    if($cat -eq 'SCL'){$d.scl+=$v};if($cat -eq 'Deo'){$d.deo+=$v}
  }
  Release-Com $u;Release-Com $ws

  try {
    $clear=$wb.Worksheets.Item('Clear');$cm=ColMap $clear 1;$cu=$clear.UsedRange;$ca=$cu.Value2;$cr=$cu.Rows.Count
    if($cm.ContainsKey('Master Outlet Code') -and $cm.ContainsKey('KPI')){
      $latest=-1;$monthCol=if($cm.ContainsKey('Year Month')){$cm['Year Month']}else{0}
      if($monthCol){for($r=2;$r -le $cr;$r++){$mv=Num (GetValue $ca $r $monthCol);if($mv -gt $latest){$latest=$mv}}}
      for($r=2;$r -le $cr;$r++){
        if($monthCol -and (Num (GetValue $ca $r $monthCol)) -ne $latest){continue}
        $code=Code (GetValue $ca $r $cm['Master Outlet Code']);if(!$code){continue};$d=EnsureStore $code '' '' '';$v=(Num (GetValue $ca $r $cm['KPI']))*1000000
        $sub=if($cm.ContainsKey('Sub Brand')){(Text (GetValue $ca $r $cm['Sub Brand'])).ToUpperInvariant()}else{''};$view=if($cm.ContainsKey('CD View')){(Text (GetValue $ca $r $cm['CD View'])).ToUpperInvariant()}else{''}
        if($sub -match 'HERBAL'){$d.herbal+=$v};if($sub -match 'MEN' -and $view -notmatch 'SACHET'){$d.men+=$v}
      }
    }
    Release-Com $cu;Release-Com $clear
  } catch { Write-Host 'Khong co sheet Clear, bo qua Men/Herbal.' }
  $wb.Close($false);Release-Com $wb;$wb=$null

  Write-Host '[2/5] Doc Trung bay hien tai...'
  $tb=$excel.Workbooks.Open($DisplayPath,0,$true);$tw=$tb.Worksheets.Item(1);$tu=$tw.UsedRange;$ta=$tu.Value2;$tr=$tu.Rows.Count;$tc=$tu.Columns.Count
  $current=@{};$asm=@{}
  for($r=3;$r -le $tr;$r++){
    $code=Code (GetValue $ta $r 6);if(!$code){continue};$av=Text (GetValue $ta $r 1);if($av){$asm[$code]=$av};$vals=New-Object System.Collections.Generic.List[string]
    for($c=10;$c -le [Math]::Min(28,$tc);$c++){
      $v=GetValue $ta $r $c;if($null-eq $v -or $v -eq '' -or $v -eq 0){continue};$head=Text (GetValue $ta 2 $c);if($head -eq 'Super FG'){$vals.Add('Super FG')}else{$vals.Add((Text $v))}
    }
    if($vals.Count){$current[$code]=@($vals)}
  }
  $tb.Close($false);Release-Com $tu;Release-Com $tw;Release-Com $tb;$tb=$null
} finally {
  if($wb){$wb.Close($false);Release-Com $wb};if($tb){$tb.Close($false);Release-Com $tb};if($excel){$excel.Quit();Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}

Write-Host '[3/5] Tao HTML tong va tung NPP...'
$data=New-Object System.Collections.Generic.List[object];$extra=@{};$hc100=@{}
foreach($k in $stores.Keys){$d=$stores[$k];$data.Add([ordered]@{s=$d.s;c=$d.c;n=$d.n;ch=$d.ch;t=[math]::Round($d.t);hair=[math]::Round($d.hair);food=[math]::Round($d.food);skin=[math]::Round($d.skin);oral=[math]::Round($d.oral);tp=[math]::Round($d.tp);tb=[math]::Round($d.tb);sachet=[math]::Round($d.sachet)});$extra[$k]=[ordered]@{liquid=[math]::Round($d.liquid);scl=[math]::Round($d.scl);deo=[math]::Round($d.deo);herbal=[math]::Round($d.herbal);men=[math]::Round($d.men)};$hc100[$k]=[math]::Round($d.hc100)}
$data=@($data|Sort-Object t -Descending);$template=Get-Content $TemplatePath -Raw -Encoding UTF8
function Render($rows,$cur,$ex,$am,$h100){$h=$template;$h=ReplaceBlock $h 'HC100' $h100 'ASM';$h=ReplaceBlock $h 'ASM' $am 'EXTRA';$h=ReplaceBlock $h 'EXTRA' $ex 'CURRENT';$h=ReplaceBlock $h 'CURRENT' $cur 'DATA';$h=ReplaceBlock $h 'DATA' $rows 'RULES';return $h}
[IO.File]::WriteAllText((Join-Path $Docs 'index.html'),(Render $data $current $extra $asm $hc100),[Text.UTF8Encoding]::new($false))
$nppRoot=Join-Path $Docs 'npp';if(Test-Path $nppRoot){Remove-Item $nppRoot -Recurse -Force};New-Item -ItemType Directory -Force $nppRoot|Out-Null
$base="https://$($Config.github_username).github.io/$($Config.repository_name)"
$links=New-Object System.Collections.Generic.List[object];$links.Add([pscustomobject]@{Loai='ASM';NPP='TONG';SoCH=$data.Count;Link="$base/"})
$sites=@($data|ForEach-Object{$_.s}|Where-Object{$_}|Sort-Object -Unique)
foreach($site in $sites){$rows=@($data|Where-Object{$_.s -eq $site});$ids=@{};$rows|ForEach-Object{$ids[$_.c]=$true};$cur=@{};$ex=@{};$am=@{};$h100=@{};foreach($id in $ids.Keys){if($current.ContainsKey($id)){$cur[$id]=$current[$id]};if($extra.ContainsKey($id)){$ex[$id]=$extra[$id]};if($asm.ContainsKey($id)){$am[$id]=$asm[$id]};if($hc100.ContainsKey($id)){$h100[$id]=$hc100[$id]}};$folder=Join-Path $nppRoot (Slug $site);New-Item -ItemType Directory -Force $folder|Out-Null;[IO.File]::WriteAllText((Join-Path $folder 'index.html'),(Render $rows $cur $ex $am $h100),[Text.UTF8Encoding]::new($false));$links.Add([pscustomobject]@{Loai='NPP';NPP=$site;SoCH=$rows.Count;Link="$base/npp/$(Slug $site)/"})}
$links|Export-Csv (Join-Path $Root 'LINK_NPP.csv') -NoTypeInformation -Encoding UTF8
Set-Content (Join-Path $Docs '.nojekyll') '' -Encoding ascii
Write-Host "Da tao $($sites.Count) link NPP va 1 link ASM."

Write-Host '[4/5] Kiem tra output...'
if(!(Test-Path (Join-Path $Docs 'index.html'))){throw 'Khong tao duoc docs\index.html'}
if((Get-ChildItem $nppRoot -Directory).Count -ne $sites.Count){throw 'So folder NPP khong khop'}

Write-Host '[5/5] Hoan tat build.'
