$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)

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
function ColMap($sheet,$row=1) { $used=$sheet.UsedRange; try { $cols=$used.Columns.Count; $m=@{}; for($c=1;$c -le $cols;$c++){ $h=Text $sheet.Cells.Item($row,$c).Value2; if($h){$m[$h]=$c} }; return $m } finally { Release-Com $used } }
function GetValue($arr,$r,$c) { if($null -eq $arr){return $null}; return $arr[$r,$c] }
function ReplaceBlock($html,$name,$obj,$next) { $json=Json $obj; $pattern='(?s)const '+[regex]::Escape($name)+'=.*?;const '+[regex]::Escape($next)+'='; $replacement='const '+$name+'='+$json+';const '+$next+'='; $out=[regex]::Replace($html,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($m)$replacement},1); if($out -eq $html){throw "Khong tim thay block $name trong template"}; return $out }
function Release-Com($o){ if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)} }
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
function NormalizeResult($v){
  $x=(Text $v).ToUpperInvariant()
  if($x -in @('ĐẠT','DAT','PASS')){ return 'Đạt' }
  if($x -in @('RỚT','ROT','FAIL')){ return 'Rớt' }
  return ''
}

$required=[ordered]@{'template.html'=$TemplatePath;'New Microsoft Excel Worksheet.xlsx'=$SalesPath;'Storelist.xlsx'=$DisplayPath;'config.json'=$ConfigPath}
foreach($name in $required.Keys){$p=$required[$name];if([string]::IsNullOrWhiteSpace($p)-or !(Test-Path -LiteralPath $p)){throw "Thieu file: $name | $p"}}
$Config=Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $Docs | Out-Null

$excel=$null;$wb=$null;$tb=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
  Write-Host '[1/5] Refresh Query T9-T10-T11, Clear va Ket qua trung bay...'
  $wb=$excel.Workbooks.Open($SalesPath,0,$false)
  $wb.RefreshAll();$excel.CalculateUntilAsyncQueriesDone()
  while($excel.CalculationState -ne 0){Start-Sleep -Milliseconds 500}

  $stores=@{};$displayResult=@{}
  function EnsureStore($code,$site,$name,$channel){
    if(!$stores.ContainsKey($code)){$stores[$code]=[ordered]@{s=$site;c=$code;n=$name;ch=$channel;proposed=(NewBucket);current=(NewBucket)}}
    $d=$stores[$code];if($site){$d.s=$site};if($name){$d.n=$name};if($channel){$d.ch=$channel};return $d
  }

  $req=@('Site Name','Master Outlet Code','Outlet Name','Outlet Channel','Category','Segment','Pack Group','Year Month','KPI')
  $salesSheets=New-Object System.Collections.Generic.List[object]
  $monthSet=New-Object 'System.Collections.Generic.HashSet[int]'
  foreach($ws in @($wb.Worksheets)){
    try {
      if((Text $ws.Name) -in @('Clear','Kết quả trưng bày')){continue}
      $map=ColMap $ws 1;$ok=$true
      foreach($x in $req){if(!$map.ContainsKey($x)){$ok=$false;break}}
      if(!$ok){continue}
      $u=$ws.UsedRange
      try {
        $a=$u.Value2;$rows=$u.Rows.Count;$sheetMonths=New-Object 'System.Collections.Generic.HashSet[int]'
        for($r=2;$r -le $rows;$r++){$mk=MonthKey(GetValue $a $r $map['Year Month']);if($null-ne$mk -and $mk-gt 0){[void]$monthSet.Add($mk);[void]$sheetMonths.Add($mk)}}
        $salesSheets.Add([pscustomobject]@{Name=(Text $ws.Name);Map=$map;Data=$a;Rows=$rows;Months=@($sheetMonths)})
        Write-Host ("  Sales sheet: {0} | months: {1}" -f (Text $ws.Name),(@($sheetMonths|Sort-Object)-join ','))
      } finally {Release-Com $u}
    } finally {Release-Com $ws}
  }

  $months=@($monthSet|Sort-Object)
  if($months.Count -lt 3){throw "Can it nhat 3 thang sales. Tim thay: $($months -join ', ')"}
  $N2=$months[-3];$N1=$months[-2];$N=$months[-1]
  Write-Host "Mapping: N-2=$N2 | N-1=$N1 | N=$N"
  foreach($ss in $salesSheets){
    $map=$ss.Map;$a=$ss.Data;$rows=$ss.Rows
    for($r=2;$r -le $rows;$r++){
      $code=Code(GetValue $a $r $map['Master Outlet Code']);if(!$code){continue}
      $mk=MonthKey(GetValue $a $r $map['Year Month']);if($mk -notin @($N2,$N1,$N)){continue}
      $d=EnsureStore $code (Text(GetValue $a $r $map['Site Name'])) (Text(GetValue $a $r $map['Outlet Name'])) (Text(GetValue $a $r $map['Outlet Channel']))
      $cat=Text(GetValue $a $r $map['Category']);$seg=Text(GetValue $a $r $map['Segment']);$pack=Text(GetValue $a $r $map['Pack Group']);$v=(Num(GetValue $a $r $map['KPI']))*1000000
      if($mk -in @($N2,$N1)){AddSales $d.proposed $cat $seg $pack $v}
      if($mk -in @($N1,$N)){AddSales $d.current $cat $seg $pack $v}
    }
  }

  try {
    $clear=$wb.Worksheets.Item('Clear');$cm=ColMap $clear 1;$cu=$clear.UsedRange
    try {
      $ca=$cu.Value2;$cr=$cu.Rows.Count
      if($cm.ContainsKey('Master Outlet Code') -and $cm.ContainsKey('KPI')){
        $clearMonthCol=if($cm.ContainsKey('Year Month')){$cm['Year Month']}else{0}
        for($r=2;$r -le $cr;$r++){
          $mk=if($clearMonthCol){MonthKey(GetValue $ca $r $clearMonthCol)}else{$N};if($mk -notin @($N2,$N1,$N)){continue}
          $code=Code(GetValue $ca $r $cm['Master Outlet Code']);if(!$code){continue}
          $d=EnsureStore $code '' '' '';$v=(Num(GetValue $ca $r $cm['KPI']))*1000000
          $sub=if($cm.ContainsKey('Sub Brand')){(Text(GetValue $ca $r $cm['Sub Brand'])).ToUpperInvariant()}else{''}
          $view=if($cm.ContainsKey('CD View')){(Text(GetValue $ca $r $cm['CD View'])).ToUpperInvariant()}else{''}
          if($mk -in @($N2,$N1)){if($sub-match'HERBAL'){$d.proposed.herbal+=$v};if($sub-match'MEN' -and $view-notmatch'SACHET'){$d.proposed.men+=$v}}
          if($mk -in @($N1,$N)){if($sub-match'HERBAL'){$d.current.herbal+=$v};if($sub-match'MEN' -and $view-notmatch'SACHET'){$d.current.men+=$v}}
        }
      }
    } finally {Release-Com $cu;Release-Com $clear}
  } catch {Write-Warning "Khong doc duoc sheet Clear: $($_.Exception.Message)"}

  Write-Host '  Doc sheet Ket qua trung bay...'
  try {
    $kr=$wb.Worksheets.Item('Kết quả trưng bày');$ku=$kr.UsedRange
    try {
      $ka=$ku.Value2;$kRows=$ku.Rows.Count;$kCols=$ku.Columns.Count;$headerRow=0;$km=@{}
      for($hr=1;$hr -le [Math]::Min(10,$kRows);$hr++){
        $tmp=@{};for($c=1;$c -le $kCols;$c++){$h=Text(GetValue $ka $hr $c);if($h){$tmp[$h]=$c}}
        if($tmp.ContainsKey('Master Outlet Code') -or $tmp.ContainsKey('Outlet Code') -or $tmp.ContainsKey('Code')){$headerRow=$hr;$km=$tmp;break}
      }
      if($headerRow -eq 0){throw 'Khong tim thay header sheet Ket qua trung bay'}
      $codeCol=0;foreach($name in @('Master Outlet Code','Outlet Code','Code')){if($km.ContainsKey($name)){$codeCol=$km[$name];break}}
      if(!$codeCol){throw 'Khong tim thay cot Master Outlet Code'}
      $setCol=0;foreach($name in @('Set','Set trưng bày','Tên Set','Display Set')){if($km.ContainsKey($name)){$setCol=$km[$name];break}}
      $resultCol=0;foreach($name in @('Kết quả','Kết quả trưng bày','Đạt/Rớt','Result')){if($km.ContainsKey($name)){$resultCol=$km[$name];break}}
      if($setCol -and $resultCol){
        Write-Host '  Format: dang doc (CH + Set + Ket qua).'
        for($r=$headerRow+1;$r -le $kRows;$r++){
          $code=Code(GetValue $ka $r $codeCol);$setName=Text(GetValue $ka $r $setCol);$result=NormalizeResult(GetValue $ka $r $resultCol)
          if(!$code -or !$setName -or !$result){continue};if(!$displayResult.ContainsKey($code)){$displayResult[$code]=@{}};$displayResult[$code][$setName]=$result
        }
      } else {
        Write-Host '  Format: dang ngang (moi Set la mot cot).'
        $ignore=@('Master Outlet Code','Outlet Code','Code','Outlet Name','Site Name','ASM','NVTT','NPP','Channel')
        for($r=$headerRow+1;$r -le $kRows;$r++){
          $code=Code(GetValue $ka $r $codeCol);if(!$code){continue};if(!$displayResult.ContainsKey($code)){$displayResult[$code]=@{}}
          foreach($columnName in $km.Keys){if($columnName -in $ignore){continue};$result=NormalizeResult(GetValue $ka $r $km[$columnName]);if($result){$displayResult[$code][$columnName]=$result}}
        }
      }
      Write-Host ("  Da doc ket qua trung bay cho {0} CH" -f $displayResult.Count)
    } finally {Release-Com $ku;Release-Com $kr}
  } catch {Write-Warning "Khong doc duoc sheet Ket qua trung bay: $($_.Exception.Message)"}

  $wb.Save();$wb.Close($false);Release-Com $wb;$wb=$null

Write-Host '[2/5] Doc Storelist / Trung bay hien tai...'

$tb = $excel.Workbooks.Open($DisplayPath, 0, $true)
$tw = $tb.Worksheets.Item(1)
$tu = $tw.UsedRange

$ta = $tu.Value2
$tr = $tu.Rows.Count
$tc = $tu.Columns.Count

$current = @{}
$asm = @{}

for ($r = 3; $r -le $tr; $r++) {

    $code = Code (GetValue $ta $r 6)

    if (!$code) {
        continue
    }

    $av = Text (GetValue $ta $r 1)

    if ($av) {
        $asm[$code] = $av
    }

    $vals = New-Object System.Collections.Generic.List[string]

    for ($c = 10; $c -le [Math]::Min(28, $tc); $c++) {

        $v = GetValue $ta $r $c

        if ($null -eq $v -or $v -eq '' -or $v -eq 0) {
            continue
        }

        $setName = Text $v

        if ($setName) {
            $vals.Add($setName)
        }
    }

    if ($vals.Count -gt 0) {
        $current[$code] = @($vals)
    }
}

$tb.Close($false)
Release-Com $tu
Release-Com $tw
Release-Com $tb
$tb = $null

} finally {

    if ($wb) {
        $wb.Close($false)
        Release-Com $wb
    }

    if ($tb) {
        $tb.Close($false)
        Release-Com $tb
    }

    if ($excel) {
        $excel.Quit()
        Release-Com $excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host '[3/5] Tao HTML tong va tung NPP...'
$data=New-Object System.Collections.Generic.List[object];$currentSales=@{};$extra=@{};$hc100=@{}
foreach($k in $stores.Keys){
  $d=$stores[$k];$p=$d.proposed;$c=$d.current
  $data.Add([ordered]@{s=$d.s;c=$d.c;n=$d.n;ch=$d.ch;t=[math]::Round($p.t);hair=[math]::Round($p.hair);food=[math]::Round($p.food);skin=[math]::Round($p.skin);oral=[math]::Round($p.oral);tp=[math]::Round($p.tp);tb=[math]::Round($p.tb);sachet=[math]::Round($p.sachet)})
  $currentSales[$k]=[ordered]@{t=[math]::Round($c.t);hair=[math]::Round($c.hair);food=[math]::Round($c.food);skin=[math]::Round($c.skin);oral=[math]::Round($c.oral);tp=[math]::Round($c.tp);tb=[math]::Round($c.tb);sachet=[math]::Round($c.sachet);liquid=[math]::Round($c.liquid);scl=[math]::Round($c.scl);deo=[math]::Round($c.deo);herbal=[math]::Round($c.herbal);men=[math]::Round($c.men);hc100=[math]::Round($c.hc100)}
  $extra[$k]=[ordered]@{liquid=[math]::Round($c.liquid);scl=[math]::Round($c.scl);deo=[math]::Round($c.deo);herbal=[math]::Round($c.herbal);men=[math]::Round($c.men)}
  $hc100[$k]=[math]::Round($c.hc100)
}
$data=@($data|Sort-Object t -Descending);$template=Get-Content $TemplatePath -Raw -Encoding UTF8

function Add-CurrentSalesSupport($html,$sales){
  $json=Json $sales
  if($html -match '(?s)const CURRENT_SALES=.*?;const RULES='){$html=[regex]::Replace($html,'(?s)const CURRENT_SALES=.*?;const RULES=',('const CURRENT_SALES='+$json+';const RULES='),1)}
  elseif($html -match 'const RULES='){$html=$html.Replace('const RULES=',('const CURRENT_SALES='+$json+';const RULES='))}
  else{throw 'Template khong co const RULES= de chen CURRENT_SALES'}
  if($html -notmatch 'function currentMetric\('){
    $helper=@"
const ZERO_CURRENT={t:0,hair:0,food:0,skin:0,oral:0,tp:0,tb:0,sachet:0,liquid:0,scl:0,deo:0,herbal:0,men:0,hc100:0};
function currentMetric(o,type){const x=CURRENT_SALES[String(o.c)]||ZERO_CURRENT;return metric(x,type);}
function normalizeDisplayText(value){return String(value==null?'':value).replace(/_/g,' ').replace(/\s+/g,' ').trim().toLowerCase();}
function compactDisplayText(value){return normalizeDisplayText(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/[^a-z0-9]/g,'');}
function findDisplayResultValue(o,label){const store=DISPLAY_RESULT[String(o.c)]||{};if(Object.prototype.hasOwnProperty.call(store,label))return store[label];const wanted=compactDisplayText(label);const key=Object.keys(store).find(function(k){return compactDisplayText(k)===wanted;});return key==null?null:store[key];}
function currentSetResult(o,label){const value=findDisplayResultValue(o,label);if(value==null||String(value).trim()==='')return 'na';const x=compactDisplayText(value);if(x==='dat'||x==='pass')return 'pass';if(x==='rot'||x==='fail')return 'fail';return 'na';}
function currentSetResultHTML(o,label){const x=currentSetResult(o,label);if(x==='pass')return '<span class="tb-result tb-pass">Kết quả trưng bày: Đạt</span>';if(x==='fail')return '<span class="tb-result tb-fail">Kết quả trưng bày: Rớt</span>';return '<span class="tb-result tb-na">Chưa đánh giá</span>';}
function currentOverallResult(o){const sets=CURRENT[String(o.c)]||[];const evaluated=sets.map(function(x){return currentSetResult(o,x);}).filter(function(x){return x!=='na';});if(!evaluated.length)return 'na';return evaluated.every(function(x){return x==='pass';})?'pass':'fail';}
"@
    if($html -notmatch 'function currentBankerHTML'){throw 'Khong tim thay currentBankerHTML trong template'}
    $html=$html.Replace('function currentBankerHTML',$helper+'function currentBankerHTML')
  }
  $start=$html.IndexOf('function currentBankerHTML');if($start -lt 0){throw 'Khong tim thay currentBankerHTML trong template'};$finish=$html.IndexOf('function ',$start+30);if($finish -lt 0){$finish=$html.Length}
  $block=$html.Substring($start,$finish-$start).Replace('metric(o,r.type)','currentMetric(o,r.type)');$html=$html.Substring(0,$start)+$block+$html.Substring($finish)
  $old='${esc(x.replaceAll(''_'','' ''))}${currentBankerHTML(o,x)}';$new='${esc(x.replaceAll(''_'','' ''))} ${currentSetResultHTML(o,x)}${currentBankerHTML(o,x)}'
  if($html.Contains($old)){$html=$html.Replace($old,$new)}elseif($html -notmatch 'currentSetResultHTML\(o,x\)'){Write-Warning 'Khong tim thay vi tri chen badge Dat/Rot trong template.'}
  if($html -notmatch '\.tb-result\{'){$css='.tb-result{display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:900;vertical-align:middle}.tb-pass{background:#D9FBEA;color:#087A4F}.tb-fail{background:#FFE1E1;color:#B42318}.tb-na{background:#EDF2F4;color:#60717B}';if($html -match '</style>'){$html=$html.Replace('</style>',$css+'</style>')}else{Write-Warning 'Template khong co </style>; CSS badge chua duoc chen.'}}
  return $html
}
function Render($rows,$cur,$sales,$ex,$am,$h100,$results){
  $h=$template;$h=ReplaceBlock $h 'HC100' $h100 'ASM';$h=ReplaceBlock $h 'ASM' $am 'EXTRA';$h=ReplaceBlock $h 'EXTRA' $ex 'CURRENT';$h=ReplaceBlock $h 'CURRENT' $cur 'DATA';$h=ReplaceBlock $h 'DATA' $rows 'RULES'
  $resultJson=Json $results
  if($h -match '(?s)const DISPLAY_RESULT=.*?;'){$h=[regex]::Replace($h,'(?s)const DISPLAY_RESULT=.*?;',('const DISPLAY_RESULT='+$resultJson+';'),1)}
  elseif($h -match 'const RULES='){$h=$h.Replace('const RULES=',('const DISPLAY_RESULT='+$resultJson+';const RULES='))}
  else{throw 'Khong tim thay const RULES= de chen DISPLAY_RESULT'}
  return (Add-CurrentSalesSupport $h $sales)
}

[IO.File]::WriteAllText((Join-Path $Docs 'index.html'),(Render $data $current $currentSales $extra $asm $hc100 $displayResult),(New-Object Text.UTF8Encoding($false)))
$nppRoot=Join-Path $Docs 'npp';if(Test-Path $nppRoot){Remove-Item $nppRoot -Recurse -Force};New-Item -ItemType Directory -Force $nppRoot|Out-Null
$base="https://$($Config.github_username).github.io/$($Config.repository_name)"
$links=New-Object System.Collections.Generic.List[object];$links.Add([pscustomobject]@{Loai='ASM';NPP='TONG';SoCH=$data.Count;Link="$base/"})
$sites=@($data|ForEach-Object{$_.s}|Where-Object{$_}|Sort-Object -Unique)
foreach($site in $sites){
  $rows=@($data|Where-Object{$_.s -eq $site});$ids=@{};$rows|ForEach-Object{$ids[$_.c]=$true};$cur=@{};$sales=@{};$ex=@{};$am=@{};$h100=@{};$results=@{}
  foreach($id in $ids.Keys){if($displayResult.ContainsKey($id)){$results[$id]=$displayResult[$id]};if($current.ContainsKey($id)){$cur[$id]=$current[$id]};if($currentSales.ContainsKey($id)){$sales[$id]=$currentSales[$id]};if($extra.ContainsKey($id)){$ex[$id]=$extra[$id]};if($asm.ContainsKey($id)){$am[$id]=$asm[$id]};if($hc100.ContainsKey($id)){$h100[$id]=$hc100[$id]}}
  $folder=Join-Path $nppRoot (Slug $site);New-Item -ItemType Directory -Force $folder|Out-Null
  [IO.File]::WriteAllText((Join-Path $folder 'index.html'),(Render $rows $cur $sales $ex $am $h100 $results),(New-Object Text.UTF8Encoding($false)))
  $links.Add([pscustomobject]@{Loai='NPP';NPP=$site;SoCH=$rows.Count;Link="$base/npp/$(Slug $site)/"})
}
$linkPath=Join-Path $Root 'LINK_NPP.csv';$linkTemp=Join-Path $Root 'LINK_NPP.__new__.csv';if(Test-Path $linkTemp){Remove-Item $linkTemp -Force -ErrorAction SilentlyContinue}
$links|Export-Csv $linkTemp -NoTypeInformation -Encoding UTF8
try{Move-Item $linkTemp $linkPath -Force;Write-Host 'Da cap nhat LINK_NPP.csv.'}catch{$fallback=Join-Path $Root 'LINK_NPP_NEW.csv';Move-Item $linkTemp $fallback -Force;Write-Warning 'LINK_NPP.csv dang bi khoa. Da ghi LINK_NPP_NEW.csv.'}
Set-Content (Join-Path $Docs '.nojekyll') '' -Encoding ascii
Write-Host "Da tao $($sites.Count) link NPP va 1 link ASM."
Write-Host '[4/5] Kiem tra output...'
if(!(Test-Path (Join-Path $Docs 'index.html'))){throw 'Khong tao duoc docs\index.html'}
if((Get-ChildItem $nppRoot -Directory).Count -ne $sites.Count){throw 'So folder NPP khong khop'}
Write-Host '[5/5] Hoan tat build.'
