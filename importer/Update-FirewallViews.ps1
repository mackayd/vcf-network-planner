#requires -Version 7.0
<#
.SYNOPSIS
Rebuilds the engineer worklist and native blueprint diagrams in a new saved copy.
.DESCRIPTION
The input workbook is never saved. Original input and calculation sheets retain
their formulas, inputs and protection. Blueprint view is regenerated from these
inputs using BlueprintDiagram.ps1. Generated views reset to All component/network selectors,
Both endpoint sides, and Candidate, Review and Ready work queues. Only a separate owned Excel process
is used. A watchdog bounds the run and can terminate only that owned process.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][string]$OutputWorkbook,
    [string]$ViewTemplatePath = (Join-Path $PSScriptRoot 'EngineerViews.xlsx'),
    [ValidateRange(60,1800)][int]$TimeoutSeconds = 600,
    [switch]$KeepDiagnosticCopy,
    [switch]$Worker,
    [string]$StatePath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$viewNames = @('Engineer rules','Engineer guide')
$ownedViewNames = $viewNames
$marker = 'VCF_ENGINEER_VIEWS_V1'
$tableCacheNames=@('EV_Workqueue','EV_Domain','EV_Cluster','EV_Feature')
$legacyPivotCacheNames=@('EV_PivotDomain','EV_PivotCluster')

function Get-SharedHash([string]$Path) {
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try { return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
    finally { $stream.Dispose() }
}
function Release-Com($Object) {
    if($null-ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)){
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}
function Quote-PowerShell([string]$Value) { return "'"+$Value.Replace("'","''")+"'" }
function Stop-OwnedExcel($State) {
    if(-not $State -or -not $State.ownedPid -or -not $State.ownedStartTicks){return}
    $p=Get-Process -Id ([int]$State.ownedPid) -ErrorAction SilentlyContinue
    if($p -and $p.ProcessName -ieq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq [long]$State.ownedStartTicks){
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
}

$source=(Get-Item -LiteralPath $WorkbookPath).FullName
$scaffold=(Get-Item -LiteralPath $ViewTemplatePath).FullName
$output=[IO.Path]::GetFullPath($OutputWorkbook)
$reportPath=[IO.Path]::ChangeExtension($output,'views-report.json')
if([IO.Path]::GetExtension($source) -ine '.xlsx' -or [IO.Path]::GetExtension($output) -ine '.xlsx'){
    throw 'Source and output must be .xlsx files.'
}
if($output -ieq $source -or $output -ieq $scaffold -or (Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $reportPath)){
    throw 'Choose a distinct, new output filename; existing workbooks and reports are never overwritten.'
}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output))

if(-not $Worker){
    $runDir=Join-Path ([IO.Path]::GetDirectoryName($output)) ('.engineer-views-'+[guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($runDir)
    $stateFile=Join-Path $runDir 'state.json'
    $stdout=Join-Path $runDir 'stdout.txt';$stderr=Join-Path $runDir 'stderr.txt'
    $command='& '+(Quote-PowerShell $PSCommandPath)+' -Worker -WorkbookPath '+(Quote-PowerShell $source)+' -OutputWorkbook '+(Quote-PowerShell $output)+' -ViewTemplatePath '+(Quote-PowerShell $scaffold)+' -StatePath '+(Quote-PowerShell $stateFile)
    if($KeepDiagnosticCopy){$command+=' -KeepDiagnosticCopy'}
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $process=$null
    try {
        $process=Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList @('-STA','-NoProfile','-NonInteractive','-OutputFormat','Text','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        if(-not $process.WaitForExit($TimeoutSeconds*1000)){
            if(Test-Path -LiteralPath $stateFile){Stop-OwnedExcel (Get-Content -LiteralPath $stateFile -Raw|ConvertFrom-Json)}
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            throw "Engineer view refresh exceeded $TimeoutSeconds seconds. The source workbook was not saved."
        }
        $process.Refresh()
        if($process.ExitCode -ne 0){
            $detail=if(Test-Path -LiteralPath $stderr){[IO.File]::ReadAllText($stderr)}else{''}
            throw ('Engineer view refresh failed. '+$detail.Trim())
        }
        if(-not (Test-Path -LiteralPath $output) -or -not (Test-Path -LiteralPath $reportPath)){throw 'Refresh finished without verified output and report.'}
        [IO.File]::ReadAllText($reportPath)
    }
    finally {
        if(Test-Path -LiteralPath $stateFile){
            $state=Get-Content -LiteralPath $stateFile -Raw|ConvertFrom-Json
            Stop-OwnedExcel $state
            if($state.temporaryPath -and (Test-Path -LiteralPath $state.temporaryPath)){
                $tmpFull=[IO.Path]::GetFullPath([string]$state.temporaryPath)
                if([IO.Path]::GetDirectoryName($tmpFull) -ieq [IO.Path]::GetDirectoryName($output) -and [IO.Path]::GetFileName($tmpFull).StartsWith('.engineer-view-')){Remove-Item -LiteralPath $tmpFull -Force -ErrorAction SilentlyContinue}
            }
        }
        # Remove only this explicitly created scratch directory after resolving it.
        $resolved=[IO.Path]::GetFullPath($runDir)
        if([IO.Path]::GetDirectoryName($resolved) -ieq [IO.Path]::GetDirectoryName($output) -and [IO.Path]::GetFileName($resolved).StartsWith('.engineer-views-')){Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue}
    }
    return
}

if(-not $StatePath){throw 'Worker state path is required.'}
Add-Type -TypeDefinition @'
using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface EngineerOleFilter {
 [PreserveSig] int HandleInComingCall(int type,IntPtr task,int tick,IntPtr info);
 [PreserveSig] int RetryRejectedCall(IntPtr task,int tick,int reject);
 [PreserveSig] int MessagePending(IntPtr task,int tick,int type);
}
public sealed class EngineerMessageFilter : EngineerOleFilter {
 [DllImport("ole32.dll")] static extern int CoRegisterMessageFilter(EngineerOleFilter filter,out EngineerOleFilter old);
 static EngineerOleFilter previous; static EngineerMessageFilter current;
 public static void Register(){current=new EngineerMessageFilter();Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(current,out previous));}
 public static void Revoke(){EngineerOleFilter ignored;CoRegisterMessageFilter(previous,out ignored);current=null;previous=null;}
 public int HandleInComingCall(int type,IntPtr task,int tick,IntPtr info){return 0;}
 public int RetryRejectedCall(IntPtr task,int tick,int reject){return tick<90000 && (reject==1 || reject==2) ? 250 : -1;}
 public int MessagePending(IntPtr task,int tick,int type){return 2;}
}
public static class EngineerViewNative {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  public static string MatrixHash(object raw) {
    using(var h=IncrementalHash.CreateHash(HashAlgorithmName.SHA256)){
      Action<object> append = value => {
        var s=value==null ? "N" : value is string ? "S"+(string)value : value is bool ? "B"+value.ToString() : "V"+Convert.ToString(value,CultureInfo.InvariantCulture);
        var b=Encoding.UTF8.GetBytes(s); h.AppendData(BitConverter.GetBytes(b.Length)); h.AppendData(b);
      };
      var a=raw as Array;
      if(a==null){append(raw);}else{
        append(a.GetLength(0));append(a.GetLength(1));
        foreach(object v in a){append(v);}
      }
      return Convert.ToHexString(h.GetHashAndReset());
    }
  }
}
'@
$excel=$null;$book=$null;$viewBook=$null;$owned=$false;$ownedPid=0;$ownedTicks=0
$closeRecoveries=[Collections.Generic.List[string]]::new()
$temporary=Join-Path ([IO.Path]::GetDirectoryName($output)) ('.engineer-view-'+[guid]::NewGuid().ToString('N')+'.xlsx')
$state=[ordered]@{ownedPid=0;ownedStartTicks=0;temporaryPath=$temporary;stage='copy'}
function Save-State([string]$Stage){$script:state.stage=$Stage;$script:state|ConvertTo-Json|Set-Content -LiteralPath $StatePath -Encoding utf8}
function Assert([bool]$Passed,[string]$Message){if(-not $Passed){throw $Message}}
function Start-OwnedExcel {
    $existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
    $script:excel=New-Object -ComObject Excel.Application
    [uint32]$newPid=0;[void][EngineerViewNative]::GetWindowThreadProcessId([IntPtr]$script:excel.Hwnd,[ref]$newPid)
    Assert ($newPid -gt 0 -and $newPid -notin $existing) 'Could not establish ownership of a separate Excel process.'
    $script:owned=$true;$script:ownedPid=[int]$newPid;$script:ownedTicks=(Get-Process -Id $newPid).StartTime.ToUniversalTime().Ticks
    $script:state.ownedPid=$script:ownedPid;$script:state.ownedStartTicks=$script:ownedTicks;Save-State 'owned_excel_ready'
    Assert ($script:excel.Workbooks.Count -eq 0) 'Owned Excel unexpectedly contains a workbook.'
    $script:excel.Visible=$false;$script:excel.ScreenUpdating=$false;$script:excel.DisplayAlerts=$false;$script:excel.EnableEvents=$false;$script:excel.AskToUpdateLinks=$false;$script:excel.AutomationSecurity=3
}
function Close-SavedBook([string]$Phase){
    try{[void]$script:book.Close($false)}
    catch{
        if($_.Exception.Message -notmatch '0x800706BE|remote procedure call failed|RPC server is unavailable'){throw}
        # Some Office builds crash while closing a saved workbook containing
        # copied sheets and slicers. Publication still requires a fresh-process
        # reopen and all content/protection checks; the source is never saved.
        $script:closeRecoveries.Add($Phase)
        Stop-OwnedExcel $script:state
        Release-Com $script:excel;$script:excel=$null
    }
    Release-Com $script:book;$script:book=$null
}
function Wait-Calculation {
    $deadline=[DateTime]::UtcNow.AddSeconds(90)
    while([int]$script:excel.CalculationState -eq 1){
        if([DateTime]::UtcNow -gt $deadline){throw 'Excel calculation did not finish within 90 seconds.'}
        Start-Sleep -Milliseconds 100
    }
}
function Get-CoreSnapshot($Workbook){
    $sheets=[Collections.Generic.List[object]]::new()
    foreach($sheet in $Workbook.Worksheets){
        if([string]$sheet.Name -cin $ownedViewNames -or [string]$sheet.Name -ceq 'Blueprint view'){Release-Com $sheet;continue}
        $range=$sheet.UsedRange;$values=$range.Formula
        $tables=[Collections.Generic.List[object]]::new()
        foreach($t in $sheet.ListObjects){
            $filterStates=[Collections.Generic.List[object]]::new();$auto=$t.AutoFilter
            for($f=1;$f-le $auto.Filters.Count;$f++){
                $filter=$auto.Filters.Item($f);$criteria1=$null;$criteria2=$null;$op=0
                if($filter.On){$criteria1=$filter.Criteria1;$op=[int]$filter.Operator;try{$criteria2=$filter.Criteria2}catch{}}
                $filterStates.Add([ordered]@{on=[bool]$filter.On;operator=$op;criteria1=$criteria1;criteria2=$criteria2});Release-Com $filter
            }
            $tables.Add([ordered]@{name=[string]$t.Name;range=[string]$t.Range.Address();filters=$filterStates.ToArray()});Release-Com $auto;Release-Com $t
        }
        $protection=$sheet.Protection
        $permissions=[ordered]@{}
        foreach($key in @('AllowFormattingCells','AllowFormattingColumns','AllowFormattingRows','AllowInsertingColumns','AllowInsertingRows','AllowInsertingHyperlinks','AllowDeletingColumns','AllowDeletingRows','AllowSorting','AllowFiltering','AllowUsingPivotTables')){$permissions[$key]=[bool]$protection.$key}
        $sheets.Add([ordered]@{name=[string]$sheet.Name;range=[string]$range.Address();contentHash=[EngineerViewNative]::MatrixHash($values);visible=[int]$sheet.Visible;protected=[bool]$sheet.ProtectContents;drawings=[bool]$sheet.ProtectDrawingObjects;scenarios=[bool]$sheet.ProtectScenarios;selection=[int]$sheet.EnableSelection;permissions=$permissions;tables=$tables.ToArray()})
        $values=$null;Release-Com $protection;Release-Com $range;Release-Com $sheet
    }
    $names=[Collections.Generic.List[string]]::new()
    $ownedCaches=@(Get-OwnedCacheNames $Workbook)
    $ownedSheetReferences=(@($ownedViewNames|ForEach-Object {"(?:'"+[regex]::Escape($_.Replace("'","''"))+"'|"+[regex]::Escape($_)+")!"}) -join '|')
    foreach($name in $Workbook.Names){
        $n=[string]$name.Name;$ref=[string]$name.RefersTo
        $ownedMetadata=$n -cin $ownedCaches -and $ref -cin @('=""','=#N/A','#N/A')
        $diagramPrintName=$n -match "^'?Blueprint view'?!(_xlnm\.)?Print_(Area|Titles)$"
        if(-not $ownedMetadata -and -not $diagramPrintName -and $n -notmatch $ownedSheetReferences -and $ref -notmatch $ownedSheetReferences){$names.Add($n+'='+$ref)}
        Release-Com $name
    }
    return [ordered]@{sheets=$sheets.ToArray();names=@($names|Sort-Object);structure=[bool]$Workbook.ProtectStructure;windows=[bool]$Workbook.ProtectWindows;externalLinks=@($Workbook.LinkSources(1))}
}
function Get-OwnedCacheNames($Workbook){
    foreach($cache in $Workbook.SlicerCaches){
        $cacheName=[string]$cache.Name
        if($cacheName -cin $tableCacheNames){
            Assert ([string]$cache.ListObject.Name -ceq 'EngineerRules' -and [string]$cache.ListObject.Parent.Name -ceq 'Engineer rules') ('Slicer cache name collides with another table: '+$cacheName)
            $cacheName
        }
        elseif($cacheName -cin $legacyPivotCacheNames -and 'Pivot summary' -cin $ownedViewNames){
            Assert ($cache.PivotTables.Count -eq 1 -and [string]$cache.PivotTables.Item(1).Name -ceq 'EV_RuleContextSummary' -and [string]$cache.PivotTables.Item(1).Parent.Name -ceq 'Pivot summary') ('Slicer cache name collides with another pivot: '+$cacheName)
            $cacheName
        }
        Release-Com $cache
    }
}
function Get-Queue([string]$Status){
    switch -CaseSensitive($Status){
        'Excluded' {return 'Excluded'}
        'Context not used' {return 'Not applicable'}
        'Feature off / binding mismatch' {return 'Not applicable'}
        'Management bring-up only' {return 'Not applicable'}
        'No peer pair' {return 'Not applicable'}
        'Example only - review ports' {return 'Candidate'}
        'Example only' {return 'Candidate'}
        'Ready for firewall review' {return 'Ready'}
        default {return 'Review'}
    }
}
function Set-DefaultQueueFilter($Cache){
    $Cache.ClearManualFilter()
    $allowed=@('Candidate','Review','Ready')
    $available=@($Cache.SlicerItems|ForEach-Object {[string]$_.Name})
    if(@($available|Where-Object {$_ -cin $allowed}).Count -gt 0){
        foreach($item in $Cache.SlicerItems){if([string]$item.Name -cnotin $allowed){$item.Selected=$false};Release-Com $item}
    }
    else{
        # Excel cannot deselect every slicer item. A normal table filter safely
        # produces an empty worklist when the source has no active/review rows.
        [void]$script:engineerTable.Range.AutoFilter(2,'=__NO_ACTIVE_ENGINEER_QUEUE__')
    }
}
function Get-ExpectedFocus([string]$Component,[string]$Side){
    # This oracle reads unchanged source fields. Bidirectional traffic can be
    # initiated at either endpoint, but each rule context still counts once.
    $matching=0;$reverse=0
    if($Side -cnotin @('Source','Destination','Both')){return @{rows=0;oppositeEndpointBidirectionalRows=0}}
    for($r=1;$r-le $script:count;$r++){
        if((Get-Queue ([string]$script:raw[$r,$script:map['Status']])) -cnotin @('Candidate','Review','Ready')){continue}
        $atSource=[string]$script:raw[$r,$script:map['Source role']] -ieq $Component
        $atDestination=[string]$script:raw[$r,$script:map['Destination role']] -ieq $Component
        $bidirectional=[string]$script:raw[$r,$script:map['Direction']] -ieq 'Bidirectional'
        if($Component -ceq 'All'){$matching++;continue}
        $hit=switch($Side){
            'Source' {$atSource -or ($bidirectional -and $atDestination)}
            'Destination' {$atDestination -or ($bidirectional -and $atSource)}
            'Both' {$atSource -or $atDestination}
        }
        if($hit){$matching++}
        if(($Side -ceq 'Source' -and -not $atSource -and $atDestination -and $bidirectional) -or ($Side -ceq 'Destination' -and -not $atDestination -and $atSource -and $bidirectional)){$reverse++}
    }
    return @{rows=$matching;oppositeEndpointBidirectionalRows=$reverse}
}
function Fit-EngineerDetailRows($Rules,[int]$FirstRow,[int]$LastRow){
    # AutoFit uses the actual fonts, column widths and wrapped text in Purpose,
    # Port review and Conditions. Never reduce an endpoint-driven row height.
    $rowCount=$LastRow-$FirstRow+1;$hidden=[bool[]]::new($rowCount);$heights=[double[]]::new($rowCount)
    for($i=0;$i-lt $rowCount;$i++){$row=$Rules.Rows.Item($FirstRow+$i);$hidden[$i]=[bool]$row.Hidden;Release-Com $row}
    $Rules.Range('A'+$FirstRow+':AD'+$LastRow).EntireRow.Hidden=$false
    for($i=0;$i-lt $rowCount;$i++){$row=$Rules.Rows.Item($FirstRow+$i);$heights[$i]=[double]$row.RowHeight;Release-Com $row}
    [void]$Rules.Range('R'+$FirstRow+':T'+$LastRow).Rows.AutoFit()
    $changed=0;$maximum=0.0
    for($i=0;$i-lt $rowCount;$i++){
        $row=$Rules.Rows.Item($FirstRow+$i);$fitted=[double]$row.RowHeight
        if($fitted -gt $heights[$i]+0.01){$changed++}
        $heights[$i]=[Math]::Max($heights[$i],$fitted);$row.RowHeight=[double]($heights[$i])
        $row.Hidden=$hidden[$i];$maximum=[Math]::Max($maximum,$heights[$i]);Release-Com $row
    }
    return @{summary=[ordered]@{sizedRows=$rowCount;increasedRows=$changed;maximumHeightPoints=$maximum;columns='Purpose, Port review, Rule conditions';existingEndpointHeightsPreserved=$true;hiddenStatesPreserved=$true};rowHeights=$heights;hiddenRows=$hidden}
}
try{
    [EngineerMessageFilter]::Register()
    Save-State 'copy'
    $sourceHash=Get-SharedHash $source
    $stream=[IO.File]::Open($source,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try{$dest=[IO.File]::Open($temporary,[IO.FileMode]::CreateNew);try{$stream.CopyTo($dest)}finally{$dest.Dispose()}}finally{$stream.Dispose()}
    Assert ((Get-SharedHash $temporary) -ceq $sourceHash) 'Source changed while making the private copy.'
    (Get-Item -LiteralPath $temporary).IsReadOnly=$false
    Save-State 'start_owned_excel'
    Start-OwnedExcel
    Save-State 'open_copy'
    $book=$excel.Workbooks.Open($temporary,0,$false)
    $excel.Calculation=-4105;[void]$excel.CalculateFullRebuild();Wait-Calculation
    $initialMatrix=$book.Worksheets.Item('Firewall matrix');$initialTable=$initialMatrix.ListObjects.Item('FirewallRules')
    [void]$initialTable.AutoFilter.ApplyFilter();Release-Com $initialTable;Release-Com $initialMatrix
    $existingViews=@($book.Worksheets|Where-Object {[string]$_.Name -cin $viewNames}|ForEach-Object {[string]$_.Name})
    if($existingViews.Count){
        Assert ('Engineer guide' -cin $existingViews) 'Generated-view sheet names collide with an unmarked workbook.'
        Assert ([string]$book.Worksheets.Item('Engineer guide').Range('B1').Value2 -ceq $marker) 'Engineer guide marker is absent; existing sheets will not be replaced.'
        # Earlier marked versions owned this summary sheet. An unrelated sheet
        # with this name in an unmarked source remains part of the original core.
        if(@($book.Worksheets|Where-Object {[string]$_.Name -ceq 'Pivot summary'}).Count){$ownedViewNames=$viewNames+@('Pivot summary');$existingViews+=@('Pivot summary')}
    }
    Save-State 'snapshot_core'
    . (Join-Path $PSScriptRoot 'BlueprintDiagram.ps1')
    $commandsAdded=[bool](Initialize-VcfCommands $book)
    $baselineOwnedCaches=@(Get-OwnedCacheNames $book)
    $baseline=Get-CoreSnapshot $book
    $baselineJson=$baseline|ConvertTo-Json -Depth 12 -Compress
    Save-State 'render_blueprint'
    $diagramReport=Update-BlueprintDiagram $book
    $excel.Calculation=-4135
    $matrix=$book.Worksheets.Item('Firewall matrix');$matrixTable=$matrix.ListObjects.Item('FirewallRules')
    $sourceHeaders=$matrixTable.HeaderRowRange.Value2;$raw=$matrixTable.DataBodyRange.Value2
    Assert ($null-ne $raw -and $raw.GetLength(0) -gt 0) 'FirewallRules must contain at least one defined rule context.'
    $count=$raw.GetLength(0);$map=@{};for($c=1;$c-le $sourceHeaders.GetLength(1);$c++){$map[[string]$sourceHeaders[1,$c]]=$c}
    $required=@('Rule ID','Status','Source role','Source resources and addresses','Destination role','Destination resources and addresses','Protocol','Destination ports','Direction','Feature','Purpose','Sources','Destinations','Port review','Rule conditions','Source IP / FQDN objects','Destination IP / FQDN objects','Source networks','Destination networks','Domain ID','Cluster ID','Binding class','Source scope key','Destination scope key')
    foreach($h in $required){Assert ($map.ContainsKey($h)) ('FirewallRules is missing field '+$h)}
    $firstSourceRow=[int]$matrixTable.DataBodyRange.Row
    $matrixHash=[EngineerViewNative]::MatrixHash($raw)
    $rows=[object[,]]::new($count,30)
    $keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $components=[Collections.Generic.SortedSet[string]]::new([StringComparer]::Ordinal)
    $networks=[Collections.Generic.SortedSet[string]]::new([StringComparer]::Ordinal)
    $queues=@{};$statuses=@{}
    $defaultCount=0
    $copyFields=@{3='Rule ID';4='Domain ID';5='Cluster ID';6='Source role';7='Destination role';8='Protocol';9='Destination ports';10='Direction';11='Feature';12='Status';13='Source IP / FQDN objects';14='Destination IP / FQDN objects';15='Source networks';16='Destination networks';17='Purpose';18='Port review';19='Rule conditions';20='Source resources and addresses';21='Destination resources and addresses';22='Sources';23='Destinations';24='Binding class';25='Source scope key';26='Destination scope key'}
    for($r=1;$r-le $count;$r++){
        $id=[string]$raw[$r,$map['Rule ID']];$domain=[string]$raw[$r,$map['Domain ID']];$cluster=[string]$raw[$r,$map['Cluster ID']]
        Assert (-not [string]::IsNullOrWhiteSpace($id)) ('Blank Rule ID at source row '+($firstSourceRow+$r-1))
        $key=$id+'|'+$domain+'|'+$cluster;Assert ($keys.Add($key)) ('Duplicate rule context key: '+$key)
        $status=[string]$raw[$r,$map['Status']];$queue=Get-Queue $status
        if(-not $queues.ContainsKey($queue)){$queues[$queue]=0};$queues[$queue]++
        if(-not $statuses.ContainsKey($status)){$statuses[$status]=0};$statuses[$status]++
        if($queue -cin @('Candidate','Review','Ready')){$defaultCount++}
        $rows[($r-1),1]=$queue;$rows[($r-1),2]=$key;$rows[($r-1),27]=$firstSourceRow+$r-1
        foreach($c in $copyFields.Keys){$rows[($r-1),[int]$c]=$raw[$r,$map[$copyFields[$c]]]}
        foreach($field in @('Source role','Destination role')){if($raw[$r,$map[$field]]){[void]$components.Add([string]$raw[$r,$map[$field]])}}
        foreach($field in @('Source networks','Destination networks')){foreach($n in ([string]$raw[$r,$map[$field]] -split "`n")){if($n){[void]$networks.Add($n)}}}
    }
    Release-Com $matrixTable;Release-Com $matrix
    Save-State 'replace_owned_views'
    if($book.ProtectStructure -or $book.ProtectWindows){[void]$book.Unprotect('')}
    Assert (-not $book.ProtectStructure) 'Workbook structure could not be unlocked with the planner blank password.'
    if($existingViews.Count){
        $ownedExistingCaches=@(Get-OwnedCacheNames $book)
        for($i=$book.SlicerCaches.Count;$i-ge 1;$i--){
            $cache=$book.SlicerCaches.Item($i)
            if([string]$cache.Name -cin $ownedExistingCaches){[void]$cache.Delete()}
            Release-Com $cache
        }
        # Remove only verified generated cache metadata if this Excel version
        # leaves its sentinel name after deleting the owning cache.
        for($i=$book.Names.Count;$i-ge 1;$i--){
            $name=$book.Names.Item($i)
            if([string]$name.Name -cin $ownedExistingCaches -and [string]$name.RefersTo -cin @('=""','=#N/A','#N/A')){[void]$name.Delete()}
            Release-Com $name
        }
        foreach($name in $existingViews){[void]$book.Worksheets.Item($name).Delete()}
    }
    $viewBook=$excel.Workbooks.Open($scaffold,0,$true)
    $selection=$viewBook.Worksheets.Item($viewNames)
    [void]$selection.Copy([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
    Release-Com $selection;[void]$viewBook.Close($false);Release-Com $viewBook;$viewBook=$null
    [void]$book.Worksheets.Item('Commands').Move([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
    $rules=$book.Worksheets.Item('Engineer rules');$guide=$book.Worksheets.Item('Engineer guide')
    [void]$rules.Select()
    foreach($sheet in @($rules,$guide)){if($sheet.ProtectContents){[void]$sheet.Unprotect('')}}
    $guide.Range('B1').Value2=$marker
    Save-State 'populate_worklist'
    $last=$count+18
    if($rules.ListObjects.Count -gt 0){$engineerTable=$rules.ListObjects.Item('EngineerRules');[void]$engineerTable.Resize($rules.Range("A18:AD$last"))}
    else{$engineerTable=$rules.ListObjects.Add(1,$rules.Range("A18:AD$last"),[Type]::Missing,1);$engineerTable.Name='EngineerRules'}
    $destinationRange=$rules.Range("A19:AD$last");[void]$destinationRange.ClearContents()
    $destinationRange.NumberFormat='@'
    for($dr=0;$dr-lt $count;$dr++){for($dc=0;$dc-lt 30;$dc++){
        $value=$rows[$dr,$dc];if($null-eq $value){continue}
        $cell=$destinationRange.Cells.Item($dr+1,$dc+1)
        if($value -is [string]){$cell.Value2=if($value -match '^[=+@-]'){"'"+$value}else{[string]$value}}
        elseif($value -is [bool]){$cell.Value2=$value}
        else{$cell.Value2=[double]$value}
        Release-Com $cell
    }}
    Release-Com $destinationRange
    $rules.Range('B6').Value2='All';$rules.Range('E6').Value2='Both';$rules.Range('H6').Value2='All'
    $rules.Range("A19:A$last").NumberFormat='General';$rules.Range("AC19:AD$last").NumberFormat='General';$rules.Range("W19:X$last").NumberFormat='0';$rules.Range("AB19:AB$last").NumberFormat='0'
    $rules.Range("A19:A$last").FormulaR1C1='=IF(AND(RC29="Yes",RC30="Yes"),"Yes","No")'
    $rules.Range("AC19:AC$last").FormulaR1C1='=IF(AND(OR(R6C5="Both",R6C5="Source",R6C5="Destination"),OR(R6C2="All",AND(OR(R6C5="Both",R6C5="Source",RC11="Bidirectional"),RC7=R6C2),AND(OR(R6C5="Both",R6C5="Destination",RC11="Bidirectional"),RC8=R6C2))),"Yes","No")'
    $rules.Range("AD19:AD$last").FormulaR1C1='=IF(OR(R6C8="All",ISNUMBER(FIND(CHAR(10)&R6C8&CHAR(10),CHAR(10)&RC16&CHAR(10))),ISNUMBER(FIND(CHAR(10)&R6C8&CHAR(10),CHAR(10)&RC17&CHAR(10)))),"Yes","No")'
    $rules.Range('B7').Formula='=COUNTIFS(A19:A'+$last+',"Yes",B19:B'+$last+',"Candidate")+COUNTIFS(A19:A'+$last+',"Yes",B19:B'+$last+',"Review")+COUNTIFS(A19:A'+$last+',"Yes",B19:B'+$last+',"Ready")'
    $rules.Range('B8').Formula='=SUBTOTAL(103,D19:D'+$last+')'
    # Excel can remap scaffold structured references to an unrelated existing
    # table during a cross-workbook sheet copy. Bind these only after the final
    # EngineerRules table exists in the destination workbook.
    for($guideRow=9;$guideRow-le 13;$guideRow++){$guide.Range('B'+$guideRow).Formula='=COUNTIFS(EngineerRules[Work queue],A'+$guideRow+')'}
    $rules.Range('C7').Value2='Candidate + Review + Ready; excludes other filters'
    $rules.Range('A4').Value2='Generated from this planner after recalculation. Ready means ready for firewall review, not approved or deployed.'
    $listComponents=@('All')+@($components);$listNetworks=@('All')+@($networks)
    $listEnd=[Math]::Max($listComponents.Count,$listNetworks.Count)+7
    $guide.Range("H8:I$listEnd").ClearContents()
    $guide.Range("H8:I$listEnd").NumberFormat='@'
    for($i=0;$i-lt $listComponents.Count;$i++){$s=[string]$listComponents[$i];$guide.Cells.Item($i+8,8).Value2=if($s-match '^[=+@-]'){"'"+$s}else{$s}}
    for($i=0;$i-lt $listNetworks.Count;$i++){$s=[string]$listNetworks[$i];$guide.Cells.Item($i+8,9).Value2=if($s-match '^[=+@-]'){"'"+$s}else{$s}}
    foreach($spec in @(@('B6','H',$listComponents.Count),@('H6','I',$listNetworks.Count))){$dv=$rules.Range($spec[0]).Validation;[void]$dv.Delete();[void]$dv.Add(3,1,1,"='Engineer guide'!`$"+$spec[1]+'$8:$'+$spec[1]+'$'+([int]$spec[2]+7));$dv.InCellDropdown=$true;$dv.ShowError=$true;Release-Com $dv}
    $sideValidation=$rules.Range('E6').Validation;[void]$sideValidation.Delete();[void]$sideValidation.Add(3,1,1,'Source,Destination,Both');$sideValidation.InCellDropdown=$true;$sideValidation.ShowError=$true;$sideValidation.IgnoreBlank=$false;$sideValidation.ErrorTitle='Choose an endpoint side';$sideValidation.ErrorMessage='Choose Source, Destination or Both.';Release-Com $sideValidation
    $refreshed=[DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')
    $guide.Range('B18').Value2=$refreshed
    $guide.Range('B19').Value2=[IO.Path]::GetFileName($source)
    $guide.Range('B20').Value2=$matrixHash
    $guide.Range('B21').Value2='Verified in Excel: source fields, core formulas and inputs, protection, endpoint-side filters, save and reopen.'
    $guide.Rows.Item(1).Hidden=$true
    $rules.Range('C8').Value2='After changing a selector, open Focus match filter and reselect Yes. Slicers update immediately.'
    # Keep all original detail fields visible; only formula helpers are hidden.
    $rules.Range("A19:AD$last").WrapText=$true
    $rules.Range("A19:AD$last").RowHeight=44
    for($i=0;$i-lt $count;$i++){
        $lines=1
        foreach($c in @(13,14,15,16)){$text=[string]$rows[$i,$c];$width=if($c-lt15){42}else{22};$n=0;foreach($line in ($text -split "`n")){$n+=[Math]::Max(1,[Math]::Ceiling($line.Length/$width))};$lines=[Math]::Max($lines,$n)}
        if($lines -gt 2){$rules.Rows.Item($i+19).RowHeight=[Math]::Min(409,14*$lines+8)}
    }
    $detailSizing=Fit-EngineerDetailRows $rules 19 $last
    $rules.Range('U:AB').EntireColumn.Hidden=$false;$rules.Range('AC:AD').EntireColumn.Hidden=$true
    Save-State 'checkpoint_populated_views'
    [void]$book.Save()
    Save-State 'native_worklist_filters'
    $cacheMap=@{};$fields=@('Work queue','Domain','Cluster','Feature')
    for($i=0;$i-lt $fields.Count;$i++){
        $field=$fields[$i];$id=$field.Replace(' ','');$cache=$book.SlicerCaches.Add2($engineerTable,$field,'EV_'+$id)
        # Excel's table-slicer Add2 ignores its optional name on some builds.
        # Assign the documented Name property explicitly for stable ownership.
        $cache.Name='EV_'+$id;$cacheMap[$field]=$cache
        $slicer=$cache.Slicers.Add($rules,[Type]::Missing,'EV_Filter'+$id,$field,[double]$rules.Range('A10').Top,([double]$rules.Range('A10').Left+200*$i),190,([double]$rules.Range('A17').Top-[double]$rules.Range('A10').Top-5))
        $slicer.NumberOfColumns=1;$slicer.RowHeight=14;$slicer.Style='SlicerStyleLight2';Release-Com $slicer
    }
    [void]$engineerTable.Range.AutoFilter(1,'Yes');Set-DefaultQueueFilter $cacheMap['Work queue']
    $excel.Calculation=-4105;$book.RefreshAll();[void]$excel.CalculateFullRebuild();Wait-Calculation
    [void]$engineerTable.AutoFilter.ApplyFilter();[void]$excel.Calculate();Wait-Calculation
    Assert ([int]$rules.Range('B8').Value2 -eq $defaultCount) 'The default worklist count differs from the eligible source rows.'
    $rules.Activate();$rules.Range('E19').Select();$excel.ActiveWindow.FreezePanes=$false;$excel.ActiveWindow.SplitRow=18;$excel.ActiveWindow.SplitColumn=4;$excel.ActiveWindow.FreezePanes=$true;$excel.ActiveWindow.Zoom=80;$rules.Range('A1').Select()
    Save-State 'protect_generated_sheets'
    foreach($sheet in @($rules,$guide)){$sheet.Cells.Locked=$true}
    $rules.Range('B6').Locked=$false;$rules.Range('E6').Locked=$false;$rules.Range('H6').Locked=$false;$rules.Range('A18:AD18').Locked=$false
    # Sheet contents/formulas stay protected; drawing objects remain selectable
    # so Excel table slicers can filter. Original sheet protection is untouched.
    [void]$rules.Protect('',$false,$true,$true,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$true,$true)
    [void]$guide.Protect('',$true,$true,$true)
    if($baseline.structure -or $baseline.windows){[void]$book.Protect('',$baseline.structure,$baseline.windows)}
    Save-State 'save_final'
    [void]$book.Save()
    Save-State 'close_saved_copy'
    Close-SavedBook 'close_after_save'
    if($null-eq $excel){Start-OwnedExcel}
    Save-State 'reopen_saved_copy'
    $book=$excel.Workbooks.Open($temporary,0,$true);Wait-Calculation
    $after=Get-CoreSnapshot $book
    [void](Test-BlueprintDiagram $book $diagramReport)
    [void](Test-VcfCommands $book)
    # Get-CoreSnapshot excludes only exact metadata names backed by verified
    # generated caches. Other names, including unrelated EV_* names, must match.
    $slicerNameAdditions=@($tableCacheNames|Where-Object {$_ -cnotin $baselineOwnedCaches})
    if(($after|ConvertTo-Json -Depth 12 -Compress) -cne $baselineJson){
        $differences=[Collections.Generic.List[string]]::new()
        foreach($k in @('names','structure','windows','externalLinks')){if(($baseline[$k]|ConvertTo-Json -Depth 6 -Compress) -cne ($after[$k]|ConvertTo-Json -Depth 6 -Compress)){$differences.Add($k+': before='+($baseline[$k]|ConvertTo-Json -Depth 4 -Compress)+' after='+($after[$k]|ConvertTo-Json -Depth 4 -Compress))}}
        for($i=0;$i-lt $baseline.sheets.Count;$i++){if(($baseline.sheets[$i]|ConvertTo-Json -Depth 8 -Compress) -cne ($after.sheets[$i]|ConvertTo-Json -Depth 8 -Compress)){$changed=@();foreach($key in $baseline.sheets[$i].Keys){if(($baseline.sheets[$i][$key]|ConvertTo-Json -Depth 8 -Compress) -cne ($after.sheets[$i][$key]|ConvertTo-Json -Depth 8 -Compress)){$changed+=$key}};$differences.Add('sheet '+$baseline.sheets[$i].name+' fields '+($changed-join','))}}
        throw ('Original sheet values/formulas, names, table ranges, links, visibility or protection changed. '+($differences-join'; '))
    }
    $rules=$book.Worksheets.Item('Engineer rules');$engineerTable=$rules.ListObjects.Item('EngineerRules')
    $saved=$engineerTable.DataBodyRange.Value2
    Assert ($saved.GetLength(0) -eq $count) 'Saved worklist row count changed.'
    $mismatch=0
    for($r=1;$r-le $count;$r++){
        foreach($c in $copyFields.Keys){
            $actual=$saved[$r,([int]$c+1)];$expected=$raw[$r,$map[$copyFields[$c]]]
            if($null-eq $actual -or $null-eq $expected){if([string]$actual -cne [string]$expected){$mismatch++}}
            elseif(($actual -is [string]) -ne ($expected -is [string]) -or ($actual -is [bool]) -ne ($expected -is [bool])){$mismatch++}
            elseif([string]$actual -cne [string]$expected){$mismatch++}
        }
        Assert ([string]$saved[$r,3] -ceq [string]$rows[($r-1),2]) 'Saved rule keys differ from the source.'
    }
    Assert ($mismatch -eq 0) 'Saved worklist differs from source matrix fields.'
    Assert ([int]$rules.Range('B8').Value2 -eq $defaultCount) 'Default filter did not survive saving and reopening.'
    $savedGuide=$book.Worksheets.Item('Engineer guide');$verifiedGuideCounts=[ordered]@{}
    for($guideRow=9;$guideRow-le 13;$guideRow++){
        $queueLabel=[string]$savedGuide.Range('A'+$guideRow).Value2
        Assert ($queueLabel -cin @('Ready','Candidate','Review','Not applicable','Excluded')) ('Unexpected guide queue label at row '+$guideRow)
        $expectedQueueCount=if($queues.ContainsKey($queueLabel)){[int]$queues[$queueLabel]}else{0}
        $actualQueueCount=[int]$savedGuide.Range('B'+$guideRow).Value2
        Assert ($actualQueueCount -eq $expectedQueueCount) ("Saved guide count differs for '$queueLabel': expected $expectedQueueCount, actual $actualQueueCount.")
        $verifiedGuideCounts[$queueLabel]=$actualQueueCount
    }
    Assert ($verifiedGuideCounts.Count -eq 5 -and ($verifiedGuideCounts.Values|Measure-Object -Sum).Sum -eq $count) 'Guide queue counts do not cover every source context exactly once.'
    Release-Com $savedGuide
    $ownedSlicers=@(Get-OwnedCacheNames $book)
    Assert ($ownedSlicers.Count -eq 4 -and @($ownedSlicers|Where-Object {$_ -cnotin $tableCacheNames}).Count -eq 0) ('Expected four generated table slicer caches after reopening. Actual cache names: '+(@($book.SlicerCaches|ForEach-Object {[string]$_.Name}) -join ', '))
    if('Pivot summary' -cin $ownedViewNames){foreach($name in $legacyPivotCacheNames){Assert (@($book.SlicerCaches|Where-Object {[string]$_.Name -ceq $name}).Count -eq 0) ('Legacy generated pivot cache survived refresh: '+$name)}}
    $savedPivotCacheCount=[int]$book.PivotCaches().Count
    foreach($name in $viewNames){Assert ([bool]$book.Worksheets.Item($name).ProtectContents) ('Generated sheet is not protected: '+$name)}
    Assert ([bool]$rules.Range("A19:A$last").Locked -and [bool]$rules.Range("AC19:AD$last").Locked) 'Generated helper formulas must remain locked.'
    foreach($cell in @('B6','E6','H6')){Assert (-not [bool]$rules.Range($cell).Locked) ('Selector must be editable under protection: '+$cell)}
    Assert ([string]$rules.Range('B6').Value2 -ceq 'All' -and [string]$rules.Range('E6').Value2 -ceq 'Both' -and [string]$rules.Range('H6').Value2 -ceq 'All') 'Default selector values did not survive saving and reopening.'
    $savedSideValidation=$rules.Range('E6').Validation
    Assert ([int]$savedSideValidation.Type -eq 3 -and [int]$savedSideValidation.AlertStyle -eq 1 -and [bool]$savedSideValidation.ShowError -and -not [bool]$savedSideValidation.IgnoreBlank) 'Endpoint-side validation must reject values outside its list.'
    Release-Com $savedSideValidation
    $rules.Activate();$rules.Range('A18').Select();$ribbonReapplyEnabled=[bool]$excel.CommandBars.GetEnabledMso('FilterReapply')
    $sideChecks=[Collections.Generic.List[object]]::new()
    $testComponents=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    [void]$testComponents.Add('All')
    if($components.Contains('ESXi Management')){[void]$testComponents.Add('ESXi Management')}
    elseif($components.Count -gt 0){[void]$testComponents.Add(@($components)[0])}
    for($r=1;$r-le $count;$r++){
        if((Get-Queue ([string]$raw[$r,$map['Status']])) -cnotin @('Candidate','Review','Ready')){continue}
        if([string]$raw[$r,$map['Direction']] -ieq 'Bidirectional' -and [string]$raw[$r,$map['Source role']] -ine [string]$raw[$r,$map['Destination role']]){
            [void]$testComponents.Add([string]$raw[$r,$map['Source role']]);[void]$testComponents.Add([string]$raw[$r,$map['Destination role']]);break
        }
    }
    foreach($testComponent in $testComponents){foreach($side in @('Source','Destination','Both')){
        $expectedFocus=Get-ExpectedFocus $testComponent $side
        $rules.Range('B6').Value2=$testComponent;$rules.Range('E6').Value2=$side;[void]$excel.Calculate();Wait-Calculation
        [void]$engineerTable.AutoFilter.ApplyFilter();[void]$excel.Calculate();Wait-Calculation
        $actualFocus=[int]$rules.Range('B8').Value2
        Assert ($actualFocus -eq $expectedFocus.rows) ("Protected endpoint-side filter differs for component '$testComponent', side '$side': expected "+$expectedFocus.rows+", actual $actualFocus.")
        $sideChecks.Add([ordered]@{component=$testComponent;side=$side;expectedRows=$expectedFocus.rows;actualRows=$actualFocus;oppositeEndpointBidirectionalRows=$expectedFocus.oppositeEndpointBidirectionalRows})
    }}
    # Programmatic writes bypass Excel validation; formulas still fail closed.
    $rules.Range('B6').Value2='All';$rules.Range('E6').Value2='Invalid side';[void]$excel.Calculate();Wait-Calculation;[void]$engineerTable.AutoFilter.ApplyFilter();[void]$excel.Calculate();Wait-Calculation
    Assert ([int]$rules.Range('B8').Value2 -eq 0) 'Invalid endpoint side must produce no matches.'
    $rules.Range('E6').Value2='Both';[void]$excel.Calculate();Wait-Calculation;[void]$engineerTable.AutoFilter.ApplyFilter();[void]$excel.Calculate();Wait-Calculation
    Assert ([int]$rules.Range('B8').Value2 -eq $defaultCount) 'Protected selector reset did not restore the default worklist.'
    $queueCache=$book.SlicerCaches.Item('EV_Workqueue');$queueCache.ClearManualFilter()
    [void]$excel.Calculate();Wait-Calculation
    Assert ([int]$rules.Range('B8').Value2 -eq $count) 'Protected slicer clearing did not reveal all contexts.'
    Set-DefaultQueueFilter $queueCache;[void]$excel.Calculate();Wait-Calculation
    Assert ([int]$rules.Range('B8').Value2 -eq $defaultCount) 'Protected work queue slicer did not restore the default view.'
    Assert ((Get-SharedHash $source) -ceq $sourceHash) 'The source changed during refresh; the generated copy is not published.'
    $excelVersion=[string]$excel.Version
    Close-SavedBook 'close_after_verification'
    Save-State 'verified'
    $report=[ordered]@{success=$true;viewSchema=$marker;viewType='Worklist';refreshedUtc=$refreshed;outputPath=$output;sourceSha256=$sourceHash;matrixFingerprint=$matrixHash;sourceRows=$count;defaultVisibleRows=$defaultCount;queueCounts=$queues;guideQueueCounts=$verifiedGuideCounts;guideQueueCountsVerified=$true;statusCounts=$statuses;sourceFieldComparisons=($count*$required.Count);sourceFieldMismatches=$mismatch;corePreserved=$true;generatedSlicerNamesAdded=$slicerNameAdditions;slicers=4;worklistSlicers=4;pivotSlicers=0;pivotCacheCount=$savedPivotCacheCount;saveReopenPassed=$true;generatedSheetsProtected=$true;generatedFormulasLocked=$true;protectedSlicersPassed=$true;protectedSelectorReapplyPassed=$true;protectedEndpointSideChecks=$sideChecks.ToArray();invalidEndpointSideHidesAll=$true;bidirectionalRulesUseBothEndpoints=$true;ribbonReapplyEnabled=$ribbonReapplyEnabled;selectorsReset=[ordered]@{component='All';endpointSide='Both';network='All'};defaultQueues=@('Candidate','Review','Ready');generatedSheets=$viewNames;excelVersion=$excelVersion;ownedExcelCloseRecoveries=$closeRecoveries.ToArray()}
    $report.detailRowSizing=$detailSizing.summary
    $report.blueprintDiagram=$diagramReport
    $report.commandsTab=[ordered]@{schema='VCF_COMMANDS_V1';added=$commandsAdded;lastTab=$true;protectedFormula=$true;editablePaths=$true;diagramNavigation=$true}
    $report.preservationScope='All original sheets except the regenerated Blueprint view; engineer views are regenerated separately.'
    if(Test-Path -LiteralPath $output){throw 'The output appeared during refresh; it was not overwritten.'}
    [IO.File]::Move($temporary,$output,$false)
    $report.outputSha256=Get-SharedHash $output
    $report|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $reportPath -Encoding utf8
}
catch {
    if($KeepDiagnosticCopy -and (Test-Path -LiteralPath $temporary)){
        $diagnostic=$output+'.diagnostic.xlsx'
        if(-not (Test-Path -LiteralPath $diagnostic)){try{[IO.File]::Copy($temporary,$diagnostic,$false)}catch{}}
    }
    Write-Error ($_|Out-String);exit 1
}
finally {
    if($viewBook){try{[void]$viewBook.Close($false)}catch{};Release-Com $viewBook}
    if($book){try{[void]$book.Close($false)}catch{};Release-Com $book}
    if($owned -and $excel){try{[void]$excel.Quit()}catch{}}
    Release-Com $excel
    [EngineerMessageFilter]::Revoke()
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($owned){Stop-OwnedExcel $state}
    if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue}
}
