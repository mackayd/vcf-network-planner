# Native, formula-linked diagrams. Loaded by Update-FirewallViews.ps1.
# All writes are confined to the generated Blueprint view canvas and bp_ shapes.
Set-StrictMode -Version Latest

function Get-BpValue($Data,[string]$Sheet,[string]$Cell){return $Data[$Sheet+'!'+$Cell]}
function Get-BpInputs($Book){
    $data=[ordered]@{}
    foreach($spec in @(@('Design','A6:H26'),@('Domains & clusters','A8:S48'),@('Resources','A9:R168'),@('Deployment','D5:I12'),@('Deployment','A19:L162'),@('Network plan','A8:I55'))){
        $range=$Book.Worksheets.Item($spec[0]).Range($spec[1]);$raw=$range.Value2
        for($r=1;$r-le $raw.GetLength(0);$r++){for($c=1;$c-le $raw.GetLength(1);$c++){
            $n=$c+[int]$range.Column-1;$col='';while($n-gt 0){$n--;$col=[char](65+$n%26)+$col;$n=[Math]::Floor($n/26)}
            $data[$spec[0]+'!'+$col+($r+[int]$range.Row-1)]=$raw[$r,$c]
        }}
    }
    return $data
}
function New-BpModel($Data){
    $domains=@();$clusters=@();$nodes=@();$issues=@();$relationships=@();$seen=@{}
    foreach($r in 8..11){if((Get-BpValue $Data 'Domains & clusters' "A$r") -eq 'Yes'){
        $id=[string](Get-BpValue $Data 'Domains & clusters' "B$r");if(-not $id -or $id -in @($domains|ForEach-Object {$_.id})){throw 'Active diagram domains need unique IDs.'}
        $domains+=@{id=$id;row=$r;name=[string](Get-BpValue $Data 'Domains & clusters' "C$r");nsx=[string](Get-BpValue $Data 'Domains & clusters' "F$r")}
    }}
    foreach($r in 17..20){if((Get-BpValue $Data 'Domains & clusters' "A$r") -eq 'Yes'){
        $id=[string](Get-BpValue $Data 'Domains & clusters' "B$r");if(-not $id -or $id -in @($clusters|ForEach-Object {$_.id})){throw 'Active diagram clusters need unique IDs.'}
        $domain=[string](Get-BpValue $Data 'Domains & clusters' "C$r")
        $azs=Get-BpValue $Data 'Domains & clusters' "E$r"
        $clusters+=@{id=$id;row=$r;domain=$domain;azs=$azs;name=[string](Get-BpValue $Data 'Domains & clusters' "D$r");site1=[string](Get-BpValue $Data 'Domains & clusters' "F$r");site2=[string](Get-BpValue $Data 'Domains & clusters' "G$r")}
        if($domain -notin @($domains|ForEach-Object {$_.id})){$issues+="Cluster $id has no active owning domain."}
        if($azs -notin @(1,2)){$issues+="Cluster $id has an unresolved data-AZ count."}
        $relationships+=@{source=$id;target=$domain;kind='Owned by domain'}
    }}
    $skipRoles=@('ESXi vMotion','ESXi vSAN','ESXi TEP','ESXi NFS','NSX Edge TEP','NSX Edge Uplink 1','NSX Edge Uplink 2','Witness source interface')
    foreach($r in 9..168){
        if((Get-BpValue $Data 'Resources' "A$r") -ne 'Yes'){continue}
        $role=[string](Get-BpValue $Data 'Resources' "C$r");if($role -in $skipRoles){continue}
        $fqdn=[string](Get-BpValue $Data 'Resources' "D$r");$location=[string](Get-BpValue $Data 'Resources' "L$r")
        $hosted=[string](Get-BpValue $Data 'Resources' "O$r");$managed=[string](Get-BpValue $Data 'Resources' "N$r")
        if($role -eq 'ESXi Management'){$hosted=$managed}
        $id=[string](Get-BpValue $Data 'Resources' "B$r");if(-not $id){$id="row-$r"}
        $key=$role+'|'+$(if($fqdn){$fqdn.ToLowerInvariant()}else{$id})
        if($seen.ContainsKey($key)){
            $old=$seen[$key]
            if($old.hosted -ne $hosted -or $old.location -ne $location){$old.unresolved=$true;$old.hosted='';$issues+="Conflicting placement for $role $fqdn."}
            $old.rows+=,$r;continue
        }
        $unresolved=$false;$gated=$false
        if(-not $fqdn){$unresolved=$true;$issues+="$role $id has no resource name or FQDN."}
        if(($role -eq 'Automation' -and (Get-BpValue $Data 'Deployment' 'I9') -ne 'Yes') -or ($role -eq 'Supervisor' -and (Get-BpValue $Data 'Deployment' 'D10') -ne 'Yes')){$gated=$true;$unresolved=$true;$issues+="$role inventory is present while its feature is disabled."}
        if($role -eq 'NSX Edge'){
            $owner=@($clusters|Where-Object {$_.id -eq $(if($managed){$managed}else{$hosted})})
            if($owner.Count -ne 1 -or (Get-BpValue $Data 'Domains & clusters' ('O'+$owner[0].row)) -in @($null,'','None')){$gated=$true;$unresolved=$true;$issues+="NSX Edge $fqdn lacks an enabled target-cluster Edge model."}
        }
        if($hosted -and $hosted -notin @($clusters|ForEach-Object {$_.id})){$unresolved=$true;$issues+="$role $fqdn references an inactive or unknown hosting cluster."}
        if(-not $hosted -and $location -notin @('External','Shared service','Witness')){$unresolved=$true;$issues+="$role $fqdn has unresolved hosting."}
        if($hosted -and $location -notin @('AZ1','AZ2')){$unresolved=$true;$issues+="$role $fqdn lacks a data-AZ location for its hosting cluster."}
        $owner=@($clusters|Where-Object {$_.id -eq $hosted})
        if($owner.Count -eq 1 -and $location -eq 'AZ2' -and $owner[0].azs -ne 2){$unresolved=$true;$issues+="$role $fqdn references AZ2 in a single-AZ cluster."}
        $nsx=[string](Get-BpValue $Data 'Resources' "P$r");$served=[string](Get-BpValue $Data 'Resources' "M$r")
        $servedDomains=@();if($served){$servedDomains+=,$served}
        if($nsx){$servedDomains+=@($domains|Where-Object {$_.nsx -eq $nsx}|ForEach-Object {$_.id})}
        $servedDomains=@($servedDomains|Sort-Object -Unique)
        $node=@{id=$id;role=$role;fqdn=$fqdn;row=$r;rows=@($r);hosted=$hosted;managed=$managed;location=$location;nsx=$nsx;served=$servedDomains;unresolved=$unresolved;gated=$gated}
        $nodes+=,$node;$seen[$key]=$node
        if($hosted){$relationships+=@{source=$id;target=$hosted;kind='Hosted by cluster'}}
        foreach($domain in $servedDomains){$relationships+=@{source=$id;target=$domain;kind='Serves domain'}}
    }
    foreach($gate in @(@('Automation','I9'),@('Supervisor','D10'))){if((Get-BpValue $Data 'Deployment' $gate[1]) -eq 'Yes' -and @($nodes|Where-Object {$_.role -eq $gate[0]}).Count -eq 0){$issues+=($gate[0]+' is enabled but has no active resource inventory.')}}
    return @{domains=@($domains);clusters=@($clusters);nodes=@($nodes);issues=@($issues|Sort-Object -Unique);relationships=@($relationships)}
}
function Get-BpColor([string]$Role){
    switch($Role){
        'Automation'{return 'F3BA16'}
        {$_ -in @('Operations','Cloud Proxy','Fleet Lifecycle','Salt Master','Salt RaaS','Runtime Pool','License Server')}{return '61A60E'}
        {$_ -in @('ESXi Management','vCenter','NSX Manager','NSX VIP','NSX Edge','Supervisor')}{return '0098C7'}
        default{return '000000'}
    }
}
function Convert-BpColor([string]$Hex){return [Convert]::ToInt32($Hex.Substring(0,2),16)+256*[Convert]::ToInt32($Hex.Substring(2,2),16)+65536*[Convert]::ToInt32($Hex.Substring(4,2),16)}
function Bp-Ref([string]$Sheet,[string]$Cell){return "'"+$Sheet.Replace("'","''")+"'!"+$Cell}
function Bp-Label([string]$Sheet,[string]$Cell){$ref=Bp-Ref $Sheet $Cell;return 'IF('+ $ref+'="","Unspecified",'+$ref+')'}
function Set-BpMatrix($Range,[string]$Property,$Matrix){
    $arguments=[object[]]::new(1);$arguments[0]=$Matrix
    [void]$Range.GetType().InvokeMember($Property,[Reflection.BindingFlags]::SetProperty,$null,$Range,$arguments)
}
function Add-BpBox($Sheet,[string]$Id,[double]$X,[double]$Y,[double]$W,[double]$H,[string]$Text,[string]$Color='000000',[double]$Size=12,[switch]$Dashed,[switch]$Heading,[switch]$Formula,[switch]$NoBorder){
    $shape=$Sheet.Shapes.AddShape(1,$X,$Y,$W,$H);$shape.Name='bp_'+$Id;$shape.AlternativeText='VCF_BLUEPRINT_V1:'+ $Id
    $shape.Fill.ForeColor.RGB=16777215;$shape.Line.ForeColor.RGB=Convert-BpColor $Color;$shape.Line.Weight=1
    $shape.Shadow.Visible=0;$shape.Placement=3;$shape.Locked=$true
    if($Dashed){$shape.Line.DashStyle=4};if($NoBorder){$shape.Line.Visible=0}
    $shape.TextFrame2.MarginLeft=8;$shape.TextFrame2.MarginRight=8;$shape.TextFrame2.MarginTop=5;$shape.TextFrame2.MarginBottom=4
    $shape.TextFrame2.WordWrap=-1;$shape.TextFrame2.VerticalAnchor=1
    $shape.TextFrame2.TextRange.ParagraphFormat.Alignment=1
    $shape.TextFrame2.TextRange.Font.Name='Arial';$shape.TextFrame2.TextRange.Font.Size=$Size
    $shape.TextFrame.Characters().Font.Color=0
    if($Formula){
        $script:bpLabelRow++;$cell=$Sheet.Range('AF'+$script:bpLabelRow);$cell.Formula=$Text;$cell.Font.Name='Arial';$cell.Font.Size=$Size;$cell.Font.Color=0
        $shape.DrawingObject.Formula="='Blueprint view'!`$AF`$$script:bpLabelRow"
    }else{$shape.TextFrame2.TextRange.Text=$Text}
    if($Heading){$shape.TextFrame.Characters().Font.Bold=$true;$shape.TextFrame.Characters().Font.Color=Convert-BpColor $Color}
    return $shape
}
function Add-BpLine($Sheet,[string]$Id,[double]$X1,[double]$Y1,[double]$X2,[double]$Y2){
    $line=$Sheet.Shapes.AddConnector(1,$X1,$Y1,$X2,$Y2);$line.Name='bp_'+$Id;$line.AlternativeText='VCF_BLUEPRINT_V1:'+ $Id;$line.Line.ForeColor.RGB=0;$line.Line.Weight=1;$line.Placement=3;$line.Locked=$true
}
function Start-BpPage($Sheet,[string]$Title,[string]$Subtitle){
    $script:bpPage++;$top=[double]$Sheet.Cells.Item(($script:bpNextRow+1),1).Top;$script:bpPageTop=$top;$script:bpPageStarts.Add($script:bpNextRow+1)
    $shape=Add-BpBox $Sheet "page$script:bpPage-title" 12 ($top+8) 950 45 $Title '005C8A' 20 -Heading -Formula -NoBorder
    $Sheet.Calculate();$h=[Math]::Max(45,[double]$shape.TextFrame2.TextRange.BoundHeight+12);$shape.Height=[double]$h
    $context=Add-BpBox $Sheet "page$script:bpPage-context" 12 ($top+$h+12) 950 28 $Subtitle '000000' 11 -NoBorder
    $context.Height=[Math]::Max(28,[double]$context.TextFrame2.TextRange.BoundHeight+12)
    return ($context.Top+$context.Height+12)
}
function End-BpPage($Sheet,[double]$Bottom){
    $footer=[Math]::Max($script:bpPageTop+664,$Bottom+12)
    $null=Add-BpBox $Sheet "page$script:bpPage-status" 12 $footer 730 28 "='Blueprint view'!`$B`$4" '000000' 11 -Formula -NoBorder
    $back=Add-BpBox $Sheet "page$script:bpPage-back" 760 $footer 202 28 'Back to Design' '005C8A' 11 -NoBorder
    [void]$Sheet.Hyperlinks.Add($back,'',"'Design'!A1",'Return to the design inputs')
    $script:bpNextRow=[int][Math]::Ceiling(($footer+36)/$script:bpRowHeight)
}
function Measure-BpText($Sheet,[string]$Text,[double]$Width,[double]$Size=11,[double]$Minimum=28){
    $probe=Add-BpBox $Sheet 'measure' 0 0 $Width 400 $Text '000000' $Size -NoBorder
    try {return [Math]::Max($Minimum,[double]$probe.TextFrame2.TextRange.BoundHeight+12)} finally {$probe.Delete()}
}
function Get-BpRelation($Node){
    $relation=if($Node.nsx){'NSX '+$Node.nsx+'; serves '+($Node.served -join ', ')}elseif($Node.served.Count){'Serves '+($Node.served -join ', ')}else{'Resource '+$Node.id}
    if($Node.unresolved){$relation='Unresolved placement / feature; '+$relation}
    return $relation
}
function Group-BpNodes($Sheet,$Nodes,[double]$Width){
    $batches=[Collections.Generic.List[object]]::new();$batch=[Collections.Generic.List[object]]::new();$used=0
    foreach($node in $Nodes){
        $label=$node.role+"`n"+$(if($node.fqdn){$node.fqdn}else{'Unspecified'})+"`n"+(Get-BpRelation $node)
        $height=Measure-BpText $Sheet $label $Width 11 68
        if($height -gt 380){throw ('Resource label exceeds a full diagram page: '+$node.id)}
        if($batch.Count -and $used+$height+5 -gt 385){$batches.Add($batch.ToArray());$batch.Clear();$used=0}
        $batch.Add($node);$used+=$height+5
    }
    if($batch.Count){$batches.Add($batch.ToArray())}
    return @{batches=$batches.ToArray()}
}
function Add-BpReference($Sheet,$Data){
    $model=[string](Get-BpValue $Data 'Design' 'C6');$mixed=[string](Get-BpValue $Data 'Design' 'C14')
    $minimal=$model -eq 'Single site - minimal';$stretch=$model -in @('Multiple sites - one region','Stretched region + other regions');$multi=$model -in @('Multiple regions','Stretched region + other regions')
    $y=Start-BpPage $Sheet '=Design!C6&" / fleet reference pattern"' 'Representative topology. Peer instances and components below are illustrative, not this workbook''s inventory.'
    if($model -notin @('Single site - minimal','Single site - standard','Multiple sites - one region','Multiple regions','Stretched region + other regions')){$null=Add-BpBox $Sheet 'unknown-blueprint' 20 $y 930 80 'Select a supported blueprint on Design before drawing the reference topology.';End-BpPage $Sheet ($y+80);return}
    $n=if($multi -and $stretch){3}elseif($multi -or $stretch){2}else{1};$w=(930-16*($n-1))/$n
    for($i=0;$i-lt $n;$i++){$label=if($multi -and $i -eq $n-1){'Representative region B / site'}elseif($stretch){'Representative region A / AZ '+($i+1)}else{'Representative region A / site'};$null=Add-BpBox $Sheet "ref-site-$i" (20+$i*($w+16)) $y $w 468 $label '000000' 12 -Dashed}
    $null=Add-BpBox $Sheet 'ref-ops' 32 ($y+38) 906 30 'VCF Operations / fleet services (reference scope)' '61A60E'
    $null=Add-BpBox $Sheet 'ref-auto' 32 ($y+76) 906 30 'VCF Automation where selected (reference scope)' 'F3BA16'
    $count=if($multi){2}else{1}
    for($i=0;$i-lt $count;$i++){
        $x=32;$iw=906;if($multi){$iw=if($stretch){2*$w+4}else{$w-24};if($i -eq 1){$x=32+($n-1)*($w+16);$iw=$w-24}}
        $null=Add-BpBox $Sheet "ref-instance-$i" $x ($y+120) $iw 245 ('Representative instance '+($i+1)) '000000' 13 -Heading
        $text=if($minimal){'Consolidated management / workload cluster'}else{'Management domain cluster'}
        $null=Add-BpBox $Sheet "ref-management-$i" ($x+12) ($y+157) ($iw-24) 92 $text '0098C7' 12
        $null=Add-BpBox $Sheet "ref-controllers-$i" ($x+22) ($y+192) ($iw-44) 44 'vCenter and NSX management' '0098C7' 11
        if(-not $minimal){$null=Add-BpBox $Sheet "ref-workload-$i" ($x+12) ($y+264) ($iw-24) 80 'Workload domain clusters; AZ span follows each cluster design' '0098C7' 12}
    }
    foreach($i in 0..($n-1)){
        $x=32+$i*($w+16);$null=Add-BpBox $Sheet "ref-virtual-$i" $x ($y+381) ($w-24) 30 'Virtual networking' '0098C7'
        Add-BpLine $Sheet "ref-link-$i" ($x+($w-24)/2) ($y+411) ($x+($w-24)/2) ($y+425)
        $null=Add-BpBox $Sheet "ref-physical-$i" $x ($y+425) ($w-24) 30 'Physical network' '000000'
    }
    $note=if($multi){'Other regional instances are representative. Replication is optional and is not inferred as enabled.'}elseif($stretch){'Two data AZs share this instance. A vSAN witness is outside both data-AZ boundaries.'}else{'One site. Availability and deployed components follow the selected-instance inventory.'}
    if($model -eq 'Stretched region + other regions'){$note+=' Selected instance model: '+$mixed+'.'}
    $null=Add-BpBox $Sheet 'ref-note' 20 ($y+480) 930 58 $note '000000' 12 -NoBorder
    End-BpPage $Sheet ($y+538)
}
function Add-BpNode($Sheet,$Node,[string]$Id,[double]$X,[double]$Y,[double]$W){
    $r=$Node.row;$formula='='+ (Bp-Label 'Resources' "C$r")+'&CHAR(10)&'+(Bp-Label 'Resources' "D$r")
    $relation=Get-BpRelation $Node
    $formula+='&CHAR(10)&"'+$relation.Replace('"','""')+'"'
    $shape=Add-BpBox $Sheet $Id $X $Y $W 68 $formula (Get-BpColor $Node.role) 11 -Formula
    $Sheet.Calculate();$height=[Math]::Max(68,[double]$shape.TextFrame2.TextRange.BoundHeight+12);$shape.Height=[double]$height
    return $height
}
function Invoke-BpRender($Book){
    $data=Get-BpInputs $Book;$model=New-BpModel $data;$sheet=$Book.Worksheets.Item('Blueprint view')
    # Use a fixed zoom and pixel-exact row height before placing shapes. Changing
    # zoom afterwards can round Excel's saved anchors across manual page breaks.
    $sheet.Activate();$Book.Application.ActiveWindow.Zoom=100
    $oldProtection=@{objects=[bool]$sheet.ProtectDrawingObjects;contents=[bool]$sheet.ProtectContents;scenarios=[bool]$sheet.ProtectScenarios}
    $unowned=@($sheet.Shapes|Where-Object {$_.Name -notlike 'bp_*'}|ForEach-Object {$_.Name})
    if(@($sheet.Shapes|Where-Object {$_.Name -like 'bp_*'}).Count -gt 0 -and [string]$sheet.Range('AE1').Value2 -ne 'VCF_BLUEPRINT_V1'){throw 'Blueprint shape prefix collides with an unmarked diagram.'}
    $sheet.Unprotect('');$sheet.Range('A1:AN52').UnMerge();$sheet.Range('A1:AN52').Clear()
    for($i=$sheet.Shapes.Count;$i-ge 1;$i--){if($sheet.Shapes.Item($i).Name -like 'bp_*'){$sheet.Shapes.Item($i).Delete()}}
    $sheet.Range('AA:AH').Clear();$sheet.Range('AA:AH').EntireColumn.Hidden=$false
    # Printable columns stay separate from hidden helpers: A:X is the canvas.
    $sheet.Range('A:X').ColumnWidth=7;$sheet.Range('1:1800').RowHeight=15
    $sheet.Range('A1:X1800').Interior.Color=16777215;$sheet.Range('A1:X1800').Borders.LineStyle=0
    $sheet.Range('AE1').Value2='VCF_BLUEPRINT_V1';$sheet.Range('AE2').Value2=[DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss UTC')
    $script:bpLabelRow=10;$script:bpPage=0;$script:bpNextRow=0;$script:bpPageStarts=[Collections.Generic.List[int]]::new();$script:bpRowHeight=[double]$sheet.Rows.Item(1).Height
    $n=$data.Count;$current=[object[,]]::new($n,1);$snapshot=[object[,]]::new($n,1);$labels=[object[,]]::new($n,1);$diff=[object[,]]::new($n,1);$i=0
    foreach($key in $data.Keys){
        $split=$key.LastIndexOf('!');$ref=Bp-Ref $key.Substring(0,$split) $key.Substring($split+1);$r=$i+2000
        $labels[$i,0]=$key;$current[$i,0]='=IF('+ $ref+'="","",'+$ref+')';$value=$data[$key];$snapshot[$i,0]=if($null-eq $value){''}elseif($value -is [string] -and $value -match '^[=+@-]'){"'"+$value}else{$value}
        $diff[$i,0]='=IF(EXACT(AB'+$r+'&"",AC'+$r+'&""),0,1)';$i++
    }
    $last=1999+$n
    Set-BpMatrix $sheet.Range("AA2000:AA$last") 'Value2' $labels
    Set-BpMatrix $sheet.Range("AB2000:AB$last") 'Formula' $current
    $sheet.Range("AC2000:AC$last").NumberFormat='@'
    Set-BpMatrix $sheet.Range("AC2000:AC$last") 'Value2' $snapshot
    Set-BpMatrix $sheet.Range("AD2000:AD$last") 'Formula' $diff
    $sheet.Range('B4').Formula='=IF(SUM(AD2000:AD'+$last+')=0,"Diagram layout current","Diagram layout needs refresh")';$sheet.Range('B4').Font.Color=16777215
    Add-BpReference $sheet $data
    $y=Start-BpPage $sheet '=Design!C7&" / "&Design!C8' 'Selected-instance inventory. Domain ownership and appliance hosting are separate relationships.'
    $regionBox=Add-BpBox $sheet 'actual-region' 20 $y 930 480 ('="Region: "&Design!C9&" / "&Design!C13') '000000' 13 -Dashed -Formula
    $regionHeading=Measure-BpText $sheet ('Region: '+[string](Get-BpValue $data 'Design' 'C9')+' / '+[string](Get-BpValue $data 'Design' 'C13')) 930 13 35
    $instanceBox=Add-BpBox $sheet 'actual-instance' 34 ($y+$regionHeading) 902 430 '=Design!C8&" / active domain and cluster inventory"' '000000' 14 -Formula
    $instanceHeading=Measure-BpText $sheet ([string](Get-BpValue $data 'Design' 'C8')+' / active domain and cluster inventory') 902 14 35
    $dy=$y+$regionHeading+$instanceHeading+8
    foreach($domain in $model.domains){
        $r=$domain.row;$domainBox=Add-BpBox $sheet ('domain-'+$domain.id) 48 $dy 874 200 ('='+ (Bp-Label 'Domains & clusters' "C$r")+'&" / '+$domain.id+'"') '0098C7' 13 -Formula
        $ny=$dy+(Measure-BpText $sheet ($domain.name+' / '+$domain.id) 874 13 32)
        $owned=@($model.clusters|Where-Object {$_.domain -eq $domain.id})
        foreach($cluster in $owned){
            $cr=$cluster.row;$f='='+ (Bp-Label 'Domains & clusters' "D$cr")+'&" / '+$cluster.id+' / "&'+(Bp-Label 'Domains & clusters' "E$cr")+'&" data AZ(s)"'
            $height=Measure-BpText $sheet ($cluster.name+' / '+$cluster.id+' / '+$cluster.azs+' data AZ(s)') 854 11 28
            $null=Add-BpBox $sheet ('summary-'+$cluster.id) 58 $ny 854 $height $f '0098C7' 11 -Formula;$ny+=$height+6
        }
        if($owned.Count -eq 0){$null=Add-BpBox $sheet ('no-clusters-'+$domain.id) 58 $ny 854 35 'No active cluster inventory for this domain.' '000000' 11 -NoBorder;$ny+=41}
        $domainBox.Height=[double]($ny-$dy+4);$dy=$ny+16
    }
    if($model.domains.Count -eq 0){$null=Add-BpBox $sheet 'no-domains' 50 $dy 850 70 'No active domain inventory. Complete Domains & clusters.' '000000';$dy+=80}
    $instanceBox.Height=[double]($dy-$instanceBox.Top);$regionBox.Height=[double]($dy-$y+14)
    $null=Add-BpBox $sheet 'inventory-scope' 20 ($dy+24) 930 50 'Cluster detail pages place appliances by hosted cluster. Other instances in this fleet are not inventoried by this file.' '000000' 12 -NoBorder
    End-BpPage $sheet ($dy+74)
    foreach($cluster in $model.clusters){
        $members=@($model.nodes|Where-Object {$_.hosted -eq $cluster.id -and -not $_.unresolved -and $_.role -ne 'vSAN Witness'})
        $sides=@(if($cluster.azs -eq 2){'AZ1';'AZ2'}else{'AZ1'});$sideData=@{}
        $nodeWidth=if($sides.Count -eq 2){418}else{886}
        foreach($side in $sides){$sideData[$side]=Group-BpNodes $sheet @($members|Where-Object {$_.location -eq $side}) $nodeWidth}
        $chunks=1;foreach($side in $sides){$chunks=[Math]::Max($chunks,$sideData[$side].batches.Count)}
        for($part=0;$part-lt $chunks;$part++){
            $cr=$cluster.row;$title='='+ (Bp-Label 'Domains & clusters' "D$cr")+'&" / '+$cluster.id+' / inventory '+($part+1)+' of '+$chunks+'"'
            $y=Start-BpPage $sheet $title ('Hosted in instance '+[string](Get-BpValue $data 'Design' 'C8')+'; owning domain '+$cluster.domain+'. Interface rows are deduplicated into nodes.')
            $null=Add-BpBox $sheet ('cluster-'+$cluster.id+'-'+$part) 20 $y 930 480 ('='+ (Bp-Label 'Domains & clusters' "I$cr")+'&" / "&'+(Bp-Label 'Domains & clusters' "H$cr")) '0098C7' 12 -Formula
            foreach($side in $sides){
                $si=if($side -eq 'AZ2'){1}else{0};$x=32+$si*456;$width=if($sides.Count -eq 2){438}else{906}
                $siteCell=if($si){"G$cr"}else{"F$cr"};$null=Add-BpBox $sheet ($cluster.id+'-'+$side+'-'+$part) $x ($y+35) $width 428 ('="'+$side+' / "&'+(Bp-Label 'Domains & clusters' $siteCell)) '000000' 12 -Dashed -Formula
                $ny=$y+70;$slice=@();if($part -lt $sideData[$side].batches.Count){$slice=@($sideData[$side].batches[$part])}
                foreach($node in $slice){$height=Add-BpNode $sheet $node ('node-'+$node.id) ($x+10) $ny ($width-20);$ny+=$height+5}
                if($slice.Count -eq 0){
                    $emptyText=if($part -gt 0){'No additional nodes on this page.'}else{'No active nodes with resolved hosting in this AZ.'}
                    $null=Add-BpBox $sheet ('empty-'+$cluster.id+'-'+$side+'-'+$part) ($x+10) $ny ($width-20) 60 $emptyText '000000' 11 -NoBorder
                }
                if($ny -gt $y+459){throw 'Diagram node text exceeds its page. Shorten the label or use a smaller page batch.'}
            }
            $witness=@($model.nodes|Where-Object {$_.role -eq 'vSAN Witness' -and $_.managed -eq $cluster.id})
            $note=if($witness.Count){'Witness outside data AZs: '+(($witness|ForEach-Object {$_.fqdn}) -join ', ')}elseif($cluster.azs -eq 2 -and (Get-BpValue $data 'Domains & clusters' "I$cr") -eq 'vSAN HCI'){'Witness site: '+[string](Get-BpValue $data 'Domains & clusters' "N$cr")+'; no explicitly bound witness resource.'}else{'Storage and virtual networking follow this cluster''s recorded design.'}
            $null=Add-BpBox $sheet ('witness-note-'+$cluster.id+'-'+$part) 20 ($y+492) 930 50 $note '000000' 11 -NoBorder
            End-BpPage $sheet ($y+542)
        }
    }
    $other=@($model.nodes|Where-Object {-not $_.hosted -or $_.unresolved -or $_.role -eq 'vSAN Witness'})
    for($start=0;$start-lt [Math]::Max(1,$other.Count);$start+=6){
        $y=Start-BpPage $sheet '=Design!C8&" / shared, external and unresolved inventory"' 'These objects are outside the cluster placement diagrams. Shared services are references, not additional local deployments.'
        $slice=@($other|Select-Object -Skip $start -First 6);$j=0;$maxHeight=0;$ny=$y
        foreach($node in $slice){$x=20+($j%2)*472;$height=Add-BpNode $sheet $node ('external-'+$node.id) $x $ny 458;$maxHeight=[Math]::Max($maxHeight,$height);if($j%2 -eq 1){$ny+=$maxHeight+8;$maxHeight=0};$j++}
        if($slice.Count -eq 0){$null=Add-BpBox $sheet 'no-external' 20 $y 930 70 'No active external, shared or unresolved resource inventory.' '000000' 12 -NoBorder}
        End-BpPage $sheet ($ny+$maxHeight+70)
    }
    if($model.issues.Count){
        foreach($start in (0..([Math]::Floor(($model.issues.Count-1)/12)))){
            $y=Start-BpPage $sheet '="Inventory items requiring review"' 'The drawing retains unresolved source data explicitly; it does not infer missing placement or enabled products.'
            $text=(@($model.issues|Select-Object -Skip ($start*12) -First 12) -join "`n`n")
            $null=Add-BpBox $sheet ('issues-'+$start) 20 $y 930 500 $text '000000' 12 -NoBorder
            End-BpPage $sheet ($y+500)
        }
    }
    $lastPageRow=$script:bpNextRow;$sheet.Range('AA:AH').EntireColumn.Hidden=$true;$sheet.Cells.Locked=$true
    $sheet.ResetAllPageBreaks();foreach($r in @($script:bpPageStarts|Select-Object -Skip 1)){$null=$sheet.HPageBreaks.Add($sheet.Cells.Item($r,1))}
    $page=$sheet.PageSetup;$page.PrintArea='A1:X'+$lastPageRow;$page.Orientation=2;$page.PaperSize=8;$page.Zoom=$false;$page.FitToPagesWide=1;$page.FitToPagesTall=$false
    $page.LeftMargin=12;$page.RightMargin=12;$page.TopMargin=12;$page.BottomMargin=12;$page.PrintGridlines=$false;$page.PrintHeadings=$false
    $sheet.Activate();$Book.Application.ActiveWindow.DisplayGridlines=$false;$sheet.Range('A1').Select()
    $sheet.Calculate();if([string]$sheet.Range('B4').Value2 -ne 'Diagram layout current'){
        $errors=@();$check=$sheet.Range("AA2000:AD$last").Value2
        for($q=1;$q-le $n;$q++){if($check[$q,4] -ne 0){$errors+=([string]$check[$q,1]+': current='+[string]$check[$q,2]+' snapshot='+[string]$check[$q,3]);if($errors.Count -ge 8){break}}}
        throw ('Diagram structural snapshot differs immediately after rendering: '+($errors -join '; '))
    }
    $sheet.Protect('',$true,$true,$true)
    foreach($name in $unowned){if(@($sheet.Shapes|Where-Object {$_.Name -ceq $name}).Count -ne 1){throw 'A user annotation was lost.'}}
    $count=@($sheet.Shapes|Where-Object {$_.Name -like 'bp_*'}).Count
    return [ordered]@{schema='VCF_BLUEPRINT_V1';pages=$script:bpPage;shapeCount=$count;labelFormulas=$script:bpLabelRow-10;trackedCells=$n;model=$model;status='Diagram layout current';nativeShapes=$true;unownedShapesPreserved=$unowned.Count}
}

function Update-BlueprintDiagram($Book){
    # Excel automation methods may return incidental COM values. Only the
    # renderer's explicit verification record is part of this API.
    $records=@(Invoke-BpRender $Book | Where-Object {$_ -is [Collections.IDictionary] -and $_['schema'] -eq 'VCF_BLUEPRINT_V1'})
    if($records.Count -ne 1){throw 'Diagram renderer did not produce one verification record.'}
    return $records[0]
}

function Test-BlueprintDiagram($Book,$Expected){
    $sheet=$Book.Worksheets.Item('Blueprint view')
    if([string]$sheet.Range('AE1').Value2 -ne 'VCF_BLUEPRINT_V1' -or [string]$sheet.Range('B4').Value2 -ne 'Diagram layout current'){throw 'Saved diagram marker or freshness check failed.'}
    $shapes=@($sheet.Shapes|Where-Object {$_.Name -like 'bp_*'})
    $helpers=$sheet.Range(('AF11:AF'+(10+$Expected.labelFormulas))).Value2
    for($i=1;$i-le $helpers.GetLength(0);$i++){
        if($helpers[$i,1] -is [Runtime.InteropServices.ErrorWrapper] -or ($helpers[$i,1] -is [int] -and $helpers[$i,1] -lt -2000000000)){throw ('Diagram label formula error at AF'+($i+10))}
    }
    if($shapes.Count -ne $Expected.shapeCount -or @($shapes.Name|Sort-Object -Unique).Count -ne $shapes.Count){throw 'Saved diagram shape IDs differ or are duplicated.'}
    if(-not $sheet.ProtectContents -or -not $sheet.ProtectDrawingObjects -or -not $sheet.Range('AF11').Locked){throw 'Saved diagram protection failed.'}
    foreach($shape in $shapes){
        if($shape.Left -lt 0 -or $shape.Left+$shape.Width -gt $sheet.Range('A:X').Width+1){throw ('Diagram shape exceeds printable width: '+$shape.Name)}
        if($shape.Connector -eq -1){continue}
        try {$boundHeight=[double]$shape.TextFrame2.TextRange.BoundHeight}
        catch {throw ('Cannot measure saved diagram text: '+$shape.Name+'; '+$_.Exception.Message)}
        if($boundHeight -gt [double]$shape.Height-4){throw ('Diagram text is clipped: '+$shape.Name)}
    }
    $pageEnds=@($sheet.HPageBreaks|ForEach-Object {[double]$_.Location.Top}|Sort-Object)
    $pageEnds+=([double]$sheet.Range($sheet.PageSetup.PrintArea).Top+[double]$sheet.Range($sheet.PageSetup.PrintArea).Height)
    foreach($shape in $shapes){
        if($shape.Name -match '^bp_page(\d+)-status$'){
            $pageNumber=[int]$Matches[1]
            if($pageNumber -gt $pageEnds.Count -or $shape.Top+$shape.Height -gt $pageEnds[$pageNumber-1]){throw ('Saved diagram footer crosses a printed page: '+$shape.Name)}
        }
    }
    return $true
}
