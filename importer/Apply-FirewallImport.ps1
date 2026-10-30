#requires -Version 7.0
<#
.SYNOPSIS
Applies an explicit, literal-value import plan to a new firewall planner copy.
.DESCRIPTION
Plan JSON: {"clearInputs":true,"clearCells":[{"sheet":"Resources","cell":"B9"}],
            "edits":[{"sheet":"Resources","cell":"B9","value":"host-01"}]}
clearInputs=true requires an explicit nonempty clearCells list. The script never
discovers cells to erase from their current values. Overlap between clearCells
and edits is intentional: clears run first. Every target must be one unmerged,
unlocked, non-formula cell on an allowed input sheet. Existing outputs are refused.
Only a newly owned Excel process is used. Nothing is written to the template.
The script emits one JSON result without imported values. It does not approve
designs, mark ports accepted, or infer missing source information.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TemplatePath,
    [Parameter(Mandatory)][string]$PlanPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [switch]$IncludeFormulaDiagnostics
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:FailureCode = 'UNEXPECTED_FAILURE'
$script:FailureMessage = 'Import failed.'
$stage = 'validate_paths'
$currentSheet = $null
$currentRange = $null
$calculationStateAfterRebuild = $null
$calculationStateAfterAdditionalCalculate = $null
$calculationStateBeforeSave = $null
$calculationStateAfterReopen = $null
$additionalCalculateRequested = $false
$snapshotDifference = $null
$beforeFormulaMap = $null
$afterFormulaMap = $null
if ($IncludeFormulaDiagnostics) {
    $beforeFormulaMap = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
    $afterFormulaMap = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
}
$excel = $null
$book = $null
$ownedExcel = $false
$ownedPid = 0
$ownedStartTime = $null
$temporaryPath = $null
$success = $false
$result = $null
$started = [DateTime]::UtcNow
$allowedSheets = @('Design', 'Domains & clusters', 'Deployment', 'Network plan', 'Resources')

function Fail-Import([string]$Code, [string]$Message) {
    $script:FailureCode = $Code
    $script:FailureMessage = $Message
    throw [InvalidOperationException]::new($Message)
}

function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Assert-CellSpec($Spec) {
    if ($Spec -isnot [System.Collections.IDictionary] -or
        -not $Spec.Contains('sheet') -or -not $Spec.Contains('cell') -or
        $Spec.sheet -isnot [string] -or $Spec.cell -isnot [string]) {
        Fail-Import 'INVALID_CELL_SPEC' 'Every target must provide string sheet and cell fields.'
    }
    if ($Spec.sheet -cnotin $allowedSheets) {
        Fail-Import 'SHEET_NOT_ALLOWED' 'The plan targets a sheet outside the supported input schema.'
    }
    if ($Spec.cell -cnotmatch '^[A-Za-z]{1,3}[1-9][0-9]{0,6}$') {
        Fail-Import 'INVALID_CELL_ADDRESS' 'Targets must be individual A1 cell addresses without dollar signs.'
    }
    $address = $Spec.cell.ToUpperInvariant()
    $letters = $address -replace '[0-9]', ''
    $column = 0
    foreach ($character in $letters.ToCharArray()) { $column = 26 * $column + [int]$character - 64 }
    $row = [int]($address -replace '[A-Z]', '')
    if ($column -gt 16384 -or $row -gt 1048576) {
        Fail-Import 'INVALID_CELL_ADDRESS' 'A target is outside Excel worksheet bounds.'
    }
    return @{ sheet = $Spec.sheet; cell = $address; row = $row; column = $column }
}

function Assert-LiteralValue($Value) {
    if ($null -eq $Value) { return }
    if ($Value -is [string]) {
        if ($Value.Length -gt 32767 -or $Value.IndexOf([char]0) -ge 0) {
            Fail-Import 'INVALID_TEXT_VALUE' 'An input string exceeds Excel limits or contains a null character.'
        }
        # Value2 can interpret formula-like text. Refuse it rather than changing
        # cell formatting on a protected worksheet or relying on quote prefixes.
        if ($Value -match '^[\s\x00-\x1f]*[=+@-]') {
            Fail-Import 'FORMULA_LIKE_TEXT' 'A string begins with a formula-like prefix. Supply literal source data.'
        }
        return
    }
    if ($Value -is [bool]) { return }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or
        $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        if ([double]::IsNaN([double]$Value) -or [double]::IsInfinity([double]$Value)) {
            Fail-Import 'INVALID_NUMBER' 'Numeric inputs must be finite.'
        }
        return
    }
    Fail-Import 'INVALID_VALUE_TYPE' 'Values must be JSON strings, finite numbers, booleans or null.'
}

function Get-CellKey($Spec) { return $Spec.sheet + [char]31 + $Spec.cell }

function Get-InputBatches([Collections.IEnumerable]$Specs) {
    # Excel accepts short unions on a protected sheet when every member is an
    # unlocked input. Bound both member count and address length explicitly.
    foreach ($group in ($Specs | Group-Object -Property sheet)) {
        $addresses = [Collections.Generic.List[string]]::new()
        $length = 0
        foreach ($spec in $group.Group) {
            $nextLength = $length + $spec.cell.Length + $(if ($addresses.Count -gt 0) { 1 } else { 0 })
            if ($addresses.Count -ge 20 -or $nextLength -gt 240) {
                [pscustomobject]@{ sheet = $group.Name; address = $addresses -join ','; cells = $addresses.ToArray(); count = $addresses.Count }
                $addresses.Clear()
                $length = 0
            }
            $length += $spec.cell.Length + $(if ($addresses.Count -gt 0) { 1 } else { 0 })
            $addresses.Add($spec.cell)
        }
        if ($addresses.Count -gt 0) {
            [pscustomobject]@{ sheet = $group.Name; address = $addresses -join ','; cells = $addresses.ToArray(); count = $addresses.Count }
        }
    }
}

function Test-SameLiteral($Expected, $Actual) {
    if ($null -eq $Expected -or ($Expected -is [string] -and $Expected.Length -eq 0)) {
        return $null -eq $Actual -or ($Actual -is [string] -and $Actual.Length -eq 0)
    }
    if ($Expected -is [string]) { return $Actual -is [string] -and $Expected -ceq $Actual }
    if ($Expected -is [bool]) { return $Actual -is [bool] -and $Expected -eq $Actual }
    if ($null -eq $Actual -or $Actual -is [string] -or $Actual -is [bool]) { return $false }
    return [double]$Expected -eq [double]$Actual
}

function Wait-ActiveCalculation($Application, [DateTime]$Deadline) {
    $state = [int]$Application.CalculationState
    # xlCalculating=1 means work is running. xlPending=2 may remain after the
    # synchronous calculation calls return; record it rather than treating it
    # as an indefinitely running calculation. Result correctness is checked by
    # complete formula/error scans, saved input checks and separate native QA.
    while ($state -eq 1) {
        if ([DateTime]::UtcNow -gt $Deadline) {
            Fail-Import 'CALCULATION_TIMEOUT' ('Excel remained actively calculating for five minutes. CalculationState=' + $state + '.')
        }
        Start-Sleep -Milliseconds 200
        $state = [int]$Application.CalculationState
    }
    if ($state -notin @(0, 2)) {
        Fail-Import 'UNKNOWN_CALCULATION_STATE' ('Excel returned an unexpected CalculationState=' + $state + '.')
    }
    return $state
}

function Get-WorkbookSnapshot($Workbook, $FormulaMap = $null) {
    $capturedFormulas = $FormulaMap
    if ($null -eq $capturedFormulas) {
        $capturedFormulas = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
    }
    $sheetStates = [Collections.Generic.List[object]]::new()
    $tableStates = [Collections.Generic.List[string]]::new()
    $nameStates = [Collections.Generic.List[string]]::new()
    $formulaCount = 0
    try {
        for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
            $sheet = $Workbook.Worksheets.Item($index)
            $used = $null
            try {
                $sheetName = [string]$sheet.Name
                $sheetStates.Add([ordered]@{
                    name = $sheetName; visible = [int]$sheet.Visible
                    contents = [bool]$sheet.ProtectContents; drawings = [bool]$sheet.ProtectDrawingObjects
                    scenarios = [bool]$sheet.ProtectScenarios
                })
                $used = $sheet.UsedRange
                $formulaValues = $used.Formula
                for ($readRetry = 0; [int64]$used.CountLarge -gt 1 -and $formulaValues -isnot [Array] -and $readRetry -lt 3; $readRetry++) {
                    Start-Sleep -Milliseconds 300
                    $formulaValues = $used.Formula
                }
                if ([int64]$used.CountLarge -gt 1 -and ($formulaValues -isnot [Array] -or $formulaValues.Rank -ne 2)) {
                    Fail-Import 'INCOMPLETE_FORMULA_READ' 'Excel did not return the complete formula array for a worksheet.'
                }
                $firstRow = [int]$used.Row
                $firstColumn = [int]$used.Column
                if ($formulaValues -is [Array] -and $formulaValues.Rank -eq 2) {
                    $rowLow = $formulaValues.GetLowerBound(0)
                    $colLow = $formulaValues.GetLowerBound(1)
                    for ($r = $rowLow; $r -le $formulaValues.GetUpperBound(0); $r++) {
                        for ($c = $colLow; $c -le $formulaValues.GetUpperBound(1); $c++) {
                            $formula = $formulaValues.GetValue($r, $c)
                            if ($formula -is [string] -and $formula.StartsWith('=')) {
                                $key = $sheetName + [char]31 + ($firstRow + $r - $rowLow) + ':' + ($firstColumn + $c - $colLow)
                                $capturedFormulas[$key] = [string]::new($formula.ToCharArray())
                                $formulaCount++
                            }
                        }
                    }
                }
                elseif ($formulaValues -is [string] -and $formulaValues.StartsWith('=')) {
                    $capturedFormulas[$sheetName + [char]31 + $firstRow + ':' + $firstColumn] = [string]::new($formulaValues.ToCharArray())
                    $formulaCount++
                }
                for ($t = 1; $t -le $sheet.ListObjects.Count; $t++) {
                    $table = $sheet.ListObjects.Item($t)
                    $range = $null
                    try {
                        $range = $table.Range
                        $tableStates.Add($sheetName + '!' + [string]$table.Name + ':' + [string]$range.Address())
                    }
                    finally { Release-Com $range; Release-Com $table }
                }
            }
            finally { Release-Com $used; Release-Com $sheet }
        }
        for ($n = 1; $n -le $Workbook.Names.Count; $n++) {
            $name = $Workbook.Names.Item($n)
            try { $nameStates.Add([string]$name.Name + [char]31 + [string]$name.RefersTo + [char]31 + [string]$name.Visible) }
            finally { Release-Com $name }
        }
        if ($capturedFormulas.Count -ne $formulaCount) {
            Fail-Import 'FORMULA_ADDRESS_COVERAGE' 'Formula address-map count differs from the complete scan.'
        }
        return [ordered]@{
            # Hash complete materialized address/formula pairs in ordinal key
            # order. No formula normalization or excluded cells are permitted.
            formulaHash = Get-FormulaMapDigest $capturedFormulas -SortKeys
            formulaCount = $formulaCount
            sheetStates = @($sheetStates.ToArray())
            tables = @($tableStates.ToArray() | Sort-Object)
            names = @($nameStates.ToArray() | Sort-Object)
            structureProtected = [bool]$Workbook.ProtectStructure
            windowsProtected = [bool]$Workbook.ProtectWindows
        }
    }
    finally { if ($null -eq $FormulaMap) { $capturedFormulas.Clear() } }
}

function Get-FormulaErrorCount($Workbook, [int]$ExpectedFormulaCount) {
    $count = 0
    $scannedFormulas = 0
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $sheet = $Workbook.Worksheets.Item($index)
        $used = $null
        try {
            $used = $sheet.UsedRange
            $formulas = $used.Formula
            $values = $used.Value2
            for ($readRetry = 0; [int64]$used.CountLarge -gt 1 -and ($formulas -isnot [Array] -or $values -isnot [Array]) -and $readRetry -lt 3; $readRetry++) {
                Start-Sleep -Milliseconds 300
                $formulas = $used.Formula
                $values = $used.Value2
            }
            if ([int64]$used.CountLarge -gt 1 -and ($formulas -isnot [Array] -or $values -isnot [Array] -or
                $formulas.Rank -ne 2 -or $values.Rank -ne 2 -or
                $formulas.GetLength(0) -ne $values.GetLength(0) -or $formulas.GetLength(1) -ne $values.GetLength(1))) {
                Fail-Import 'INCOMPLETE_ERROR_SCAN' 'Excel did not return complete formula/value arrays for error scanning.'
            }
            # SpecialCells can return Excel 1004 on a protected sheet even when
            # formula errors exist. Scan the same bulk Formula/Value2 arrays as
            # the separately sentinel-tested native QA harness instead.
            if ($formulas -is [Array] -and $formulas.Rank -eq 2) {
                for ($r = $formulas.GetLowerBound(0); $r -le $formulas.GetUpperBound(0); $r++) {
                    for ($c = $formulas.GetLowerBound(1); $c -le $formulas.GetUpperBound(1); $c++) {
                        $formula = $formulas.GetValue($r, $c)
                        if ($formula -is [string] -and $formula.StartsWith('=')) {
                            $scannedFormulas++
                            $value = $values.GetValue($r, $c)
                            if (($value -is [int] -and $value -lt -2146820000 -and $value -gt -2146830000) -or
                                $value -is [Runtime.InteropServices.ErrorWrapper]) { $count++ }
                        }
                    }
                }
            }
            elseif ($formulas -is [string] -and $formulas.StartsWith('=')) {
                $scannedFormulas++
                if (($values -is [int] -and $values -lt -2146820000 -and $values -gt -2146830000) -or
                    $values -is [Runtime.InteropServices.ErrorWrapper]) { $count++ }
            }
        }
        finally { Release-Com $used; Release-Com $sheet }
    }
    if ($ExpectedFormulaCount -le 0 -or $scannedFormulas -ne $ExpectedFormulaCount) {
        Fail-Import 'ERROR_SCAN_COVERAGE' 'Error-scan formula coverage differs from the full formula snapshot.'
    }
    return $count
}

function Get-SnapshotDifference($Before, $After) {
    $changed = [Collections.Generic.List[string]]::new()
    foreach ($key in $Before.Keys) {
        if (($Before[$key] | ConvertTo-Json -Depth 10 -Compress) -cne ($After[$key] | ConvertTo-Json -Depth 10 -Compress)) {
            $changed.Add([string]$key)
        }
    }
    $difference = [ordered]@{
        changedKeys = $changed.ToArray()
        before = [ordered]@{
            formulaHash = $Before.formulaHash; formulaCount = $Before.formulaCount
            sheetCount = @($Before.sheetStates).Count; nameCount = @($Before.names).Count; tableCount = @($Before.tables).Count
            structureProtected = $Before.structureProtected; windowsProtected = $Before.windowsProtected
        }
        after = [ordered]@{
            formulaHash = $After.formulaHash; formulaCount = $After.formulaCount
            sheetCount = @($After.sheetStates).Count; nameCount = @($After.names).Count; tableCount = @($After.tables).Count
            structureProtected = $After.structureProtected; windowsProtected = $After.windowsProtected
        }
    }
    if ('names' -in $changed) {
        $oldNames = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
        $newNames = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
        foreach ($record in $Before.names) {
            $parts = $record.Split([char]31)
            $oldNames[$parts[0]] = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($record)))
        }
        foreach ($record in $After.names) {
            $parts = $record.Split([char]31)
            $newNames[$parts[0]] = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($record)))
        }
        $difference.names = [ordered]@{
            removedIdentifiers = @($oldNames.Keys | Where-Object { -not $newNames.ContainsKey($_) } | Sort-Object)
            addedIdentifiers = @($newNames.Keys | Where-Object { -not $oldNames.ContainsKey($_) } | Sort-Object)
            changedBindings = @($oldNames.Keys | Where-Object { $newNames.ContainsKey($_) -and $oldNames[$_] -cne $newNames[$_] } | Sort-Object | ForEach-Object {
                [ordered]@{identifier=$_; beforeHash=$oldNames[$_]; afterHash=$newNames[$_]}
            })
        }
    }
    if ('tables' -in $changed) {
        # These contain only sheet/table identifiers and range addresses.
        $difference.tables = [ordered]@{ before = $Before.tables; after = $After.tables }
    }
    if ('sheetStates' -in $changed) {
        $difference.sheetStates = [ordered]@{ before = $Before.sheetStates; after = $After.sheetStates }
    }
    return $difference
}

function Get-FormulaMapDigest($Map, [switch]$SortKeys) {
    $keys = [Collections.Generic.List[string]]::new()
    $keys.AddRange($Map.Keys)
    if ($SortKeys) { $keys.Sort([StringComparer]::Ordinal) }
    $digest = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
    try {
        foreach ($key in $keys) {
            $record = $key + [char]31 + $Map[$key] + [char]30
            $digest.AppendData([Text.Encoding]::UTF8.GetBytes($record))
        }
        return [Convert]::ToHexString($digest.GetHashAndReset())
    }
    finally { $digest.Dispose() }
}

function Get-FormulaDifferences($Before, $After) {
    $changed = [Collections.Generic.List[string]]::new()
    $removed = [Collections.Generic.List[string]]::new()
    $added = [Collections.Generic.List[string]]::new()
    foreach ($key in $Before.Keys) {
        if (-not $After.ContainsKey($key)) { $removed.Add($key) }
        elseif (-not [string]::Equals($Before[$key], $After[$key], [StringComparison]::Ordinal)) { $changed.Add($key) }
    }
    foreach ($key in $After.Keys) { if (-not $Before.ContainsKey($key)) { $added.Add($key) } }
    $samples = [Collections.Generic.List[object]]::new()
    foreach ($key in ($changed | Select-Object -First 3)) {
        $parts = $key.Split([char]31)
        $position = $parts[1].Split(':')
        $column = [int]$position[1]
        $letters = ''
        while ($column -gt 0) {
            $remainder = ($column - 1) % 26
            $letters = [string][char](65 + $remainder) + $letters
            $column = [int][Math]::Floor(($column - 1) / 26)
        }
        $old = $Before[$key]
        $new = $After[$key]
        $samples.Add([ordered]@{
            sheet = $parts[0]; cell = $letters + $position[0]
            beforeLength = $old.Length; afterLength = $new.Length
            before = $old.Substring(0, [Math]::Min(500, $old.Length))
            after = $new.Substring(0, [Math]::Min(500, $new.Length))
        })
    }
    return [ordered]@{
        beforeMapCount=$Before.Count;afterMapCount=$After.Count
        changedCount=$changed.Count;removedCount=$removed.Count;addedCount=$added.Count
        beforeMapOrderHash=Get-FormulaMapDigest $Before
        afterMapOrderHash=Get-FormulaMapDigest $After
        beforeMapSortedHash=Get-FormulaMapDigest $Before -SortKeys
        afterMapSortedHash=Get-FormulaMapDigest $After -SortKeys
        firstDifferences=$samples.ToArray()
    }
}

function Get-NativeCellProtection($Path) {
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        function Read-EntryXml($Entry) {
            $stream = $Entry.Open()
            try { $xml = [Xml.XmlDocument]::new(); $xml.Load($stream); return ,$xml }
            finally { $stream.Dispose() }
        }
        $styles = Read-EntryXml $archive.GetEntry('xl/styles.xml')
        $stylesNs = [Xml.XmlNamespaceManager]::new($styles.NameTable)
        $stylesNs.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $xfs = $styles.SelectNodes('/s:styleSheet/s:cellXfs/s:xf', $stylesNs)
        $unlocked = [Collections.Generic.List[string]]::new()
        $unlockedFormulas = 0
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notmatch '^xl/worksheets/sheet[0-9]+\.xml$') { continue }
            $xml = Read-EntryXml $entry
            $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach ($cell in $xml.SelectNodes('/s:worksheet/s:sheetData/s:row/s:c', $ns)) {
                $styleId = if ($cell.HasAttribute('s')) { [int]$cell.GetAttribute('s') } else { 0 }
                $protection = $xfs[$styleId].SelectSingleNode('s:protection', $stylesNs)
                if ($null -ne $protection -and $protection.GetAttribute('locked') -in @('0', 'false')) {
                    $unlocked.Add($entry.FullName + '!' + $cell.GetAttribute('r'))
                    if ($null -ne $cell.SelectSingleNode('s:f', $ns)) { $unlockedFormulas++ }
                }
            }
        }
        return @{ unlockedCells = @($unlocked.ToArray() | Sort-Object); unlockedFormulas = $unlockedFormulas }
    }
    finally { $archive.Dispose() }
}

try {
    $template = (Get-Item -LiteralPath $TemplatePath -ErrorAction Stop).FullName
    $planFile = (Get-Item -LiteralPath $PlanPath -ErrorAction Stop).FullName
    $output = [IO.Path]::GetFullPath($OutputPath)
    if ([IO.Path]::GetExtension($template) -ine '.xlsx' -or [IO.Path]::GetExtension($output) -ine '.xlsx') {
        Fail-Import 'XLSX_REQUIRED' 'Template and output must be .xlsx files.'
    }
    if ($output -ieq $template -or $output -ieq $planFile -or (Test-Path -LiteralPath $output)) {
        Fail-Import 'OUTPUT_EXISTS' 'Output must be a new path distinct from the template and plan.'
    }
    $templateHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash
    $stage = 'validate_plan'
    $plan = Get-Content -LiteralPath $planFile -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable -Depth 20
    if ($plan -isnot [Collections.IDictionary] -or -not $plan.Contains('edits') -or $plan.edits -isnot [Array]) {
        Fail-Import 'INVALID_PLAN' 'Plan must be an object containing an edits array.'
    }
    $clearInputs = $plan.Contains('clearInputs') -and $plan.clearInputs -eq $true
    if ($plan.Contains('clearInputs') -and $plan.clearInputs -isnot [bool]) {
        Fail-Import 'INVALID_CLEAR_FLAG' 'clearInputs must be a JSON boolean.'
    }
    $clearSpecs = @()
    if ($plan.Contains('clearCells')) {
        if ($plan.clearCells -isnot [Array]) {
            Fail-Import 'INVALID_CLEAR_LIST' 'clearCells must be a JSON array.'
        }
        $clearSpecs = @($plan.clearCells)
    }
    if ($clearInputs -and $clearSpecs.Count -eq 0) {
        Fail-Import 'EXPLICIT_CLEARS_REQUIRED' 'clearInputs=true requires a nonempty explicit clearCells array.'
    }
    $targets = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    $clears = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    $edits = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach ($spec in $clearSpecs) {
        $cellSpec = Assert-CellSpec $spec
        $key = Get-CellKey $cellSpec
        $clears[$key] = $cellSpec
        $targets[$key] = $cellSpec
    }
    foreach ($spec in $plan.edits) {
        $cellSpec = Assert-CellSpec $spec
        if (-not $spec.Contains('value')) { Fail-Import 'MISSING_VALUE' 'Every edit must include value, including explicit null for a clear.' }
        Assert-LiteralValue $spec.value
        $key = Get-CellKey $cellSpec
        if ($edits.ContainsKey($key)) { Fail-Import 'DUPLICATE_EDIT' 'The edit list contains a repeated sheet/cell target.' }
        $cellSpec.value = $spec.value
        $edits.Add($key, $cellSpec)
        $targets[$key] = $cellSpec
    }
    if ($targets.Count -eq 0) { Fail-Import 'EMPTY_PLAN' 'The import plan contains no cell changes.' }
    $baselineNative = Get-NativeCellProtection $template
    if ($baselineNative.unlockedFormulas -ne 0) { Fail-Import 'UNLOCKED_SOURCE_FORMULA' 'The template contains an unlocked formula.' }
    $parentDirectory = [IO.Path]::GetDirectoryName($output)
    [void][IO.Directory]::CreateDirectory($parentDirectory)
    $temporaryPath = Join-Path $parentDirectory ('.firewall-import-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
    [IO.File]::Copy($template, $temporaryPath, $false)
    # A reusable template may carry the filesystem read-only attribute. Clear
    # that attribute on this new temporary copy only, never on the source.
    (Get-Item -LiteralPath $temporaryPath).IsReadOnly = $false

    $stage = 'start_owned_excel'
    if ($null -eq ('FirewallImportNativeProcess' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FirewallImportNativeProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
    }
    $existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcess = 0
    [void][FirewallImportNativeProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcess)
    if ($excelProcess -eq 0 -or $excelProcess -in $existingExcelIds) {
        Fail-Import 'EXCEL_OWNERSHIP_UNPROVEN' 'Excel did not provide a separate process. Existing user Excel instances were not changed.'
    }
    $ownedPid = [int]$excelProcess
    $ownedStartTime = (Get-Process -Id $ownedPid).StartTime
    $ownedExcel = $true
    if ($excel.Workbooks.Count -ne 0) { Fail-Import 'EXCEL_NOT_EMPTY' 'The new Excel process unexpectedly contains an open workbook.' }
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $stage = 'open_temporary_copy'
    $book = $excel.Workbooks.Open($temporaryPath, 0, $false)
    $excel.Calculation = -4135
    $baseline = Get-WorkbookSnapshot $book $beforeFormulaMap
    if ($baseline.formulaCount -le 0) { Fail-Import 'NO_FORMULAS_READ' 'The template formula snapshot is empty.' }
    if (-not $baseline.structureProtected -or @($baseline.sheetStates | Where-Object { -not $_.contents }).Count -gt 0) {
        Fail-Import 'SOURCE_PROTECTION_MISSING' 'Template workbook structure and every worksheet must be protected.'
    }

    $stage = 'preflight_all_targets'
    foreach ($group in (Get-InputBatches $targets.Values | Group-Object -Property sheet)) {
        $sheet = $book.Worksheets.Item($group.Name)
        try {
            foreach ($batch in $group.Group) {
                $currentSheet = $group.Name
                $currentRange = $batch.address
                $range = $null
                try {
                    $range = $sheet.Range($batch.address)
                    $locked = $range.Locked
                    $hasFormula = $range.HasFormula
                    $merged = $range.MergeCells
                    if ([int64]$range.CountLarge -ne $batch.count -or
                        $locked -isnot [bool] -or $locked -or $hasFormula -isnot [bool] -or $hasFormula) {
                        Fail-Import 'PROTECTED_OR_FORMULA_TARGET' ('Cannot write the input batch on ' + $group.Name + ': a target is locked, contains a formula or is ambiguous.')
                    }
                    if ($merged -is [bool] -and $merged) {
                        Fail-Import 'MERGED_TARGET' ('Cannot write a merged target on ' + $group.Name + '.')
                    }
                    if ($merged -isnot [bool]) {
                        # Some Excel builds return mixed/null MergeCells for a
                        # discontiguous union. Resolve only that property per cell.
                        foreach ($address in $batch.cells) {
                            $cell = $null
                            try {
                                $cell = $sheet.Range($address)
                                if ($cell.MergeCells -isnot [bool] -or [bool]$cell.MergeCells) {
                                    Fail-Import 'MERGED_TARGET' ('Cannot write a merged or ambiguous target on ' + $group.Name + '.')
                                }
                            }
                            finally { Release-Com $cell }
                        }
                    }
                }
                finally { Release-Com $range }
            }
        }
        finally { Release-Com $sheet }
    }
    $stage = 'clear_explicit_inputs'
    foreach ($group in (Get-InputBatches $clears.Values | Group-Object -Property sheet)) {
        $sheet = $book.Worksheets.Item($group.Name)
        try {
            foreach ($batch in $group.Group) {
                $currentSheet = $group.Name
                $currentRange = $batch.address
                $range = $null
                try { $range = $sheet.Range($batch.address); [void]$range.ClearContents() }
                finally { Release-Com $range }
            }
        }
        finally { Release-Com $sheet }
    }
    $stage = 'apply_literal_edits'
    foreach ($group in ($edits.Values | Group-Object -Property sheet)) {
        $sheet = $book.Worksheets.Item($group.Name)
        try {
            foreach ($spec in $group.Group) {
                $currentSheet = $group.Name
                $currentRange = $spec.cell
                $cell = $null
                try {
                    # Keep the native Value2 assignment that passed the first
                    # integration run. Batching preflight/clears and bulk saved
                    # verification remove the larger cross-process overhead.
                    $cell = $sheet.Range($spec.cell)
                    if ($null -eq $spec.value) { [void]$cell.ClearContents() }
                    elseif ($spec.value -is [string] -or $spec.value -is [bool]) { $cell.Value2 = $spec.value }
                    else { $cell.Value2 = [double]$spec.value }
                    $hasFormula = $cell.HasFormula
                    if ($hasFormula -isnot [bool] -or $hasFormula) { Fail-Import 'FORMULA_CREATED' 'An imported value was interpreted as a formula.' }
                }
                finally { Release-Com $cell }
            }
        }
        finally { Release-Com $sheet }
    }

    $stage = 'recalculate_and_filter'
    $currentSheet = $null
    $currentRange = $null
    # Manual mode can leave CalculationState at xlPending after a rebuild.
    # Re-enter automatic mode before requesting the complete calculation.
    $excel.Calculation = -4105
    [void]$excel.CalculateFullRebuild()
    $calculationStateAfterRebuild = [int]$excel.CalculationState
    $settledState = Wait-ActiveCalculation $excel ([DateTime]::UtcNow.AddMinutes(5))
    if ($settledState -eq 2) {
        # Application.Calculate is the supported COM method. This owned Excel
        # instance contains only the temporary workbook, so no user file runs.
        $additionalCalculateRequested = $true
        [void]$excel.Calculate()
        $calculationStateAfterAdditionalCalculate = Wait-ActiveCalculation $excel ([DateTime]::UtcNow.AddMinutes(5))
    }
    $firewallSheet = $book.Worksheets.Item('Firewall matrix')
    $firewallTable = $null
    $filter = $null
    try {
        $firewallTable = $firewallSheet.ListObjects.Item('FirewallRules')
        $filter = $firewallTable.AutoFilter
        [void]$filter.ApplyFilter()
    }
    finally { Release-Com $filter; Release-Com $firewallTable; Release-Com $firewallSheet }
    [void](Wait-ActiveCalculation $excel ([DateTime]::UtcNow.AddMinutes(5)))
    if ((Get-FormulaErrorCount $book $baseline.formulaCount) -ne 0) { Fail-Import 'FORMULA_ERRORS' 'Recalculation produced formula error cells.' }
    # Return the file to automatic calculation for interactive editing.
    $excel.Calculation = -4105
    $stage = 'format_generated_output'
    $currentSheet = 'Deployment'
    $currentRange = 'L:L'
    $layoutSheet = $book.Worksheets.Item('Deployment')
    $layoutColumn = $null
    $protection = $null
    try {
        # Only this new temporary output is changed. The verified generic
        # template has no protection password and disallows column formatting.
        $protection = $layoutSheet.Protection
        $permissions = [ordered]@{}
        foreach ($permission in @('AllowFormattingCells','AllowFormattingColumns','AllowFormattingRows',
            'AllowInsertingColumns','AllowInsertingRows','AllowInsertingHyperlinks','AllowDeletingColumns',
            'AllowDeletingRows','AllowSorting','AllowFiltering','AllowUsingPivotTables')) {
            $permissions[$permission] = [bool]$protection.$permission
        }
        $drawingsProtected = [bool]$layoutSheet.ProtectDrawingObjects
        $contentsProtected = [bool]$layoutSheet.ProtectContents
        $scenariosProtected = [bool]$layoutSheet.ProtectScenarios
        $interfaceOnly = [bool]$layoutSheet.ProtectionMode
        $selectionPermission = [int]$layoutSheet.EnableSelection
        Release-Com $protection
        $protection = $null
        [void]$layoutSheet.Unprotect('')
        try {
            $layoutColumn = $layoutSheet.Range('L:L')
            $layoutColumn.ColumnWidth = 40
        }
        finally {
            [void]$layoutSheet.Protect('', $drawingsProtected, $contentsProtected, $scenariosProtected,
                $interfaceOnly, $permissions.AllowFormattingCells, $permissions.AllowFormattingColumns,
                $permissions.AllowFormattingRows, $permissions.AllowInsertingColumns, $permissions.AllowInsertingRows,
                $permissions.AllowInsertingHyperlinks, $permissions.AllowDeletingColumns, $permissions.AllowDeletingRows,
                $permissions.AllowSorting, $permissions.AllowFiltering, $permissions.AllowUsingPivotTables)
            $layoutSheet.EnableSelection = $selectionPermission
        }
        $protection = $layoutSheet.Protection
        foreach ($permission in $permissions.Keys) {
            if ([bool]$protection.$permission -ne $permissions[$permission]) {
                Fail-Import 'LAYOUT_PROTECTION_CHANGED' 'Deployment protection permissions changed while resizing column L.'
            }
        }
        if ([int]$layoutSheet.EnableSelection -ne $selectionPermission) {
            Fail-Import 'LAYOUT_SELECTION_CHANGED' 'Deployment cell-selection permissions changed while resizing column L.'
        }
    }
    finally { Release-Com $protection; Release-Com $layoutColumn; Release-Com $layoutSheet }
    $stage = 'save_and_reopen'
    $currentSheet = $null
    $currentRange = $null
    $calculationStateBeforeSave = Wait-ActiveCalculation $excel ([DateTime]::UtcNow.AddMinutes(5))
    [void]$book.Save()
    [void]$book.Close($false)
    Release-Com $book
    $book = $null
    $book = $excel.Workbooks.Open($temporaryPath, 0, $true)
    $calculationStateAfterReopen = Wait-ActiveCalculation $excel ([DateTime]::UtcNow.AddMinutes(5))
    $after = Get-WorkbookSnapshot $book $afterFormulaMap
    if (($baseline | ConvertTo-Json -Depth 10 -Compress) -cne ($after | ConvertTo-Json -Depth 10 -Compress)) {
        $snapshotDifference = Get-SnapshotDifference $baseline $after
        if ($IncludeFormulaDiagnostics -and 'formulaHash' -in $snapshotDifference.changedKeys) {
            # Explicit opt-in for a generic-template diagnostic only. Normal
            # import reports never include formula text or imported values.
            $snapshotDifference.formulaDifferences = Get-FormulaDifferences $beforeFormulaMap $afterFormulaMap
        }
        Fail-Import 'PRESERVATION_CHECK_FAILED' 'Formula expressions, table bounds, names, visibility or protection changed.'
    }
    $errorCount = Get-FormulaErrorCount $book $after.formulaCount
    if ($errorCount -ne 0) { Fail-Import 'REOPENED_FORMULA_ERRORS' 'The saved workbook contains formula error cells.' }
    $stage = 'verify_saved_input_values'
    foreach ($group in ($targets.Values | Group-Object -Property sheet)) {
        $currentSheet = $group.Name
        $currentRange = 'UsedRange'
        $sheet = $book.Worksheets.Item($group.Name)
        $used = $null
        try {
            $used = $sheet.UsedRange
            $formulas = $used.Formula
            $values = $used.Value2
            for ($readRetry = 0; ($formulas -isnot [Array] -or $values -isnot [Array]) -and $readRetry -lt 3; $readRetry++) {
                Start-Sleep -Milliseconds 300
                $formulas = $used.Formula
                $values = $used.Value2
            }
            if ($formulas -isnot [Array] -or $values -isnot [Array] -or $formulas.Rank -ne 2 -or $values.Rank -ne 2 -or
                $formulas.GetLength(0) -ne $values.GetLength(0) -or $formulas.GetLength(1) -ne $values.GetLength(1)) {
                Fail-Import 'INCOMPLETE_SAVED_INPUT_READ' 'Excel did not return complete saved input arrays.'
            }
            $firstRow = [int]$used.Row
            $firstColumn = [int]$used.Column
            foreach ($spec in $group.Group) {
                $currentRange = $spec.cell
                $key = Get-CellKey $spec
                $expected = if ($edits.ContainsKey($key)) { $edits[$key].value } else { $null }
                $r = $spec.row - $firstRow + $values.GetLowerBound(0)
                $c = $spec.column - $firstColumn + $values.GetLowerBound(1)
                if ($r -lt $values.GetLowerBound(0) -or $r -gt $values.GetUpperBound(0) -or
                    $c -lt $values.GetLowerBound(1) -or $c -gt $values.GetUpperBound(1)) {
                    Fail-Import 'SAVED_INPUT_OUTSIDE_RANGE' 'A saved input target is outside the returned used range.'
                }
                $formula = $formulas.GetValue($r, $c)
                if (($formula -is [string] -and $formula.StartsWith('=')) -or -not (Test-SameLiteral $expected $values.GetValue($r, $c))) {
                    Fail-Import 'SAVED_VALUE_MISMATCH' ('Saved input differs at ' + $spec.sheet + '!' + $spec.cell + '.')
                }
            }
        }
        finally { Release-Com $used; Release-Com $sheet }
    }
    [void]$book.Close($false)
    Release-Com $book
    $book = $null
    $stage = 'verify_native_protection'
    $currentSheet = $null
    $currentRange = $null
    $afterNative = Get-NativeCellProtection $temporaryPath
    if ($afterNative.unlockedFormulas -ne 0 -or
        ($baselineNative.unlockedCells -join [char]30) -cne ($afterNative.unlockedCells -join [char]30)) {
        Fail-Import 'CELL_PROTECTION_CHANGED' 'Saved cell lock states differ from the template.'
    }
    if ((Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash -cne $templateHash) {
        Fail-Import 'TEMPLATE_CHANGED' 'The source template changed during import; no output will be published.'
    }
    $success = $true
    $result = [ordered]@{
        success = $true; outputPath = $output; clearedCells = $clears.Count; editedCells = $edits.Count
        verifiedTargetCells = $targets.Count; formulaCount = $after.formulaCount; formulaErrors = 0
        formulasPreserved = $true; protectionPreserved = $true; namesAndTablesPreserved = $true
        templateUnchanged = $true; sourceSha256 = $templateHash; ownedExcelPid = $ownedPid
        calculationStateAfterRebuild = $calculationStateAfterRebuild
        additionalCalculateRequested = $additionalCalculateRequested
        calculationStateAfterAdditionalCalculate = $calculationStateAfterAdditionalCalculate
        snapshotDifference = $snapshotDifference
        calculationStateBeforeSave = $calculationStateBeforeSave
        calculationStateAfterReopen = $calculationStateAfterReopen
    }
}
catch {
    $result = [ordered]@{
        success = $false; stage = $stage
        errorCode = $script:FailureCode
        message = if ($script:FailureCode -eq 'UNEXPECTED_FAILURE') { 'Import failed during ' + $stage + '. No output was published.' } else { $script:FailureMessage }
        hresult = $_.Exception.HResult
        exceptionType = $_.Exception.GetType().FullName
        innerExceptionType = if ($null -ne $_.Exception.InnerException) { $_.Exception.InnerException.GetType().FullName } else { $null }
        innerHresult = if ($null -ne $_.Exception.InnerException) { $_.Exception.InnerException.HResult } else { $null }
        scriptLineNumber = if ($null -ne $_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { $null }
        scriptStackTrace = $_.ScriptStackTrace
        currentSheet = $currentSheet
        currentRange = $currentRange
        calculationStateAfterRebuild = $calculationStateAfterRebuild
        additionalCalculateRequested = $additionalCalculateRequested
        calculationStateAfterAdditionalCalculate = $calculationStateAfterAdditionalCalculate
        calculationStateBeforeSave = $calculationStateBeforeSave
        calculationStateAfterReopen = $calculationStateAfterReopen
        snapshotDifference = $snapshotDifference
    }
}
finally {
    if ($null -ne $book) {
        try { [void]$book.Close($false) } catch { }
        Release-Com $book
        $book = $null
    }
    if ($null -ne $excel) {
        if ($ownedExcel) { try { [void]$excel.Quit() } catch { } }
        Release-Com $excel
        $excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownedExcel -and $ownedPid -gt 0) {
        $ownedProcess = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($null -ne $ownedProcess -and $ownedProcess.StartTime -eq $ownedStartTime -and -not $ownedProcess.WaitForExit(3000)) {
            # This PID was absent before creation and its start time still matches.
            Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue
        }
    }
}

if ($success) {
    try {
        if (Test-Path -LiteralPath $output) { Fail-Import 'OUTPUT_RACE' 'The output appeared during import; it was not overwritten.' }
        [IO.File]::Move($temporaryPath, $output, $false)
        $temporaryPath = $null
        $result.outputSha256 = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
        $result.bytes = (Get-Item -LiteralPath $output).Length
    }
    catch {
        $success = $false
        $result = [ordered]@{ success = $false; stage = 'publish_new_output'; errorCode = $script:FailureCode; message = 'The new output could not be published without overwrite.' }
    }
}
if ($null -ne $temporaryPath -and (Test-Path -LiteralPath $temporaryPath)) {
    try { Remove-Item -LiteralPath $temporaryPath -Force }
    catch { $result.temporaryCleanupRequired = $true }
}
$result.elapsedSeconds = [Math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 2)
$result | ConvertTo-Json -Depth 10 -Compress
if (-not $success) { exit 1 }
