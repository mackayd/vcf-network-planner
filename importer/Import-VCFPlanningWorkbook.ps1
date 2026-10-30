#requires -Version 7.0
<#
.SYNOPSIS
Imports entered VCF 9.1.1 planning workbook data into a new firewall planner.
.EXAMPLE
./importer/Import-VCFPlanningWorkbook.ps1 -PlanningWorkbook C:/path/vcf-9.1.1.xlsx -FirewallTemplate ./Firewall-Planner-Template.xlsx -OutputWorkbook C:/path/My-Firewall-Planner.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PlanningWorkbook,
    [Parameter(Mandatory)][string]$FirewallTemplate,
    [Parameter(Mandatory)][string]$OutputWorkbook,
    [string]$Config,
    [switch]$DryRun,
    [string]$PythonPath
)
$ErrorActionPreference='Stop'
$entry=Join-Path $PSScriptRoot 'vcf_import.py'
$pythonArgs=@()
if($PythonPath){$python=$PythonPath}
elseif(Get-Command py -ErrorAction SilentlyContinue){$python=(Get-Command py).Source;$pythonArgs=@('-3')}
elseif(Get-Command python -ErrorAction SilentlyContinue){$python=(Get-Command python).Source}
else{throw 'Python 3.10 or later is required. Install Python or supply -PythonPath.'}
& $python @pythonArgs -c 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)'
if($LASTEXITCODE -ne 0){throw 'A working Python 3.10 or later interpreter is required.'}
$argsToImport=@($entry,'--source',$PlanningWorkbook,'--template',$FirewallTemplate,'--output',$OutputWorkbook)
if($Config){$argsToImport+=@('--config',$Config)}
if($DryRun){$argsToImport+='--dry-run'}
& $python @pythonArgs @argsToImport
if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}
