<#  load-dev.ps1  —  quick dev loader (no module packaging)  #>
param(
  [string]$DataRoot = 'Y:\Stat_blocks\Data'
)

Import-Module ImportExcel -ErrorAction Stop

# Shared helpers so functions can find your spreadsheets
$script:DataRoot = $DataRoot

function Resolve-StatPath {
  param([Parameter(Mandatory)][string]$Name)
  if ([System.IO.Path]::IsPathRooted($Name)) {
    if (Test-Path -LiteralPath $Name) { return (Resolve-Path -LiteralPath $Name).Path }
    return $null
  }
  $candidates = @($script:DataRoot, $PSScriptRoot, (Split-Path $PSScriptRoot -Parent), $pwd.Path) | Where-Object { $_ }
  foreach ($root in ($candidates | Select-Object -Unique)) {
    $p = Join-Path $root $Name
    if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
  }
  return $null
}
function Set-RQGDataRoot {
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Data root not found: $Path" }
  $script:DataRoot = (Resolve-Path -LiteralPath $Path).Path
}
function Show-RQGDataRoot { [pscustomobject]@{ DataRoot = $script:DataRoot } }

# Dot-source your authoring functions (edit these files freely)
. "$PSScriptRoot\New-StatDiceRow.ps1"
. "$PSScriptRoot\New-HitLocationSheet.ps1"
. "$PSScriptRoot\Import-CreatureWeaponsFromText.ps1"
. "$PSScriptRoot\New-Creature.ps1"

Write-Host "RQG authoring functions loaded from source. DataRoot = $script:DataRoot"
