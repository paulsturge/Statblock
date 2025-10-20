# RQG.Authoring.psm1
Import-Module ImportExcel -ErrorAction Stop

# ---- Data root (external) ----
$script:DataRoot = 'Y:\Stat_blocks\Data'

function Set-RQGDataRoot {
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Data root path not found: $Path" }
  $script:DataRoot = (Resolve-Path -LiteralPath $Path).Path
}
function Show-RQGDataRoot { [pscustomobject]@{ DataRoot = $script:DataRoot } }

function Resolve-StatPath {
  param([Parameter(Mandatory)][string]$Name)
  if ([System.IO.Path]::IsPathRooted($Name)) {
    if (Test-Path -LiteralPath $Name) { return (Resolve-Path -LiteralPath $Name).Path }
    return $null
  }
  $candidates = @()
  if ($script:DataRoot) { $candidates += $script:DataRoot }
  if ($PSScriptRoot)    { $candidates += $PSScriptRoot; $candidates += (Split-Path $PSScriptRoot -Parent) }
  if ($pwd)             { $candidates += $pwd.Path }
  foreach ($root in ($candidates | Select-Object -Unique)) {
    $p = Join-Path $root $Name
    if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
  }
  return $null
}

$here = Split-Path -Parent $PSCommandPath

# Dot-source function files
. "$here\New-StatDiceRow.ps1"
. "$here\New-HitLocationSheet.ps1"
. "$here\Import-CreatureWeaponsFromText.ps1"
. "$here\New-Creature.ps1"

# Export public functions
Export-ModuleMember -Function `
  New-StatDiceRow, `
  New-HitLocationSheet, `
  Import-CreatureWeaponsFromText, `
  New-Creature, `
  Set-RQGDataRoot, `
  Show-RQGDataRoot, `
  Resolve-StatPath
