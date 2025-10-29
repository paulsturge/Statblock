# Minimal dev loader — no recursion, no long scans
$ErrorActionPreference = 'Stop'
$PSModuleAutoLoadingPreference = 'None'

$root = "Y:\Stat_blocks\Scripts\Authoring"

function _ok ($m)  { Write-Host "[devLoader] $m" -ForegroundColor Green }
function _err($m)  { Write-Host "[devLoader] $m" -ForegroundColor Red }

# --- Load modules first (compiled/packaged functionality) ---
# --- Core tools (context + constructors) ---
# Statblock-tools.psm1 lives one folder up from Authoring
$modCore = Join-Path (Split-Path $PSScriptRoot -Parent) 'Statblock-tools.psm1'
if (-not (Test-Path $modCore)) {
    _err "Missing module: $modCore"
    throw
}
Import-Module $modCore -Force
_ok "Module loaded: $modCore"

# Sanity: confirm core exports we rely on
$coreExports = @('Initialize-StatblockContext','New-Statblock')
$missing = $coreExports | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) }
if ($missing) {
    _err "Core export(s) missing: $($missing -join ', ')"
} else {
    _ok  "Available: $($coreExports -join ', ')"
}

# SpiritMagicRandomizer.psm1 (spirit-magic catalog + randomizer + viewers)
$modSpiritMagic = Join-Path $root 'SpiritMagicRandomizer.psm1'
if (-not (Test-Path $modSpiritMagic)) {
    _err "Missing module: $modSpiritMagic"
    throw
}
Import-Module $modSpiritMagic -Force
_ok "Module loaded: $modSpiritMagic"

# Get-CultData.psm1 (data access helpers: Get-CultNames, Get-CultMagic, Get-CultRoles, Get-CultAssociations)
$modCultData = Join-Path $root 'Get-CultData.psm1'
if (-not (Test-Path $modCultData)) { _err "Missing module: $modCultData"; throw }
Import-Module $modCultData -Force
_ok "Module loaded: $modCultData"

# Add-CultInfoToStatblock.psm1 (decorator for the $sb object)
$modAddCult = Join-Path $root 'Add-CultInfoToStatblock.psm1'
if (-not (Test-Path $modAddCult)) { _err "Missing module: $modAddCult (did you rename it from .ps1?)"; throw }
Import-Module $modAddCult -Force
_ok "Module loaded: $modAddCult"

# --- Dot-source script helpers (utilities that don't need Export-ModuleMember) ---
$helpers = Join-Path $root '_AuthoringHelpers.ps1'
if (-not (Test-Path $helpers)) { _err "Missing helpers: $helpers"; throw }
. $helpers
_ok "Helpers loaded. DataRoot = $((Show-RQGDataRoot).DataRoot)"

# --- Dot-source authoring scripts in dependency order ---
$ordered = @(
  'New-StatDiceRow.ps1',
  'New-HitLocationSheet.ps1',
  'Import-CreatureWeaponsFromText.ps1',
  'New-Creature.ps1'
)

foreach ($file in $ordered) {
  $path = Join-Path $root $file
  if (-not (Test-Path $path)) { Write-Warning "Missing: $path"; continue }
  . $path
  _ok "Loaded: $file"
}

# --- Quick sanity checks ---
$mustHave = @(
  'Resolve-StatPath',
  'Set-RQGDataRoot',
  'Show-RQGDataRoot',
  'New-StatDiceRow',
  'New-HitLocationSheet',
  'Import-CreatureWeaponsFromText',
  'New-Creature',
  'Add-CultInfoToStatblock',   # from Add-CultInfoToStatblock.psm1
  'Get-CultNames',             # from Get-CultData.psm1
  'Get-CultMagic',             # from Get-CultData.psm1
  'Get-CultRoles',             # from Get-CultData.psm1
  'Get-CultAssociations'       # from Get-CultData.psm1
)
foreach ($n in $mustHave) {
  if (-not (Get-Command $n -ErrorAction SilentlyContinue)) {
    Write-Warning "MISSING: $n"
  } else {
    _ok "Available: $n"
  }
}

_ok "Done."
