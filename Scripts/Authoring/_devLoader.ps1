# Minimal dev loader — no recursion, no long scans
$ErrorActionPreference = 'Stop'
$PSModuleAutoLoadingPreference = 'None'

$root = "Y:\Stat_blocks\Scripts\Authoring"

function _ok($m){ Write-Host "[devLoader] $m" }

# Load helpers first
. "$root\_AuthoringHelpers.ps1"
_ok "Helpers loaded. DataRoot = $((Show-RQGDataRoot).DataRoot)"

# Dot-source function files in dependency order
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

# Quick sanity
foreach ($n in 'Resolve-StatPath','Set-RQGDataRoot','Show-RQGDataRoot','New-StatDiceRow','New-Creature') {
  if (-not (Get-Command $n -ErrorAction SilentlyContinue)) { Write-Warning "MISSING: $n" } else { _ok "Available: $n" }
}

_ok "Done."
