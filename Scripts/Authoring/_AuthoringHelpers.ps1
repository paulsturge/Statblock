# --- helpers expected by New-* commands when not using the module ---

# Optional (used all over your scripts)
try { Import-Module ImportExcel -ErrorAction Stop } catch {}

# Data root for Resolve-StatPath probing
if (-not (Get-Variable -Name DataRoot -Scope Script -ErrorAction SilentlyContinue)) {
  $script:DataRoot = 'Y:\Stat_blocks\Data'
}

function Set-RQGDataRoot {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Data root not found: $Path" }
  $script:DataRoot = (Resolve-Path -LiteralPath $Path).Path
}
function Show-RQGDataRoot { [pscustomobject]@{ DataRoot = $script:DataRoot } }

function Resolve-StatPath {
  [CmdletBinding()]
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


