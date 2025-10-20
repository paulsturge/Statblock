param(
  [string]$ModuleName = 'RQG.Authoring',
  [string]$Version    = '1.0.0',
  [string]$DataRoot   = 'Y:\Stat_blocks\Data'
)

$src = Split-Path -Parent $PSCommandPath
$dst = Join-Path $HOME "Documents\PowerShell\Modules\$ModuleName\$Version"
New-Item -ItemType Directory -Force -Path $dst | Out-Null

# copy function files
$funcs = @(
  'New-StatDiceRow.ps1',
  'New-HitLocationSheet.ps1',
  'Import-CreatureWeaponsFromText.ps1',
  'New-Creature.ps1'
)
foreach ($f in $funcs) { Copy-Item (Join-Path $src $f) -Destination $dst -Force }

# compose PSM1 content (single-line Export-ModuleMember)
$psm1Content = @"
Import-Module ImportExcel -ErrorAction Stop

`$script:DataRoot = "$DataRoot"

function Set-RQGDataRoot {
  param([Parameter(Mandatory)][string]`$Path)
  if (-not (Test-Path -LiteralPath `"$Path`")) { throw "Data root not found: `$Path" }
  `$script:DataRoot = (Resolve-Path -LiteralPath `"$Path`").Path
}
function Show-RQGDataRoot { [pscustomobject]@{ DataRoot = `$script:DataRoot } }

function Resolve-StatPath {
  param([Parameter(Mandatory)][string]`$Name)
  if ([System.IO.Path]::IsPathRooted(`$Name)) {
    if (Test-Path -LiteralPath `"$Name`") { return (Resolve-Path -LiteralPath `"$Name`").Path }
    return `$null
  }
  `$candidates = @(`$script:DataRoot, `$PSScriptRoot, (Split-Path `$PSScriptRoot -Parent), `$pwd.Path) | Where-Object { `$_ }
  foreach (`$root in (`$candidates | Select-Object -Unique)) {
    `$p = Join-Path `$root `$Name
    if (Test-Path -LiteralPath `$p) { return (Resolve-Path -LiteralPath `$p).Path }
  }
  return `$null
}

. "`$PSScriptRoot\New-StatDiceRow.ps1"
. "`$PSScriptRoot\New-HitLocationSheet.ps1"
. "`$PSScriptRoot\Import-CreatureWeaponsFromText.ps1"
. "`$PSScriptRoot\New-Creature.ps1"

Export-ModuleMember -Function New-StatDiceRow,New-HitLocationSheet,Import-CreatureWeaponsFromText,New-Creature,Set-RQGDataRoot,Show-RQGDataRoot,Resolve-StatPath
"@

# validate PSM1 syntax before writing
try { [ScriptBlock]::Create($psm1Content) | Out-Null }
catch {
  Write-Error "Generated PSM1 failed to parse: $($_.Exception.Message)"
  throw
}

# write PSM1
$psm1Path = Join-Path $dst "$ModuleName.psm1"
$psm1Content | Set-Content -Encoding UTF8 -Path $psm1Path

# write manifest with correct RootModule filename
$psd1Path = Join-Path $dst "$ModuleName.psd1"
if (Test-Path $psd1Path) { Remove-Item $psd1Path -Force }
New-ModuleManifest -Path $psd1Path `
  -RootModule "$ModuleName.psm1" `
  -ModuleVersion $Version `
  -Guid (New-Guid) `
  -Author 'You' `
  -Description 'RQG creature authoring tools' `
  -PowerShellVersion '7.0' `
  -RequiredModules @('ImportExcel') `
  -FunctionsToExport @(
    'New-StatDiceRow','New-HitLocationSheet','Import-CreatureWeaponsFromText','New-Creature',
    'Set-RQGDataRoot','Show-RQGDataRoot','Resolve-StatPath'
  ) `
  -CmdletsToExport @() -VariablesToExport @() -AliasesToExport @()

# import by PATH to avoid any lookup quirks
Remove-Module $ModuleName -ErrorAction SilentlyContinue
Import-Module $psm1Path -Force -Verbose

# smoke test
Get-Command -Module $ModuleName | Select-Object Name
Show-RQGDataRoot
