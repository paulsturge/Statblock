# Demo: Spirit Magic randomizer (Thed Broo Initiate), using longhand CHA
# ----------------------------------------------------------------------

# 0) Start clean and load your environment
Set-Location (Split-Path $PSCommandPath -Parent)  # enter the _demo folder
. "$PSScriptRoot\..\_devLoader.ps1"               # dot-source loader so modules land in session

# 1) Make sure the core + randomizer functions are present
$need = 'Import-SpiritMagicCatalog','New-RandomSpiritMagicLoadout','Set-StatblockSpiritMagic','Add-CultInfoToStatblock','New-Statblock'
$missing = $need | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) }
if ($missing) { throw "Missing required functions: $($missing -join ', ')" }

# 2) Ensure the row picker exists (older copies might not export it)
if (-not (Get-Command Get-WeightedRandomRow -ErrorAction SilentlyContinue)) {
    function Get-WeightedRandomRow {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][array]$Catalog,
            [int]$Seed
        )
        if ($Seed) { $null = Get-Random -SetSeed $Seed }
        if (-not $Catalog -or $Catalog.Count -eq 0) { return $null }

        # If a Weight column is present, expand by weight; else uniform pick
        if ($Catalog[0].PSObject.Properties.Match('Weight').Count -gt 0) {
            $expanded = foreach ($row in $Catalog) {
                $w = 1
                try { $w = [math]::Max(1, [int]$row.Weight) } catch {}
                1..$w | ForEach-Object { $row }
            }
            if ($expanded -and $expanded.Count -gt 0) { return ($expanded | Get-Random) }
        }
        return ($Catalog | Get-Random)
    }
}

# 3) Load the catalog
$catalogPath = "Y:\Stat_blocks\Data\spirit_magic_catalog.csv"
if (-not (Test-Path $catalogPath)) { throw "Catalog not found at $catalogPath" }
$cat = Import-SpiritMagicCatalog -CsvPath $catalogPath
if (-not $cat -or $cat.Count -eq 0) { throw "Catalog loaded empty from $catalogPath" }

# 4) Build the Broo, add Thed role
$ctx = Initialize-StatblockContext
$sb  = New-Statblock -Creature 'Broo' -Context $ctx -AddArmor 0
$sb  = Add-CultInfoToStatblock -Statblock $sb -CultName 'Thed' -Role 'Initiate'

# 5) Roll 7 spirit-magic points, capped by CHA (longhand)
$cha = $sb.Characteristics.CHA
if (-not $cha -or $cha -le 0) { throw "Statblock CHA is invalid: '$cha'" }

$rand = New-RandomSpiritMagicLoadout `
            -PointsBudget 7 `
            -CHA $cha `
            -Role 'Initiate' `
            -Catalog $cat `
            -Seed 99 `
            -Trace

# 6) Apply to statblock
$sb = Set-StatblockSpiritMagic -Statblock $sb -Spells $rand

# 7) Show results
$sb = Apply-SpiritMagic-RQG $sb $rand   # positional, CHA cap enforced
$sp = @($sb.Magic['Spirit'])
$sp | Sort-Object Name | ft Name, Points -Auto
"Total Spirit Points: " + (($sp | Measure-Object Points -Sum).Sum) + " / CHA " + $sb.Characteristics.CHA

