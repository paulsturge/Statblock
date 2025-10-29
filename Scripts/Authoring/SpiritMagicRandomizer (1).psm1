# =========================
# Spirit Magic Randomizer (Role-cap aware)
# =========================

function Import-SpiritMagicCatalog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CsvPath
    )
    $rows = Import-Csv -Path $CsvPath
    $rows | ForEach-Object {
        # Capture all RoleMax_* columns dynamically into a hashtable
        $roleMax = @{}
        $_.PSObject.Properties |
            Where-Object { $_.Name -like 'RoleMax_*' -and $_.Value -ne $null -and $_.Value -ne '' } |
            ForEach-Object {
                $role = ($_.Name -replace '^RoleMax_', '')
                $roleMax[$role] = [int]$_.Value
            }

        [pscustomobject]@{
            Name      = $_.Name
            Min       = [int]$_.Min
            Max       = [int]$_.Max
            Notes     = $_.Notes
            RoleMax   = $roleMax
        }
    }
}
function Apply-SpiritMagic-RQG {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] $Statblock,
        [Parameter(Mandatory, Position=1)] [object[]]$Spells,
        [switch]$Force
    )

    # Normalize spells (simple PSCustomObjects)
    $normalized = @(
        $Spells | ForEach-Object {
            [pscustomobject]@{
                Name   = $_.Name
                Points = [int]$_.Points
                Notes  = $_.Notes
            }
        }
    )

    # Ensure .Magic exists WITHOUT Add-Member
    if (-not $Statblock.PSObject.Properties['Magic']) {
        $Statblock.PSObject.Properties.Add(
            [System.Management.Automation.PSNoteProperty]::new('Magic', @{})
        )
    }

    # If .Magic is not a dictionary, convert it to hashtable (no Add-Member used)
    if ($Statblock.Magic -isnot [System.Collections.IDictionary]) {
        $tmp = @{}
        foreach ($p in $Statblock.Magic.PSObject.Properties) { $tmp[$p.Name] = $p.Value }
        $Statblock.Magic = $tmp
    }

    # Assign 'Spirit' directly on the hashtable
    $Statblock.Magic['Spirit'] = $normalized

    # CHA cap check (longhand)
    $sum = ($normalized | Measure-Object Points -Sum).Sum
    $cha = $Statblock.Characteristics.CHA
    if (-not $Force -and $cha -is [int] -and $cha -gt 0 -and $sum -gt $cha) {
        throw "Spirit Magic exceeds CHA cap. Total=$sum, CHA=$cha."
    }

    return $Statblock
}

function Get-IntensityRangeForSpell {
    <#
      Returns [Min, Max] for a spell considering role caps and any bespoke rules.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$SpellRow,
        [Parameter(Mandatory)][string]$Role
    )
    $min = [int]$SpellRow.Min
    $max = [int]$SpellRow.Max

    # Apply dynamic RoleMax_* if present
    $normRole = $Role -replace '\s',''
    if ($SpellRow.RoleMax.ContainsKey($normRole)) {
        $max = [int]$SpellRow.RoleMax[$normRole]
    }

    # Bespoke rule: Initiate + Bladesharp should not roll lower than 2
    if ($normRole -eq 'Initiate' -and $SpellRow.Name -eq 'Bladesharp' -and $min -lt 2) {
        $min = 2
        if ($max -lt 2) { $max = 2 }
    }

    if ($max -lt $min) { $max = $min }
    return @($min, $max)
}

function Get-WeightedRandomRow {
    param([Parameter(Mandatory)][array]$Catalog)
    # If you later add a Weight column, you can expand it here. For now: plain random.
    return ($Catalog | Get-Random)
}

function New-RandomSpiritMagicLoadout {
    <#
      Allocates a PointsBudget of spirit magic for a role, not breaking CHA.
      - Picks new spells; if picked again, raises intensity +1 up to role-capped max.
      - Honors fixed-point spells (Min=Max).
      - Always returns an array (possibly empty), never $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$PointsBudget,
        [Parameter(Mandatory)][int]$CHA,
        [Parameter(Mandatory)][ValidateSet('Lay','Initiate','Rune','Lord','RuneLord','Priest','Acolyte','Hero','Other')]
        [string]$Role,
        [Parameter(Mandatory)][array]$Catalog,
        [int]$Seed,
        [switch]$Trace
    )

    # ---------- TOP-OF-FUNCTION GUARDS ----------
    if (-not $Catalog -or $Catalog.Count -eq 0) {
        throw "New-RandomSpiritMagicLoadout: Catalog is empty. Check your CSV path and contents."
    }
    if ($PointsBudget -le 0 -or $CHA -le 0) {
        return @()  # empty loadout, not $null
    }
    if ($Seed) { $null = Get-Random -SetSeed $Seed }

    $target = [Math]::Min($PointsBudget, $CHA)
    $total  = 0
    $out    = @()
    $byName = @{}

    # ---------- MAIN LOOP ----------
    while ($total -lt $target) {

        # Row pick and null-check belong INSIDE the loop:
        $row = Get-WeightedRandomRow -Catalog $Catalog
        if (-not $row) { break }  # nothing to pick

        if ($byName.ContainsKey($row.Name)) {
            # Try to raise intensity on an existing pick
            $existing = $byName[$row.Name]
            $range    = Get-IntensityRangeForSpell -SpellRow $row -Role $Role
            $roleMax  = $range[1]

            if ($existing.Points -lt $roleMax) {
                if (($total + 1) -le $target) {
                    $existing.Points++
                    $total++
                    if ($Trace) { Write-Host "↑ $($row.Name) -> $($existing.Points) (total $total/$target)" -f DarkCyan }
                } else {
                    break
                }
            } else {
                continue  # can't raise more; try another pick next loop
            }

        } else {
            # New spell: roll an intensity in the role-aware range
            $range  = Get-IntensityRangeForSpell -SpellRow $row -Role $Role
            $min    = [int]$range[0]
            $max    = [int]$range[1]
            $rolled = if ($min -eq $max) { $min } else { Get-Random -Minimum $min -Maximum ($max + 1) }

            # Fit into remaining budget; shrink if needed
            $pts = $rolled
            while ($pts -gt 0 -and ($total + $pts) -gt $target) { $pts-- }
            if ($pts -le 0) {
                if ($Trace) { Write-Host "skip $($row.Name) (rolled $rolled won't fit; total $total/$target)" -f DarkYellow }
                continue
            }

            $item = [pscustomobject]@{
                Name   = $row.Name
                Points = $pts
                Notes  = $row.Notes
            }
            $out   += $item
            $byName[$row.Name] = $item
            $total += $pts
            if ($Trace) { Write-Host "＋ $($row.Name) $pts (total $total/$target)" -f Green }
        }
    }

    # ---------- ALWAYS RETURN AN ARRAY ----------
    ,@($out | Sort-Object Name)
}

# helper stays OUTSIDE the while, anywhere in your module:
function Get-WeightedRandomRow {
    param([Parameter(Mandatory)][array]$Catalog)
    if (-not $Catalog -or $Catalog.Count -eq 0) { return $null }
    # If you add a Weight column later, expand by weight here.
    return ($Catalog | Get-Random)
}


function Set-StatblockSpiritMagic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] $Statblock,
        [Parameter(Mandatory, Position=1)] [object[]]$Spells,
        [switch]$Force
    )

    # normalize
    $normalized = @(
        $Spells | ForEach-Object {
            [pscustomobject]@{ Name=$_.Name; Points=[int]$_.Points; Notes=$_.Notes }
        }
    )

    # ensure .Magic exists (no Add-Member collisions)
    if (-not $Statblock.PSObject.Properties['Magic']) {
        $Statblock.PSObject.Properties.Add(
            [System.Management.Automation.PSNoteProperty]::new('Magic', @{})
        )
    }

    # convert to hashtable if needed, then assign Spirit
    if ($Statblock.Magic -isnot [System.Collections.IDictionary]) {
        $tmp = @{}; foreach ($p in $Statblock.Magic.PSObject.Properties){ $tmp[$p.Name]=$p.Value }; $Statblock.Magic = $tmp
    }
    $Statblock.Magic['Spirit'] = $normalized

    # CHA cap (longhand)
    $sum = ($normalized | Measure-Object Points -Sum).Sum
    $cha = $Statblock.Characteristics.CHA
    if (-not $Force -and $cha -is [int] -and $cha -gt 0 -and $sum -gt $cha) {
        throw "Spirit Magic exceeds CHA cap. Total=$sum, CHA=$cha."
    }
    return $Statblock
}

function Show-SpiritMagic {
    param([Parameter(Mandatory)][pscustomobject]$Statblock)
    $Statblock.Magic.Spirit |
        Sort-Object Name |
        Format-Table Name, Points, Notes -Auto
}

function Get-SpiritBudgetByRole {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][int]$CHA
    )
    # Conservative defaults from your notes; always cap by CHA.
    switch ($Role) {
        'Lay'       { $min=2;  $max=4 }
        'Initiate'  { $min=5;  $max=10 }  # powerful initiates may reach 12–13 if CHA allows
        'Rune' { $min=8;  $max=[math]::Max(10,[int]([math]::Round($CHA*0.8))) }
        'Lord' { $min=10; $max=$CHA }
        'RuneLord'  { $min=10; $max=$CHA }
        'Priest'    { $min=10; $max=$CHA }
        default     { $min=4;  $max=[math]::Min(8,$CHA) }
    }
    if ($max -lt 1) { return 0 }
    $roll = if ($min -ge $max) { [math]::Min($min,$CHA) } else { Get-Random -Minimum $min -Maximum ($max+1) }
    return [math]::Min($roll, $CHA)
}

function Test-CultGrantsFullPriceSpirit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CultName
    )
    # If Get-CultMagic returns structured data, swap this for a property check.
    $info = Get-CultMagic -CultName $CultName -ErrorAction SilentlyContinue
    if (-not $info) { return $false }

    $txt = ($info | Out-String)
    return ($txt -match 'all\s+spirit\s+magic\s+at\s+full\s+price')
}

function Get-RoleRunePointRange {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Role)

    # Normalize: keep letters only, lowercase. e.g., "Rune Lord" -> "runelord"
    $norm = ('' + $Role).ToLower() -replace '[^a-z]', ''

    switch ($norm) {
        'runelord' { 5,10 }   # 5 + up to 5
        'runelady' { 5,10 }   # treat same as runelord
        'runepriest' { 6,14 } # 6 + up to 8
        'priest'     { 6,14 } # accept shorthand
        'initiate'   { 3,6 }  # 3 + up to 3
        default      { 0,0 }  # others: none for now
    }
}

function New-RunePoints {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][int]$INT
    )
    if ($INT -le 0) { return 0 }  # INT gate

    $min,$max = Get-RoleRunePointRange -Role $Role
    if ($max -le $min) { return $min }
    Get-Random -Minimum $min -Maximum ($max + 1)
}

function Resolve-CultSheetName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CultName,
        [Parameter(Mandatory)][ValidateSet('Magic','Associations','Roles')]$Suffix
    )
    "{0}_{1}" -f $CultName, $Suffix
}

function Get-CultRuneSpellCatalog {
    <#
      Returns a PSCustomObject with two arrays:
        .Common  = @({ Name, FromCult }...)
        .Special = @({ Name, FromCult }...)
      Reads the *same* per-cult Magic sheet used for Spirit, filters by MagicType =~ '^rune'.
      Classification:
        - Common if Tags or Access contain 'common' (case-insensitive)
        - Else Special
      If -IncludeAssociates, we also pull associates’ rune rows (single hop).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CultName,
        [string]$WorkbookPath = "Y:\Stat_blocks\Data\Cults.xlsx",
        [switch]$IncludeAssociates
    )

    if (-not (Test-Path $WorkbookPath)) { throw "Cults workbook not found: $WorkbookPath" }

    $common  = New-Object System.Collections.Generic.List[object]
    $special = New-Object System.Collections.Generic.List[object]

    function Add-FromRows([object[]]$rows, [string]$fromCult) {
        foreach ($r in $rows) {
            $mt = '' + $r.MagicType
            $sp = ('' + $r.Spell).Trim()
            if ($sp -eq '') { continue }
            if ($mt -notmatch '^(?i)rune') { continue }  # only Rune rows from this sheet

            $tags   = '' + $r.Tags
            $access = '' + $r.Access
            $isCommon = ($tags -match '(?i)\bcommon\b') -or ($access -match '(?i)\bcommon\b')

            $obj = [pscustomobject]@{ Name = $sp; FromCult = $fromCult }
            if ($isCommon) { $common.Add($obj) } else { $special.Add($obj) }
        }
    }

    # main cult
    $magicSheet = Resolve-CultSheetName -CultName $CultName -Suffix 'Magic'
    $rows = Import-Excel -Path $WorkbookPath -WorksheetName $magicSheet -ErrorAction SilentlyContinue
    if ($rows) { Add-FromRows $rows $CultName }

    # associates (optional)
    if ($IncludeAssociates) {
        $assocSheet = Resolve-CultSheetName -CultName $CultName -Suffix 'Associations'
        $assoc = Import-Excel -Path $WorkbookPath -WorksheetName $assocSheet -ErrorAction SilentlyContinue
        if ($assoc) {
            $assocCults = @(
                $assoc | Where-Object { ('' + $_.FromCult).Trim() -ne '' } |
                Select-Object -ExpandProperty FromCult -Unique
            )
            foreach ($ac in $assocCults) {
                $rows2 = Import-Excel -Path $WorkbookPath -WorksheetName (Resolve-CultSheetName -CultName $ac -Suffix 'Magic') -ErrorAction SilentlyContinue
                if ($rows2) { Add-FromRows $rows2 $ac }
            }
        }
    }

    # dedupe by Name within each bucket
    $dedupe = {
        param($list)
        $byName = @{}
        foreach ($it in $list) { if (-not $byName.ContainsKey($it.Name)) { $byName[$it.Name] = $it } }
        ,($byName.Values)
    }

    [pscustomobject]@{
        Common  = & $dedupe $common
        Special = & $dedupe $special
    }
}

function New-RuneSpellLoadout {
    <#
      Picks 1 SPECIALTY rune spell per Rune Point from .Special.
      If we run out of uniques, repeats are allowed (pool refills).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$RunePoints,
        [Parameter(Mandatory)][string]$CultName,
        [string]$WorkbookPath = "Y:\Stat_blocks\Data\Cults.xlsx",
        [switch]$IncludeAssociates,
        [int]$Seed
    )
    if ($Seed) { $null = Get-Random -SetSeed $Seed }
    if ($RunePoints -le 0) { return @(), @() }

    $catalog = Get-CultRuneSpellCatalog -CultName $CultName -WorkbookPath $WorkbookPath -IncludeAssociates:$IncludeAssociates
    $special = @($catalog.Special)
    $common  = @($catalog.Common)

    if ($special.Count -eq 0) { return @(), $common }

    $out = New-Object System.Collections.Generic.List[object]
    $pool = @($special | Select-Object -ExpandProperty Name -Unique)
    for ($i=0; $i -lt $RunePoints; $i++) {
        if ($pool.Count -eq 0) { $pool = @($special | Select-Object -ExpandProperty Name -Unique) }
        $pick = Get-Random -InputObject $pool
        $pool = $pool | Where-Object { $_ -ne $pick }
        $out.Add([pscustomobject]@{ Name = $pick })
    }
    ,($out), $common
}

function Set-StatblockRuneMagic {
    [CmdletBinding()]
  param(
        [Parameter(Mandatory, Position=0)] $Statblock,
        [Parameter(Mandatory, Position=1)] [int]$RunePoints,

        # Not mandatory; default to empty arrays so empty-catalog cases work
        [Parameter(Position=2)] [object[]]$RuneSpecialSpells = @(),
        [Parameter(Position=3)] [object[]]$RuneCommonSpells  = @()
    )
    # Ensure .Magic exists and is a hashtable for safe indexer writes
    if (-not $Statblock.PSObject.Properties['Magic']) {
        $Statblock.PSObject.Properties.Add(
            [System.Management.Automation.PSNoteProperty]::new('Magic', @{})
        )
    }
    if ($Statblock.Magic -isnot [System.Collections.IDictionary]) {
        $tmp = @{}; foreach ($p in $Statblock.Magic.PSObject.Properties) { $tmp[$p.Name] = $p.Value }; $Statblock.Magic = $tmp
    }

    $spec = @($RuneSpecialSpells | ForEach-Object { [pscustomobject]@{ Name = $_.Name } })
    $comm = @($RuneCommonSpells  | ForEach-Object { [pscustomobject]@{ Name = $_.Name } })

    $Statblock.Magic['RunePoints']   = [int]$RunePoints
    $Statblock.Magic['RuneSpecial']  = $spec
    $Statblock.Magic['RuneCommon']   = $comm

    # Optional mirror for printing
    if ($Statblock.PSObject.Properties['RuneMagic']) { $null = $Statblock.PSObject.Properties.Remove('RuneMagic') }
    $Statblock.PSObject.Properties.Add(
        [System.Management.Automation.PSNoteProperty]::new('RuneMagic', [pscustomobject]@{
            Points  = [int]$RunePoints
            Special = $spec
            Common  = $comm
        })
    )
    return $Statblock
}


Export-ModuleMember -Function `
    Import-SpiritMagicCatalog, `
    Get-IntensityRangeForSpell, `
    New-RandomSpiritMagicLoadout, `
    Set-StatblockSpiritMagic, `
    Show-SpiritMagic, `
    Get-SpiritBudgetByRole, `
    Get-RoleRunePointRange, New-RunePoints, Resolve-CultSheetName, `
    Get-CultRuneSpellCatalog, New-RuneSpellLoadout, Set-StatblockRuneMagic, `
    Test-CultGrantsFullPriceSpirit -ErrorAction SilentlyContinue 


