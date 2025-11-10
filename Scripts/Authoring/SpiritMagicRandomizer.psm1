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
function Get-CultSpiritCatalogSlim {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CultName,
        [string]$WorkbookPath = "Y:\Stat_blocks\Data\Cults.xlsx"
    )

    if (-not (Test-Path $WorkbookPath)) { return @() }

    $prefix = Resolve-CultSheetName -CultName $CultName -WorkbookPath $WorkbookPath
    $sheet  = "${prefix}_Magic"

    $rows = @()
    try {
        $rows = Import-Excel -Path $WorkbookPath -WorksheetName $sheet -ErrorAction Stop
    } catch {
        $rows = @()
    }

    # keep Spirit rows that are not Prohibited
    $allowedRows = $rows | Where-Object {
        ('' + $_.MagicType) -match '^(?i)spirit$' -and
        -not (('' + $_.Access) -match '^(?i)prohibited$') -and
        -not (('' + $_.Prohibited) -match '^(?i)true$')
    }

    $out = foreach ($r in $allowedRows) {
        [pscustomobject]@{
            Name    = ('' + $r.Spell).Trim()
            Min     = 1
            Max     = 1
            Notes   = ('' + $r.Notes).Trim()
            RoleMax = @{}
        }
    }

    # dedupe by Name
    $out | Group-Object Name | ForEach-Object { $_.Group[0] }
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

    # normalize role (same logic as elsewhere)
    function _NormRole([string]$r){
        if (-not $r) { return 'Initiate' }
        switch -Regex ($r) {
            '^(?i)lay'                                 { return 'Lay' }
            '^(?i)init'                                { return 'Initiate' }
            '^(?i)doomed'                              { return 'Initiate' }
            '^(?i)rune\s*lord|doom\s*master|jaw'       { return 'RuneLord' }
            '^(?i)priest|tongue|hand|horn|breath|high' { return 'Priest' }
            '^(?i)shaman'                              { return 'Shaman' }
            default                                    { return 'Initiate' }
        }
    }

    function _Roll([int]$min,[int]$max){
        if ($max -lt $min) { return $min }
        # inclusive range [min..max]
        Get-Random -Minimum $min -Maximum ($max + 1)
    }

    $normRole = _NormRole $Role
    if ($CHA -lt 0) { $CHA = 0 }

    $pts = 0

    switch ($normRole) {
        'Lay' {
            # Lay Member: 2–4 points of Spirit Magic, flat
            $pts = _Roll 2 4
        }
        'Initiate' {
            # Initiate: 5–10 points, 30% chance to bump by +1–2 up to 13
            $pts = _Roll 5 10
            if ((Get-Random -Minimum 1 -Maximum 101) -le 30) {
                $pts = [Math]::Min($pts + (_Roll 1 2), 13)
            }
        }
        'RuneLord' {
            # Rune Lord: CHA-based, but capped between 8 and CHA
            $base = [int][Math]::Round($CHA * (_Roll 75 100) / 100.0)
            $pts  = [Math]::Max([Math]::Min($base, $CHA), 8)
        }
        'Priest' {
            # Priest: more generous than Initiate, with a higher minimum
            $base = [int][Math]::Round($CHA * (_Roll 85 100) / 100.0)
            $pts  = [Math]::Max([Math]::Min($base, $CHA), 10)
        }
        'Shaman' {
            # Shaman: very spirit-heavy, high floor
            $base = [int][Math]::Round($CHA * (_Roll 85 100) / 100.0)
            $pts  = [Math]::Max([Math]::Min($base, $CHA), 12)
        }
        default {
            # Fallback: treat as Initiate-ish
            $pts = _Roll 5 10
        }
    }

    if ($pts -lt 0) { $pts = 0 }

    return [int]$pts
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
        [string]$WorkbookPath = "Y:\Stat_blocks\Data\Cults.xlsx"
    )

    if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

    $sheets = (Get-ExcelSheetInfo -Path $WorkbookPath).Name
    $prefixes = @()
    foreach ($s in $sheets) {
        if ($s -match '^(.*)_(Magic|Associations|Roles)$') { $prefixes += $Matches[1] }
    }
    $prefixes = $prefixes | Sort-Object -Unique
    if (-not $prefixes) { throw "No cult sheets found in $WorkbookPath" }

    $norm = { param($t) ('' + $t).ToLower() -replace '[^a-z]' }
    $want = & $norm $CultName

    foreach ($p in $prefixes) { if ((& $norm $p) -eq $want) { return $p } }
    foreach ($p in $prefixes) {
        $np = & $norm $p
        if ($np.StartsWith($want) -or $want.StartsWith($np) -or $np.Contains($want) -or $want.Contains($np)) { return $p }
    }
    return $CultName
}

function Deduplicate-ByName {
    [CmdletBinding()]
    param([object[]]$Items)

    if (-not $Items -or $Items.Count -eq 0) { return @() }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $out  = New-Object 'System.Collections.Generic.List[object]'

    foreach ($it in @($Items)) {
        if ($null -eq $it) { continue }
        $name = ('' + $it.Name).Trim()
        if (-not $name) { continue }
        if ($seen.Add($name)) { $out.Add($it) }
    }
    ,$out.ToArray()
}



function Get-CultRuneSpellCatalog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CultName,
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$IncludeAssociates
    )

    if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

    # Resolve sheet prefix, e.g. "KygerLitor", "Mallia", "Thed"
    $prefix     = Resolve-CultSheetName -CultName $CultName -WorkbookPath $WorkbookPath
    $magicSheet = "${prefix}_Magic"
    $assocSheet = "${prefix}_Associations"

    #
    # STEP 1: Load this cult's own rune magic
    #
    $base = @()
    try {
        $base = Import-Excel -Path $WorkbookPath -WorksheetName $magicSheet -ErrorAction Stop
    } catch {
        $base = @()
    }

    # helper: split rows tagged MagicType=Rune into separate Common vs Special lists
    $splitRune = {
        param($rows, $fromPrefix)

        $runeRows = @(
            $rows | Where-Object {
                ('' + $_.MagicType) -match '^(?i)rune$'
            }
        )

        $commonList = @(
            $runeRows |
            Where-Object { ('' + $_.Access) -match '(?i)\bcommon\b' } |
            ForEach-Object {
                [pscustomobject]@{
                    Name     = ('' + $_.Spell).Trim()
                    FromCult = $fromPrefix
                }
            }
        )

        $specialList = @(
            $runeRows |
            Where-Object { -not ( ('' + $_.Access) -match '(?i)\bcommon\b' ) } |
            ForEach-Object {
                [pscustomobject]@{
                    Name     = ('' + $_.Spell).Trim()
                    FromCult = $fromPrefix
                }
            }
        )

        ,$commonList, $specialList
    }

    $common  = @()
    $special = @()

    $cBase, $sBase = & $splitRune $base $prefix
    $common  += $cBase
    $special += $sBase

    #
    # STEP 2: Pull in Associate spells (if requested)
    # Rule (per you): Whatever is in Provides gets injected as SPECIAL, literally,
    # and is attributed to the associate cult.
    #
    if ($IncludeAssociates) {
        $assocRows = @()
        try {
            $assocRows = Import-Excel -Path $WorkbookPath -WorksheetName $assocSheet -ErrorAction Stop
        } catch {
            $assocRows = @()
        }

        foreach ($a in @($assocRows)) {
            $fromCultRaw = ('' + $a.FromCult).Trim()
            if (-not $fromCultRaw) { continue }

            # normalize the associate cult prefix for labeling consistency
            $associatePrefix = Resolve-CultSheetName -CultName $fromCultRaw -WorkbookPath $WorkbookPath

            $providesRaw = '' + $a.Provides
            if (-not $providesRaw) { continue }

            # split "Summon Specific Ancestor; Crush" etc.
            $providedNames = @(
                $providesRaw -split '[,;]' |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ }
            )

            foreach ($spellName in $providedNames) {
                # add as a SPECIAL rune spell, verbatim name
                $special += [pscustomobject]@{
                    Name     = $spellName
                    FromCult = $associatePrefix
                }
            }
        }
    }

    #
    # STEP 3: De-dupe (case-insensitive by Name).
    #
    $common  = Deduplicate-ByName -Items $common
    $special = Deduplicate-ByName -Items $special

    [pscustomobject]@{
        Common  = $common
        Special = $special
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
    if ($RunePoints -le 0) { return ,@(), @() }   # ⬅ ensure TWO return objects

    $catalog = Get-CultRuneSpellCatalog -CultName $CultName -WorkbookPath $WorkbookPath -IncludeAssociates:$IncludeAssociates
    $special = @($catalog.Special)
    $common  = @($catalog.Common)

    if ($special.Count -eq 0) { return ,@(), @($common) }  # ⬅ TWO objects (empty special, common array)

    $out = New-Object System.Collections.Generic.List[object]
# Build a normalized, case-insensitive unique pool of names
# Build a normalized, case-insensitive unique pool of names
$pool = @(
    $special | ForEach-Object { ('' + $_.Name).Trim() } |
    Sort-Object -Unique -CaseSensitive:$false
)

for ($i=0; $i -lt $RunePoints; $i++) {
    if ($pool.Count -eq 0) { break }  # 🚫 no repeats; stop when uniques are exhausted
    $pick = Get-Random -InputObject $pool
    $pool = $pool | Where-Object { $_ -ne $pick }
    $out.Add([pscustomobject]@{ Name = $pick })
}



    # ⬅ FINAL RETURN: two separate array objects
    $specialArr = @($out.ToArray())   # array of {Name=...}
    $commonArr  = @($common)          # array of {Name=..., FromCult=...} etc.
    return ,$specialArr, $commonArr
}
function ConvertTo-SpellObjectList {
    [CmdletBinding()]
    param($Value)

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($Value)) {
        if ($null -eq $item) { continue }

        # Plain string -> wrap
        if ($item -is [string]) {
            $out.Add([pscustomobject]@{ Name = ('' + $item) })
            continue
        }

        # If it has a Name property, normalize
        $nameProp = $item | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
        if ($nameProp) {
            if ($nameProp -is [System.Array]) {
                foreach ($n in $nameProp) { if ($n) { $out.Add([pscustomobject]@{ Name = ('' + $n) }) } }
            } else {
                $out.Add([pscustomobject]@{ Name = ('' + $nameProp) })
            }
            continue
        }

        # If it's an enumerable of strings/objects, expand
        if ($item -is [System.Collections.IEnumerable] -and $item -isnot [string]) {
            foreach ($x in $item) {
                if ($null -eq $x) { continue }
                if ($x -is [string]) { $out.Add([pscustomobject]@{ Name = ('' + $x) }) }
                else {
                    $n2 = $x | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
                    if ($n2) { $out.Add([pscustomobject]@{ Name = ('' + $n2) }) }
                    else     { $out.Add([pscustomobject]@{ Name = ('' + $x) }) }
                }
            }
            continue
        }

        # Fallback
        $out.Add([pscustomobject]@{ Name = ('' + $item) })
    }

    ,$out.ToArray()
}


function Set-StatblockRuneMagic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] $Statblock,
        [Parameter(Mandatory, Position=1)] [int]$RunePoints,
        [Parameter(Position=2)] [object[]]$RuneSpecialSpells = @(),
        [Parameter(Position=3)] [object[]]$RuneCommonSpells  = @()
    )

    if (-not $Statblock.PSObject.Properties['Magic']) {
        $Statblock.PSObject.Properties.Add(
            [System.Management.Automation.PSNoteProperty]::new('Magic', @{})
        )
    }
    if ($Statblock.Magic -isnot [System.Collections.IDictionary]) {
        $tmp = @{}; foreach ($p in $Statblock.Magic.PSObject.Properties) { $tmp[$p.Name] = $p.Value }; $Statblock.Magic = $tmp
    }

    $spec = ConvertTo-SpellObjectList $RuneSpecialSpells
    $comm = ConvertTo-SpellObjectList $RuneCommonSpells

    $Statblock.Magic['RunePoints']  = [int]$RunePoints
    $Statblock.Magic['RuneSpecial'] = $spec
    $Statblock.Magic['RuneCommon']  = $comm

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
    Test-CultGrantsFullPriceSpirit, `
    ConvertTo-SpellObjectList, `
    Get-CultSpiritCatalogSlim `
    -ErrorAction SilentlyContinue 


