$ErrorActionPreference = 'Stop'

function Add-CultInfoToStatblock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Statblock,
        [Parameter(Mandatory)][string]$CultName,
        [string]$Role = 'Initiate'
    )

    if (-not $Statblock) { return $Statblock }

    #
    # --- helpers ---------------------------------------------------------
    #
    function _NormRole([string]$r){
        if (-not $r) { return 'Initiate' }
        switch -Regex ($r) {
            '^(?i)lay'                                 { return 'Lay' }
            '^(?i)init'                                { return 'Initiate' }
            '^(?i)doomed'                              { return 'Initiate' }
            '^(?i)rune\s*lord|doom\s*master|jaw |karrg'       { return 'RuneLord' }
            '^(?i)priest|tongue|hand|horn|breath|high' { return 'Priest' }
            '^(?i)shaman'                              { return 'Shaman' }
            default                                    { return 'Initiate' }
        }
    }

    function _CatalogRole([string]$normRole){
    switch ($normRole) {
        'Shaman' { return 'Priest' }   # use Priest caps/columns for Shaman in the catalog
        default  { return $normRole }  # Lay, Initiate, RuneLord, Priest stay themselves
    }
}
    function _HasRoleAccess {
        param($row,$roleRaw,$roleNorm)

        if (-not $row.Access) { return $true }   # no Access = everyone

        $accessList = ('' + $row.Access).Split(';') |
                      ForEach-Object { $_.Trim() } |
                      Where-Object { $_ -ne '' }

        if ($accessList -contains $roleRaw)  { return $true }
        if ($accessList -contains $roleNorm) { return $true }

        foreach ($acc in $accessList) {
            if ($acc -match [regex]::Escape($roleRaw))  { return $true }
            if ($acc -match [regex]::Escape($roleNorm)) { return $true }
        }
        return $false
    }

   $normRole    = _NormRole $Role
$catalogRole = _CatalogRole $normRole

    #
    # --- ensure core properties on $Statblock ----------------------------
    #
    if ($Statblock.PSObject.Properties['CultName'].Count -eq 0) {
        $Statblock | Add-Member -NotePropertyName CultName -NotePropertyValue $CultName
    } else {
        $Statblock.CultName = $CultName
    }

    # Role rank text from roles table
    $rolesData = @()
    try {
        $rolesData = Get-CultRoles -CultName $CultName
    } catch {
        $rolesData = @()
    }
    $roleRow = $rolesData | Where-Object { $_.Role -like "*$Role*" } | Select-Object -First 1
    if (-not $roleRow) { $roleRow = $rolesData | Select-Object -First 1 }
    $rankText = if ($roleRow -and $roleRow.Rank) { $roleRow.Rank } else { $Role }

    if ($Statblock.PSObject.Properties['CultRole'].Count -eq 0) {
        $Statblock | Add-Member -NotePropertyName CultRole -NotePropertyValue $rankText
    } else {
        $Statblock.CultRole = $rankText
    }

    # Magic container
    if ($Statblock.PSObject.Properties['Magic'].Count -eq 0) {
        $Statblock | Add-Member -NotePropertyName Magic -NotePropertyValue @{}
    }
    if ($Statblock.Magic -isnot [System.Collections.IDictionary]) {
        $Statblock.Magic = @{}
    }

    # RuneMagic container
    if ($Statblock.PSObject.Properties['RuneMagic'].Count -eq 0) {
        $Statblock | Add-Member -NotePropertyName RuneMagic -NotePropertyValue @{}
    }
    if ($Statblock.RuneMagic -isnot [System.Collections.IDictionary] -and
        $Statblock.RuneMagic -isnot [psobject]) {
        $Statblock.RuneMagic = @{}
    }

    # reset magic slots each run
    $Statblock.Magic['Spirit'] = @()
    $Statblock.Magic['Notes']  = $null
    $Statblock.RuneMagic       = @{}

    #
    # --- pull cult magic sheet once -------------------------------------
    #
    $allMagic = @()
    try {
        $allMagic = Get-CultMagic -CultName $CultName
    } catch {
        $allMagic = @()
    }

    #
    # === SPIRIT MAGIC (using new Allowed field) ==========================
    #
    try {
        $cha = 0
        if ($Statblock.Characteristics -and $Statblock.Characteristics.CHA) {
            $cha = [int]$Statblock.Characteristics.CHA
        }

        # Points budget based on role & CHA
        $budget = 0
        try {
            $budget = Get-SpiritBudgetByRole -Role $normRole -CHA $cha
        } catch {
            $budget = 0
        }

        $catalogPath = "Y:\Stat_blocks\Data\spirit_magic_catalog.csv"
        $cat = @()
        if (Test-Path $catalogPath) {
            $cat = Import-SpiritMagicCatalog -CsvPath $catalogPath
        }

        $finalLoadout    = @()
        $remainingBudget = $budget

        if ($remainingBudget -gt 0 -and $cat -and $cat.Count -gt 0) {

            # 1) Cult spirit rows
            $spiritRows = $allMagic | Where-Object {
                $_.MagicType -match '^(?i)spirit$'
            }

            # drop Prohibited
            $spiritRows = $spiritRows | Where-Object {
                -not (('' + $_.Allowed) -match '^(?i)prohibited$')
            }

            # primary pool: Allowed/Common/Special
            $primaryCult = $spiritRows | Where-Object {
                ('' + $_.Allowed) -match '^(?i)(allowed|common|special)$'
            }

            # fallback pool: everything else (blank Allowed), still not prohibited
            $fallbackCult = $spiritRows | Where-Object {
                -not (('' + $_.Allowed) -match '^(?i)(allowed|common|special|prohibited)$')
            }

            $primaryNames = @(
                $primaryCult |
                Select-Object -ExpandProperty Spell -Unique |
                ForEach-Object { ('' + $_).Trim() } |
                Where-Object { $_ -ne '' }
            )

            $fallbackNames = @(
                $fallbackCult |
                Select-Object -ExpandProperty Spell -Unique |
                ForEach-Object { ('' + $_).Trim() } |
                Where-Object { $_ -ne '' }
            )

            # 1a) spend budget on cult-primary spirit spells (via global catalog)
            if ($primaryNames.Count -gt 0) {
                $catPrimary = @(
                    $cat | Where-Object { $primaryNames -contains $_.Name }
                )

                if ($catPrimary -and $catPrimary.Count -gt 0) {
                  $loadPrimary = New-RandomSpiritMagicLoadout `
                    -PointsBudget $remainingBudget `
                    -CHA          $cha `
                    -Role         $catalogRole `
                    -Catalog      $catPrimary `
                    -Seed         (Get-Random)

                    if ($loadPrimary -and $loadPrimary.Count -gt 0) {
                        $finalLoadout += $loadPrimary

                        $spent = ($loadPrimary | Measure-Object -Property Points -Sum).Sum
                        if ($null -eq $spent) { $spent = 0 }
                        $remainingBudget = [Math]::Max(0, $remainingBudget - [int]$spent)
                    }
                }
            }

            # 1b) spend remaining budget on cult-fallback spirit spells
            if ($remainingBudget -gt 0 -and $fallbackNames.Count -gt 0) {

                $usedNames = @(
                    $finalLoadout |
                    Where-Object { $_.Name } |
                    Select-Object -ExpandProperty Name -Unique
                )

                $catFallback = @(
                    $cat | Where-Object {
                        $fallbackNames -contains $_.Name -and
                        -not ($usedNames -contains $_.Name)
                    }
                )

                if ($catFallback -and $catFallback.Count -gt 0) {
                    $loadFallback = New-RandomSpiritMagicLoadout `
                    -PointsBudget $remainingBudget `
                    -CHA          $cha `
                    -Role         $catalogRole `
                    -Catalog      $catFallback `
                    -Seed         (Get-Random)


                    if ($loadFallback -and $loadFallback.Count -gt 0) {
                        $finalLoadout += $loadFallback

                        $spent2 = ($loadFallback | Measure-Object -Property Points -Sum).Sum
                        if ($null -eq $spent2) { $spent2 = 0 }
                        $remainingBudget = [Math]::Max(0, $remainingBudget - [int]$spent2)
                    }
                }
            }

            # 1c) if still budget left, fall back to full catalog
            if ($remainingBudget -gt 0) {
                $usedNames = @(
                    $finalLoadout |
                    Where-Object { $_.Name } |
                    Select-Object -ExpandProperty Name -Unique
                )

                $catGlobal = @(
                    $cat | Where-Object { -not ($usedNames -contains $_.Name) }
                )

                if ($catGlobal -and $catGlobal.Count -gt 0) {
                  $loadGlobal = New-RandomSpiritMagicLoadout `
                    -PointsBudget $remainingBudget `
                    -CHA          $cha `
                     -Role         $catalogRole `
                    -Catalog      $catGlobal `
                    -Seed         (Get-Random)

                    if ($loadGlobal -and $loadGlobal.Count -gt 0) {
                        $finalLoadout += $loadGlobal
                    }
                }
            }
        }

        if ($finalLoadout -and $finalLoadout.Count -gt 0) {
            $Statblock = Set-StatblockSpiritMagic $Statblock $finalLoadout
            $Statblock.Magic['Notes'] = "Spirit magic loadout auto-budgeted for $rankText of $CultName (cult-first)."
        }
        else {
            $Statblock.Magic['Notes'] = "No spirit magic assigned."
        }

    } catch {
        $Statblock.Magic['Notes'] = "No spirit magic assigned (generation error)."
    }

    #
    # === RUNE MAGIC (Common vs Special via Allowed) ======================
    #
    $runeRows = $allMagic | Where-Object {
        $_.MagicType -match '(?i)rune|ritual' -and (_HasRoleAccess $_ $Role $normRole)
    }
    if (-not $runeRows -or $runeRows.Count -eq 0) {
        $runeRows = $allMagic | Where-Object {
            $_.MagicType -match '(?i)rune|ritual'
        }
    }

    $specialList = @()
    $commonList  = @()

    foreach ($row in $runeRows) {
        $spellName = ('' + $row.Spell).Trim()
        if (-not $spellName) { continue }

        # drop Prohibited rune spells
        if (('' + $row.Allowed) -match '^(?i)prohibited$') { continue }

        $pts  = ('' + $row.Points).Trim()
        $flag = ('' + $row.Allowed).Trim()

        $assembled = $spellName
        if ($pts) { $assembled = ('{0} ({1}pt)' -f $assembled, $pts) }

        if ($flag -match '^(?i)special$') {
            $specialList += $assembled
        }
        elseif ($flag -match '^(?i)common|allowed$' -or $flag -eq '') {
            $commonList  += $assembled
        }
        else {
            # any other flag we treat as common-ish for now
            $commonList  += $assembled
        }
    }

    $specialList = $specialList | Select-Object -Unique
    $commonList  = $commonList  | Select-Object -Unique

    $finalRuneList = @()
    if ($specialList.Count -gt 0) {
        $finalRuneList = $specialList
    } elseif ($commonList.Count -gt 0) {
        $finalRuneList = $commonList
    }

    if ($finalRuneList.Count -gt 0) {
        $Statblock.RuneMagic = [pscustomobject]@{
            Spells  = @(
                foreach ($rn in $finalRuneList) {
                    [pscustomobject]@{ Name = $rn }
                }
            )
            Special = $null
        }
    } else {
        $Statblock.RuneMagic = @{}
    }

    #
    # --- MagicNotes ------------------------------------------------------
    #
    $noteText = "Spirit magic is a rolled loadout (cult-specific lists first, then general). Rune Magic shows cult-legal rune spells by role."
    if ($Statblock.PSObject.Properties['MagicNotes'].Count -eq 0) {
        $Statblock | Add-Member -NotePropertyName MagicNotes -NotePropertyValue $noteText
    } else {
        $Statblock.MagicNotes = $noteText
    }

    return $Statblock
}

Export-ModuleMember -Function Add-CultInfoToStatblock
