$ErrorActionPreference = 'Stop'

function Add-CultInfoToStatblock {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][object]$Statblock,   # your $sb object
    [Parameter(Mandatory)][string]$CultName,
    [string]$Role = 'Initiate'                  # Lay | Initiate | RuneLord | Priest | Shaman
  )

  if (-not $Statblock) { return $Statblock }

  # ----------------------------
  # helpers (local/private)
  # ----------------------------
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
    Get-Random -Minimum $min -Maximum ($max + 1)
  }

  function _SpiritBudget([string]$normRole, [int]$CHA, [int]$INT){
    if ($INT -le 0) { return [pscustomobject]@{ Points=0; Slots=0 } }

    switch ($normRole) {
      'Lay' {
        $pts = _Roll 2 4
      }
      'Initiate' {
        $pts = _Roll 5 10
        if ((Get-Random -Minimum 1 -Maximum 101) -le 30) {
          $pts = [Math]::Min($pts + (_Roll 1 3), 13)
        }
      }
      'RuneLord' {
        $base = [int][Math]::Round($CHA * (_Roll 75 100) / 100.0)
        $pts  = [Math]::Max([Math]::Min($base, $CHA), 8)
      }
      'Priest' {
        $base = [int][Math]::Round($CHA * (_Roll 85 100) / 100.0)
        $pts  = [Math]::Max([Math]::Min($base, $CHA), 10)
      }
      'Shaman' {
        $base = [int][Math]::Round($CHA * (_Roll 85 100) / 100.0)
        $pts  = [Math]::Max([Math]::Min($base, $CHA), 12)
      }
      default {
        $pts = _Roll 5 10
      }
    }

    if ($pts -lt 0) { $pts = 0 }
    $slots = [Math]::Min($pts, [Math]::Max($CHA,0))
    [pscustomobject]@{ Points = $pts; Slots = $slots }
  }

  # ----------------------------
  # pull data
  # ----------------------------
  $magic = Get-CultMagic -CultName $CultName
  $roles = Get-CultRoles -CultName $CultName

  $normRole = _NormRole $Role

  $roleRow = $roles | Where-Object { $_.Role -like "*$Role*" } | Select-Object -First 1
  if (-not $roleRow) { $roleRow = $roles | Select-Object -First 1 }

  # Filter by access
  $hasRole = {
    param($row,$r)
    if (-not $row.Access) { return $true }
    ($row.Access -split '\s*;\s*') -contains $r
  }

  $spiritRows = $magic | Where-Object { $_.MagicType -match '^(?i)spirit$' -and (& $hasRole $_ $Role) }
  $runeRows   = $magic | Where-Object { $_.MagicType -match '^(?i)rune|ritual$' -and (& $hasRole $_ $Role) }

  # ----------------------------
  # allocate Spirit Magic per your rules
  # ----------------------------
  $cha = [int]$Statblock.Characteristics.CHA
  $int = [int]$Statblock.Characteristics.INT
  $budget = _SpiritBudget -normRole $normRole -CHA $cha -INT $int

  $hasAll = $false
  if ($spiritRows) {
    $hasAll = ($spiritRows | Where-Object { $_.Spell -match '^\s*All\s+spirit\s+magic' }).Count -gt 0
  }

  $spiritTxt = $null
  if ($budget.Points -gt 0 -and $budget.Slots -gt 0) {
    if ($hasAll) {
      $plural = if ($budget.Slots -eq 1) { '' } else { 's' }
      $spiritTxt = ('All spirit magic available; allocating ~{0} points across {1} spell{2} (placeholder)' -f $budget.Points, $budget.Slots, $plural)
    } else {
      $allowed = @()
      if ($spiritRows) {
        $allowed = $spiritRows |
          Where-Object { $_.Spell -notmatch '^\s*All\s+spirit\s+magic' } |
          Select-Object -ExpandProperty Spell -Unique
      }
      $plural = if ($budget.Slots -eq 1) { '' } else { 's' }
      if ($allowed.Count -gt 0) {
        $spiritTxt = ('{0} points / {1} slot{2} from: {3} (placeholder)' -f $budget.Points, $budget.Slots, $plural, ($allowed -join ', '))
      } else {
        $spiritTxt = ('Allocating ~{0} points across {1} spell{2} (placeholder)' -f $budget.Points, $budget.Slots, $plural)
      }
    }
  }

  # ----------------------------
  # Rune spells (no budgeting yet)
  # ----------------------------
  $runeTxt = $null
  if ($runeRows) {
    $buf = New-Object System.Collections.Generic.List[string]
    foreach ($row in $runeRows) {
      $spell = ($row.Spell   | ForEach-Object { '' + $_ }).Trim()
      $pts   = ($row.Points  | ForEach-Object { '' + $_ }).Trim()
      $tags  = ($row.Tags    | ForEach-Object { '' + $_ }).Trim()

      $piece = $spell
      if ($pts)  { $piece = ('{0} ({1}pt)' -f $piece, $pts) }
      if ($tags) { $piece = ('{0} [{1}]'  -f $piece, $tags) }
      if ($piece) { [void]$buf.Add($piece) }
    }
    if ($buf.Count -gt 0) { $runeTxt = ($buf -join '; ') }
  }

  # ----------------------------
  # patch statblock
  # ----------------------------
  $rankText = if ($roleRow -and $roleRow.Rank) { $roleRow.Rank } else { $Role }
  $cultHeader = ('Cult: {0} ({1})' -f $CultName, $rankText)

  $magicLines = @()
  if ($spiritTxt) { $magicLines += ('Spirit: {0}' -f $spiritTxt) }
  if ($runeTxt)   { $magicLines += ('Rune: {0}'   -f $runeTxt) }
  $magicText = ($magicLines -join ' | ')

  if ($magicText) {
    $Statblock.Magic = @($Statblock.Magic, $cultHeader, $magicText) -ne $null -join "`n"
  } else {
    $Statblock.Magic = @($Statblock.Magic, $cultHeader) -ne $null -join "`n"
  }

  $Statblock.MagicNotes = @($Statblock.MagicNotes, 'Spirit magic allocation auto-budgeted by role/CHA.') -ne $null -join ' '
# --- Auto-assign Spirit Magic (always-on for now) -----------------------------
try {
    # 1) Load catalog
    $catalogPath = "Y:\Stat_blocks\Data\spirit_magic_catalog.csv"
    if (Test-Path $catalogPath) {
        $cat = Import-SpiritMagicCatalog -CsvPath $catalogPath
    } else {
        throw "Spirit-magic catalog not found at $catalogPath"
    }

    # 2) Skip if we've already assigned Spirit (avoid duplicates on re-run)
    $already = $false
    if ($Statblock.Magic -is [System.Collections.IDictionary]) {
        $already = ($Statblock.Magic.ContainsKey('Spirit') -and $Statblock.Magic['Spirit'] -and $Statblock.Magic['Spirit'].Count -gt 0)
    } elseif ($Statblock.Magic) {
        $already = ($null -ne $Statblock.Magic.Spirit -and $Statblock.Magic.Spirit.Count -gt 0)
    }

    if (-not $already -and $cat -and $cat.Count -gt 0) {
        # 3) Budget: simple role-based default, capped by CHA (longhand)
        $cha    = [int]$Statblock.Characteristics.CHA
        $role   = $Role
        $budget = Get-SpiritBudgetByRole -Role $role -CHA $cha  # tweak ranges in the helper as you like

        if ($budget -gt 0) {
            # 4) Roll & apply (role caps respected via CSV RoleMax_* columns)
            $rolls = New-RandomSpiritMagicLoadout -PointsBudget $budget -CHA $cha -Role $role -Catalog $cat -Seed (Get-Random)
            $Statblock = Set-StatblockSpiritMagic $Statblock $rolls
        }
    }
} catch {
    Write-Verbose "Spirit-magic auto-assign skipped: $($_.Exception.Message)"
}
# ----------------------------------------------------------------------------- 

  return $Statblock
}

Export-ModuleMember -Function Add-CultInfoToStatblock
