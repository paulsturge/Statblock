#requires -Modules ImportExcel
Import-Module ImportExcel -ErrorAction Stop

$script:ModuleRoot = $PSScriptRoot
$script:DataRoot   = Split-Path -Parent $script:ModuleRoot

function Resolve-StatPath { param([Parameter(Mandatory)][string]$Name) Join-Path -Path $script:DataRoot -ChildPath $Name }

function Initialize-StatblockContext {
  param([string]$DataRootOverride)
  if ($PSBoundParameters.ContainsKey('DataRootOverride')) { $script:DataRoot = $DataRootOverride }
  $hp      = Import-Excel (Resolve-StatPath 'hp_modifier_table.xlsx')      | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $chaos   = Import-Excel (Resolve-StatPath 'Chaotic_features.xlsx')       | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $stat    = Import-Excel (Resolve-StatPath 'Stat_Dice_Source.xlsx')       | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $melee   = Import-Excel (Resolve-StatPath 'weapons_melee_table.xlsx')    | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $missile = Import-Excel (Resolve-StatPath 'weapons_missile_table.xlsx')  | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $shields = Import-Excel (Resolve-StatPath 'weapons_shields_table.xlsx')  | Where-Object { $_.PSObject.Properties.Value -ne $null }
  [pscustomobject]@{
    HpTable  = $hp
    Chaos    = $chaos
    StatDice = $stat
    Weapons  = @{ Melee = $melee; Missile = $missile; Shields = $shields }
    Folder   = $script:DataRoot
  }
}

function Invoke-DiceRoll {
  param([Parameter(Mandatory)][int]$Count, [int]$Faces = 6)
  $t = 0
  1..$Count | ForEach-Object { $t += Get-Random -Minimum 1 -Maximum ($Faces + 1) }
  $t
}


function Get-StatRow { param($Context,[Parameter(Mandatory)][string]$Creature)
  $row = $Context.StatDice | Where-Object { $_.Creature -eq $Creature } | Select-Object -First 1
  if (-not $row) { throw "Creature '$Creature' not found in Stat_Dice_Source.xlsx" }
  $row
}

function Get-StatRoll { param($Row,[Parameter(Mandatory)][string]$StatName)
  $n = [int]($Row.$StatName)
  $m = [int]($Row.("$StatName`Mod"))
  if ($n -le 0) { return 0 }
  (Invoke-DiceRoll -Count $n -Faces 6) + $m
}

function New-Characteristics { param($Row)
  [ordered]@{
    STR = Get-StatRoll $Row 'STR'
    CON = Get-StatRoll $Row 'CON'
    SIZ = Get-StatRoll $Row 'SIZ'
    DEX = Get-StatRoll $Row 'DEX'
    INT = Get-StatRoll $Row 'INT'
    POW = Get-StatRoll $Row 'POW'
    CHA = Get-StatRoll $Row 'CHA'
  }
}

function Get-StrikeRanks { param([Parameter(Mandatory)][int]$Dex,[Parameter(Mandatory)][int]$Siz)
  $dexSR = switch ($Dex) { {$_ -in 1..5}{5}; {$_ -in 6..8}{4}; {$_ -in 9..12}{3}; {$_ -in 13..15}{2}; {$_ -in 16..18}{1}; default {0} }
  $sizSR = switch ($Siz) { {$_ -in 1..6}{3}; {$_ -in 7..14}{2}; {$_ -in 15..21}{1}; default {0} }
  [pscustomobject]@{ DexSR = $dexSR; SizSR = $sizSR; Base = $dexSR + $sizSR }
}

function Get-HitPoints { param([Parameter(Mandatory)][int]$CON,[Parameter(Mandatory)][int]$SIZ,[Parameter(Mandatory)][int]$POW,$HpTable)
  $hp = $CON
  if ($SIZ -gt 28) { $sizMod = 4 + [math]::Ceiling(($SIZ - 28)/4) } else { $sizMod = [int]($HpTable | Where-Object Stat_value -eq $SIZ | Select-Object -First 1).SIZ }
  if ($POW -gt 28) { $powMod = 3 + [math]::Ceiling(($POW - 28)/4) } else { $powMod = [int]($HpTable | Where-Object Stat_value -eq $POW | Select-Object -First 1).POW }
  $hp + $sizMod + $powMod
}

function Get-DamageBonus { param([Parameter(Mandatory)][int]$STR,[Parameter(Mandatory)][int]$SIZ)
  $sum = $STR + $SIZ
  switch ($sum) {
    {$_ -le 12}     { '-1d4' }
    {$_ -in 13..24} { '' }
    {$_ -in 25..32} { '+1D4' }
    {$_ -in 33..40} { '+1D6' }
    {$_ -in 41..56} { '+2D6' }
    default { "+$((2 + [math]::Ceiling(($sum - 56)/16)))D6" }
  }
}

function Get-HalfDamageBonus { param([string]$Db)
  if ($null -eq $Db) { $db = '' } else { $db = $Db.Trim() }
  if ([string]::IsNullOrWhiteSpace($db)) { return '' }
  if ($db -match '^-1[dD]4$') { return '-1d2' }  # special case
  $m = [regex]::Match($db, '^[+]?(\\d+)[dD](\\d+)$')
  if (-not $m.Success) { return '' }
  $n     = [int]$m.Groups[1].Value
  $faces = [int]$m.Groups[2].Value
  $half  = [int][math]::Ceiling($faces / 2.0); if ($half -lt 2) { $half = 2 }
  "+${n}D$half"
}

function Get-SpiritCombatDamage { param([Parameter(Mandatory)][int]$POW,[Parameter(Mandatory)][int]$CHA)
  $pc = $POW + $CHA
  switch ($pc) {
    {$_ -in 2..12}  { '1d3' }
    {$_ -in 13..24} { '1d6' }
    {$_ -in 25..32} { '1d6+1' }
    {$_ -in 33..40} { '1d6+3' }
    {$_ -in 41..56} { '2d6+3' }
    default { "$((2 + [math]::Floor(($pc - 56)/16)))d6+$((3 + [math]::Floor(($pc - 56)/16)))" }
  }
}

function Get-HitLocations {
  param($Context,[Parameter(Mandatory)][string]$Sheet,[Parameter(Mandatory)][int]$HP,[int]$AddArmor = 0)
  $path  = Resolve-StatPath 'Hit_Location_Source.xlsx'
  $locs  = Import-Excel -Path $path -WorksheetName $Sheet | Where-Object { $_.PSObject.Properties.Value -ne $null } | ForEach-Object { $_ | Select-Object * }
  if ($HP -lt 13 -or $HP -gt 15) {
    if ($HP -lt 13) { $delta = [math]::Floor(($HP - 13) / 3) } else { $delta = [math]::Ceiling(($HP - 15) / 3) }
    foreach ($l in $locs) { $l.HP = [int]$l.HP + $delta }
  }
  if ($AddArmor -ne 0) { foreach ($l in $locs) { $l.armor = [int]$l.armor + $AddArmor } }
  $locs
}

function Get-Weapons {
  param(
    $Context,
    [Parameter(Mandatory)][string]$Creature,
    [Parameter(Mandatory)][int]$BaseSR,

    # allow empty/no-DB
    [AllowNull()][AllowEmptyString()]
    [string]$DamageBonus = ''
  )

  $useDB = -not [string]::IsNullOrWhiteSpace($DamageBonus)

  $file = Resolve-StatPath ("{0}_weapons.txt" -f $Creature)
  if (-not (Test-Path $file)) { return @() }
  $lines = Get-Content -Path $file | Where-Object { $_ -and (-not $_.StartsWith('#')) }
  $out   = New-Object System.Collections.Generic.List[object]

  foreach ($line in $lines) {
    $parts = $line.Split('_', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -lt 2) { continue }
    $name = $parts[0]
    $type = $parts[1].ToLowerInvariant()

    if ($type -eq 'me') {
      $row = $Context.Weapons.Melee | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
      if ($null -ne $row) {
        $row = $row | Select-Object *
        if ($useDB) { $row.Damage = "$($row.Damage)$DamageBonus" }
        $row.SR = [int]$row.SR + $BaseSR
        $out.Add($row)
      }
    }
    elseif ($type -eq 'mi') {
      $row = $Context.Weapons.Missile | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
      if ($null -ne $row) {
        $row = $row | Select-Object *
        $row.SR = 0 + $BaseSR
        if ($parts.Count -ge 3 -and $parts[2].ToLowerInvariant() -eq 'th' -and $useDB) {
          $half = Get-HalfDamageBonus $DamageBonus
          if ($half) { $row.Damage = "$($row.Damage)$half" }
        }
        $out.Add($row)
      }
    }
    elseif ($type -eq 'sh') {
      $row = $Context.Weapons.Shields | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
      if ($null -ne $row) {
        $row = $row | Select-Object *
        if ($useDB) { $row.Damage = "$($row.Damage)$DamageBonus" }
        $row.SR = [int]$row.SR + $BaseSR
        $out.Add($row)
      }
    }
  }
  $out.ToArray()
}


function New-Statblock {
  param([Parameter(Mandatory)][string]$Creature,[Parameter(Mandatory)]$Context,[int]$AddArmor = 0,[string]$OverrideHitLocationSheet)
  $row   = Get-StatRow -Context $Context -Creature $Creature
  $chars = New-Characteristics -Row $row
  $sr    = Get-StrikeRanks -Dex $chars.DEX -Siz $chars.SIZ
  $hp    = Get-HitPoints -CON $chars.CON -SIZ $chars.SIZ -POW $chars.POW -HpTable $Context.HpTable
  $db    = Get-DamageBonus -STR $chars.STR -SIZ $chars.SIZ
  $scd   = Get-SpiritCombatDamage -POW $chars.POW -CHA $chars.CHA
  if ($OverrideHitLocationSheet) { $sheet = $OverrideHitLocationSheet } else {
    $sheet = [string]$row.Hit_location
    if ($Creature -eq 'Dragonsnail') { $r = Get-Random -Minimum 1 -Maximum 101; if ($r -gt 65) { $sheet = 'Dragonsnail1' } else { $sheet = 'Dragonsnail' } }
  }
  $hitLocs = Get-HitLocations -Context $Context -Sheet $sheet -HP $hp -AddArmor $AddArmor
  $weapons = Get-Weapons -Context $Context -Creature $Creature -BaseSR $sr.Base -DamageBonus $db
  [pscustomobject]@{
    Creature        = $Creature
    Move            = [int]$row.Move
    Runes1          = $row.Runes1
    Rune1Score      = $row.Rune1score
    Runes2          = $row.Runes2
    Rune2Score      = $row.Rune2score
    Characteristics = [pscustomobject]([ordered]@{ STR=$chars.STR; CON=$chars.CON; SIZ=$chars.SIZ; DEX=$chars.DEX; INT=$chars.INT; POW=$chars.POW; CHA=$chars.CHA })
    HP              = $hp
    StrikeRanks     = $sr
    DamageBonus     = $db
    SpiritCombat    = $scd
    HitLocations    = $hitLocs
    Weapons         = $weapons
  }
}

Export-ModuleMember -Function Initialize-StatblockContext, New-Statblock, Roll-Dice, Get-StrikeRanks, Get-StatRow, Get-StatRoll, New-Characteristics, Get-HitPoints, Get-DamageBonus, Get-HalfDamageBonus, Get-SpiritCombatDamage, Get-HitLocations, Get-Weapons
