#requires -Modules ImportExcel
Import-Module ImportExcel -ErrorAction Stop

$script:ModuleRoot = $PSScriptRoot
$script:DataRoot   = Split-Path -Parent $script:ModuleRoot

function Resolve-StatPath {
  param([Parameter(Mandatory)][string]$Name)

  if ([System.IO.Path]::IsPathRooted($Name)) {
    if (Test-Path -LiteralPath $Name) { return (Resolve-Path -LiteralPath $Name).Path }
    return $null
  }

  $candidates = @()
  if ($script:DataRoot) { $candidates += $script:DataRoot }
  if ($PSScriptRoot) { $candidates += $PSScriptRoot; $candidates += (Split-Path $PSScriptRoot -Parent) }
  if ($pwd) { $candidates += $pwd.Path }

  foreach ($root in ($candidates | Select-Object -Unique)) {
    $p = Join-Path $root $Name
    if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
  }
  return $null
}


function Initialize-StatblockContext {
  param([string]$DataRootOverride)
  if ($PSBoundParameters.ContainsKey('DataRootOverride')) { $script:DataRoot = $DataRootOverride }
  $hp      = Import-Excel (Resolve-StatPath 'hp_modifier_table.xlsx')      | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $chaos   = Import-Excel (Resolve-StatPath 'Chaotic_features.xlsx')       | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $stat    = Import-Excel (Resolve-StatPath 'Stat_Dice_Source.xlsx')       | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $melee   = Import-Excel (Resolve-StatPath 'weapons_melee_table.xlsx')    | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $missile = Import-Excel (Resolve-StatPath 'weapons_missile_table.xlsx')  | Where-Object { $_.PSObject.Properties.Value -ne $null }
  $shields = Import-Excel (Resolve-StatPath 'weapons_shields_table.xlsx')  | Where-Object { $_.PSObject.Properties.Value -ne $null }
 
  $curse = $null
  $cfPath = Resolve-StatPath 'Chaotic_features.xlsx'
  if ($cfPath) {
    $curse = Import-Excel -Path $cfPath -WorksheetName 'Curse of Thed' -ErrorAction Stop |
             Where-Object { $_.PSObject.Properties.Value -ne $null }
  }
  
  [pscustomobject]@{
    HpTable  = $hp
    Chaos    = $chaos
    StatDice = $stat
    Weapons  = @{ Melee = $melee; Missile = $missile; Shields = $shields }
    Folder   = $script:DataRoot
    CurseOfThed = $curse
  }
}
function ConvertFrom-DiceSpec {
  param([Parameter(Mandatory)][string]$Spec)

  $m = [regex]::Match($Spec.Trim(), '^(?<n>\d+)\s*[dD]\s*(?<f>\d+)\s*(?:(?<sign>[+-])\s*(?<k>\d+))?\s*(?:[x\*]\s*(?<mult>\d+))?\s*$')
  if (-not $m.Success) { return $null }

  $count = [int]$m.Groups['n'].Value
  $faces = [int]$m.Groups['f'].Value
  $mod   = if ($m.Groups['k'].Success) {
             $val = [int]$m.Groups['k'].Value
             if ($m.Groups['sign'].Value -eq '-') { -$val } else { $val }
           } else { 0 }
  $mult  = if ($m.Groups['mult'].Success) { [int]$m.Groups['mult'].Value } else { 1 }

  $sum = 0
  1..$count | ForEach-Object { $sum += (Get-Random -Minimum 1 -Maximum ($faces + 1)) }
  [pscustomobject]@{ Total = [int](($sum + $mod) * $mult) }
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
function Get-RunesFromRow {
  param([Parameter(Mandatory)]$Row)

  # normalize property names for loose lookup
  $norm = @{}
  foreach ($p in $Row.PSObject.Properties) {
    $k = ($p.Name -replace '\s','' -replace '_','').ToLower()
    $norm[$k] = $p.Name
  }
  function _get([string]$key) {
    $k = ($key -replace '\s','' -replace '_','').ToLower()
    if ($norm.ContainsKey($k)) { return $Row.$($norm[$k]) }
    return $null
  }

  # try numbered, then generic
  $r1 = _get 'Runes1'; if (-not $r1) { $r1 = _get 'Rune1' }
  $s1 = _get 'Rune1score'; if (-not $s1) { $s1 = _get 'Runes1score' }

  $r2 = _get 'Runes2'; if (-not $r2) { $r2 = _get 'Rune2' }
  $s2 = _get 'Rune2score'; if (-not $s2) { $s2 = _get 'Runes2score' }

  $r3 = _get 'Runes3'; if (-not $r3) { $r3 = _get 'Rune3' }
  $s3 = _get 'Rune3score'; if (-not $s3) { $s3 = _get 'Runes3score' }

  if (-not $r1) { $r1 = _get 'Runes' }
  if (-not $s1) { $s1 = _get 'Runescore' }

  [pscustomobject]@{
    R1=[string]$r1; S1=$s1
    R2=[string]$r2; S2=$s2
    R3=[string]$r3; S3=$s3
  }
}

function Get-StatRoll {
  param(
    [Parameter(Mandatory)]$Row,
    [Parameter(Mandatory)][ValidateSet('STR','CON','SIZ','DEX','INT','POW','CHA')] [string]$Stat
  )

  # Prefer STATSpec if present and non-empty
  $specCol = "${Stat}Spec"
  if ($Row.PSObject.Properties.Match($specCol)) {
    $spec = [string]$Row.$specCol
    if ($spec -and -not [string]::IsNullOrWhiteSpace($spec)) {
      $r = ConvertFrom-DiceSpec -Spec $spec
      if ($r) { return [int]$r.Total }
    }
  }

  # Legacy fallback: STAT (count of d6) + STATMod
  $n   = 0
  $mod = 0
  if ($Row.PSObject.Properties.Match($Stat))       { $n   = [int]$Row.$Stat }
  if ($Row.PSObject.Properties.Match("${Stat}Mod")){ $mod = [int]$Row."${Stat}Mod" }

  if ($n -le 0) { return [int]$mod }   # fixed stat case
  $sum = 0
  1..$n | ForEach-Object { $sum += (Get-Random -Minimum 1 -Maximum 7) } # d6
  return [int]($sum + $mod)
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
  param($Context,[string]$Sheet,[int]$HP,[int]$AddArmor = 0)

  $path  = Resolve-StatPath 'Hit_Location_Source.xlsx'
  $locs  = Import-Excel -Path $path -WorksheetName $Sheet |
           Where-Object { $_.PSObject.Properties.Value -ne $null } |
           ForEach-Object { $_ | Select-Object * }

  # --- normalize to ints up-front ---
  foreach ($l in $locs) {
    if ($l.PSObject.Properties.Match('HP'))    { $l.HP    = [int]([double]$l.HP) }
    if ($l.PSObject.Properties.Match('armor')) { $l.armor = [int]([double]$l.armor) }
  }

  if ($HP -lt 13 -or $HP -gt 15) {
    if     ($HP -lt 13) { $delta = [math]::Floor(($HP - 13) / 3) }
    else                { $delta = [math]::Ceiling(($HP - 15) / 3) }
    foreach ($l in $locs) { $l.HP = [int]$l.HP + $delta }
  }

  if ($AddArmor -ne 0) {
    foreach ($l in $locs) { $l.armor = [int]$l.armor + $AddArmor }
  }

  # --- normalize again (in case adjustments promoted to doubles) ---
  foreach ($l in $locs) {
    if ($l.PSObject.Properties.Match('HP'))    { $l.HP    = [int]([double]$l.HP) }
    if ($l.PSObject.Properties.Match('armor')) { $l.armor = [int]([double]$l.armor) }
  }

  $locs
}

function Find-WeaponRow {
  param(
    [Parameter(Mandatory)]$Table,
    [Parameter(Mandatory)][string]$Name,
    [Parameter(Mandatory)][string]$Creature
  )
  $cand = @($Table | Where-Object { $_.Name -like "*$Name*" })
  if ($cand.Count -eq 0) { return $null }

  $exact = @($cand | Where-Object { [string]$_.Name -eq $Name })
  if ($exact.Count -gt 0) { $cand = $exact }

  $hasCreature = $cand[0].PSObject.Properties.Match('Creature')
  if ($hasCreature) {
    $spec = @($cand | Where-Object { [string]$_.Creature -eq $Creature })
    if ($spec.Count -gt 0) { return ($spec | Select-Object -First 1) }

    $generic = @($cand | Where-Object { -not $_.PSObject.Properties['Creature'] -or [string]::IsNullOrWhiteSpace([string]$_.Creature) })
    if ($generic.Count -gt 0) { return ($generic | Select-Object -First 1) }
  }
  return ($cand | Select-Object -First 1)
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
     $row = Find-WeaponRow -Table $Context.Weapons.Melee -Name $name -Creature $Creature

      if ($null -ne $row) {
        $row = $row | Select-Object *
        # normalize numeric columns that should be whole numbers
        $intProps = @('HP','Base %','Base','Base%','SR')
        foreach ($p in $intProps) {
        if ($row.PSObject.Properties.Match($p)) {
        try { $row.$p = [int]([double]$row.$p) } catch { }
  }
}
        
        # before adding DB, ensure the row actually has damage
$rawDmg   = ("$($row.Damage)").Trim()
$hasDmg   = (-not [string]::IsNullOrWhiteSpace($rawDmg)) -and ($rawDmg -notmatch '^(0|—|-|n/?a)$')

if ($useDB -and $hasDmg) {
  $row.Damage = "$rawDmg$DamageBonus"
}
        $row.SR = [int]$row.SR + $BaseSR
        $out.Add($row)
      }
    }
    elseif ($type -eq 'mi') {
      $row = Find-WeaponRow -Table $Context.Weapons.Missile -Name $name -Creature $Creature

      if ($null -ne $row) {
        $row = $row | Select-Object *
        # normalize numeric columns that should be whole numbers
        $intProps = @('HP','Base %','Base','Base%','SR')
        foreach ($p in $intProps) {
        if ($row.PSObject.Properties.Match($p)) {
        try { $row.$p = [int]([double]$row.$p) } catch { }
  }
}
   $rawDmg   = ("$($row.Damage)").Trim()
$hasDmg   = (-not [string]::IsNullOrWhiteSpace($rawDmg)) -and ($rawDmg -notmatch '^(0|—|-|n/?a)$')

$row.SR = 0 + $BaseSR

# thrown? only add half-DB if the weapon actually has base damage
if ($parts.Count -ge 3 -and $parts[2].ToLowerInvariant() -eq 'th' -and $useDB -and $hasDmg) {
  $half = Get-HalfDamageBonus $DamageBonus
  if ($half) { $row.Damage = "$rawDmg$half" }
}
        $out.Add($row)
      }
    }
    elseif ($type -eq 'sh') {
      $row = Find-WeaponRow -Table $Context.Weapons.Shields -Name $name -Creature $Creature

      if ($null -ne $row) {
        $row = $row | Select-Object *
                $intProps = @('HP','Base %','Base','Base%','SR')
        foreach ($p in $intProps) {
        if ($row.PSObject.Properties.Match($p)) {
        try { $row.$p = [int]([double]$row.$p) } catch { }
  }
}
        $rawDmg   = ("$($row.Damage)").Trim()
$hasDmg   = (-not [string]::IsNullOrWhiteSpace($rawDmg)) -and ($rawDmg -notmatch '^(0|—|-|n/?a)$')

if ($useDB -and $hasDmg) {
  $row.Damage = "$rawDmg$DamageBonus"
}

        $row.SR = [int]$row.SR + $BaseSR
        $out.Add($row)
      }
    }
  }
  $out.ToArray()
}

function Get-ChaosFeature {
  param(
    $Context,
    [int]$Roll  # optional: allow forcing a specific roll
  )
  $roll = if ($PSBoundParameters.ContainsKey('Roll')) { $Roll } else { Get-Random -Minimum 1 -Maximum 101 }

  $tbl = $Context.Chaos
  if (-not $tbl -or $tbl.Count -eq 0) {
    return [pscustomobject]@{ Roll = $roll; Feature = '(no chaos table loaded)' }
  }

  # Roll column may be imported as double → cast to [int] for comparison
  $row = $tbl | Where-Object { [int]$_.Roll -eq $roll } | Select-Object -First 1
  if (-not $row) { $row = $tbl | Get-Random }  # safety fallback

  [pscustomobject]@{ Roll = $roll; Feature = [string]$row.Feature }
}

function Get-UniqueChaosPicks {
  param(
    $Context,
    [Parameter(Mandatory)][int]$Count,
    [int]$MaxTries = 100
  )
  $out  = New-Object System.Collections.Generic.List[string]
  $seen = @{}
  $tries = 0

  while ($out.Count -lt $Count -and $tries -lt $MaxTries) {
    $pick = (Get-ChaosFeature -Context $Context).Feature
    $tries++

    $allowDup = $pick -match '(?i)\bCurse of Thed\b' -or $pick -match '^\s*Roll twice\b'
    if ($allowDup -or -not $seen.ContainsKey($pick)) {
      $out.Add($pick)
      if (-not $allowDup) { $seen[$pick] = $true }
    }
  }

  # Safety: if table is too small, fill remaining slots (can include dups)
  while ($out.Count -lt $Count) {
    $out.Add((Get-ChaosFeature -Context $Context).Feature)
  }

  $out.ToArray()
}


function Get-ChaosFeaturesForCreature {
  param(
    $Context,
    [Parameter(Mandatory)][string]$Creature,
    [Parameter(Mandatory)][int]$POW,
    [switch]$Force
  )

  $out = New-Object System.Collections.Generic.List[string]

  # Gate by Chaos rune; Force ONLY skips POWx5, not the rune requirement
  $row = Get-StatRow -Context $Context -Creature $Creature
  if (-not (Test-IsChaosCreature -Row $row)) { return @() }

  switch ($Creature) {
    'Broo' {
      if ($Force -or (Get-Random -Minimum 1 -Maximum 101) -le ($POW * 5)) {
        $out.Add((Get-ChaosFeature -Context $Context).Feature)
      }
    }
    'Scorpion Man' {
      if ($Force -or (Get-Random -Minimum 1 -Maximum 101) -le ($POW * 5)) {
        $out.Add((Get-ChaosFeature -Context $Context).Feature)
      }
    }
    'Dragonsnail' {
      # 1–3 features regardless; Force doesn’t change the count
      $n = Get-Random -Minimum 1 -Maximum 4
      1..$n | ForEach-Object { $out.Add((Get-ChaosFeature -Context $Context).Feature) }
    }
    default {
      # NEW: any other Chaos creature → give 1 feature when forced
      if ($Force) {
        $out.Add((Get-ChaosFeature -Context $Context).Feature)
      }
      # If not forced, leave empty (no POWx5 rule for these by default)
    }
  }

  if ($out.Count -eq 0) { return @() }
  Resolve-ChaosFeatures -Context $Context -Features ($out.ToArray())
}




function ConvertFrom-ChaosStatBoostText {
  param([Parameter(Mandatory)][string]$Text)
  # matches: +2D6 DEX, -1d6 STR, +3d6 POW, etc.
  $m = [regex]::Match($Text, '^\s*(?<sign>[+\-])?\s*(?<n>\d+)\s*[dD]\s*(?<faces>\d+)\s*(?<stat>STR|CON|SIZ|DEX|INT|POW|CHA)\b')
  if (-not $m.Success) { return $null }
  [pscustomobject]@{
    Sign  = if ($m.Groups['sign'].Value -eq '-') { -1 } else { +1 }
    Count = [int]$m.Groups['n'].Value
    Faces = [int]$m.Groups['faces'].Value
    Stat  = $m.Groups['stat'].Value.ToUpperInvariant()
    Spec  = "$($m.Groups['n'].Value)D$($m.Groups['faces'].Value)"
  }
}


function Update-CharacteristicsForChaos {
  param(
    [Parameter(Mandatory)]$Characteristics,     # ordered hashtable {STR,CON,SIZ,DEX,INT,POW,CHA}
    [Parameter(Mandatory)][string[]]$Features
  )
  # clone so we don't mutate the original
  $new = [ordered]@{}
  foreach ($k in 'STR','CON','SIZ','DEX','INT','POW','CHA') { $new[$k] = [int]$Characteristics[$k] }

  $applied = New-Object System.Collections.Generic.List[string]

  foreach ($f in $Features) {
    $p = ConvertFrom-ChaosStatBoostText -Text $f
    if ($null -eq $p) { continue }
    $roll  = Invoke-DiceRoll -Count $p.Count -Faces $p.Faces
    $delta = $p.Sign * $roll
    $new[$p.Stat] = [int]$new[$p.Stat] + $delta
    $signTxt = if ($delta -ge 0) { '+' } else { '' }
    $specSign = if ($p.Sign -lt 0) { '-' } else { '+' }
    $applied.Add("$($p.Stat) ${signTxt}$delta (from $specSign$($p.Spec))")
  }

  [pscustomobject]@{
    Updated = $new
    Applied = $applied.ToArray()
  }
}



function Get-ChaosEffects {
  param([string[]]$Features)

  $extraArmor = 0
  $specials   = New-Object System.Collections.Generic.List[object]

  foreach ($f in $Features) {
    if ([string]::IsNullOrWhiteSpace($f)) { continue }

    # Armor: "6-point skin", "9-point skin", "12-point skin"
    if ($f -match '(?i)\b(\d+)\s*-?\s*point\s+skin\b') {
      $val = [int]$matches[1]
      if ($val -gt $extraArmor) { $extraArmor = $val }
      continue
    }

    # Special attacks / effects (keep original text for rules detail)
    if     ($f -match '(?i)\bspits?\s+acid\b')                { $specials.Add([pscustomobject]@{ Name='Acid Spit';           Description=$f }); continue }
    elseif ($f -match '(?i)\bbreathes?\s+.*fire\b')           { $specials.Add([pscustomobject]@{ Name='Fire Breath';         Description=$f }); continue }
    elseif ($f -match '(?i)\bpoison\s+touch\b')               { $specials.Add([pscustomobject]@{ Name='Poison Touch';        Description=$f }); continue }
    elseif ($f -match '(?i)\bexplodes?\s+at\s+death\b')       { $specials.Add([pscustomobject]@{ Name='Death Explosion';     Description=$f }); continue }
    elseif ($f -match '(?i)\bregenerates?\b')                 { $specials.Add([pscustomobject]@{ Name='Regeneration';        Description=$f }); continue }
    elseif ($f -match '(?i)\bagonizing\s+screams\b')          { $specials.Add([pscustomobject]@{ Name='Agonizing Screams';   Description=$f }); continue }
    elseif ($f -match '(?i)\bstench\b')                       { $specials.Add([pscustomobject]@{ Name='Overpowering Stench'; Description=$f }); continue }
    elseif ($f -match '(?i)\bhideous\b')                      { $specials.Add([pscustomobject]@{ Name='Hideous Presence';    Description=$f }); continue }
    elseif ($f -match '(?i)\bbefuddle')                       { $specials.Add([pscustomobject]@{ Name='Befuddle (extra attack)'; Description=$f }); continue }
    elseif ($f -match '(?i)\b(leaping|leap)\b.*\bDEX\b')      { $specials.Add([pscustomobject]@{ Name='Great Leap';          Description=$f }); continue }
    elseif ($f -match '(?i)\bhypnotic\s+appearance\b')        { $specials.Add([pscustomobject]@{ Name='Hypnotic Appearance'; Description=$f }); continue }
    elseif ($f -match '(?i)\bappears?\s+to\s+be\s+a\s+harmless\b') { $specials.Add([pscustomobject]@{ Name='Deceptive Appearance'; Description=$f }); continue }
  }

  [pscustomobject]@{
    ExtraArmor     = $extraArmor
    SpecialAttacks = $specials.ToArray()
  }
}

function Resolve-ChaosFeatures {
 param(
    $Context,
    [object[]]$Features = @(),   # not Mandatory
    [int]$MaxExtra = 6
  )
  if (-not $Features -or $Features.Count -eq 0) { return @() }
  $result = New-Object System.Collections.Generic.List[string]
  $queue  = New-Object System.Collections.Generic.Queue[string]
  foreach ($f in $Features) { $queue.Enqueue($f) }
  $extras = 0
  $prevReal = $null

  while ($queue.Count -gt 0) {
    $f = $queue.Dequeue()
    if ([string]::IsNullOrWhiteSpace($f)) { continue }

    # Roll twice → enqueue two fresh rolls
    if ($f -match '^\s*Roll twice\b') {
      if ($extras -lt $MaxExtra) {
        1..2 | ForEach-Object {
          $queue.Enqueue((Get-ChaosFeature -Context $Context).Feature)
          $extras++
        }
      }
      continue
    }

    # Doubled → duplicate the previous real feature (or roll one and duplicate)
    if ($f -match '(?i)creature.?s chaos feature is doubled') {
      if ($prevReal) {
        $result.Add($prevReal)
      } else {
        if ($extras -lt $MaxExtra) {
          $new = (Get-ChaosFeature -Context $Context).Feature
          $result.Add($new); $result.Add($new); $extras++
        }
      }
      continue
    }

    # (Curse of Thed handled later — leave as-is for now)
    $result.Add($f)
    $prevReal = $f
  }

  $result.ToArray()
}

function Get-CurseOfThedFeature {
  param($Context,[int]$Roll)
  $roll = if ($PSBoundParameters.ContainsKey('Roll')) { $Roll } else { Get-Random -Minimum 1 -Maximum 101 }
  $t = $Context.CurseOfThed
  if (-not $t -or $t.Count -eq 0) { return [pscustomobject]@{ Roll=$roll; Feature='(Curse of Thed not loaded)' } }

  $rows = foreach ($r in $t) {
    $raw = [string]$r.Roll; if (-not $raw) { continue }
    $s = ($raw -replace '[–—]', '-' -replace '\bto\b', '-').Trim()
    if ($s -match '^\s*(\d{1,3})\s*-\s*(\d{1,3})\s*$') { $min=[int]$matches[1]; $max=[int]$matches[2] }
    elseif ($s -match '^\s*(\d{1,3})\s*$')            { $min=[int]$matches[1]; $max=$min }
    else { continue }
    [pscustomobject]@{ Min=$min; Max=$max; Text=[string]$r.Feature }
  }

  $hit = $rows | Where-Object { $_.Min -le $roll -and $_.Max -ge $roll } | Select-Object -First 1
  if (-not $hit) { $hit = $rows | Get-Random }
  [pscustomobject]@{ Roll=$roll; Feature=$hit.Text }
}

function Resolve-ChaosFeatures {
  param($Context,[Parameter(Mandatory)][string[]]$Features,[int]$MaxExtra=6)
  $result = New-Object System.Collections.Generic.List[string]
  $queue  = New-Object System.Collections.Generic.Queue[string]
  foreach ($f in $Features) { if ($f) { $queue.Enqueue($f) } }
  $extras=0; $prevReal=$null

  while ($queue.Count -gt 0) {
    $f = $queue.Dequeue(); if ([string]::IsNullOrWhiteSpace($f)) { continue }

    if ($f -match '^\s*Roll twice\.?\s*$') {
      if ($extras -le $MaxExtra - 2) {
        $queue.Enqueue((Get-ChaosFeature -Context $Context).Feature)
        $queue.Enqueue((Get-ChaosFeature -Context $Context).Feature)
        $extras += 2
      }
      continue
    }

    if ($f -match '(?i)chaos feature is doubled') {
      if ($prevReal) { $result.Add($prevReal) }
      else {
        if ($extras -le $MaxExtra - 1) {
          $n = (Get-ChaosFeature -Context $Context).Feature
          $result.Add($n); $result.Add($n); $extras += 1
        }
      }
      continue
    }

    # NEW: replace “Curse of Thed” with a rolled entry
    if ($f -match '(?i)Curse of Thed') {
      $c = Get-CurseOfThedFeature -Context $Context
      $text = "Curse of Thed: $($c.Feature)"
      $result.Add($text); $prevReal = $text
      continue
    }

    $result.Add($f); $prevReal=$f
  }
  $result.ToArray()
}

function New-SpecialWeapons {
  param(
    $Specials,
    [Parameter(Mandatory)][int]$BaseSR,
    $ExistingWeapons,
    [Parameter(Mandatory)][int]$Dex
  )

  $list = New-Object System.Collections.Generic.List[object]
  if (-not $Specials) { return @() }

  # Column names to match your weapons table
  $baseProp  = 'Base %'
  $hpProp    = 'HP'
  $rangeProp = 'Range'
  if ($ExistingWeapons -and $ExistingWeapons.Count -gt 0) {
    $cols = $ExistingWeapons[0].PSObject.Properties.Name
    $bp = ($cols | Where-Object { $_ -match '^(Base ?%?|Skill)$' } | Select-Object -First 1)
    $hp = ($cols | Where-Object { $_ -match '^(HP|Hp|hp)$' }       | Select-Object -First 1)
    $rp = ($cols | Where-Object { $_ -match '^(Range|Rng)$' }      | Select-Object -First 1)
    if ($bp) { $baseProp = $bp }
    if ($hp) { $hpProp   = $hp }
    if ($rp) { $rangeProp= $rp }
  }

  # helper to build a row aligned with existing columns
$mk = {
  param(
    [string]$name,
    [string]$damage,
    [int]$sr,
    [string]$range,
    [string]$notes,
    [int]$basePercent,
    [string]$hpProp,
    [string]$rangeProp
  )
  # Build with 'Base %' present from the start
  $o = [pscustomobject]@{
    Name    = $name
    'Base %' = $basePercent
    Damage  = $damage
    SR      = $sr
    Notes   = $notes
  }
  # HP + Range (use detected column names)
  $o | Add-Member -NotePropertyName $hpProp -NotePropertyValue 0
  if ($range) { $o | Add-Member -NotePropertyName $rangeProp -NotePropertyValue $range }
  return $o
}

  foreach ($s in $Specials) {
    $text = if ($s -is [string]) { $s }
            elseif ($s.PSObject.Properties.Match('Description')) { [string]$s.Description }
            elseif ($s.PSObject.Properties.Match('Text')) { [string]$s.Text }
            else { [string]$s }

    if ([string]::IsNullOrWhiteSpace($text)) { continue }

    # ---------- Acid Spit ----------
    if ($text -match '(?i)\bspits?\s+acid\b') {
      $pot  = '2D10'
      if ($text -match '(?i)\b(\d+)\s*[dD]\s*(\d+)\s*POT') { $pot = "$($matches[1])D$($matches[2])" }
      $uses = $null
      if ($text -match '(?i)\b(\d+)\s*[dD]\s*(\d+)\s*times per day') { $uses = "$($matches[1])D$($matches[2])/day" }
      $rangeVal = $null
      if ($text -match '(?i)\b(\d+)\s*-\s*meter') { $rangeVal = $matches[1] }
      elseif ($text -match '(?i)\b(\d+)\s*meters?') { $rangeVal = $matches[1] }
      $range = if ($rangeVal) { "${rangeVal}m" } else { $null }

      $list.Add( (& $mk 'Acid Spit' ("POT $pot") $BaseSR $range $uses $baseVal $hpProp $rangeProp) )
      continue
    }

    # ---------- Fire Breath ----------
    if ($text -match '(?i)\bbreathes?\s+(\d+)\s*[dD]\s*(\d+)\s*fire\b') {
      $dmg = "$($matches[1])D$($matches[2])"
      $uses = $null
      if ($text -match '(?i)\b(\d+)\s*[dD]\s*(\d+)\s*times per day') { $uses = "$($matches[1])D$($matches[2])/day" }
      $rangeVal = $null
      if ($text -match '(?i)\b(\d+)\s*-\s*meter') { $rangeVal = $matches[1] }
      elseif ($text -match '(?i)\b(\d+)\s*meters?') { $rangeVal = $matches[1] }
      $range = if ($rangeVal) { "${rangeVal}m" } else { $null }
      $notes = $null
      if ($text -match '(?i)\bsingle\s+target\b') { $notes = 'single target' }
      if ($uses) { $notes = if ($notes) { "$notes; $uses" } else { $uses } }

      $list.Add( (& $mk 'Fire Breath' $dmg $BaseSR $range $notes $baseVal $hpProp $rangeProp) )
      continue
    }

    # ---------- Poison Touch ----------
    if ($text -match '(?i)\bpoison\s+touch\b') {
      $pot = $null
      if ($text -match '(?i)\b(\d+)\s*[dD]\s*(\d+)\s*POT') { $pot = "$($matches[1])D$($matches[2])" }
      $dmg = if ($pot) { "POT $pot" } else { 'POT —' }
      $list.Add( (& $mk 'Poison Touch' $dmg $BaseSR 'touch' 'Must penetrate armor' $baseVal $hpProp $rangeProp) )
      continue
    }

    # ---------- Death Explosion ----------
    if ($text -match '(?i)\bexplodes?\s+at\s+death\b') {
      $rangeVal = $null
      if ($text -match '(?i)\b(\d+)\s*-\s*meter') { $rangeVal = $matches[1] }
      elseif ($text -match '(?i)\b(\d+)\s*meters?') { $rangeVal = $matches[1] }
      $range = if ($rangeVal) { "${rangeVal}m" } else { '3m' }
      $list.Add( (& $mk 'Death Explosion' '1-6D6' 0 $range 'Triggers on death' $baseVal $hpProp $rangeProp) )
      continue
    }
  }

  return $list.ToArray()
}

function Test-IsChaosCreature {
  param([Parameter(Mandatory)]$Row)
  $r = Get-RunesFromRow -Row $Row
  return (@($r.R1,$r.R2,$r.R3) | Where-Object { $_ -match '(?i)\bchaos\b' }).Count -gt 0
}


function Normalize-WeaponsColumns {
  param(
    [Parameter(Mandatory)]$Weapons,
    [Parameter(Mandatory)][int]$Dex
  )

  if (-not $Weapons) { return @() }
  $out = New-Object System.Collections.Generic.List[object]
  $baseDefault = [int]($Dex * 5)

  foreach ($w in @($Weapons)) {
    $props = $w.PSObject.Properties

    # find the real base% column name (handles NBSP, odd spacing, or 'Skill')
    $baseName = $props.Name |
      Where-Object {
        # normalize by removing all non-word & non-% chars
        (($_ -replace '[^\w%]', '') -as [string]).ToLower() -eq 'base%' -or
        $_ -match '^(?i)skill$'
      } |
      Select-Object -First 1

    # pull the value (if present)
    $baseVal = $null
    if ($baseName) { $baseVal = $props[$baseName].Value }
    if ($null -eq $baseVal -and $props['Base %']) { $baseVal = $w.'Base %' }

    # coerce to int if numeric
    $baseInt = $null
    $tmp = 0.0
    if ($baseVal -ne $null -and [double]::TryParse(("$baseVal" -replace '[^\d\.-]',''), [ref]$tmp)) {
      $baseInt = [int]$tmp
    }

    # if still null and it's one of our specials → DEX×5
    if ($null -eq $baseInt) {
      if ($w.Name -in @('Acid Spit','Fire Breath','Poison Touch','Death Explosion')) {
        $baseInt = $baseDefault
      } else {
        $baseInt = 0
      }
    }

    # ensure a standard 'Base %' property exists and is int
    if ($props['Base %']) { $w.'Base %' = $baseInt }
    else { Add-Member -InputObject $w -NotePropertyName 'Base %' -NotePropertyValue $baseInt }

    # normalize SR/HP to ints too
    if ($props['SR']) { $w.SR = [int]([double]$w.SR) }
    if ($props['HP']) { $w.HP = [int]([double]$w.HP) }

    $out.Add($w)
  }
  return $out.ToArray()
}
function Ensure-BasePercent {
  param(
    [Parameter(Mandatory)]$Weapons,
    [Parameter(Mandatory)][int]$Dex
  )

  if (-not $Weapons) { return @() }

  $specialNames = @('Acid Spit','Fire Breath','Poison Touch','Death Explosion')
  $baseVal = [int]($Dex * 5)
  $out = New-Object System.Collections.Generic.List[object]

  foreach ($w in @($Weapons)) {
    $props = $w.PSObject.Properties

    # Find any existing "Base %" column (handles NBSP/odd spacing) or "Skill"
    $realBaseNames = @(
      $props.Name |
        Where-Object {
          (($_ -replace '\u00A0',' ') -replace '\s+',' ') -match '^(?i)base %$' -or
          $_ -match '^(?i)skill$'
        }
    )

    $isSpecial = $specialNames -contains ([string]$w.Name)

    if ($isSpecial) {
      # For specials: force Base % = DEX×5 and mirror to whatever base column exists
      if ($props['Base %']) { $w.'Base %' = $baseVal } else { Add-Member -InputObject $w -NotePropertyName 'Base %' -NotePropertyValue $baseVal }
      foreach ($bn in $realBaseNames) {
        Add-Member -InputObject $w -NotePropertyName $bn -NotePropertyValue $baseVal -Force
      }
    }
    else {
      # For normal weapons, copy from existing base column into a standard 'Base %' (int)
      if (-not $props['Base %'] -or [string]::IsNullOrWhiteSpace("$($w.'Base %')")) {
        $src = $realBaseNames | Select-Object -First 1
        if ($src) {
          $v = $w.$src
          $tmp = 0.0
          if ([double]::TryParse(("$v" -replace '[^\d\.-]',''), [ref]$tmp)) {
            Add-Member -InputObject $w -NotePropertyName 'Base %' -NotePropertyValue ([int]$tmp) -Force
          } else {
            Add-Member -InputObject $w -NotePropertyName 'Base %' -NotePropertyValue 0 -Force
          }
        } else {
          Add-Member -InputObject $w -NotePropertyName 'Base %' -NotePropertyValue 0 -Force
        }
      } else {
        $tmp = 0.0
        if ([double]::TryParse(("$($w.'Base %')" -replace '[^\d\.-]',''), [ref]$tmp)) {
          $w.'Base %' = [int]$tmp
        }
      }
    }

    # Numeric-only coercion for SR/HP (skip text like 'Head', 'Arm', '—')
    if ($props['SR']) {
      $tmp = 0.0; $s = "$($w.SR)"
      if ([double]::TryParse(($s -replace '[^\d\.-]',''), [ref]$tmp)) { $w.SR = [int]$tmp }
    }
    if ($props['HP']) {
      $tmp = 0.0; $s = "$($w.HP)"
      if ([double]::TryParse(($s -replace '[^\d\.-]',''), [ref]$tmp)) { $w.HP = [int]$tmp }
    }

    $out.Add($w)
  }

  $out.ToArray()
}


function New-Statblock {
  param(
    [Parameter(Mandatory)][string]$Creature,
    [Parameter(Mandatory)]$Context,
    [int]$AddArmor = 0,
    [string]$OverrideHitLocationSheet,
    [switch]$ForceChaos
  )

  # 1) Base rolls from the stat-dice table
  $row   = Get-StatRow -Context $Context -Creature $Creature
  $runes = Get-RunesFromRow -Row $row
  $isChaos = Test-IsChaosCreature -Row $row



  $base  = New-Characteristics -Row $row

  # 2) Roll chaos features (ForceChaos guarantees at least one for Chaos-rune creatures except Dragonsnail logic)
$chaos = @()
if ($isChaos) {
  $chaos = Get-ChaosFeaturesForCreature -Context $Context -Creature $Creature -POW $base.POW -Force:$ForceChaos
}

  # 3) Apply any "+xDy STAT" chaos boosts to the characteristics
  $chars = $base
  $appliedBoosts = @()
  if ($chaos -and $chaos.Count -gt 0) {
    $r = Update-CharacteristicsForChaos -Characteristics $base -Features $chaos
    if ($r) {
      $chars = $r.Updated
      $appliedBoosts = $r.Applied
    }
  }
# Chaos-driven armor & specials
$effects     = Get-ChaosEffects -Features $chaos
$armorFromCF = if ($effects) { [int]$effects.ExtraArmor } else { 0 }
  # 4) Derive everything from the (possibly boosted) stats
  $sr   = Get-StrikeRanks       -Dex $chars.DEX -Siz $chars.SIZ
  $hp   = Get-HitPoints         -CON $chars.CON -SIZ $chars.SIZ -POW $chars.POW -HpTable $Context.HpTable
  $db   = Get-DamageBonus       -STR $chars.STR -SIZ $chars.SIZ
  $scd  = Get-SpiritCombatDamage -POW $chars.POW -CHA $chars.CHA




  # 5) Choose hit-location sheet (override wins; default Dragonsnail1 if none)
  $sheet = $null
  if ($PSBoundParameters.ContainsKey('OverrideHitLocationSheet') -and -not [string]::IsNullOrWhiteSpace($OverrideHitLocationSheet)) {
    $sheet = $OverrideHitLocationSheet
  }
  elseif ($Creature -eq 'Dragonsnail') {
    $sheet = 'Dragonsnail1'
  }
  else {
    $sheet = [string]$row.Hit_location
  }

  # 6) Build locations & weapons with the computed derived values
  $hitLocs = Get-HitLocations -Context $Context -Sheet $sheet -HP $hp -AddArmor ($AddArmor + $armorFromCF)
  $weapons = Get-Weapons      -Context $Context -Creature $Creature -BaseSR $sr.Base -DamageBonus $db
$weapons = @($weapons)   # <-- force array
  # Turn special effects into weapon entries and append
$specialWeapons = New-SpecialWeapons -Specials $effects.SpecialAttacks -BaseSR $sr.Base -ExistingWeapons $weapons -Dex $chars.Dex
if ($specialWeapons -and $specialWeapons.Count -gt 0) {
  $weapons = @($weapons + $specialWeapons)
}
# Harvest specials defined on the matched weapon rows (tolerant of header variants)
$weaponTableSpecials = @()
foreach ($w in @($weapons)) {
  $props = $w.PSObject.Properties

  # Find a name column: SpecialName / Special Name / Special_Name
  $nameProp = ($props.Name | Where-Object { $_ -match '^(?i)special[\s_]*name$' } | Select-Object -First 1)
  $specialName = if ($nameProp) { [string]$w.$nameProp } else { $null }

  # Find a text column: SpecialText / Special Text / Special_Text / Special
  $textProp = ($props.Name | Where-Object { $_ -match '^(?i)special([\s_]*text)?$' } | Select-Object -First 1)
  $specialText = if ($textProp) { [string]$w.$textProp } else { $null }

  # Fallback: if Notes looks rule-like and no SpecialText was found
  if ((-not $specialText) -and $props.Match('Notes') -and $w.Notes) {
    $notesStr = [string]$w.Notes
    if ($notesStr -match '(?i)\b(special|on hit|penetrat|mp vs mp|times per day|blood|drain|poison|pot)\b') {
      $specialText = $notesStr
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($specialText)) {
    if (-not $specialName -or [string]::IsNullOrWhiteSpace($specialName)) {
      $specialName = [string]$w.Name
    }
    $weaponTableSpecials += [pscustomobject]@{
      Name        = $specialName
      Description = ($specialText -replace '\s+', ' ').Trim()
    }
  }
}

# Merge chaos-driven specials with table-driven specials
$allSpecials = @($effects.SpecialAttacks + $weaponTableSpecials)

$effects = Get-ChaosEffects
# Ensure Base % is present & correct (DEX×5 for specials; mirror to real column)
$weapons = Ensure-BasePercent -Weapons $weapons -Dex $chars.DEX
  # 7) Return the full block
  [pscustomobject]@{
    Creature           = $Creature
    Move               = [int]$row.Move
    Runes1       = $runes.R1
Rune1Score   = $runes.S1
Runes2       = $runes.R2
Rune2Score   = $runes.S2
Runes3       = $runes.R3
Rune3Score   = $runes.S3
    Characteristics    = [pscustomobject]([ordered]@{ STR=$chars.STR; CON=$chars.CON; SIZ=$chars.SIZ; DEX=$chars.DEX; INT=$chars.INT; POW=$chars.POW; CHA=$chars.CHA })
    # Optional: show pre-chaos for debugging; remove if you don't want it
    BaseCharacteristics= [pscustomobject]([ordered]@{ STR=$base.STR; CON=$base.CON; SIZ=$base.SIZ; DEX=$base.DEX; INT=$base.INT; POW=$base.POW; CHA=$base.CHA })
    HP                 = $hp
    StrikeRanks        = $sr
    DamageBonus        = $db
    SpiritCombat       = $scd
    HitLocations       = $hitLocs
    Weapons            = $weapons
    ChaosFeatures      = $chaos
    ChaosApplied       = $appliedBoosts
    HitLocationSheet   = $sheet
    ChaosArmorBonus  = $armorFromCF
    SpecialAttacks = $allSpecials

  }
}


Export-ModuleMember -Function Initialize-StatblockContext, New-Statblock, Roll-Dice, Get-StrikeRanks, Get-StatRow, Get-StatRoll, New-Characteristics, Get-HitPoints, Get-DamageBonus, Get-HalfDamageBonus, Get-SpiritCombatDamage, Get-HitLocations, Get-Weapons, Get-ChaosFeature, Get-ChaosFeaturesForCreature, ConvertFrom-ChaosStatBoostText, Update-CharacteristicsForChaos, Get-ChaosEffects, Resolve-ChaosFeatures, Get-CurseOfThedFeature, New-SpecialWeapons, Get-RunesFromRow, Test-IsChaosCreature

