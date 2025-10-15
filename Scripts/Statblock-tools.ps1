# Statblock-tools.psm1 (PowerShell module)
'me' {
$row = $Context.Weapons.Melee | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
if ($null -ne $row) {
$row = $row | Select-Object *
if ($DamageBonus) { $row.Damage = "$($row.Damage)$DamageBonus" }
$row.SR = [int]$row.SR + $BaseSR
$out.Add($row)
}
}
'mi' {
$row = $Context.Weapons.Missile | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
if ($null -ne $row) {
$row = $row | Select-Object *
$row.SR = [int]0 + $BaseSR
if ($parts.Count -ge 3 -and $parts[2].ToLowerInvariant() -eq 'th') {
$half = Get-HalfDamageBonus $DamageBonus
if ($half) { $row.Damage = "$($row.Damage)$half" }
}
$out.Add($row)
}
}
'sh' {
$row = $Context.Weapons.Shields | Where-Object { $_.Name -like "*$name*" } | Select-Object -First 1
if ($null -ne $row) {
$row = $row | Select-Object *
if ($DamageBonus) { $row.Damage = "$($row.Damage)$DamageBonus" }
$row.SR = [int]$row.SR + $BaseSR
$out.Add($row)
}
}
}
}
$out.ToArray()
}


function New-Statblock {
param(
[Parameter(Mandatory)][string]$Creature,
[Parameter(Mandatory)]$Context,
[int]$AddArmor = 0,
[string]$OverrideHitLocationSheet
)
$row = Get-StatRow -Context $Context -Creature $Creature
$chars = New-Characteristics -Row $row
$sr = Get-StrikeRanks -Dex $chars.DEX -Siz $chars.SIZ
$hp = Get-HitPoints -CON $chars.CON -SIZ $chars.SIZ -POW $chars.POW -HpTable $Context.HpTable
$db = Get-DamageBonus -STR $chars.STR -SIZ $chars.SIZ
$scd = Get-SpiritCombatDamage -POW $chars.POW -CHA $chars.CHA


# Hit-location worksheet: allow override; special case Dragonsnail randomization
$sheet = if ($OverrideHitLocationSheet) { $OverrideHitLocationSheet } else { [string]$row.Hit_location }
if (-not $OverrideHitLocationSheet -and $Creature -eq 'Dragonsnail') {
$sheet = if ((Get-Random -Minimum 1 -Maximum 101) -gt 65) { 'Dragonsnail1' } else { 'Dragonsnail' }
}
$hitLocs = Get-HitLocations -Context $Context -Sheet $sheet -HP $hp -AddArmor $AddArmor
$weapons = Get-Weapons -Context $Context -Creature $Creature -BaseSR $sr.Base -DamageBonus $db


[pscustomobject]@{
Creature = $Creature
Move = [int]$row.Move
Runes1 = $row.Runes1
Rune1Score = $row.Rune1score
Runes2 = $row.Runes2
Rune2Score = $row.Rune2score
Characteristics = [pscustomobject](
[ordered]@{ STR=$chars.STR; CON=$chars.CON; SIZ=$chars.SIZ; DEX=$chars.DEX; INT=$chars.INT; POW=$chars.POW; CHA=$chars.CHA }
)
HP = $hp
StrikeRanks = $sr
DamageBonus = $db
SpiritCombat = $scd
HitLocations = $hitLocs
Weapons = $weapons
}
}


Export-ModuleMember -Function * -Alias *