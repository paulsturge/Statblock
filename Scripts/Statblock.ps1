# -----------------------------
# Statblock.ps1 (entry script)
# -----------------------------
# Usage: run this script from the Scripts folder; it will import the module beside it.


Import-Module "$PSScriptRoot\Statblock-tools.psm1" -Force


$ctx = Initialize-StatblockContext # or: Initialize-StatblockContext -DataRootOverride 'C:\Full\Path\To\Stat_blocks'


$creature = 'Human' # change as needed, e.g., 'Dark Troll', 'Mistress Race Troll', 'Dragonsnail'
$sb = New-Statblock -Creature $creature -Context $ctx -AddArmor 0


# Pretty print
$chars = $sb.Characteristics
Write-Host ("{0}: STR {1} CON {2} SIZ {3} DEX {4} INT {5} POW {6} CHA {7}" -f $sb.Creature,$chars.STR,$chars.CON,$chars.SIZ,$chars.DEX,$chars.INT,$chars.POW,$chars.CHA)
Write-Host ("HP {0} Move {1} | Dex SR {2} Siz SR {3} | DB {4} | Spirit {5}" -f $sb.HP,$sb.Move,$sb.StrikeRanks.DexSR,$sb.StrikeRanks.SizSR,$sb.DamageBonus,$sb.SpiritCombat)
if ($sb.Runes1 -or $sb.Runes2) {
Write-Host ("Runes: {0} {1}, {2} {3}" -f $sb.Runes1,$sb.Rune1Score,$sb.Runes2,$sb.Rune2Score)
}


$sb.HitLocations | Format-Table -AutoSize
$sb.Weapons | Format-Table -AutoSize