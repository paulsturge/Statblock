Import-Module .\Statblock-tools.ps1
Import-Module .\d6.ps1
Import-Module .\d100
#set path for creature weapon type
$Global:folder ='C:\Users\psturge.WARWICK\My Drive\RuneQuest\RQG\GM_aids\Stat_blocks\'

#Load tables, if not already loaded
if (!$Global:hp_modifier_table){
$Global:hp_modifier_table = Import-Excel -Path 'C:\Users\psturge.WARWICK\My Drive\RuneQuest\RQG\GM_aids\Stat_blocks\hp_modifier_table.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$Global:Chaosfeaturetable = Import-Excel -Path 'C:\Users\psturge.WARWICK\My Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Chaotic_features.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$Global:statdice = Import-Excel -Path 'C:\Users\psturge.WARWICK\My Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Stat_Dice_Source.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}

}
#$statblocksraw = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Stat_Blocks_raw.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$STR =0
$CON =0
$SIZ =0
$DEX =0
$INT =0
$POW =0
$CHA =0
$chaosswitch = 0
$Global:primecult = ""
$creaturetype = "Human"
$Global:addarmor = 0
$Characteristics = [ordered]@{STR=$STR; CON=$CON; SIZ=$SIZ; DEX=$DEX; INT=$INT; POW=$POW; CHA=$CHA}


ForEach($stat in $($Characteristics.Keys)){

#get numbers of dice to roll
Statdice $creaturetype $stat



#If stat <> 0 Generate stats
if($numberdice -ne 0){
d6 $numberdice

$sum += $modifier
$Characteristics[$stat] = $sum
}




}






  $Global:move = $currentcreature.Move
  

  Switch($creaturetype){
  {$creaturetype -eq "Broo"}{Broo}
  {$creaturetype -eq "Dark Troll"}{Dark_Troll}
  {$creaturetype -eq "Great Troll"}{Great_Troll}
  {$creaturetype -eq "Mistress Race Troll"}{Mistress_Race_Troll}
  {$creaturetype -like "*Beetle*"}{Beetle}
  {$creaturetype -eq "Dragonsnail"}{Dragonsnail}
  {$creaturetype -eq "Human"}{Human}
  {$creaturetype -eq "Scorpion Man"}{Scorpion_Men}
  }
  #Else {Write-Host "No Chaotic Feature."}
  
#Write-host $Characteristics "-$($numberdice)- $($modifier)" 
#Parse stats and call hitpoint funtion
#$CON  = $Characteristics.CON
#$SIZ = $Characteristics.SIZ
#$POW = $Characteristics.POW
Hitpoints $Characteristics.CON $Characteristics.SIZ $Characteristics.POW
Hitlocations


CalculateSR $Characteristics.DEX $Characteristics.SIZ
DamageBonus $Characteristics.STR $Characteristics.SIZ
SpiritCombatDamage $Characteristics.POW $Characteristics.CHA
Weapons

$rune1 = $currentcreature.Runes1
$rune2 = $currentcreature.Runes2
$rune1score = $currentcreature.Rune1score
$rune2score = $currentcreature.Rune2score

Write-host ""
Write-Host "STR $($Characteristics.STR) CON $($Characteristics.CON) SIZ $($Characteristics.SIZ) DEX $($Characteristics.DEX) INT $($Characteristics.INT) POW $($Characteristics.POW) CHA $($Characteristics.CHA)"


Write-Host "General hp: $($Global:hp) Move $($move)"



Write-Host "DEX SR $($Global:dex_sr) SIZ SR $($Global:siz_sr)"
 
Write-Host "Damage Bonus: $($dambonus)"
Write-Host "Spirit Combat Damage: $($spiritcombadam)"

Write-Host "Runes: $($rune1) $($rune1score), $($rune2) $($rune2score)"
Write-Host ""
$Global:hit_locations
$creature_weapons |ft
$outputraw = $Characteristics
Write-Host "Primary Cult:" $primecult
$runepoints



#$Characteristics |Export-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Stat_Blocks_raw.xlsx' -Append
#$sum = 0
#$3d6 |Foreach {$sum += $_}
#$Characteristics. = $sum
#write-host $stat 