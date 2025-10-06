Function Weapons(){

$Global:weapons_melee_table = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\weapons_melee_table.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$Global:weapons_missile_table = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\weapons_missile_table.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$Global:weapons_shields_table = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\weapons_shields_table.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
$path = $folder + $filename
$seperator = "_"
$option = [System.StringSplitOptions]::RemoveEmptyEntries
$Global:weapons = Get-Content -Path $path
$Global:creature_weapons =@()
$baseSR = $dex_sr + $siz_sr


foreach ($weapon in $weapons){
$current_weapon = $weapon.split($seperator, $option)

Switch($current_weapon[1]){

{$_ -eq "me"}{
$Global:creature_weapon = $Global:weapons_melee_table | Where Name -like $current_weapon[0]
$creature_weapon.Damage += $dambonus
$creature_weapon.SR += $baseSR
$Global:creature_weapons += $creature_weapon}

{$_ -eq "mi"}{
$Global:creature_weapon = $Global:weapons_missile_table | Where Name -like $current_weapon[0]
$creature_weapon.SR = 0
    if($current_weapon[2] -eq "th"){
        Halfdamagebonus $dambonus
        $creature_weapon.Damage += $halfdambonus
    }

$creature_weapon.SR += $baseSR
$Global:creature_weapons += $creature_weapon}

{$_ -eq "sh"}{
$Global:creature_weapon = $Global:weapons_shields_table | Where Name -like "*$($current_weapon[0])*"
$creature_weapon.Damage += $dambonus
$creature_weapon.SR += $baseSR
$Global:creature_weapons += $creature_weapon}

}
}

foreach($item in $creature_weapons){

#$item.Damage += $dambonus
#$item.SR += $baseSR

}

}