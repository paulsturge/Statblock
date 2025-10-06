#############################
#Various Cults by race/Runes#
#############################

Function BrooCults{
    d100
    Write-Host "This is the roll for Cult" $d100
    Switch($d100){

        {$_ -in 01..03}{$Global:primecult = "Daka Fal"}
        {$_ -eq 4     }{$Global:primecult = "Seven Mothers"}
        {$_ -in 05..14}{$Global:primecult = "Primal Chaos"}
        {$_ -in 15..49}{$Global:primecult = "Mallia"}
        {$_ -eq 50    }{$Global:primecult = "Bagog"}
        {$_ -in 51..90}{$Global:primecult = "Thed"}
        {$_ -in 91..95}{$Global:primecult = "Thanatar"}
        {$_ -in 96..97}{$Global:primecult = "Krarsht"}
        {$_ -eq 98    }{$Global:primecult = "Gbaji"}
        {$_ -in 99..100}{$Global:primecult = "Other (gamemaster choice)"}
    }
}

Function TrollCults{
    d100
    Write-Host "This is the roll for Cult" $d100
    Switch($d100){

        {$_ -in 01..03}{$Global:primecult = "Daka Fal"}
        {$_ -eq 4     }{$Global:primecult = "Seven Mothers"}
        {$_ -in 05..14}{$Global:primecult = "Primal Chaos"}
        {$_ -in 15..49}{$Global:primecult = "Mallia"}
        {$_ -eq 50    }{$Global:primecult = "Bagog"}
        {$_ -in 51..90}{$Global:primecult = "Thed"}
        {$_ -in 91..95}{$Global:primecult = "Thanatar"}
        {$_ -in 96..97}{$Global:primecult = "Krarsht"}
        {$_ -eq 98    }{$Global:primecult = "Gbaji"}
        {$_ -in 99..100}{$Global:primecult = "Other (gamemaster choice)"}
    }
}

#############
#Rune points#
#############

Function RunePoints{
$Global:runepoints = 0
Switch($currentcreature){

    {$creaturetype -eq "Mistress Race Troll"}{
        d3
        $d3
        d6($d3)
        $sum
        $runepoints = 20 + $sum
        $runepoints
}

}
}

#add any additional armor
Function Addarmor{
Write-Host "Yep! Here is the addarmor function"
$hit_locations |ForEach-Object {$_.armor += $addarmor}
}



#weapon function

Function Weapons(){
If($creaturetype -eq "Karg Beetle"){
$Global:weapons_melee_table = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\weapons_melee_table.xlsx' -WorksheetName "Karg_Beetle"|Where-Object { $_.PSObject.Properties.Value -ne $null}
}
Else{
$Global:weapons_melee_table = Import-Excel -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\weapons_melee_table.xlsx' |Where-Object { $_.PSObject.Properties.Value -ne $null}
}
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

#Calculate hit locations
Function Hitlocations{
$hpchange = 0

if($hp -notin 13..15){
    Switch ($hp){

        {$_ -lt 13}{
            $hpchange = [math]::Floor(($hp -13) /3)
            Foreach($location in $hit_locations){
                $location.HP += $hpchange
                
            }
        }
        {$_ -gt 15}{    
            $hpchange = [math]::Ceiling(($hp -15) /3)
            Foreach($location in $hit_locations){
                $location.HP += $hpchange
            }
        }
    }
    } 
}

#calculate half damage bonus

Function Halfdamagebonus ($dambonus){
$damsep = "+","D"

$dambonus.split($damsep, $option)

if($dambonus[1] -ne "1"){

$numdamdice = [System.Convert]::ToInt32($dambonus[1], 10)
$damdicetype = [System.Convert]::ToInt32($dambonus[3], 10)
$Global:halfdambonus = "+" + "$($numdamdice)" + "D" + "$($damdicetype/2)"
}
Else {

$numdamdice = "1"
$damdicetype = [System.Convert]::ToInt32($dambonus[3], 10)
$Global:halfdambonus = "+" + "$($numdamdice)" + "D" + "$($damdicetype/2)"


}

#$halfdam
}



#dice

Function d3(){

$Global:d3 = 1..3 |Get-Random
}

Function d6 ([Int]$max){
$Global:sum = 0
$3d6 = 1..$max | ForEach-Object {
    1..6 | Get-Random 
  
    }
   # $3d6
    $3d6 |Foreach {$Global:sum += $_

 # Write-Host $_}
    }}


Function d100(){

$Global:d100 = 1..100 |Get-Random
}

#==================
#creature functions
#==================

Function Broo(){
 Write-Host $creaturetype
  $pow5 = $Characteristics.POW * 5
    
   d100
   Write-host "POWx5 = " $pow5
   Write-Host "d100 roll = " $d100
    If($d100 -le $pow5){
    Chaosfeature
   }
 Else {$Global:chaosswitch = 1
 Write-Host "No chaotic feature"}
 #$Global:weapons = Get-Content -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Broo_weapons.txt'
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 
if($addarmor -gt 0){
Write-Host "Triggered Addarmor"
Addarmor

}
BrooCults
}

Function Mistress_Race_Troll(){
Write-Host $creaturetype
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 
RunePoints
}

Function Dark_Troll(){
Write-Host $creaturetype
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 

}

Function Great_Troll(){
Write-Host $creaturetype
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 

}

Function Dragonsnail(){
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
d100
Write-Host "d100 roll for two heads: " $d100
if($d100 -gt 65){
$sheet = "Dragonsnail1"
}
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null}
Write-Host $creaturetype


Function Beetle(){
Write-Host $creaturetype
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName "Beetle" |Where-Object { $_.PSObject.Properties.Value -ne $null} 

}


d3
Write-Host "Result on the D3: " $d3
1..$d3 |Foreach {Write-Host "loop " $_
Chaosfeature
}
if($addarmor -gt 0){
Write-Host "Triggered Addarmor"
Addarmor
}

}
    
Function Human(){
Write-Host $creaturetype
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 
}

Function Scorpion_Men(){
 Write-Host $creaturetype
  $pow5 = $Characteristics.POW * 5
    
   d100
   Write-host "POWx5 = " $pow5
   Write-Host "d100 roll = " $d100
    If($d100 -le $pow5){
    Chaosfeature
   }
 Else {$Global:chaosswitch = 1
 Write-Host "No chaotic feature"}
 #$Global:weapons = Get-Content -Path 'C:\Users\psturge\Google Drive\RuneQuest\RQG\GM_aids\Stat_blocks\Broo_weapons.txt'
$filename = "Hit_Location_Source.xlsx"
$path = $folder + $filename
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 
if($addarmor -gt 0){
Write-Host "Triggered Addarmor"
Addarmor
}

}