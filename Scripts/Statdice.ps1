Function Statdice ($creaturetype, $stat){

$Global:numberdice = 0
$Global:modifier = 0



$tempmod = $stat + 'Mod'

$Global:currentcreature = $statdice |where Creature -eq $creaturetype

#Write-Host "This is the current stat: $($stat)"

$Global:numberdice = $currentcreature.$stat
$Global:modifier = $currentcreature.$tempmod

$Global:filename = $currentcreature.Creature + "_weapons.txt"
$Global:sheet = $currentcreature.Hit_location

#$creaturedice = @{}


}
