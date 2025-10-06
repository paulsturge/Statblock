Function SpiritCombatDamage ($pow, $cha){

#Spirit combat Damage 
#$pow = 39
#$cha = 35
$powcha = $pow + $cha


switch ($powcha){


{$_ -in 2..12}{$Global:spiritcombadam = "1d3"}

{$_ -in 13..24}{$Global:spiritcombadam = "1d6"}

{$_ -in 25..32}{$Global:spiritcombadam = "1d6+1"}

{$_ -in 33..40}{$Global:spiritcombadam = "+1d6+3"}

{$_ -in 41..56}{$Global:spiritcombadam = "+2d6+3"}

{$_ -gt 56}{
$dicetype = "d6"
$dicenumber = 2 + [math]::Floor(($powcha - 56)/16)
$diceadd = 3 + [math]::Floor(($powcha - 56)/16)
#{$siz_hp_modifier = 4 + [math]::Ceiling(($siz -28)/4)
$Global:spiritcombadam =  $dicenumber.ToString() + $dicetype + "+" + $diceadd.ToString()
}
}
}