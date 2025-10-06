Function DamageBonus ($str, $siz){

#Damage bonus

$Global:dambonus = ""
$strsiz = $str + $siz


switch ($strsiz){


{$_ -le 12}{$Global:dambonus = "-1d4"}

{$_ -in 13..24}{$Global:dambonus = ""}

{$_ -in 25..32}{$Global:dambonus = "+1D4"}

{$_ -in 33..40}{$Global:dambonus = "+1D6"}

{$_ -in 41..56}{$Global:dambonus = "+2D6"}

{$_ -gt 56}{
$dicetype = "d6"
$dicenumber = 2 + [math]::Ceiling(($strsiz - 56)/16)
#{$siz_hp_modifier = 4 + [math]::Ceiling(($siz -28)/4)
$Global:dambonus = "+" + $dicenumber.ToString() + $dicetype
}
}

}