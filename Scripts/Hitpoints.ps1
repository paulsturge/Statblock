Function Hitpoints($CON, $SIZ, $POW){
#siz and pow  hit point modifier

$Global:hp = 0
#$Global:chaosfeature = $Chaosfeaturetable |Where Roll -EQ $d100

#temp stats for testing
#$con= 10
#$siz= 21
#$pow= 10

#get hp modifier based on size and pow

#if size is > 28 calculate additional bonus over +4. IF siz is =< 28 calculate bonus based on modifier table.
if($siz -gt 28){$siz_hp_modifier = 4 + [math]::Ceiling(($siz -28)/4)
#$siz_hp_modifier
$Global:hp = $con + $siz_hp_modifier}
Else{$siz_hp_modifier = $Global:hp_modifier_table |Where Stat_value -EQ $siz
#$siz_hp_modifier.SIZ
$Global:hp = $con + $siz_hp_modifier.SIZ
}

#if pow is > 28 calculate additional bonus over +4. IF pow is =< 28 calculate bonus based on modifier table.
if($pow -gt 28){$pow_hp_modifier = 3 + [math]::Ceiling(($pow -28)/4)
#$pow_hp_modifier
$Global:hp += $pow_hp_modifier}
Else{$pow_hp_modifier = $Global:hp_modifier_table |Where Stat_value -EQ $pow
$Global:hp += $pow_hp_modifier.POW
#$pow_hp_modifier.POW

}




}

#defaul = 40% chest