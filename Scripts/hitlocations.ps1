Function Hitlocations{
$hpchange = 0
$filename = "Hit_Location_Source.xlsx"
$hpchange = 0
$path = $folder + $filename




$Global:hit_locations.Clear()
$Global:hit_locations = Import-Excel -Path $path -WorksheetName $sheet |Where-Object { $_.PSObject.Properties.Value -ne $null} 


if($hp -notin 13..15){
    Write-Host "Trigger location hp calculations"
    Switch ($hp){

        {$_ -lt 13}{Write-Host "Less than 13"
            $hpchange = [math]::Floor(($hp -13) /3)
            Foreach($location in $hit_locations){
                Write-Host $location.HP
                $location.HP += $hpchange
                
            }
        }
        {$_ -gt 15}{Write-Host "greater than 15"
    
            $hpchange = [math]::Ceiling(($hp -15) /3)
            Foreach($location in $hit_locations){
                Write-Host $location.hp
                $location.HP += $hpchange
            }
    
    
        }




    }

} 




$Global:hit_locations

}