Function d6 ([Int]$max){
$Global:sum = 0
$3d6 = 1..$max | ForEach-Object {
    1..6 | Get-Random 
  
    }
   # $3d6
    $3d6 |Foreach {$Global:sum += $_

 # Write-Host $_}
    }}



