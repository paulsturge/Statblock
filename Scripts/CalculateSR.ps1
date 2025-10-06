Function CalculateSR($dex, $siz){
$Global:dex_sr = 0
$Global:siz_sr = 0





switch ($dex)
{
    {$_ -in 1..5}{$Global:dex_sr = 5}
    {$_ -in 6..8}{$Global:dex_sr = 4}
    {$_ -in 9..12}{$Global:dex_sr = 3}
    {$_ -in 13..15}{$Global:dex_sr = 2}
    {$_ -in 16..18}{$Global:dex_sr = 1}
    {$_ -ge 19}{$Global:dex_sr = 0}
}

switch ($siz)
{
    {$_ -in 1..3}{$Global:siz_sr = 3}
    {$_ -in 7..14}{$Global:siz_sr = 2}
    {$_ -in 15..21}{$Global:siz_sr = 1}
    {$_ -ge 22 }{$Global:siz_sr = 0}
}

}