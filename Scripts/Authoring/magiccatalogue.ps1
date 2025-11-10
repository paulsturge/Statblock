$catMallia = Get-CultRuneSpellCatalog -CultName 'Mallia' -WorkbookPath $wb -IncludeAssociates
" Mallia: Special=$($catMallia.Special.Count)  Common=$($catMallia.Common.Count)"
$catMallia.Special | Sort-Object Name | Format-Table Name,FromCult -Auto

$catKygerLitor = Get-CultRuneSpellCatalog -CultName 'KygerLitor' -WorkbookPath $wb -IncludeAssociates
" KygerLitor: Special=$($catKyger.Special.Count)  Common=$($catKigerLitor.Common.Count)"
$catKigerLitor.Special | Sort-Object Name | Format-Table Name,FromCult -Auto


$catK = Get-CultRuneSpellCatalog -CultName 'Kyger Litor' -WorkbookPath $wb -IncludeAssociates
" Kyger Litor : Special=$($catK.Special.Count)  Common=$($catK.Common.Count)"
$catK.Special | Sort-Object Name | Format-Table Name,FromCult -Auto
$catK.Common  | Sort-Object Name | Format-Table Name,FromCult -Auto