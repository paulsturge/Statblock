# -----------------------------
# Statblock.ps1 (entry script)
# -----------------------------
# Usage: run this script from the Scripts folder; it will import the module beside it.



# Statblock.ps1 (top)
param(
  [Parameter(Position=0)]
  [Alias('c')]
  [ArgumentCompleter({
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    try {
      Import-Module "$PSScriptRoot\Statblock-tools.psm1" -ErrorAction Stop | Out-Null
      $ctx = Initialize-StatblockContext
      $ctx.StatDice.Creature |
        Sort-Object -Unique |
        Where-Object { $_ -like "$wordToComplete*" } |
        ForEach-Object {
          $display = $_
          $needsQuotes = $display -match '\s'          # quote if any whitespace
          $escaped     = $display -replace "'", "''"   # escape single quotes
          $text        = if ($needsQuotes) { "'$escaped'" } else { $display }
          [System.Management.Automation.CompletionResult]::new($text, $display, 'ParameterValue', $display)
        }
    } catch { @() }
  })]
  [string]$Creature = 'Human',

  [int]$Seed
)



Import-Module "$PSScriptRoot\Statblock-tools.psm1" -Force -ErrorAction Stop
$ctx = Initialize-StatblockContext

if ($ListCreatures) {
  $ctx.StatDice.Creature | Sort-Object -Unique | Format-Wide -AutoSize
  return
}

if ($PSBoundParameters.ContainsKey('Seed')) { Get-Random -SetSeed $Seed }

$sb = New-Statblock -Creature $Creature -Context $ctx -AddArmor 0


# Pretty print
$chars = $sb.Characteristics
Write-Host ("{0}: STR {1} CON {2} SIZ {3} DEX {4} INT {5} POW {6} CHA {7}" -f $sb.Creature,$chars.STR,$chars.CON,$chars.SIZ,$chars.DEX,$chars.INT,$chars.POW,$chars.CHA)
Write-Host ("HP {0} Move {1} | Dex SR {2} Siz SR {3} | DB {4} | Spirit {5}" -f $sb.HP,$sb.Move,$sb.StrikeRanks.DexSR,$sb.StrikeRanks.SizSR,$sb.DamageBonus,$sb.SpiritCombat)
if ($sb.Runes1 -or $sb.Runes2) {
Write-Host ("Runes: {0} {1}, {2} {3}" -f $sb.Runes1,$sb.Rune1Score,$sb.Runes2,$sb.Rune2Score)
}


$sb.HitLocations | Format-Table -AutoSize
$sb.Weapons | Format-Table -AutoSize