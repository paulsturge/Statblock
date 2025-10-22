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
          $needsQuotes = $display -match '\s'
          $escaped     = $display -replace "'", "''"
          $text        = if ($needsQuotes) { "'$escaped'" } else { $display }
          [System.Management.Automation.CompletionResult]::new($text, $display, 'ParameterValue', $display)
        }
    } catch { @() }
  })]
  [string]$Creature = 'Human',

  [switch]$TwoHeaded,
  [switch]$ListCreatures,   # 👈 add this back
  [int]$Seed,
  [switch]$ForceChaos
)
# Allow 'Dragonsnail -2' as a single value for convenience
if ($Creature -match '^\s*Dragonsnail\s*-\s*2\s*$') {
  $TwoHeaded = $true
  $Creature  = 'Dragonsnail'
}

function Format-MoveText {
  param($Sb)
  if ($Sb.MoveModes -and $Sb.MoveModes.Count -gt 0) {
    $modesTxt = ($Sb.MoveModes | ForEach-Object { "$($_.Name) $($_.Value)" }) -join ' | '
    return ("{0} | {1}" -f $Sb.Move, $modesTxt)
  } elseif ($Sb.MoveRaw) {
    return $Sb.MoveRaw
  } else {
    return $Sb.Move
  }
}


Import-Module "$PSScriptRoot\Statblock-tools.psm1" -Force -ErrorAction Stop
$ctx = Initialize-StatblockContext

if ($ListCreatures) {
  $ctx.StatDice.Creature |
    Sort-Object -Unique |
    ForEach-Object {
      if ($_ -match '\s') { "'$($_ -replace "'", "''")'" } else { $_ }
    } | Format-Wide -AutoSize
  return
}

if ($PSBoundParameters.ContainsKey('Seed')) { Get-Random -SetSeed $Seed }

$overrideSheet = $null
if ($Creature -eq 'Dragonsnail') {
  $overrideSheet = if ($TwoHeaded) { 'Dragonsnail2' } else { 'Dragonsnail1' }
}

$sb = New-Statblock -Creature $Creature -Context $ctx -AddArmor 0 -OverrideHitLocationSheet $overrideSheet -ForceChaos:$ForceChaos
Write-Host ("Hit locations sheet: {0}" -f $sb.HitLocationSheet)
#$sb | Get-Member -Name BaseCharacteristics,ChaosApplied,Characteristics
# Chaos features + what got applied
if ($sb.ChaosFeatures -and $sb.ChaosFeatures.Count) {
  Write-Host ("Chaos rolled: " + ($sb.ChaosFeatures -join '; '))
}
if ($sb.ChaosApplied -and $sb.ChaosApplied.Count) {
  Write-Host ("Applied: " + ($sb.ChaosApplied -join '; '))
}

# Base vs Final characteristics (with Delta)
$stats = 'STR','CON','SIZ','DEX','INT','POW','CHA'
$rows = foreach ($k in $stats) {
  $b = [int]($sb.BaseCharacteristics.$k)
  $f = [int]($sb.Characteristics.$k)
  if ($null -eq $b) { $b = $f } # fallback if BaseCharacteristics wasn't included
  [pscustomobject]@{
    Stat  = $k
    Base  = $b
    Delta = $f - $b
    Final = $f
  }
}
$rows | Format-Table -AutoSize
# Pretty print
$chars = $sb.Characteristics
Write-Host ("{0}: STR {1} CON {2} SIZ {3} DEX {4} INT {5} POW {6} CHA {7}" -f $sb.Creature,$chars.STR,$chars.CON,$chars.SIZ,$chars.DEX,$chars.INT,$chars.POW,$chars.CHA)
$moveTxt = Format-MoveText -Sb $sb
Write-Host ("HP {0}  Move {1} | Dex SR {2} Siz SR {3} | DB {4} | Spirit {5}" -f $sb.HP,$moveTxt,$sb.StrikeRanks.DexSR,$sb.StrikeRanks.SizSR,$sb.DamageBonus,$sb.SpiritCombat)

# print runes (handles missing values)
$runes = @()
if ($sb.Runes1) { $runes += "$($sb.Runes1) $($sb.Rune1Score)" }
if ($sb.Runes2) { $runes += "$($sb.Runes2) $($sb.Rune2Score)" }
if ($sb.Runes3) { $runes += "$($sb.Runes3) $($sb.Rune3Score)" }
Write-Host ("Runes: " + ($runes -join ', '))

if ($sb.ChaosFeatures -and $sb.ChaosFeatures.Count -gt 0) {
  Write-Host ("Chaos: " + ($sb.ChaosFeatures -join '; '))
}
if ($sb.ChaosArmorBonus -gt 0) {
  Write-Host ("Chaos Armor Bonus: +{0} (applied to all hit locations)" -f $sb.ChaosArmorBonus)
}

if ($sb.SpecialAttacks -and $sb.SpecialAttacks.Count -gt 0) {
  Write-Host "Special Attacks / Effects:"
  $sb.SpecialAttacks | ForEach-Object {
    Write-Host (" - {0}: {1}" -f $_.Name, $_.Description)
  }
}

$sb.HitLocations | Format-Table -AutoSize

# ========= WEAPONS OUTPUT (replaces your previous block) =========

# Weapons table: Damage + only short "+effect" style inline notes
$wepRows = $sb.Weapons |
  Select-Object `
    Name,
    @{ Name = 'Base %'; Expression = {
        $props = $_.PSObject.Properties
        $baseName = $props.Name |
          Where-Object {
            (($_ -replace '\u00A0',' ') -replace '\s+',' ') -match '^(?i)base %$' -or $_ -match '^(?i)skill$'
          } |
          Select-Object -First 1
        $v = if ($baseName) { $props[$baseName].Value } else { $_.'Base %' }
        $num = 0.0
        if ($null -eq $v -or "$v" -eq '' -or -not [double]::TryParse(("$v" -replace '[^\d\.-]',''), [ref]$num)) { '-' }
        elseif ([int]$num -eq 0) { '-' } else { [int]$num }
      } },
    @{ Name = 'Damage'; Expression = {
        $d = ('' + $_.Damage).Trim()
        if ([string]::IsNullOrWhiteSpace($d) -or $d -match '^(0|0\.0+|—|-)$') { $d = '-' }

        # candidate inline notes — ONLY compact "+effect" tokens (e.g., "+ poison+ acid")
        $n = ('' + $_.Notes).Trim()
        $inline = ''
        if ($n) {
          # Must look like series of "+word" tokens, and short
          $isCompactPlus = ($n -match '^\s*\+(?:\s*[A-Za-z][\w/-]*)+(?:\s*\+\s*[A-Za-z][\w/-]*)*\s*$')
          $shortEnough   = ($n.Length -le 28)
          if ($isCompactPlus -and $shortEnough) { $inline = $n }
        }

        if ($inline) { if ($d -eq '-') { $inline } else { "$d $inline" } } else { $d }
      } },
    @{ Name = 'HP'; Expression = {
        $raw = ('' + $_.HP).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { '-' }
        elseif ($raw -match '[A-Za-z]') { $raw }                 # keep body-part text as-is
        else {
          $num = 0.0
          if (-not [double]::TryParse(($raw -replace '[^\d\.-]',''), [ref]$num)) { '-' }
          elseif ([int]$num -eq 0) { '-' } else { [int]$num }
        }
      } },
    @{ Name = 'SR'; Expression = {
        $v = $_.SR; $num = 0.0
        if ($null -eq $v -or -not [double]::TryParse(("$v" -replace '[^\d\.-]',''), [ref]$num)) { '-' }
        elseif ([int]$num -eq 0) { '-' } else { [int]$num }
      } },
    @{ Name = 'Range'; Expression = {
        $r = ('' + $_.Range).Trim()
        if ([string]::IsNullOrWhiteSpace($r) -or $r -match '^(—|-)$') { '-' } else { $r }
      } }

$wepRows | Format-Table -AutoSize

# Under-table Notes: include SpecialText and any long/sentence Notes not inlined
function Write-WrappedNote {
  param(
    [Parameter(Mandatory)][string]$Label,
    [Parameter(Mandatory)][string]$Text,
    [int]$Width = 80
  )
  $label = ($Label -replace '\s+',' ').Trim()
  $text  = ($Text  -replace '\s+',' ').Trim()
  if (-not $text) { return }

  $prefix = "- $label :"
  $indent = ' ' * ($prefix.Length + 1)

  $words = $text.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
  $line = $prefix
  foreach ($w in $words) {
    if (($line.Length + 1 + $w.Length) -gt $Width) {
      Write-Host $line
      $line = "$indent$w"
    } else {
      $line = "$line $w"
    }
  }
  if ($line) { Write-Host $line }
}

# Build a list of label→text and dedupe
$notesToPrint = New-Object System.Collections.Generic.List[object]
$seen = @{}  # key = "$label|$text"

foreach ($row in $sb.Weapons) {
  $label = if ($row.SpecialName) { $row.SpecialName } else { $row.Name }

  # Always include SpecialText (rulesy/footnote/Note:)
  if ($row.SpecialText -and -not [string]::IsNullOrWhiteSpace([string]$row.SpecialText)) {
    $txt = ($row.SpecialText -replace '\s+',' ').Trim()
    $key = "$label|$txt"
    if (-not $seen.ContainsKey($key)) { $notesToPrint.Add([pscustomobject]@{ Label=$label; Text=$txt }); $seen[$key] = $true }
  }

  # Include long/sentence Notes that we didn't inline in Damage
  $n = ('' + $row.Notes).Trim()
  if ($n) {
    # Skip junky classifier fragments
    if ($n -match '^(?i)\bmeters?\s*dropped\b$' -or $n -match '^(?i)\brange\b$') { continue }

    $isCompactPlus = ($n -match '^\s*\+(?:\s*[A-Za-z][\w/-]*)+(?:\s*\+\s*[A-Za-z][\w/-]*)*\s*$')
    $shortEnough   = ($n.Length -le 28)
    if (-not ($isCompactPlus -and $shortEnough)) {
      $txt = ($n -replace '\s+',' ').Trim()
      $key = "$label|$txt"
      if (-not $seen.ContainsKey($key)) { $notesToPrint.Add([pscustomobject]@{ Label=$row.Name; Text=$txt }); $seen[$key] = $true }
    }
  }
}

if ($notesToPrint.Count -gt 0) {
  Write-Host ""
  Write-Host "Notes:"
  foreach ($n in $notesToPrint) { Write-WrappedNote -Label $n.Label -Text $n.Text -Width 45 }
}

# ========= END WEAPONS OUTPUT =========
