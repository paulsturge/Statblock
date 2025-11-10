# -----------------------------
# Statblock.ps1 (entry script)
# -----------------------------
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

  [string]$Cult,       # e.g. "Mallia"
  [string]$Role,       # e.g. "Initiate"
  [switch]$TwoHeaded,
  [switch]$ListCreatures,
  [int]$Seed,
  [switch]$ForceChaos
)

function Write-WrappedBlock {
  param(
    [Parameter(Mandatory)][string]$Title,
    [Parameter()][AllowNull()][AllowEmptyString()][string]$Text,
    [int]$Width = 80,
    [int]$Indent = 2
  )
  if ([string]::IsNullOrWhiteSpace($Text)) { return }
  Write-Host $Title
  $pad = ' ' * $Indent
  $s   = ($Text -replace '\s+', ' ').Trim()
  $lines = New-Object System.Collections.Generic.List[string]
  while ($s.Length -gt $Width) {
    $break = $s.LastIndexOf(' ', [math]::Min($Width, $s.Length-1))
    if ($break -le 0) { break }
    $lines.Add($s.Substring(0, $break))
    $s = $s.Substring($break + 1)
  }
  if ($s.Length -gt 0) { $lines.Add($s) }

  foreach ($i in 0..($lines.Count-1)) {
    if ($i -eq 0) { Write-Host ("- " + $lines[$i]) }
    else          { Write-Host ( $pad + $lines[$i]) }
  }
  Write-Host ""  # blank line
}

# Convenience alias: "Dragonsnail -2"
if ($Creature -match '^\s*Dragonsnail\s*-\s*2\s*$') {
  $TwoHeaded = $true
  $Creature  = 'Dragonsnail'
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

# Build bare statblock (no cult/magic yet)
$sb = New-Statblock -Creature $Creature -Context $ctx -AddArmor 0 -OverrideHitLocationSheet $overrideSheet -ForceChaos:$ForceChaos
Write-Host ("Hit locations sheet: {0}" -f $sb.HitLocationSheet)

# =========================
# CULT + MAGIC DECORATION
# =========================
# Goal: ALL cult / spirit magic / rune magic logic happens in Add-CultInfoToStatblock now.
# Statblock.ps1 no longer builds spell lists, spends budgets, or loops randomizers.

# Normalize inputs
$cultName = if ([string]::IsNullOrWhiteSpace($Cult)) { $null } else { $Cult }
$roleName = if ([string]::IsNullOrWhiteSpace($Role)) { $null } else { $Role }

# If they gave a cult but no role, assume Lay Member by default
if ($cultName -and -not $roleName) {
  $roleName = 'Lay Member'
}

# Import the one decorator we trust to mutate $sb safely
Import-Module "$PSScriptRoot\Authoring\Add-CultInfoToStatblock.psm1" -Force

# We will *try* to decorate with cult info. If Add-CultInfoToStatblock blows up,
# we catch it and keep going so we can STILL print a usable block.
if ($cultName) {
  try {
    $sb = Add-CultInfoToStatblock -Statblock $sb -CultName $cultName -Role $roleName
 } catch {
    # graceful fallback: make sure the fields exist in a sane shape so printing won't die

    # CultName / CultRole
    if ($sb.PSObject.Properties['CultName'].Count -eq 0) {
      $sb | Add-Member -NotePropertyName CultName -NotePropertyValue $cultName
    } else {
      $sb.CultName = $cultName
    }
    if ($sb.PSObject.Properties['CultRole'].Count -eq 0) {
      $sb | Add-Member -NotePropertyName CultRole -NotePropertyValue $roleName
    } else {
      $sb.CultRole = $roleName
    }

    # Magic should always end up as a hashtable-like object with optional Notes
    if ($sb.PSObject.Properties['Magic'].Count -eq 0) {
      # doesn't exist at all yet
      $sb | Add-Member -NotePropertyName Magic -NotePropertyValue @{
        Notes = 'Cult magic not generated (failed to attach cleanly).'
      }
    } else {
      # exists, but make sure it's not a bare string or some other scalar
      if ($sb.Magic -isnot [System.Collections.IDictionary]) {
        $sb.Magic = @{
          Notes = 'Cult magic not generated (failed to attach cleanly).'
        }
      } else {
        # it's already a dictionary/hashtable; ensure it has Notes
        if (-not $sb.Magic['Notes']) {
          $sb.Magic['Notes'] = 'Cult magic not generated (failed to attach cleanly).'
        }
      }
    }

    # RuneMagic: force it to be at least an empty hashtable so the renderer won't choke
    if ($sb.PSObject.Properties['RuneMagic'].Count -eq 0) {
      $sb | Add-Member -NotePropertyName RuneMagic -NotePropertyValue @{}
    } else {
      if ($sb.RuneMagic -isnot [System.Collections.IDictionary] -and
          $sb.RuneMagic -isnot [psobject]) {
        $sb.RuneMagic = @{}
      }
    }
}

} else {
  # No cult provided: ensure the properties exist so the printer never chokes
  if ($sb.PSObject.Properties['CultName'].Count -eq 0) { $sb | Add-Member -NotePropertyName CultName -NotePropertyValue $null }
  if ($sb.PSObject.Properties['CultRole'].Count -eq 0) { $sb | Add-Member -NotePropertyName CultRole -NotePropertyValue $null }
  if ($sb.PSObject.Properties['Magic'].Count -eq 0)    { $sb | Add-Member -NotePropertyName Magic     -NotePropertyValue @{} }
  if ($sb.PSObject.Properties['RuneMagic'].Count -eq 0){ $sb | Add-Member -NotePropertyName RuneMagic -NotePropertyValue @{} }
}

# =========================
# CHAOS / CHARACTERISTICS
# =========================

if ($sb.ChaosFeatures -and $sb.ChaosFeatures.Count) {
  Write-Host ("Chaos rolled: " + ($sb.ChaosFeatures -join '; '))
}
if ($sb.ChaosApplied -and $sb.ChaosApplied.Count) {
  Write-Host ("Applied: " + ($sb.ChaosApplied -join '; '))
}

# compare Base vs Final characteristics
$stats = 'STR','CON','SIZ','DEX','INT','POW','CHA'
$rows = foreach ($k in $stats) {
  $b = [int]($sb.BaseCharacteristics.$k)
  $f = [int]($sb.Characteristics.$k)
  if ($null -eq $b) { $b = $f }
  [pscustomobject]@{
    Stat  = $k
    Base  = $b
    Delta = $f - $b
    Final = $f
  }
}
$rows | Format-Table -AutoSize

# Movement text (use Move cell from StatDice)
$chars = $sb.Characteristics
$moveCell = ($ctx.StatDice | Where-Object { [string]$_.'Creature' -eq $sb.Creature } | Select-Object -First 1).Move
if ([string]::IsNullOrWhiteSpace([string]$moveCell)) {
  $moveText = 'Move: -'
} else {
  $core = ('' + $moveCell) -replace '^(?i)\s*Move\s*:\s*',''
  $moveText = "Move: $($core.Trim())"
}

Write-Host ("{0}: STR {1} CON {2} SIZ {3} DEX {4} INT {5} POW {6} CHA {7}" -f $sb.Creature,$chars.STR,$chars.CON,$chars.SIZ,$chars.DEX,$chars.INT,$chars.POW,$chars.CHA)
Write-Host ("HP {0}  {1} | Dex SR {2} Siz SR {3} | DB {4} | Spirit {5}" -f $sb.HP,$moveText,$sb.StrikeRanks.DexSR,$sb.StrikeRanks.SizSR,$sb.DamageBonus,$sb.SpiritCombat)

# --- Cult / Role summary line ---

$cultPrintName = @(
  $sb.CultName
  $sb.Cult
  $sb.CultInfo?.Name
  $sb.CultDetails?.Name
) | Where-Object { $_ } | Select-Object -First 1

$rolePrintName = @(
  $sb.Role
  $sb.CultRole
  $sb.CultInfo?.Role
) | Where-Object { $_ } | Select-Object -First 1

if (-not $cultPrintName -and $Cult) { $cultPrintName = $Cult }
if (-not $rolePrintName -and $Role) { $rolePrintName = $Role }

if ($cultPrintName -or $rolePrintName) {
  $cn = if ($cultPrintName) { $cultPrintName } else { '-' }
  $rn = if ($rolePrintName) { " ($rolePrintName)" } else { '' }
  Write-Host ("Cult: {0}{1}" -f $cn, $rn)
}

# Runes
$runes = @()
if ($sb.Runes1) { $runes += "$($sb.Runes1) $($sb.Rune1Score)" }
if ($sb.Runes2) { $runes += "$($sb.Runes2) $($sb.Rune2Score)" }
if ($sb.Runes3) { $runes += "$($sb.Runes3) $($sb.Rune3Score)" }
if ($runes.Count -gt 0) {
  Write-Host ("Runes: " + ($runes -join ', '))
}

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

# --- Extras: Skills / Languages / Passions ---

Write-WrappedBlock -Title 'Skills:'    -Text $sb.Skills    -Width 45 -Indent 10
Write-WrappedBlock -Title 'Languages:' -Text $sb.Languages -Width 45 -Indent 10
Write-WrappedBlock -Title 'Passions:'  -Text $sb.Passions  -Width 45 -Indent 10

# --- Magic / Spirit Magic display (with point costs) ---
"Magic:"

$magicLines = @()

if ($sb.Magic) {

    foreach ($key in $sb.Magic.Keys) {
        $val = $sb.Magic[$key]

        # Skip obvious bookkeeping buckets we don't want to show as separate magic lines themselves
        if ($null -eq $val) { continue }
        if ($key -match '^Rune' -or $key -match 'RunePoints') { continue }

        # Case 1: plain string ("Spirit magic loadout auto-budgeted...")
        if ($val -is [string]) {
            if ($val.Trim() -ne "") {
                # Treat Notes specially so it shows first
                if ($key -match '^(?i)notes$') {
                    $magicLines += (" - Notes : {0}" -f $val.Trim())
                }
                else {
                    $magicLines += (" - {0} : {1}" -f $key, $val.Trim())
                }
            }
            continue
        }

        # Case 2: collection (this is usually the Spirit spell list)
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {

            foreach ($item in $val) {
                if ($null -eq $item) { continue }

                # item is a simple string like "Strength"? just print it.
                if ($item -is [string]) {
                    $magicLines += (" - {0}" -f $item)
                    continue
                }

                # item is an object with Name / Points / Notes
                $name   = $null
                $pts    = $null
                $notes  = $null

                if ($item.PSObject.Properties['Name'])   { $name  = ('' + $item.Name).Trim() }
                if ($item.PSObject.Properties['Points']) { $pts   = [string]$item.Points }
                if ($item.PSObject.Properties['Notes'])  { $notes = ('' + $item.Notes).Trim() }

                if ($name) {
                    # Build "Strength (2 pts)" if we have a numeric Point cost
                    $line = $name
                    if ($pts -and $pts -ne '' -and $pts -ne '0') {
                        # normalize "1" -> "1 pt", "2" -> "2 pts"
                        $plural = if ([string]$pts -eq '1') { 'pt' } else { 'pts' }
                        $line = "{0} ({1} {2})" -f $line, $pts, $plural
                    }

                    # Optionally add trailing short note if present
                    if ($notes -and $notes -ne '') {
                        # keep notes short; if it's a wall of text we skip
                        if ($notes.Length -le 40) {
                            $line = "{0} - {1}" -f $line, $notes
                        }
                    }

                    $magicLines += (" - {0}" -f $line)
                    continue
                }

                # fallback: if we couldn't parse Name but we have Spell/Points pattern
                if ($item.PSObject.Properties['Spell']) {
                    $spellLine = ('' + $item.Spell).Trim()
                    $p2        = $null
                    if ($item.PSObject.Properties['Points']) {
                        $p2 = [string]$item.Points
                    }
                    if ($p2 -and $p2 -ne '' -and $p2 -ne '0') {
                        $plural2 = if ($p2 -eq '1') { 'pt' } else { 'pts' }
                        $spellLine = "{0} ({1} {2})" -f $spellLine, $p2, $plural2
                    }
                    $magicLines += (" - {0}" -f $spellLine)
                    continue
                }

                # If we get here, we don't know how to render that entry; skip it rather than dumping junk
            }

            continue
        }

        # Case 3: single hashtable / psobject (rare corner case)
        if ($val -is [hashtable] -or $val -is [psobject]) {

            $name   = $null
            $pts    = $null
            $notes  = $null

            if ($val.PSObject.Properties['Name'])   { $name  = ('' + $val.Name).Trim() }
            if ($val.PSObject.Properties['Points']) { $pts   = [string]$val.Points }
            if ($val.PSObject.Properties['Notes'])  { $notes = ('' + $val.Notes).Trim() }

            if ($name) {
                $line = $name
                if ($pts -and $pts -ne '' -and $pts -ne '0') {
                    $plural = if ($pts -eq '1') { 'pt' } else { 'pts' }
                    $line = "{0} ({1} {2})" -f $line, $pts, $plural
                }
                if ($notes -and $notes -ne '') {
                    if ($notes.Length -le 40) {
                        $line = "{0} - {1}" -f $line, $notes
                    }
                }
                $magicLines += (" - {0}" -f $line)
                continue
            }

            # Otherwise try Spell/Points combo
            if ($val.PSObject.Properties['Spell']) {
                $spellLine = ('' + $val.Spell).Trim()
                $p2        = $null
                if ($val.PSObject.Properties['Points']) {
                    $p2 = [string]$val.Points
                }
                if ($p2 -and $p2 -ne '' -and $p2 -ne '0') {
                    $plural2 = if ($p2 -eq '1') { 'pt' } else { 'pts' }
                    $spellLine = "{0} ({1} {2})" -f $spellLine, $p2, $plural2
                }
                $magicLines += (" - {0}" -f $spellLine)
                continue
            }

            # last-ditch: "Notes" only
            if ($notes -and $notes -ne '') {
                $magicLines += (" - {0}" -f $notes)
            }

            continue
        }

        # If we get here, we couldn't figure out how to print $val in a reasonable way
    }
}

if ($magicLines.Count -gt 0) {
    $magicLines
} else {
    " - None."
}

""


# --- Rune Magic display ---
$hasRuneMagic = $false
$runeLines    = @()

if ($sb.RuneMagic) {
  if ($sb.RuneMagic.PSObject.Properties['Spells']) {
    $spells = $sb.RuneMagic.Spells
    if ($spells) {
      foreach ($s in $spells) {
        if ($null -eq $s) { continue }

        if ($s -is [string]) {
          $runeLines += " - $s"
          continue
        }

        if ($s.PSObject.Properties['Name'] -and $s.Name) {
          $runeLines += " - $($s.Name)"
          continue
        }
      }
    }
  }

  $specialText = $null
  if ($sb.RuneMagic.PSObject.Properties['Special']) {
    $tmp = "$($sb.RuneMagic.Special)".Trim()
    if ($tmp -ne "") {
      $specialText = $tmp
    }
  }

  if ($runeLines.Count -gt 0 -or $specialText) {
    $hasRuneMagic = $true
    "Rune Magic:"
    if ($specialText) {
      "  Special: $specialText"
    }
    $runeLines
    ""
  }
}

Write-WrappedBlock -Title 'Magic Notes:' -Text $sb.MagicNotes -Width 45 -Indent 10

# --- Hit Locations ---
$sb.HitLocations | Format-Table -AutoSize

# --- Weapons block (unchanged from your version) ---

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

        $n = ('' + $_.Notes).Trim()
        $inline = ''
        if ($n) {
          $isCompactPlus = ($n -match '^\s*\+(?:\s*[A-Za-z][\w/-]*)+(?:\s*\+\s*[A-Za-z][\w/-]*)*\s*$')
          $shortEnough   = ($n.Length -le 28)
          if ($isCompactPlus -and $shortEnough) { $inline = $n }
        }

        if ($inline) { if ($d -eq '-') { $inline } else { "$d $inline" } } else { $d }
      } },
    @{ Name = 'HP'; Expression = {
        $raw = ('' + $_.HP).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { '-' }
        elseif ($raw -match '[A-Za-z]') { $raw }
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

$notesToPrint = New-Object System.Collections.Generic.List[object]
$seen = @{}

foreach ($row in $sb.Weapons) {
  $label = if ($row.SpecialName) { $row.SpecialName } else { $row.Name }

  if ($row.SpecialText -and -not [string]::IsNullOrWhiteSpace([string]$row.SpecialText)) {
    $txt = ($row.SpecialText -replace '\s+',' ').Trim()
    $key = "$label|$txt"
    if (-not $seen.ContainsKey($key)) {
      $notesToPrint.Add([pscustomobject]@{ Label=$label; Text=$txt })
      $seen[$key] = $true
    }
  }

  $n = ('' + $row.Notes).Trim()
  if ($n) {
    if ($n -match '^(?i)\bmeters?\s*dropped\b$' -or $n -match '^(?i)\brange\b$') { continue }

    $isCompactPlus = ($n -match '^\s*\+(?:\s*[A-Za-z][\w/-]*)+(?:\s*\+\s*[A-Za-z][\w/-]*)*\s*$')
    $shortEnough   = ($n.Length -le 28)
    if (-not ($isCompactPlus -and $shortEnough)) {
      $txt = ($n -replace '\s+',' ').Trim()
      $key = "$label|$txt"
      if (-not $seen.ContainsKey($key)) {
        $notesToPrint.Add([pscustomobject]@{ Label=$row.Name; Text=$txt })
        $seen[$key] = $true
      }
    }
  }
}

if ($notesToPrint.Count -gt 0) {
  Write-Host ""
  Write-Host "Notes:"
  foreach ($n in $notesToPrint) {
    Write-WrappedNote -Label $n.Label -Text $n.Text -Width 45
  }
}
