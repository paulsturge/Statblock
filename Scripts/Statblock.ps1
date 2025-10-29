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
  [string]$Cult,       
  [string]$Role,       
  [switch]$TwoHeaded,
  [switch]$ListCreatures,   # 👈 add this back
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
  Write-Host ""  # blank line after block
}

# Allow 'Dragonsnail -2' as a single value for convenience
if ($Creature -match '^\s*Dragonsnail\s*-\s*2\s*$') {
  $TwoHeaded = $true
  $Creature  = 'Dragonsnail'
}

# (REMOVED) Format-MoveText – we now print Move verbatim from StatDice

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

# --- Optional cult decoration + Spirit Magic allocation (only if provided) ---
if ($Cult -and $Role) {
  # Load authoring/randomizer helpers (safe to re-import)
  Import-Module "$PSScriptRoot\Authoring\Get-CultData.psm1" -Force
  Import-Module "$PSScriptRoot\Authoring\Add-CultInfoToStatblock.psm1" -Force
  Import-Module "$PSScriptRoot\Authoring\SpiritMagicRandomizer.psm1" -Force

  # Apply cult + role details
  $sb = Add-CultInfoToStatblock -Statblock $sb -CultName $Cult -Role $Role

  # Longhand CHA
  $cha = [int]$sb.Characteristics.CHA

  # Load spirit-magic catalog
  $catalogPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Data\spirit_magic_catalog.csv'
  if (Test-Path $catalogPath) {
    $cat = Import-SpiritMagicCatalog -CsvPath $catalogPath

    # Budget by role (you can tweak helper ranges later); cap by CHA
    $budget = Get-SpiritBudgetByRole -Role $Role -CHA $cha

    if ($budget -gt 0 -and $cat -and $cat.Count -gt 0) {
      # Roll spells (respects RoleMax_* caps in the CSV)
      $seed = if ($PSBoundParameters.ContainsKey('Seed')) { $Seed } else { (Get-Random) }
      $rolls = New-RandomSpiritMagicLoadout -PointsBudget $budget -CHA $cha -Role $Role -Catalog $cat -Seed $seed

      # Apply to $sb (your fixed setter in SpiritMagicRandomizer.psm1)
      $sb = Set-StatblockSpiritMagic $sb $rolls
    }
  } else {
    Write-Host "Note: Spirit-magic catalog not found at $catalogPath — skipping spells." -ForegroundColor DarkYellow
  }
}
# ---------------------------------------------------------------------------

# --- Rune Magic allocation (role + INT; pulls Rune from Cults.xlsx) ----------
# Uses the same allocator module (now contains Rune helpers too)
if (-not (Get-Module SpiritMagicRandomizer -ListAvailable | Select-Object -First 1)) {
  Import-Module "$PSScriptRoot\Authoring\SpiritMagicRandomizer.psm1" -Force
}

# Longhand stats + defaults
$intLong       = [int]$sb.Characteristics.INT
$roleForRune   = if ($PSBoundParameters.ContainsKey('Role') -and $Role) { $Role } else { 'Initiate' }
$cultForRune   = if ($PSBoundParameters.ContainsKey('Cult') -and $Cult) { $Cult } else { ($sb.CultName ?? 'Thed') }
$cultsWorkbook = "Y:\Stat_blocks\Data\Cults.xlsx"

# Points: only creatures with INT get Rune Points (your rule)
$runePoints         = New-RunePoints -Role $roleForRune -INT $intLong
$special, $common   = New-RuneSpellLoadout -RunePoints $runePoints -CultName $cultForRune -WorkbookPath $cultsWorkbook -IncludeAssociates

$sb = Set-StatblockRuneMagic $sb $runePoints $special $common
# -----------------------------------------------------------------------------


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

# Pretty print line (now uses Move verbatim from StatDice)
$chars = $sb.Characteristics

# Pull the Move cell for this creature from StatDice
$moveCell = ($ctx.StatDice | Where-Object { [string]$_.'Creature' -eq $sb.Creature } | Select-Object -First 1).Move
# Normalize: always print with "Move: " prefix, but avoid double-prefix if it’s already present
if ([string]::IsNullOrWhiteSpace([string]$moveCell)) {
  $moveText = 'Move: -'
} else {
  $core = ('' + $moveCell) -replace '^(?i)\s*Move\s*:\s*',''
  $moveText = "Move: $($core.Trim())"
}

Write-Host ("{0}: STR {1} CON {2} SIZ {3} DEX {4} INT {5} POW {6} CHA {7}" -f $sb.Creature,$chars.STR,$chars.CON,$chars.SIZ,$chars.DEX,$chars.INT,$chars.POW,$chars.CHA)
Write-Host ("HP {0}  {1} | Dex SR {2} Siz SR {3} | DB {4} | Spirit {5}" -f $sb.HP,$moveText,$sb.StrikeRanks.DexSR,$sb.StrikeRanks.SizSR,$sb.DamageBonus,$sb.SpiritCombat)

# --- Show Cult + Role/Level (robust to different property names) ---
$cultName = @(
  $sb.CultName
  $sb.Cult
  $sb.CultInfo?.Name
  $sb.CultDetails?.Name
) | Where-Object { $_ } | Select-Object -First 1

$roleName = @(
  $sb.Role
  $sb.CultRole
  $sb.CultInfo?.Role
) | Where-Object { $_ } | Select-Object -First 1

# fall back to parameters if present
if (-not $cultName -and $PSBoundParameters.ContainsKey('Cult') -and $Cult) { $cultName = $Cult }
if (-not $roleName -and $PSBoundParameters.ContainsKey('Role') -and $Role) { $roleName = $Role }

if ($cultName -or $roleName) {
  $cn = if ($cultName) { $cultName } else { '-' }
  $rn = if ($roleName) { " ($roleName)" } else { '' }
  Write-Host ("Cult: {0}{1}" -f $cn, $rn)
}


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

# --- Optional narrative/authoring sections (only print if present) ---
Write-WrappedBlock -Title 'Skills:'       -Text $sb.Skills      -Width 45 -Indent 10
Write-WrappedBlock -Title 'Languages:'    -Text $sb.Languages   -Width 45 -Indent 10
Write-WrappedBlock -Title 'Passions:'     -Text $sb.Passions    -Width 45 -Indent 10
Write-WrappedBlock -Title 'Magic:'        -Text $sb.Magic       -Width 45 -Indent 10
# Display Spirit spells if present (works whether Magic is object or hashtable)
$sp = if ($sb.Magic -is [System.Collections.IDictionary]) { @($sb.Magic['Spirit']) } else { @($sb.Magic.Spirit) }
if ($sp -and $sp.Count -gt 0) {
  Write-Host "Spirit Magic:"
  $sp | Sort-Object Name | Format-Table Name, Points -Auto
  "Total Spirit Points: " + (($sp | Measure-Object Points -Sum).Sum) + " / CHA " + $sb.Characteristics.CHA
}
# Rune Magic display (matches allocation above)
if ($sb.PSObject.Properties['RuneMagic']) {
  Write-Host ("Rune Magic: {0} Rune Points" -f $sb.RuneMagic.Points)

  if ($sb.RuneMagic.Special -and $sb.RuneMagic.Special.Count -gt 0) {
    Write-Host "  Special:"
    @($sb.RuneMagic.Special) | Sort-Object Name | Format-Table Name -Auto
  }

  if ($sb.RuneMagic.Common -and $sb.RuneMagic.Common.Count -gt 0) {
    Write-Host "  Common (always available):"
    @($sb.RuneMagic.Common) | Sort-Object Name | Format-Table Name -Auto
  }
}


Write-WrappedBlock -Title 'Magic Notes:'  -Text $sb.MagicNotes  -Width 45 -Indent 10

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
