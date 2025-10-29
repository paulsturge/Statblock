param(
  [string]$WorkbookPath = 'Y:\Stat_blocks\Data\Cults.xlsx',
  [string]$OutDir = 'Y:\Stat_blocks\Data\Cults_Clean',
  [switch]$Force,
  [switch]$VerboseLog
)

function Ensure-Folder { param([string]$Path) if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null } }
function Export-CsvUtf8 {
  param([Parameter(Mandatory)][object]$Data, [Parameter(Mandatory)][string]$Path, [switch]$Force)
  if ((-not $Force) -and (Test-Path $Path)) { Write-Host "Skip (exists): $Path"; return }
  $Data | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $Path
  Write-Host "Wrote: $Path"
}
function Normalize-Range {
  param([string]$Cell)
  if ([string]::IsNullOrWhiteSpace($Cell)) { return @{Low=$null;High=$null} }
  $s = (''+$Cell).Trim() -replace '\s',''
  $s = $s -replace '–','-'
  if ($s -match '^(\d{1,3})-(\d{1,3})$') { return @{Low=[int]$matches[1];High=[int]$matches[2]} }
  elseif ($s -match '^(\d{1,3})$')      { $n=[int]$matches[1]; return @{Low=$n;High=$n} }
  else                                  { return @{Low=$null;High=$null} }
}

function Read-Sheet {
  param([string]$Path, [string]$WorksheetName)
  try {
    Import-Excel -Path $Path -WorksheetName $WorksheetName -StartRow 1 -HeaderRow 1 -ErrorAction Stop
  } catch {
    if ($VerboseLog) { Write-Warning "Failed to read '$WorksheetName': $($_.Exception.Message)" }
    @()
  }
}

# Parse names like:
#  "Broo_Cult_Affiliation", "Broo_Affiliation", "Thanatar Guardian Summons",
#  "Vivamort-Spells", "Krasht Skills"
function Parse-SheetName {
  param([string]$Name)
  $n = $Name.Trim()

  $kinds = 'Spells','Skills','Affiliation','Gifts','Geases','GuardianSummons'
  # normalized variant for convenience
  $kindMap = @{
    'spell'='Spells'; 'spells'='Spells';
    'skill'='Skills'; 'skills'='Skills';
    'affiliation'='Affiliation'; 'cult_affiliation'='Affiliation';
    'gifts'='Gifts'; 'gift'='Gifts';
    'geases'='Geases'; 'geas'='Geases';
    'guardian summons'='GuardianSummons'; 'guardiansummons'='GuardianSummons';
  }

  # Try underscore form first (allow optional "Cult_")
  if ($n -match '^(?<cult>.+?)_(?:Cult_)?(?<kind>Spells|Skills|Affiliation|Gifts|Geases|GuardianSummons)$') {
    return @{ Cult=$matches['cult']; Kind=$matches['kind'] }
  }

  # Try loose separators: space or dash; also tolerate "Cult Affiliation"
  if ($n -match '^(?<cult>.+?)[\s\-]+(?<kind>.+)$') {
    $cult  = $matches['cult'].Trim()
    $kindR = $matches['kind'].Trim().ToLower()
    $kindR = $kindR -replace '\s+',' '
    if ($kindMap.ContainsKey($kindR)) {
      return @{ Cult=$cult; Kind=$kindMap[$kindR] }
    }
    # also support "Cult Affiliation"
    if ($kindR -match '^cult\s+affiliation$') {
      return @{ Cult=$cult; Kind='Affiliation' }
    }
  }

  return $null
}

# ---------- Normalizers ----------
function Normalize-SpellsAndSkills {
  param([object[]]$Rows)
  foreach ($r in $Rows) {
    $spell = ('' + ($r.Spell ?? $r.Name ?? $r.'Spell Name')).Trim()
    $type  = ('' + ($r.Type ?? $r.Category)).Trim()
    $notes = ('' + ($r.Notes ?? $r.Description ?? $r.Detail)).Trim()
    $cost  = ('' + ($r.Cost)).Trim()
    if (-not $type) { continue }

    $isSpell = $type -match '^(?i)(Common\s*Rune|Special\s*Rune|Spirit)\b'
    $isSkill = $type -match '^(?i)(Skill|Training|Unique|Half\s*Normal|Normal|Twice\s*Normal)$'

    if ($isSpell -and $spell) {
      [pscustomobject]@{ Spell=$spell; Type=$type; Notes=$notes }
    } elseif ($isSkill -and $spell) {
      [pscustomobject]@{ Skill=$spell; Type=$type; Cost=$cost; Notes=$notes }
    }
  }
}
function Normalize-Affiliation {
  param([object[]]$Rows)
  foreach ($r in $Rows) {
    $rangeRaw = ('' + ($r.D100 ?? $r.Range ?? $r.Roll ?? $r.'D-100' ?? $r.'d100')).Trim()
    $cult     = ('' + ($r.Cult ?? $r.Result ?? $r.Affiliation)).Trim()
    $notes    = ('' + ($r.Notes)).Trim()
    if (-not $rangeRaw) { continue }
    $rng = Normalize-Range $rangeRaw
    if ($null -eq $rng.Low -or $null -eq $rng.High) { continue }
    if (-not $cult) { continue }
    [pscustomobject]@{ RollLow=$rng.Low; RollHigh=$rng.High; Cult=$cult; Notes=$notes }
  }
}
function Normalize-Gifts {
  param([object[]]$Rows)
  foreach ($r in $Rows) {
    $gift   = ('' + ($r.Gift ?? $r.Name ?? $r.Benefit)).Trim()
    $geases = ('' + ($r.'Number of Geases' ?? $r.Geases ?? $r.GeasCount)).Trim()
    $notes  = ('' + ($r.Notes ?? $r.Detail)).Trim()
    if (-not $gift) { continue }
    $gInt = $null; [void][int]::TryParse(($geases -replace '[^\d-]',''), [ref]$gInt)
    [pscustomobject]@{ Gift=$gift; GeasCount=$gInt; Notes=$notes }
  }
}
function Normalize-Geases {
  param([object[]]$Rows)
  foreach ($r in $Rows) {
    $rangeRaw = ('' + ($r.D100 ?? $r.Range ?? $r.Roll)).Trim()
    $text     = ('' + ($r.Geas ?? $r.Text ?? $r.Requirement)).Trim()
    if (-not $rangeRaw) { continue }
    $rng = Normalize-Range $rangeRaw
    if ($null -eq $rng.Low -or $null -eq $rng.High) { continue }
    if (-not $text) { continue }
    [pscustomobject]@{ RollLow=$rng.Low; RollHigh=$rng.High; Text=$text }
  }
}
function Normalize-GuardianSummons {
  param([object[]]$Rows)
  foreach ($r in $Rows) {
    $rangeRaw   = ('' + ($r.D100 ?? $r.Range ?? $r.Roll)).Trim()
    $attitude   = ('' + ($r.Type ?? $r.Attitude)).Trim()
    $powSpec    = ('' + ($r.POW)).Trim()
    $intSpec    = ('' + ($r.INT)).Trim()
    $runeSpec   = ('' + ($r.Rune ?? $r.'Rune Spells' ?? $r.Runes)).Trim()
    $battleSpec = ('' + ($r.Battle ?? $r.'Battle Magic')).Trim()
    if (-not $rangeRaw) { continue }
    $rng = Normalize-Range $rangeRaw
    if ($null -eq $rng.Low -or $null -eq $rng.High) { continue }
    [pscustomobject]@{
      RollLow=$rng.Low; RollHigh=$rng.High; Attitude=$attitude;
      POW_Spec=$powSpec; INT_Spec=$intSpec; Rune_Spec=$runeSpec; Battle_Spec=$battleSpec
    }
  }
}

# ---------- Main ----------
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
Ensure-Folder $OutDir

Import-Module ImportExcel -ErrorAction Stop | Out-Null
$pkg = Open-ExcelPackage -Path $WorkbookPath -ErrorAction Stop
$worksheets = ,@($pkg.Workbook.Worksheets)

Write-Host "Found $($worksheets.Count) sheets in $WorkbookPath"
$recognized = 0

foreach ($ws in $worksheets) {
  $info = Parse-SheetName $ws.Name
  if (-not $info) {
    if ($VerboseLog) { Write-Host "Skip (unrecognized name): '$($ws.Name)'" }
    continue
  }
  $recognized++
  $cult = $info.Cult.Trim()
  $kind = $info.Kind.Trim()
  $cultFolder = Join-Path $OutDir ($cult -replace '\s+','')
  Ensure-Folder $cultFolder

  $rows = Read-Sheet -Path $WorkbookPath -WorksheetName $ws.Name
  if ($VerboseLog) { Write-Host "Sheet '$($ws.Name)' -> $cult/$kind : $($rows.Count) raw rows" }

  switch ($kind) {
    'Spells' {
      $normalized = Normalize-SpellsAndSkills $rows
      $spellRows = @(); $skillRows = @()
      foreach ($n in $normalized) {
        if ($n.PSObject.Properties['Spell']) { $spellRows += $n }
        elseif ($n.PSObject.Properties['Skill']) { $skillRows += $n }
      }
      if ($spellRows.Count) { Export-CsvUtf8 -Data $spellRows -Path (Join-Path $cultFolder ("{0}_Spells.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
      if ($skillRows.Count) { Export-CsvUtf8 -Data $skillRows -Path (Join-Path $cultFolder ("{0}_Skills.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
    'Skills' {
      $skillRows = Normalize-SpellsAndSkills $rows | Where-Object { $_.PSObject.Properties['Skill'] }
      if ($skillRows.Count) { Export-CsvUtf8 -Data $skillRows -Path (Join-Path $cultFolder ("{0}_Skills.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
    'Affiliation' {
      $aff = Normalize-Affiliation $rows
      if ($aff.Count) { Export-CsvUtf8 -Data $aff -Path (Join-Path $cultFolder ("{0}_Affiliation.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
    'Gifts' {
      $g = Normalize-Gifts $rows
      if ($g.Count) { Export-CsvUtf8 -Data $g -Path (Join-Path $cultFolder ("{0}_Gifts.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
    'Geases' {
      $g = Normalize-Geases $rows
      if ($g.Count) { Export-CsvUtf8 -Data $g -Path (Join-Path $cultFolder ("{0}_Geases.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
    'GuardianSummons' {
      $g = Normalize-GuardianSummons $rows
      if ($g.Count) { Export-CsvUtf8 -Data $g -Path (Join-Path $cultFolder ("{0}_GuardianSummons.csv" -f ($cult -replace '\s+',''))) -Force:$Force }
    }
  }
}

Write-Host "Recognized sheets: $recognized / $($worksheets.Count)"
Write-Host "Done. Output: $OutDir"
