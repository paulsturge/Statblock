function Import-CreatureWeaponsFromText {
  [CmdletBinding(DefaultParameterSetName='Clipboard')]
  param(
    [Parameter(Mandatory)][string]$Creature,
    [Parameter(ParameterSetName='Clipboard')][switch]$FromClipboard,
    [Parameter(ParameterSetName='Text')][string]$Text,

    # default table for names not mapped below: melee | missile | shields
    [ValidateSet('melee','missile','shields')]
    [string]$DefaultTable = 'melee',

    # name→table routing (case-insensitive exact name match)
    [hashtable]$NameToTable = @{ 'Spit'='missile' },
    [switch]$Overwrite
  )

  # --- helpers (self-contained) ---
  function Resolve-StatPath {
    param([Parameter(Mandatory)][string]$Name)

    if ([System.IO.Path]::IsPathRooted($Name)) {
      if (Test-Path -LiteralPath $Name) { return (Resolve-Path -LiteralPath $Name).Path }
      return $null
    }

    $candidates = @()
    if ($script:DataRoot) { $candidates += $script:DataRoot }
    if ($PSScriptRoot) { $candidates += $PSScriptRoot; $candidates += (Split-Path $PSScriptRoot -Parent) }
    if ($pwd) { $candidates += $pwd.Path }

    foreach ($root in ($candidates | Select-Object -Unique)) {
      $p = Join-Path $root $Name
      if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
    }
    return $null
  }
  function _resolve($name) {
    if (Get-Command Resolve-StatPath -ErrorAction SilentlyContinue) { return (Resolve-StatPath $name) }
    $candidates = @($PSScriptRoot, (Split-Path $PSScriptRoot -Parent), $pwd.Path)
    foreach ($r in $candidates) {
      $p = Join-Path $r $name
      if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
    }
    return $null
  }
  function _loadTable($which) {
    $file = switch ($which) {
      'melee'   { 'weapons_melee_table.xlsx' }
      'missile' { 'weapons_missile_table.xlsx' }
      'shields' { 'weapons_shields_table.xlsx' }
    }
    $path = _resolve $file
    if (-not $path) { throw "Weapons table not found: $file" }
    $sheetName = (Get-ExcelSheetInfo -Path $path | Select-Object -First 1 -ExpandProperty Name)
    $data = Import-Excel -Path $path -WorksheetName $sheetName | Where-Object { $_.PSObject.Properties.Value -ne $null }
    [pscustomobject]@{ Path=$path; Sheet=$sheetName; Data=$data }
  }
  function _ensureCols([ref]$data) {
    $need = @('Name','Base %','Damage','HP','SR','Range','Notes','Creature','SpecialName','SpecialText')
    foreach ($col in $need) {
      if (-not ($data.Value | Get-Member -Name $col -MemberType NoteProperty)) {
        $data.Value | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName $col -NotePropertyValue $null -Force }
      }
    }
  }
  function _saveTable($tobj) {
    Export-Excel -Path $tobj.Path -WorksheetName $tobj.Sheet -ClearSheet -AutoSize -FreezeTopRow -BoldTopRow -InputObject $tobj.Data
  }
  function _whichTable([string]$name) {
    foreach ($k in $NameToTable.Keys) { if ($name -eq $k) { return $NameToTable[$k] } }
    return $DefaultTable
  }

  # --- get text ---
  if ($PSCmdlet.ParameterSetName -eq 'Clipboard') {
    try { $Text = Get-Clipboard -Raw } catch { throw "Clipboard read failed: $_" }
  }
  if ([string]::IsNullOrWhiteSpace($Text)) { throw "No input text provided." }

# --- split lines + find header robustly ---
$lines = (($Text -replace '\r','') -split '\n') | ForEach-Object { ($_ -replace '[\u00A0]',' ') -replace '\t',' ' }

$hdrIdx = $null
for ($j = 0; $j -lt $lines.Count; $j++) {
  $norm = ($lines[$j] -replace '\s+',' ').Trim().ToLower()
  # look for a line that contains: weapon, %, damage, sr (case/spacing tolerant)
  if ($norm -match 'weapon' -and $norm -match '%' -and $norm -match 'damage' -and $norm -match '(^| )sr( |$)') {
    Write-Verbose "Header detected on line $($j+1): $($lines[$j])"
    $hdrIdx = $j + 1  # start reading rows after the header line
    break
  }
}

if ($null -eq $hdrIdx) { throw "Weapons header line (contains 'Weapon', '%', 'Damage', 'SR') not found." }
$start = $hdrIdx

  # --- parse rows under "Weapon % Damage SR" (handles stars on name/damage/after) ---
  $weaponRows = @()
  $i = $start
  for ($i = $start; $i -lt $lines.Count; $i++) {
    $ln = $lines[$i].Trim()
    if (-not $ln) { break }              # blank ends table
    if ($ln -match '^\*+') { break }     # notes start

    # Name  %  Damage-ish  SR
    $m = [regex]::Match(
      $ln,
      '^(?<name>.+?)\s+(?<pct>\d{1,3})\s+(?<dmgAndMore>.+?)\s+(?<sr>\d+)\s*$'
    )
    if (-not $m.Success) { break }

    $nameRaw = $m.Groups['name'].Value.Trim()
    $pct     = [int]$m.Groups['pct'].Value
    $tail    = $m.Groups['dmgAndMore'].Value.Trim()
    $sr      = [int]$m.Groups['sr'].Value

    # stars on name?
    $nkFromName = ([regex]::Match($nameRaw, '\*+$')).Value
    $name = ($nameRaw -replace '\*+$','').Trim()

    # split damage-ish token and trailing text (“After”)
    $parts = ($tail -split '\s+', 2)
    $dmgTok = $parts[0]
    $after  = if ($parts.Count -gt 1) { $parts[1] } else { '' }

    # normalize minus; peel stars from damage/after
    $dmgTok = ($dmgTok -replace '[–−]','-').Trim()
    $nkFromDmg   = ([regex]::Match($dmgTok, '\*+$')).Value; if ($nkFromDmg) { $dmgTok = $dmgTok.Substring(0, $dmgTok.Length - $nkFromDmg.Length) }
    $after       = $after.Trim()
    $nkFromAfter = ([regex]::Match($after, '\*+$')).Value; if ($nkFromAfter) { $after = $after.Substring(0, $after.Length - $nkFromAfter.Length).Trim() }

    # note key priority: name > damage > after
    $noteKey = if ($nkFromName) { $nkFromName } elseif ($nkFromDmg) { $nkFromDmg } elseif ($nkFromAfter) { $nkFromAfter } else { '' }

    # clean damage if it looks like dice; otherwise none
    $dmgClean = ''
    $dm = [regex]::Match($dmgTok, '^[+-]?\d+\s*[dD]\s*\d+(?:\s*[+-]\s*\d+)?(?:/\d+)?$')
    if ($dm.Success) { $dmgClean = ($dmgTok -replace '\s+','').ToUpper() }

    $weaponRows += [pscustomobject]@{
      Name    = $name
      Base    = $pct
      Damage  = $dmgClean
      SR      = $sr
      NoteKey = $noteKey
      After   = $after
    }
  }
  if ($weaponRows.Count -eq 0) {
    Write-Warning "No weapon rows parsed under the header. Check the clipboard format."
    return
  } else {
    Write-Verbose ("Parsed {0} rows:" -f $weaponRows.Count)
    foreach ($w in $weaponRows) {
      Write-Verbose ("  {0}  Base={1}  Dmg='{2}'  Key='{3}'  SR={4}  After='{5}'" -f $w.Name,$w.Base,$w.Damage,$w.NoteKey,$w.SR,$w.After)
    }
  }

 # --- parse footnotes (*, **, ***) — tolerant of blank lines and multi-line notes ---
$notes = @{}
# advance $i to the first footnote line (starting with *), if we aren't on one already
while ($i -lt $lines.Count -and $lines[$i] -notmatch '^\s*\*+') { $i++ }

$curKey = $null
$buf    = New-Object System.Text.StringBuilder

for (; $i -lt $lines.Count; $i++) {
  $line = ($lines[$i] -replace '[\u00A0]',' ').Trim()

  if ($line -match '^\s*\*+\s*(.+)$') {
    # flush prior note
    if ($curKey) {
      $notes[$curKey] = ($buf.ToString().Trim())
      $buf.Clear() | Out-Null
    }
    # start new note
    $curKey = ($line -replace '^(\s*\*+).*$','$1').Trim()  # the exact key of * / ** / ***
    $first  = ($line -replace '^\s*\*+\s*','').Trim()
    [void]$buf.Append($first)
    continue
  }

  # tolerate blank lines inside a note
  if ($curKey) {
    if ($line) { [void]$buf.Append(' ' + $line) }
    continue
  }

  # if we encounter non-note text with no current note, we’re done
  if (-not $curKey) { break }
}

# flush the last note
if ($curKey) { $notes[$curKey] = ($buf.ToString().Trim()) }

Write-Verbose ("Footnotes detected: " + (($notes.Keys | Sort-Object | ForEach-Object { "'$_'" }) -join ', '))


  # --- load tables lazily ---
  $tables = @{ melee=$null; missile=$null; shields=$null }
  function _getTable($kind) {
    if ($null -eq $tables[$kind]) {
      $tables[$kind] = _loadTable $kind
      _ensureCols ([ref]$tables[$kind].Data)
    }
    $tables[$kind]
  }

  # --- write/update rows ---
  $changes = New-Object System.Collections.Generic.List[object]
  foreach ($w in $weaponRows) {
  $kind = _whichTable $w.Name
if ($kind -eq $DefaultTable) {
  if ($w.After -match '(?i)\bmeters?\b|\brange\b|\bdropped\b' -or
      $w.Name  -match '^(?i)(spit|stone|throw|dart|arrow|javelin)$') {
    $kind = 'missile'
  }
}
$t = _getTable $kind

    # generic row (Creature blank) exists?
    $generic = $t.Data | Where-Object { $_.Name -like "*$($w.Name)*" -and ([string]::IsNullOrWhiteSpace([string]$_.Creature)) } | Select-Object -First 1

  # NEW: write if generic missing, OR we saw a star, OR we have trailing "After" text, OR -Overwrite
$hasKey      = -not [string]::IsNullOrWhiteSpace($w.NoteKey)
$specialText = if ($hasKey -and $notes.ContainsKey($w.NoteKey)) { $notes[$w.NoteKey] } else { $null }

$needsOwnRow = (-not $generic) -or $hasKey -or (-not [string]::IsNullOrWhiteSpace($w.After)) -or $Overwrite
if (-not $needsOwnRow) { continue }


    # upsert creature-specific row
    $row = $t.Data | Where-Object { $_.Name -eq $w.Name -and [string]$_.Creature -eq $Creature } | Select-Object -First 1
    $isNew = $false
    if (-not $row) {
      $rowOrdered = [ordered]@{}
      foreach ($col in $t.Data[0].PSObject.Properties.Name) { $rowOrdered[$col] = $null }
      $row = [pscustomobject]$rowOrdered
      $t.Data += $row
      $isNew = $true
    }

    # basic fields
    $row.Name        = $w.Name
    $row.'Base %'    = [int]$w.Base
    $row.Damage      = if ($w.Damage) { $w.Damage } else { $null }   # allow blank damage rows
    $row.SR          = [int]$w.SR
    $row.Creature    = $Creature

# if After says "from XXX" (CHA/STR/POW/etc), suffix the stat to Damage for clarity
if ($row.Damage -and $w.After -match '(?i)\bfrom\s+([A-Z]{3})\b') {
  $attr = $matches[1].ToUpper()
  if ($row.Damage -notmatch "\b$attr\b") {
    $row.Damage = "$($row.Damage) $attr"
  }
}


    # combine footnote special text with free-text "After" bits
    $notesText = $null
    if ($specialText) { $notesText = $specialText }
    if ($w.After)     { $notesText = if ($notesText) { "$notesText  " + $w.After } else { $w.After } }

    if ($notesText) {
      $row.SpecialName = $w.Name
      $row.SpecialText = $specialText  # keep pure footnote here
      $row.Notes       = $notesText    # user-facing combined note
    } elseif ($Overwrite) {
      $row.SpecialName = $null
      $row.SpecialText = $null
      if ($row.Notes -eq 'special') { $row.Notes = $null }
    }

    # record change for summary
    $changes.Add([pscustomobject]@{
      Table      = $kind
      Name       = $w.Name
      HasSpecial = ($hasKey -or [bool]$specialText)
      Base       = $row.'Base %'
      Damage     = $row.Damage
      SR         = $row.SR
      New        = $isNew
    })
  }

  # --- save only modified tables + summarize ---
  if ($changes.Count -eq 0) {
    Write-Warning "Parsed $($weaponRows.Count) rows, but no changes were required (generic rows already exist and no specials, or identical rows present)."
    return
  }

  foreach ($k in 'melee','missile','shields') {
    if ($tables[$k] -and ($changes | Where-Object { $_.Table -eq $k })) {
      Write-Verbose "Saving $k table → $($tables[$k].Path)"
      try { _saveTable $tables[$k] } catch { Write-Error "Failed to save $k table: $_" }
    }
  }

  Write-Host "Imported weapons for '$Creature'. Changes:"
  $changes | Sort-Object Table, Name | Format-Table Table, Name, HasSpecial, New, Base, Damage, SR -AutoSize
}
