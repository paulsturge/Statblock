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
    [switch]$Overwrite,

    # NEW: debug aid for damage parsing
    [switch]$DebugDamage
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
    if ($PSScriptRoot)    { $candidates += $PSScriptRoot; $candidates += (Split-Path $PSScriptRoot -Parent) }
    if ($pwd)             { $candidates += $pwd.Path }

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
      if (Test-Path -LiteralPath $p) { return (Resolve-Path $p).Path }
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
    if ($norm -match 'weapon' -and $norm -match '%' -and $norm -match 'damage' -and $norm -match '(^| )sr( |$)') {
      Write-Verbose "Header detected on line $($j+1): $($lines[$j])"
      $hdrIdx = $j + 1
      break
    }
  }

  if ($null -eq $hdrIdx) { throw "Weapons header line (contains 'Weapon', '%', 'Damage', 'SR') not found." }
  $start = $hdrIdx

  # --- parse rows ---
  $weaponRows = @()
  $i = $start
  for ($i = $start; $i -lt $lines.Count; $i++) {
    $ln = $lines[$i].Trim()
    if (-not $ln) { break }
    if ($ln -match '^\*+') { break }

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

    # ----------------- ROBUST damage + After parsing -----------------
    # normalize fancy minus
    $tail = ($tail -replace '[–−]','-').Trim()

    # Dice pattern (very tolerant): optional +/- then NdN, optional +/-N (but not if it starts a new NdN), optional /N
    $diceRegex = '([+\-]?\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*(?!\d*\s*[dD]\s*\d+)\d+)?(?:\s*/\s*\d+)?)'

    # Try: first dice anywhere in TAIL
    $mTail = [regex]::Match($tail, $diceRegex)

    # If not found in tail, try in the full line BEFORE SR
    $preSr = ($ln -replace '\s+\d+\s*$','').Trim()
    $mLine = if (-not $mTail.Success) { [regex]::Match($preSr, $diceRegex) } else { $null }

    $dmgTok = ''
    $after  = $tail

    if ($mTail.Success) {
      $dmgTok = $mTail.Value.Trim()
      # carve out just the first occurrence from tail to form 'after'
      $idx = $tail.IndexOf($mTail.Value)
      $after = ($tail.Substring(0,$idx) + $tail.Substring($idx + $mTail.Length)).Trim()
    } elseif ($mLine -and $mLine.Success) {
      $dmgTok = $mLine.Value.Trim()
      # leave $after = $tail
    } else {
      # no dice found at all; Damage remains blank, After=tail as-is
    }

    if ($DebugDamage) {
      Write-Verbose ("[DMG DEBUG] line: " + $ln)
      Write-Verbose ("[DMG DEBUG] tail: " + $tail)
      Write-Verbose ("[DMG DEBUG] matched-in-tail='{0}', matched-in-line='{1}'" -f ($mTail.Value), ($mLine.Value))
      Write-Verbose ("[DMG DEBUG] dmgTok='{0}'  after(pre-clean)='{1}'" -f $dmgTok, $after)
    }

    # peel stars
    $nkFromDmg   = ([regex]::Match($dmgTok, '\*+$')).Value; if ($nkFromDmg) { $dmgTok = $dmgTok.Substring(0, $dmgTok.Length - $nkFromDmg.Length) }
    $after       = $after.Trim()
    $nkFromAfter = ([regex]::Match($after, '\*+$')).Value; if ($nkFromAfter) { $after = $after.Substring(0, $after.Length - $nkFromAfter.Length).Trim() }

    # strip leading DB dice from After (e.g., "+2D6", "+ 1D4/2", repeated)
    if ($after) {
      $after = ($after -replace '^(?:\+\s*\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?(?:/\s*\d+)?\s*)+','').Trim()
    }

    # choose note key priority: name > damage > after
    $noteKey = if ($nkFromName) { $nkFromName } elseif ($nkFromDmg) { $nkFromDmg } elseif ($nkFromAfter) { $nkFromAfter } else { '' }

    # finalize damage
    $dmgClean = ''
    if ($dmgTok) { $dmgClean = ($dmgTok -replace '\s+','').ToUpper() }

    if ($DebugDamage) {
      Write-Verbose ("[DMG DEBUG] dmgClean='{0}'  after(clean)='{1}'" -f $dmgClean, $after)
    }
    # ----------------------------------------------------------------

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

  # --- parse footnotes (*, **, ***) ---
  $notes = @{}
  while ($i -lt $lines.Count -and $lines[$i] -notmatch '^\s*\*+') { $i++ }
  $curKey = $null
  $buf    = New-Object System.Text.StringBuilder
  for (; $i -lt $lines.Count; $i++) {
    $line = ($lines[$i] -replace '[\u00A0]',' ').Trim()
    if ($line -match '^\s*\*+\s*(.+)$') {
      if ($curKey) { $notes[$curKey] = ($buf.ToString().Trim()); $buf.Clear() | Out-Null }
      $curKey = ($line -replace '^(\s*\*+).*$','$1').Trim()
      $first  = ($line -replace '^\s*\*+\s*','').Trim()
      [void]$buf.Append($first)
      continue
    }
    if ($curKey) { if ($line) { [void]$buf.Append(' ' + $line) }; continue }
    if (-not $curKey) { break }
  }
  if ($curKey) { $notes[$curKey] = ($buf.ToString().Trim()) }
  Write-Verbose ("Footnotes detected: " + (($notes.Keys | Sort-Object | ForEach-Object { "'$_'" }) -join ', '))

  # --- parse plain Notes: / Note: ---
  $globalNotes = New-Object System.Collections.Generic.List[string]
  for ($k = 0; $k -lt $lines.Count; $k++) {
    if ($lines[$k] -match '^\s*Notes?\s*[:：-]\s*(.*)$') {
      $cur = $matches[1].Trim()
      $k++
      while ($k -lt $lines.Count) {
        $n = ($lines[$k] -replace '[\u00A0]',' ')
        if ($n -match '^\s*$') { break }
        if ($n -match '^\s*(Skills?:|Armor|Powers?:|Spells?:|Hit\s*Points?|Characteristics|Description)\b') { break }
        $cur = ($cur + ' ' + $n.Trim()).Trim()
        $k++
      }
      if ($cur) { $globalNotes.Add($cur) }
      break
    }
  }
  Write-Verbose ("Notes: block detected: " + ($globalNotes.Count))
  if ($globalNotes.Count -gt 0) { Write-Verbose ("Notes[0]: " + $globalNotes[0]) }

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
          $w.Name  -match '^(?i)(spit|stone|throw|dart|arrow|javelin)$') { $kind = 'missile' }
    }
    $t = _getTable $kind

    $generic = $t.Data | Where-Object { $_.Name -like "*$($w.Name)*" -and ([string]::IsNullOrWhiteSpace([string]$_.Creature)) } | Select-Object -First 1

    $hasKey      = -not [string]::IsNullOrWhiteSpace($w.NoteKey)
    $specialText = if ($hasKey -and $notes.ContainsKey($w.NoteKey)) { $notes[$w.NoteKey] } else { $null }
    $needsOwnRow = (-not $generic) -or $hasKey -or (-not [string]::IsNullOrWhiteSpace($w.After)) -or $Overwrite
    if (-not $needsOwnRow) { continue }

    # upsert row
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
    $row.Damage      = if ($w.Damage) { $w.Damage } else { $null }
    $row.SR          = [int]$w.SR
    $row.Creature    = $Creature

    # >>> Attribute suffixing on damage + clean the attribute token out of the incoming Notes
    if ($row.Damage) {
      $attrMatch = $null
      if ($w.After -match '(?i)\bfrom\s+([A-Z]{3})\b') {
        $attrMatch = $matches[1].ToUpper()
        # do not remove here; we'll sanitize incoming below so missile inference still works
      } elseif ($w.After -match '^(?i)\s*([A-Z]{3})\b') {
        $attrMatch = $matches[1].ToUpper()
      }
      if ($attrMatch -and $row.Damage -notmatch "\b$attrMatch\b") {
        $row.Damage = "$($row.Damage) $attrMatch"
      }
    }

    # >>> INLINE AFTER → Notes (strip any DB dice, remove attr fragments & junk like 'meters dropped' / lone 'range', collapse spaces)
    if ($row.Notes) {
      $row.Notes = ($row.Notes -replace '(?:^|\s)\+\s*\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?(?:/\s*\d+)?','').Trim()
      $row.Notes = ($row.Notes -replace '\s{2,}',' ').Trim()
    }
    if ($w.After) {
      $incoming = ($w.After -replace '^(?:\+\s*\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?(?:/\s*\d+)?\s*)+','').Trim()
      # remove attribute fragments (after we've already suffixed to Damage)
      $incoming = ($incoming -replace '(?i)\bfrom\s+[A-Z]{3}\b','').Trim()
      # remove junky classifier fragments
      $incoming = ($incoming `
        -replace '^(?i)\s*range\s*$','' `
        -replace '(?i)\bmeters?\s*dropped\b','' `
      ).Trim()
      # collapse whitespace
      $incoming = ($incoming -replace '\s{2,}',' ').Trim()

      if ($isNew -or $Overwrite -or [string]::IsNullOrWhiteSpace($row.Notes)) {
        $row.Notes = if ($incoming) { $incoming } else { $null }
      } elseif ($incoming -and $row.Notes -notmatch [regex]::Escape($incoming)) {
        $row.Notes = ($row.Notes + '  ' + $incoming).Trim()
      }
    }

    # FOOTNOTE SPECIAL → SpecialText
    if ($specialText) {
      $row.SpecialName = $w.Name
      if ([string]::IsNullOrWhiteSpace($row.SpecialText)) {
        $row.SpecialText = $specialText
      } elseif ($row.SpecialText -notmatch [regex]::Escape($specialText)) {
        $row.SpecialText = ($row.SpecialText.Trim() + '  ' + $specialText).Trim()
      }
    } elseif ($Overwrite) {
      $row.SpecialName = $null
      $row.SpecialText = $null
    }

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

  # apply global Notes → SpecialText on one row (prefer Bite)
  if ($globalNotes -and $globalNotes.Count -gt 0) {
    $targetRow  = $null
    $targetKind = $null
    function _firstCreatureRowOfKind($kind) {
      if (-not $tables[$kind]) { return $null }
      $tables[$kind].Data | Where-Object { [string]$_.Creature -eq $Creature } | Select-Object -First 1
    }
    foreach ($kind in @('melee','missile','shields')) {
      if (-not $tables[$kind]) { continue }
      $hit = $tables[$kind].Data | Where-Object { $_.Name -eq 'Bite' -and [string]$_.Creature -eq $Creature } | Select-Object -First 1
      if ($hit) { $targetRow = $hit; $targetKind = $kind; break }
    }
    if (-not $targetRow) {
      foreach ($kind in @('melee','missile','shields')) {
        $hit = _firstCreatureRowOfKind $kind
        if ($hit) { $targetRow = $hit; $targetKind = $kind; break }
      }
    }
    foreach ($noteLine in $globalNotes) {
      if ($targetRow) {
        if ([string]::IsNullOrWhiteSpace($targetRow.SpecialText)) {
          $targetRow.SpecialText = $noteLine
        } elseif ($targetRow.SpecialText -notmatch [regex]::Escape($noteLine)) {
          $targetRow.SpecialText = ($targetRow.SpecialText.Trim() + '  ' + $noteLine).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($targetRow.SpecialName)) {
          $targetRow.SpecialName = $targetRow.Name
        }
        $changes.Add([pscustomobject]@{
          Table=$targetKind; Name=$targetRow.Name; HasSpecial=$false; Base=$targetRow.'Base %'; Damage=$targetRow.Damage; SR=$targetRow.SR; New=$false
        })
        Write-Verbose ("Notes (plain) → SpecialText on {0} ({1})" -f $targetRow.Name, $targetKind)
      } else {
        $t = _getTable 'melee'
        $rowOrdered = [ordered]@{} ; foreach ($col in $t.Data[0].PSObject.Properties.Name) { $rowOrdered[$col] = $null }
        $row = [pscustomobject]$rowOrdered
        $row.Name        = 'Special: Note'
        $row.Creature    = $Creature
        $row.Notes       = $null
        $row.SpecialName = 'Note'
        $row.SpecialText = $noteLine
        $t.Data += $row
        $changes.Add([pscustomobject]@{
          Table='melee'; Name=$row.Name; HasSpecial=$false; Base=$row.'Base %'; Damage=$row.Damage; SR=$row.SR; New=$true
        })
        Write-Verbose ("Notes (plain) created synthetic row in melee")
      }
    }
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
