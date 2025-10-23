function New-StatDiceRow {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Creature,
    [string]$HitLocation = $Creature,
    [switch]$FromClipboard
  )

  # --- read text ---
  $text = $null
  if ($FromClipboard) {
    try { $text = Get-Clipboard -Raw } catch { throw "Clipboard read failed: $_" }
  }
  if ([string]::IsNullOrWhiteSpace($text)) {
    $text = Read-Host "Paste the STAT BLOCK text here"
  }
  if ([string]::IsNullOrWhiteSpace($text)) { throw "No stat text provided." }

  # --- helpers ---
  function _norm($s) {
    # normalize unicode ×/X → x, fancy dashes → '-', collapse spaces
    (($s -replace '[×X]', 'x') -replace '[\u2012-\u2015\u2212]', '-' -replace '\s+', ' ').Trim()
  }

  # dice spec: NdF [±K] [x MULT]
  $specRx = '(?<spec>\d+\s*[dD]\s*\d+(?:\s*[+\-]\s*\d+)?\s*(?:[xX×\*]\s*\d+)?)'
  $lineRx = '^(?<stat>STR|CON|SIZ|DEX|INT|POW|CHA)\b.*?' + $specRx

  $rawLines = ($text -replace '\r','') -split '\n'
  $lines    = $rawLines | ForEach-Object { _norm $_ }

  # ---------- 1) Stats (Specs) ----------
  $specs = @{}
  foreach ($ln in $lines) {
    if ($ln -match $lineRx) {
      $stat = $matches['stat'].ToUpper()
      $spec = _norm $matches['spec']
      $spec = ($spec -replace '[×X]', 'x')
      $specs[$stat] = $spec
    }
  }
  if ($specs.Count -eq 0) {
    throw "No STAT specs found (expected lines like 'STR 3D6×2', 'SIZ 2D6+6', etc.)."
  }

  # ---------- 2) Parse ancillary sections ----------
  $moveRaw      = $null
  $magicPoints  = $null
  $baseSR       = $null
  $armorText    = $null
  $skillsText   = $null
  $languagesTxt = $null
  $passionsTxt  = $null
  $runesTxt     = $null   # <- capture "Runes:" here
  $magicTxt     = $null   # <- goes to 'Magic'

  # Known headers (stop list) – allow with/without colon
  $hdrRegex = '^(?i)(Hit\s*Points|Move|Magic\s*Points|Base\s*SR|Armor|Skills|Languages|Passions|Runes|Magic)\b'

  function Collect-Paragraph {
    param([int]$startIndex,[string[]]$raw)
    $headLineNorm = _norm $raw[$startIndex]

    # If there's a colon, take the text after it; otherwise remove the header token itself (e.g., "Move 8" -> "8")
    if ($headLineNorm -match '^[^:]*:\s*') {
      $firstPart = ($headLineNorm -replace '^[^:]*:\s*','')
    } else {
      $firstPart = ($headLineNorm -replace '^(?i)(Hit\s*Points|Move|Magic\s*Points|Base\s*SR|Armor|Skills|Languages|Passions|Runes|Magic)\b\s*','')
    }

    $sb = New-Object System.Text.StringBuilder
    if ($firstPart) { [void]$sb.Append($firstPart) }

    $i = $startIndex + 1
    while ($i -lt $raw.Count) {
      $lineNorm = _norm $raw[$i]
      if ($lineNorm -match '^\s*$') { break }
      if ($lineNorm -match $hdrRegex) { break }
      if ($sb.Length -gt 0) { [void]$sb.Append(' ') }
      [void]$sb.Append($lineNorm)
      $i++
    }
    return @{ Text = ($sb.ToString().Trim()); Next = $i }
  }

  $i = 0
  while ($i -lt $rawLines.Count) {
    $norm = _norm $rawLines[$i]
    if ($norm -match $hdrRegex) {
      if ($norm -match '^(?i)Hit\s*Points\b') {
        $r = Collect-Paragraph $i $rawLines

        # Inline Move on same line or within the collected segment
        if ($norm -match '(?i)\bMove\s*:?\s*(?<mv>.+)$') {
          $moveRaw = $matches['mv'].Trim()
        } elseif ($r.Text -match '(?i)\bMove\s*:?\s*(?<mv>.+)$') {
          $moveRaw = $matches['mv'].Trim()
        }

        $i = $r.Next; continue
      }
      elseif ($norm -match '^(?i)Move\b') {
        $r = Collect-Paragraph $i $rawLines
        $moveRaw = $r.Text
        $i = $r.Next; continue
      }
      elseif ($norm -match '^(?i)Magic\s*Points\b') {
        $r = Collect-Paragraph $i $rawLines
        if ($r.Text -match '(\d+)') { $magicPoints = [int]$matches[1] }
        $i = $r.Next; continue
      }
      elseif ($norm -match '^(?i)Base\s*SR\b') {
        $r = Collect-Paragraph $i $rawLines
        if ($r.Text -match '(\d+)') { $baseSR = [int]$matches[1] }
        $i = $r.Next; continue
      }
      elseif ($norm -match '^(?i)Armor\b')     { $r = Collect-Paragraph $i $rawLines; $armorText    = $r.Text; $i = $r.Next; continue }
      elseif ($norm -match '^(?i)Skills\b')    { $r = Collect-Paragraph $i $rawLines; $skillsText   = $r.Text; $i = $r.Next; continue }
      elseif ($norm -match '^(?i)Languages\b') { $r = Collect-Paragraph $i $rawLines; $languagesTxt = $r.Text; $i = $r.Next; continue }
      elseif ($norm -match '^(?i)Passions\b')  { $r = Collect-Paragraph $i $rawLines; $passionsTxt  = $r.Text; $i = $r.Next; continue }
      elseif ($norm -match '^(?i)Runes\b')     { $r = Collect-Paragraph $i $rawLines; $runesTxt     = $r.Text; $i = $r.Next; continue }
      elseif ($norm -match '^(?i)Magic\b')     { $r = Collect-Paragraph $i $rawLines; $magicTxt     = $r.Text; $i = $r.Next; continue }
    }
    $i++
  }

  # Failsafe: if Move still not found, scan whole block (ignore "Move Quietly" in Skills)
  if (-not $moveRaw) {
    $block = ($rawLines -join "`n")
    $m = [regex]::Match($block, '(?im)\bMove(?!\s*Quietly)\s*:?\s*(?<mv>[^\r\n]+)')
    if ($m.Success) { $moveRaw = $m.Groups['mv'].Value.Trim() }
  }

  # ---------- 3) Open sheet ----------
  $path = Resolve-StatPath 'Stat_Dice_Source.xlsx'
  if (-not $path) { throw "Stat_Dice_Source.xlsx not found." }
  $sheet = (Get-ExcelSheetInfo -Path $path | Select-Object -First 1 -ExpandProperty Name)
  $data  = Import-Excel -Path $path -WorksheetName $sheet | Where-Object { $_.PSObject.Properties.Value -ne $null }

  # ensure columns exist
  $needCols = @(
    'Creature','Hit_location',
    'STRSpec','CONSpec','SIZSpec','DEXSpec','INTSpec','POWSpec','CHASpec',
    'Runes1','Rune1score','Runes2','Rune2score','Runes3','Rune3score',
    'Move','MagicPoints','BaseSR','Armor','Skills','Languages','Passions','Magic'
  )
  foreach ($c in $needCols) {
    if (-not ($data | Get-Member -Name $c -MemberType NoteProperty)) {
      $data | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName $c -NotePropertyValue $null -Force }
    }
  }

  # upsert row
  $row = $data | Where-Object { [string]$_.Creature -eq $Creature } | Select-Object -First 1
  if (-not $row) {
    $row = [pscustomobject]@{
      Creature=$Creature; Hit_location=$null
      STRSpec=$null; CONSpec=$null; SIZSpec=$null; DEXSpec=$null; INTSpec=$null; POWSpec=$null; CHASpec=$null
      Runes1=$null; Rune1score=$null; Runes2=$null; Rune2score=$null; Runes3=$null; Rune3score=$null
      Move=$null; MagicPoints=$null; BaseSR=$null; Armor=$null; Skills=$null; Languages=$null; Passions=$null; Magic=$null
    }
    $data += $row
  }

  # write specs
  foreach ($k in 'STR','CON','SIZ','DEX','INT','POW','CHA') {
    if ($specs.ContainsKey($k)) { $row."${k}Spec" = $specs[$k] }
  }

  # hit location override (from New-Creature’s -HumanoidHL switch mapping)
  if ($PSBoundParameters.ContainsKey('HitLocation') -and $HitLocation) {
    $row.Hit_location = $HitLocation
  }

  # ancillary fields (Move kept verbatim; Statblock.ps1 will parse/format)
  if ($moveRaw)      { $row.Move         = $moveRaw }
  if ($magicPoints)  { $row.MagicPoints  = [int]$magicPoints }
  if ($baseSR)       { $row.BaseSR       = [int]$baseSR }
  if ($armorText)    { $row.Armor        = $armorText }
  if ($skillsText)   { $row.Skills       = $skillsText }
  if ($languagesTxt) { $row.Languages    = $languagesTxt }
  if ($passionsTxt)  { $row.Passions     = $passionsTxt }
  if ($magicTxt)     { $row.Magic        = $magicTxt }

  # ---------- 4) Parse & write runes into Runes1/2/3 (+scores) ----------
  if ($runesTxt) {
    # Split on commas, normalize odd chars, strip trailing punctuation
    $parts = $runesTxt -split ',' |
      ForEach-Object {
        $p = $_ -replace '[\u00A0]', ' '              # NBSP -> space
        $p = $p -replace '[%﹪]', '%'                  # any percent -> %
        $p = $p -replace '[\u2012-\u2015\u2212]', '-'  # fancy dashes -> '-'
        $p = $p.Trim()
        $p
      } |
      Where-Object { $_ }

    $parsed = @()
    foreach ($p in $parts) {
      # remove trailing punctuation/spaces like ".", ";", ":" after the % or at end
      $pClean = $p.Trim().TrimEnd(' ','.',';','、',':')

      # Pass A: strict "... <name> <score>%"
      $m = [regex]::Match($pClean, '^(?<name>[^\d%]+?)\s*(?<score>\d{1,3})\s*%\s*$', 'IgnoreCase')

      if (-not $m.Success) {
        # Pass B: tolerant "... <name> <score> [% optional]"
        $m = [regex]::Match($pClean, '^(?<name>.+?)\s*(?<score>\d{1,3})(?:\s*%)?\s*$', 'IgnoreCase')
      }

      if ($m.Success) {
        $name  = ($m.Groups['name'].Value -replace '\s+$','') -replace '[:\-\s]+$',''
        $score = [int]$m.Groups['score'].Value
        $parsed += [pscustomobject]@{ Name = $name.Trim(); Score = $score }
      } else {
        # No numeric score; keep the name only
        $parsed += [pscustomobject]@{ Name = $pClean; Score = $null }
      }
    }

    # assign up to three runes + scores
    if ($parsed.Count -ge 1) { $row.Runes1     = $parsed[0].Name; $row.Rune1score = $parsed[0].Score }
    if ($parsed.Count -ge 2) { $row.Runes2     = $parsed[1].Name; $row.Rune2score = $parsed[1].Score }
    if ($parsed.Count -ge 3) { $row.Runes3     = $parsed[2].Name; $row.Rune3score = $parsed[2].Score }
  }

  # save
  Export-Excel -Path $path -WorksheetName $sheet -ClearSheet -AutoSize -FreezeTopRow -BoldTopRow -InputObject $data

  $specMsg = ('STR','CON','SIZ','DEX','INT','POW','CHA' | ForEach-Object {
    $v = if ($specs[$_]) { $specs[$_] } else { '-' }; "$_=$v"
  }) -join ' '
  Write-Host ("Saved '{0}' to Stat_Dice_Source.xlsx [{1}].`n  Specs: {2}" -f $Creature,$sheet,$specMsg)
}
