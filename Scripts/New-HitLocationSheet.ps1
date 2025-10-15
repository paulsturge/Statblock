function New-HitLocationSheet {
  <#
    .SYNOPSIS
      Create a hit-location worksheet in Hit_Location_Source.xlsx.

    .DESCRIPTION
      Sources:
        - Clone an existing sheet: -FromSheet 'Humanoid'
        - Import from file (CSV/TSV): -FromCsvPath '...\file.csv'
        - Paste from clipboard:
            - With headers (Location D20 [Armor] [HP] or Armor/HP): -FromClipboard
            - Without headers (e.g. "Right Leg 02–04 5"):          -FromClipboard -NoHeader

      Required columns: Location, D20
      Optional columns: Armor (defaults 0), HP (defaults 0; your runtime adjusts these).

    .PARAMETER Name
      New worksheet name to create/replace.

    .PARAMETER FromSheet
      Existing sheet name to clone.

    .PARAMETER FromCsvPath
      CSV/TSV path (auto-detects delimiter by extension; .tsv => tab).

    .PARAMETER FromClipboard
      Read table from clipboard. Supports TSV/CSV; falls back to whitespace-delimited.

    .PARAMETER NoHeader
      Only for -FromClipboard. Treat clipboard as headerless lines like: "Tail 01 5".
  #>
  [CmdletBinding(DefaultParameterSetName='Clone')]
  param(
    [Parameter(Mandatory)]
    [string]$Name,

    [Parameter(ParameterSetName='Clone', Mandatory)]
    [string]$FromSheet,

    [Parameter(ParameterSetName='FromFile', Mandatory)]
    [string]$FromCsvPath,

    [Parameter(ParameterSetName='FromClipboard', Mandatory)]
    [switch]$FromClipboard,

    [Parameter(ParameterSetName='FromClipboard')]
    [switch]$NoHeader
  )

# --- Find workbook robustly ---
$hitPath = $null

# 1) Prefer the user's resolver if it exists, but guard errors/empty
$cmd = Get-Command -Name Resolve-StatPath -ErrorAction SilentlyContinue
if ($cmd) {
  try {
    $hitPath = Resolve-StatPath 'Hit_Location_Source.xlsx'
  } catch { $hitPath = $null }
}

# 2) Fallback candidates
$candidates = @(
  (Join-Path $PSScriptRoot 'Hit_Location_Source.xlsx'),
  (Join-Path (Split-Path $PSScriptRoot -Parent) 'Hit_Location_Source.xlsx'),
  (Join-Path (Get-Location).Path 'Hit_Location_Source.xlsx')
)

# (Optional) add a quick recursive sniff one level up if still not found
if (-not $hitPath) {
  $probe = Get-ChildItem -Path (Split-Path $PSScriptRoot -Parent) -Filter 'Hit_Location_Source.xlsx' -Recurse -ErrorAction SilentlyContinue |
           Select-Object -First 1 -ExpandProperty FullName
  if ($probe) { $candidates += $probe }
}

# 3) Pick the first that exists
foreach ($p in $candidates) {
  if ($p -and (Test-Path $p)) { $hitPath = (Resolve-Path $p).Path; break }
}

if (-not $hitPath) {
  throw "Hit_Location_Source.xlsx not found. Tried:`n - " + ($candidates -join "`n - ")
}


  # --- helpers ---

  function Normalize-Columns($rows) {
    foreach ($r in $rows) {
      # map incoming headers (case/space tolerant)
      $map = @{}
      foreach ($p in $r.PSObject.Properties) {
        $k = ($p.Name -replace '\s','').ToLower()
        switch ($k) {
          'location'  { $map['Location'] = $p.Name }
          'd20'       { $map['D20']      = $p.Name }
          'armor'     { $map['Armor']    = $p.Name }  # optional
          'hp'        { $map['HP']       = $p.Name }  # optional
          'armor/hp'  { $map['HP']       = $p.Name }  # treat combined header as HP; Armor=0
        }
      }

      foreach ($need in 'Location','D20') {
        if (-not $map.ContainsKey($need)) {
          throw "Input is missing required column: $need (got: $($r.PSObject.Properties.Name -join ', '))"
        }
      }

      $armorVal = if ($map.ContainsKey('Armor')) { $r.$($map['Armor']) } else { 0 }
      $hpVal    = if ($map.ContainsKey('HP'))    { $r.$($map['HP'])    } else { 0 }

      # build canonical object (normalize en dash)
      [pscustomobject]@{
        Location = [string]$r.$($map['Location'])
        D20      = ([string]$r.$($map['D20'])) -replace '–','-'
        Armor    = $armorVal
        HP       = $hpVal
      }
    }
  }

function Cast-NumericColumns($rows) {
  foreach ($x in $rows) {
    $armor = $x.Armor
    $hp    = $x.HP

    # normalize any odd slashes (just in case)
    $armor = if ($armor -is [string]) { $armor -replace '⁄','/' } else { $armor }
    $hp    = if ($hp    -is [string]) { $hp    -replace '⁄','/' } else { $hp }

    # 1) Handle combined "Armor/HP" like "6/6" (in either column)
    $rx = '^\s*(\d+)\s*/\s*(\d+)\s*$'
    if ($armor -is [string] -and ($armor -match $rx)) {
      $armor = [int]$matches[1]
      $hp    = [int]$matches[2]
    }
    elseif ($hp -is [string] -and ($hp -match $rx)) {
      $armor = [int]$matches[1]
      $hp    = [int]$matches[2]
    }

    # 2) Numeric-only coercion (leave text like '—' as-is)
    $tmp = 0.0
    if ($null -ne $armor -and [double]::TryParse("$armor".Trim(), [ref]$tmp)) { $armor = [int]$tmp }
    if ($null -ne $hp    -and [double]::TryParse("$hp".Trim(),    [ref]$tmp)) { $hp    = [int]$tmp }

    [pscustomobject]@{
      Location = [string]$x.Location
      D20      = [string]$x.D20
      Armor    = $armor
      HP       = $hp
    }
  }
}



  # --- load source rows ---

  $rows = $null

  switch ($PSCmdlet.ParameterSetName) {
    'Clone' {
      $rows = Import-Excel -Path $hitPath -WorksheetName $FromSheet |
              Where-Object { $_.PSObject.Properties.Value -ne $null }
    }
    'FromFile' {
      if (-not (Test-Path $FromCsvPath)) { throw "File not found: $FromCsvPath" }
      $delim = if ($FromCsvPath -match '\.tsv$') { "`t" } else { ',' }
      $rows  = Import-Csv -Path $FromCsvPath -Delimiter $delim
    }
    'FromClipboard' {
      $raw = Get-Clipboard
      if (-not $raw) { throw "Clipboard is empty. Copy your table first." }

      if ($NoHeader) {
        # headerless lines: "Location  D20  HP"
        $lines = $raw -split "(`r`n|`n|`r)" | Where-Object { $_.Trim().Length -gt 0 }
        $rows = foreach ($line in $lines) {
          $line = $line.Trim()
          $toks = $line -split '\s+'
          if ($toks.Count -lt 2) { continue }

          $hpTok  = $toks[-1]
          $d20Tok = if ($toks.Count -ge 3) { $toks[-2] } else { '' }
          $locTks = if ($toks.Count -ge 3) { $toks[0..($toks.Count-3)] } else { @($toks[0]) }
          $loc    = ($locTks -join ' ').Trim()

          [pscustomobject]@{
            Location = $loc
            D20      = ($d20Tok -replace '–','-')
            Armor    = 0
            HP       = $hpTok
          }
        }
      }
      else {
        # Try TSV then CSV with headers
        try { $rows = ConvertFrom-Csv -Delimiter "`t" -InputObject $raw } catch { $rows = $null }
        if (-not $rows -or ($rows.Count -gt 0 -and $rows[0].PSObject.Properties.Name.Count -eq 1)) {
          try { $rows = ConvertFrom-Csv -Delimiter ',' -InputObject $raw } catch { $rows = $null }
        }
       # Fallback: whitespace-delimited with headers like "Location D20 Armor/HP"
if (-not $rows -or ($rows.Count -gt 0 -and $rows[0].PSObject.Properties.Name.Count -eq 1)) {
  $lines = $raw -split "(`r`n|`n|`r)" | Where-Object { $_.Trim().Length -gt 0 }
  if ($lines.Count -lt 2) { throw "Need a header line + at least one data line." }

  $hdr = ($lines[0] -split '\s+').ForEach({ $_.Trim() })
  $data = $lines[1..($lines.Count-1)]

  # normalized header keys for mapping
  $norm = $hdr.ForEach({ ($_ -replace '\s','').ToLower() })

  $rows = foreach ($line in $data) {
    $toks = ($line -split '\s+').ForEach({ $_.Trim() })
    if ($toks.Count -lt 2) { continue }

    # Right-align tokens to non-Location headers; Location gets "the rest"
    $values = @{}
    $tokIdx = $toks.Count - 1
    for ($h = $hdr.Count-1; $h -ge 0; $h--) {
      $hName = $hdr[$h]
      $hKey  = $norm[$h]
      if ($hKey -eq 'location') {
        # Location = everything left
        $locEnd = [math]::Max($tokIdx, 0)
        $loc = ($toks[0..$locEnd] -join ' ')
        $values[$hName] = $loc
      }
      else {
        $values[$hName] = ($tokIdx -ge 0) ? $toks[$tokIdx] : ''
        $tokIdx--
      }
    }

    # normalize en dash in D20 if present
    if ($values.ContainsKey('D20')) { $values['D20'] = ($values['D20'] -replace '–','-') }

    [pscustomobject]$values
  }
}

      }
    }
  }

  if (-not $rows) { throw "No rows loaded from the source." }

  # Normalize & cast
  $clean = Normalize-Columns -rows $rows
  $clean = Cast-NumericColumns -rows $clean

  # Write (create or replace the sheet)
  $clean | Export-Excel -Path $hitPath -WorksheetName $Name -ClearSheet
  Write-Host "Hit-location sheet '$Name' created in $(Split-Path $hitPath -Leaf)." -ForegroundColor Green
}
