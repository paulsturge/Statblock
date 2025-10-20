function New-StatDiceRow {
  [CmdletBinding(DefaultParameterSetName='Clipboard')]
  param(
    [Parameter(Mandatory)][string]$Creature,
    [Parameter(ParameterSetName='Clipboard')][switch]$FromClipboard,
    [Parameter(ParameterSetName='Text')][string]$Text,
    [string]$HitLocation = $Creature,
    [switch]$Overwrite
  )
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
  # --- get source text ---
  if ($PSCmdlet.ParameterSetName -eq 'Clipboard') {
    try { $Text = Get-Clipboard -Raw } catch { throw "Clipboard read failed: $_" }
  }
  if ([string]::IsNullOrWhiteSpace($Text)) { throw "No input text provided." }

  # --- parse target structure ---
  $statKeys = 'STR','CON','SIZ','DEX','INT','POW','CHA'
  $out = [ordered]@{
    Creature     = $Creature
    STRSpec=$null; CONSpec=$null; SIZSpec=$null; DEXSpec=$null; INTSpec=$null; POWSpec=$null; CHASpec=$null
    Move         = $null
    Skills       = $null
    Hit_location = $HitLocation
  }

  # normalize whitespace/dashes
  $lines = (($Text -replace '\r','') -replace '[\u00A0]',' ') -split '\n'
  foreach ($ln in $lines) {
    $line = ($ln -replace '\s+',' ').Trim()
    if (-not $line) { continue }

    # STAT lines like "STR 2D6+6 13" (ignore trailing average)
    foreach ($k in $statKeys) {
      if ($line -match ("^(?i)$k\s+([0-9]+\s*[dD]\s*[0-9]+(?:\s*[+\-]\s*[0-9]+)?(?:\s*[x\*]\s*[0-9]+)?)\b")) {
        $spec = $matches[1] -replace '\s+',''
        $out["${k}Spec"] = $spec
        break
      }
    }

    # Move: keep verbatim
    if ($line -match '^(?i)Hit Points:\s*\d+.*?\bMove:\s*(.+)$') {
      $out.Move = $matches[1].Trim()
      continue
    }
    if (-not $out.Move -and $line -match '^(?i)Move\s*:\s*(.+)$') {
      $out.Move = $matches[1].Trim()
      continue
    }

    # Skills: keep verbatim after label
    if ($line -match '^(?i)Skills?\s*:\s*(.+)$') {
      $out.Skills = $matches[1].Trim()
      continue
    }

    # Ignore: Magic Points, Base SR, Armor, etc. (derived elsewhere)
  }

  if (-not ($statKeys | Where-Object { $out["${_}Spec"] })) {
    throw "No stat specs (e.g., 'STR 2D6+6') found in the input."
  }

  # --- write to Stat_Dice_Source.xlsx ---
  if (-not (Get-Command Resolve-StatPath -ErrorAction SilentlyContinue)) {
    throw "Resolve-StatPath not found (import your module first)."
  }
  $path = Resolve-StatPath 'Stat_Dice_Source.xlsx'
  if (-not $path) { throw "Stat_Dice_Source.xlsx not found." }

  # Use the first worksheet (or change to a specific one if you prefer)
  $sheetName = (Get-ExcelSheetInfo -Path $path | Select-Object -First 1 -ExpandProperty Name)
  $data = Import-Excel -Path $path -WorksheetName $sheetName

  # Ensure needed columns exist
  foreach ($col in @('Creature','STRSpec','CONSpec','SIZSpec','DEXSpec','INTSpec','POWSpec','CHASpec','Move','Skills','Hit_location')) {
    if (-not ($data | Get-Member -Name $col -MemberType NoteProperty)) {
      $data | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName $col -NotePropertyValue $null -Force }
    }
  }

  $existing = $data | Where-Object { $_.Creature -eq $Creature } | Select-Object -First 1
  if ($existing) {
    if (-not $Overwrite) { throw "Creature '$Creature' already exists. Use -Overwrite to update." }
    foreach ($k in $out.Keys) { $existing.$k = $out[$k] }
  } else {
    # build row preserving sheet columns
    $row = [ordered]@{}
    foreach ($col in $data[0].PSObject.Properties.Name) { $row[$col] = $null }
    foreach ($k in $out.Keys) { $row[$k] = $out[$k] }
    $data += [pscustomobject]$row
  }

  Export-Excel -Path $path -WorksheetName $sheetName -ClearSheet -AutoSize -FreezeTopRow -BoldTopRow -InputObject $data

  Write-Host "Saved '$Creature' to $(Split-Path -Leaf $path) [$sheetName]."
  Write-Host ("  Specs: " + ($statKeys | ForEach-Object { if ($out["${_}Spec"]) { "$_=$($out["${_}Spec"])" } } | Where-Object {$_}) -join ', ')
  if ($out.Move)   { Write-Host "  Move:   $($out.Move)" }
  if ($out.Skills) { Write-Host "  Skills: $($out.Skills)" }
}
