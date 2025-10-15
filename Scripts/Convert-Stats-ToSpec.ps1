param(
  [Parameter(Mandatory)]
  [string]$Path,                         # e.g. Y:\Stat_blocks\Stat_Dice_Source.xlsx
  [string]$WorksheetName = 'Sheet1',     # change if needed
  [switch]$OverwriteExisting,            # overwrite existing *Spec cells
  [switch]$RemoveLegacy,                 # drop legacy STAT / STATMod columns
  [switch]$WhatIf                        # dry run (no file changes)
)

# ---- safety: backup ----
if (-not (Test-Path $Path)) { throw "File not found: $Path" }
$bak = "$Path.bak"
Copy-Item -Path $Path -Destination $bak -Force
Write-Host "Backup created: $bak" -ForegroundColor Green

# ---- helpers ----
$stats = 'STR','CON','SIZ','DEX','INT','POW','CHA'

function New-SpecFromLegacy([int]$n,[int]$mod) {
  # Encode legacy Nd6 +/- Mod as a dice spec your parser accepts
  if ($n -lt 0) { $n = 0 }
  $spec = "${n}d6"
  if     ($mod -gt 0) { "$spec+$mod" }
  elseif ($mod -lt 0) { "$spec$mod" }   # $mod already includes '-'
  else                 { $spec }
}

function Get-CellInt($row, [string]$name) {
  if ($row.PSObject.Properties.Name -contains $name -and $row.$name -ne $null -and "$($row.$name)".Trim() -ne '') {
    try { return [int]$row.$name } catch { return 0 }
  }
  return 0
}

function HasText([object]$v) {
  return ($null -ne $v -and "$v".Trim() -ne '')
}

# ---- load rows ----
$rows = Import-Excel -Path $Path -WorksheetName $WorksheetName -ErrorAction Stop
if (-not $rows) { throw "Worksheet '$WorksheetName' is empty or not found in $Path" }

# ---- convert ----
$changes = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
  $creature = if ($row.PSObject.Properties.Name -contains 'Creature') { [string]$row.Creature } else { '<Unnamed>' }

  foreach ($st in $stats) {
    $specCol = "${st}Spec"

    $hasSpecCol = $row.PSObject.Properties.Name -contains $specCol
    $hasSpecVal = $hasSpecCol -and (HasText $row.$specCol)

    $n   = Get-CellInt $row $st
    $mod = Get-CellInt $row "${st}Mod"

    if ($hasSpecVal -and -not $OverwriteExisting) {
      # keep existing spec
      continue
    }

    # Only build/overwrite when we have any legacy signal (including zeros—encode 0d6+K)
    if ($hasSpecVal -and $OverwriteExisting) {
      $old = [string]$row.$specCol
      $new = New-SpecFromLegacy $n $mod
      $changes.Add([pscustomobject]@{ Creature=$creature; Stat=$st; From=$old; To=$new })
      if (-not $WhatIf) { $row.$specCol = $new }
    }
    elseif (-not $hasSpecVal) {
      # create the spec column if missing
      if (-not $hasSpecCol) {
        Add-Member -InputObject $row -NotePropertyName $specCol -NotePropertyValue '' -Force
      }
      $new = New-SpecFromLegacy $n $mod
      # Skip writing if both n=0 and mod=0 and no prior spec — nothing to infer
      if ($n -ne 0 -or $mod -ne 0) {
        $changes.Add([pscustomobject]@{ Creature=$creature; Stat=$st; From='(empty)'; To=$new })
        if (-not $WhatIf) { $row.$specCol = $new }
      }
    }
  }
}

# ---- optionally remove legacy columns ----
if ($RemoveLegacy) {
  foreach ($st in $stats) {
    $nCol = $st
    $mCol = "${st}Mod"
    if ($rows[0].PSObject.Properties.Name -contains $nCol)  {
      Write-Host "Removing column: $nCol" -ForegroundColor Yellow
      foreach ($r in $rows) { $r.PSObject.Properties.Remove($nCol) | Out-Null }
    }
    if ($rows[0].PSObject.Properties.Name -contains $mCol)  {
      Write-Host "Removing column: $mCol" -ForegroundColor Yellow
      foreach ($r in $rows) { $r.PSObject.Properties.Remove($mCol) | Out-Null }
    }
  }
}

# ---- save or preview ----
if ($WhatIf) {
  Write-Host "Dry run (-WhatIf): no file changes written." -ForegroundColor Cyan
} else {
  # Clear & rewrite the worksheet with modified rows
  $rows | Export-Excel -Path $Path -WorksheetName $WorksheetName -ClearSheet
  Write-Host "Converted specs written to $Path" -ForegroundColor Green
}

# ---- show summary ----
if ($changes.Count -gt 0) {
  $changes | Sort-Object Creature, Stat | Format-Table -AutoSize
} else {
  Write-Host "No rows needed conversion (either already had *Spec or no legacy values found)." -ForegroundColor DarkGray
}
