
function New-Creature {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Creature,

    # If provided, we’ll prompt you to copy relevant text to clipboard before each step
    [switch]$FromClipboard,

    # Optional override: the worksheet name to create/use for hit locations
    [string]$HitLocationName = $Creature,

    # Choose which parts to run. If none are specified, all three run.
    [switch]$Stats,
    [switch]$HitLocations,
    [switch]$Weapons,

    # When true, skip prompts and just run (useful if you already primed the clipboard)
    [switch]$NoPause,

    # NEW: force Humanoid hit locations (no clipboard prompt for HL)
    [switch]$HumanoidHL
  )

  # If no specific parts were requested, do all of them
  if (-not ($Stats -or $HitLocations -or $Weapons)) {
    $Stats = $true; $HitLocations = $true; $Weapons = $true
  }

  # Effective HL name (Humanoid override wins)
  $effectiveHL = if ($HumanoidHL) { 'Humanoid' } else { $HitLocationName }

  function _PressToContinue([string]$what) {
    if ($NoPause) { return }
    Write-Host ""
    Write-Host "👉  Copy the $what text to the clipboard, then press Enter to continue..." -ForegroundColor Cyan
    [void](Read-Host)
  }

  Write-Host ""
  Write-Host "=== New Creature: $Creature ===" -ForegroundColor Green

  # 1) Stat dice row
  if ($Stats) {
    try {
      if ($FromClipboard) {
        Write-Host "`n[1/3] Stat Dice Row" -ForegroundColor Yellow
        _PressToContinue "STAT BLOCK (the lines like 'STR 2D6+6', 'CON 3D6', etc.)"
        New-StatDiceRow -Creature $Creature -HitLocation $effectiveHL -FromClipboard -ErrorAction Stop
      } else {
        New-StatDiceRow -Creature $Creature -HitLocation $effectiveHL -ErrorAction Stop
      }
      Write-Host "✓ Stat dice created/updated for '$Creature' (Hit locations = '$effectiveHL')." -ForegroundColor Green
    } catch {
      Write-Warning "Stat dice step failed: $($_.Exception.Message)"
    }
  }

  # 2) Hit locations
  if ($HitLocations) {
    try {
      if ($HumanoidHL) {
        # Explicitly skip HL import when forcing Humanoid
        Write-Host "`n[2/3] Hit Locations" -ForegroundColor Yellow
        Write-Host "Skipping hit-location import; using worksheet '$effectiveHL'." -ForegroundColor DarkYellow
      } else {
        if ($FromClipboard) {
          Write-Host "`n[2/3] Hit Locations" -ForegroundColor Yellow
          _PressToContinue "HIT LOCATION TABLE (header row like 'Location D20 Armor/HP' or 'Location D20 HP')"
          New-HitLocationSheet -Name $effectiveHL -FromClipboard -ErrorAction Stop
        } else {
          New-HitLocationSheet -Name $effectiveHL -ErrorAction Stop
        }
        Write-Host "✓ Hit locations sheet '$effectiveHL' created/updated." -ForegroundColor Green
      }
    } catch {
      Write-Warning "Hit locations step failed: $($_.Exception.Message)"
    }
  }

  # 3) Weapons
  if ($Weapons) {
    try {
      if ($FromClipboard) {
        Write-Host "`n[3/3] Weapons" -ForegroundColor Yellow
        _PressToContinue "WEAPONS TABLE (header row 'Weapon % Damage SR' + any footnotes)"
        Import-CreatureWeaponsFromText -Creature $Creature -FromClipboard -ErrorAction Stop
      } else {
        Import-CreatureWeaponsFromText -Creature $Creature -ErrorAction Stop
      }
      Write-Host "✓ Weapons imported for '$Creature'." -ForegroundColor Green
    } catch {
      Write-Warning "Weapons step failed: $($_.Exception.Message)"
    }
  }

  Write-Host ""
  Write-Host "Done." -ForegroundColor Green
}
