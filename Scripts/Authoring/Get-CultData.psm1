$ErrorActionPreference = 'Stop'

# --- Config ---
$script:CultsDir      = 'Y:\Stat_blocks\Data\Cults'
$script:CultsWorkbook = 'Y:\Stat_blocks\Data\Cults.xlsx'
$script:UseWorkbook   = Test-Path $script:CultsWorkbook

if ($script:UseWorkbook) {
  Import-Module ImportExcel -ErrorAction Stop
}

function Get-RQGCultsRoot {
  [CmdletBinding()]
  param()
  [pscustomobject]@{
    Mode     = if ($script:UseWorkbook) { 'Workbook' } else { 'Folders' }
    DataRoot = if ($script:UseWorkbook) { $script:CultsWorkbook } else { $script:CultsDir }
    Workbook = $script:CultsWorkbook
    CultsDir = $script:CultsDir
  }
}

function Get-ExcelSheetsLike {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Pattern   # e.g. 'Magic|Roles|Associations'
  )
  $all = Get-ExcelSheetInfo -Path $Path
  $all | Where-Object { $_.Name -match $Pattern } | Select-Object -ExpandProperty Name
}

function Read-ExcelWithCult {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$SheetName
  )
  $rows = Import-Excel -Path $Path -WorksheetName $SheetName
  if (-not $rows) { return @() }

  # Ensure a Cult column exists; if not, infer from sheet prefix before first underscore
  $hasCult = ($rows | Select-Object -First 1 | Get-Member -MemberType NoteProperty | Where-Object Name -eq 'Cult')
  if (-not $hasCult) {
    $cult = $SheetName
    if ($SheetName -match '^(.*?)_') { $cult = $Matches[1] }
    # add Cult property to each row
    $rows = $rows | ForEach-Object {
      $h = $_ | Select-Object * -ExcludeProperty Cult
      Add-Member -InputObject $h -NotePropertyName Cult -NotePropertyValue $cult -Force
      $h
    }
  }
  $rows
}

function Get-CultNames {
  [CmdletBinding()] param()
  if ($script:UseWorkbook) {
    $names = @()
    $sheets = Get-ExcelSheetsLike -Path $script:CultsWorkbook -Pattern 'Magic|Roles|Associations'
    foreach ($ws in $sheets) {
      $rows = Read-ExcelWithCult -Path $script:CultsWorkbook -SheetName $ws
      $names += ($rows | Where-Object { $_.Cult } | Select-Object -ExpandProperty Cult)
    }
    $names | Sort-Object -Unique
  } else {
    if (-not (Test-Path $script:CultsDir)) { return @() }
    Get-ChildItem -Path $script:CultsDir -Directory | Select-Object -ExpandProperty Name | Sort-Object
  }
}

function Get-CultMagic {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$CultName)
  if ($script:UseWorkbook) {
    $out = @()
    $sheets = Get-ExcelSheetsLike -Path $script:CultsWorkbook -Pattern 'Magic'
    foreach ($ws in $sheets) {
      $rows = Read-ExcelWithCult -Path $script:CultsWorkbook -SheetName $ws
      $out  += ($rows | Where-Object { $_.Cult -eq $CultName -or -not $_.Cult })
    }
    $out
  } else {
    $path = Join-Path $script:CultsDir "$CultName\Magic.csv"
    if (Test-Path $path) { Import-Csv $path | Where-Object { $_.Cult -eq $CultName -or -not $_.Cult } } else { @() }
  }
}

function Get-CultRoles {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$CultName)
  if ($script:UseWorkbook) {
    $out = @()
    $sheets = Get-ExcelSheetsLike -Path $script:CultsWorkbook -Pattern 'Roles?'
    foreach ($ws in $sheets) {
      $rows = Read-ExcelWithCult -Path $script:CultsWorkbook -SheetName $ws
      $out  += ($rows | Where-Object { $_.Cult -eq $CultName -or -not $_.Cult })
    }
    $out
  } else {
    $path = Join-Path $script:CultsDir "$CultName\Roles.csv"
    if (Test-Path $path) { Import-Csv $path | Where-Object { $_.Cult -eq $CultName -or -not $_.Cult } } else { @() }
  }
}

function Get-CultAssociations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CultName,

        [string]$WorkbookPath = "Y:\Stat_blocks\Data\Cults.xlsx"
    )

    if (-not (Test-Path $WorkbookPath)) {
        return @()
    }

    # Figure out the sheet prefix the same way we do for _Magic
    $prefix = $CultName
    if (Get-Command Resolve-CultSheetName -ErrorAction SilentlyContinue) {
        try {
            $prefix = Resolve-CultSheetName -CultName $CultName -WorkbookPath $WorkbookPath
        } catch {
            # fall back to raw CultName if resolution fails
            $prefix = $CultName
        }
    }

    $sheetName = "${prefix}_Associations"

    $rows = @()
    try {
        $rows = Import-Excel -Path $WorkbookPath -WorksheetName $sheetName -ErrorAction Stop
    } catch {
        # no associations sheet for this cult
        return @()
    }

    $out = foreach ($r in $rows) {

        # --- map columns from new or old schema --------------------------
        # New schema: Cult, Spell, Points, Notes
        # Old schema: FromCult, Provides, Notes

        $assocCult = $null
        $spellName = $null
        $pointsVal = $null
        $notesVal  = $null

        # Associated cult name
        if ($r.PSObject.Properties['Cult']) {
            $assocCult = ('' + $r.Cult).Trim()
        } elseif ($r.PSObject.Properties['FromCult']) {
            $assocCult = ('' + $r.FromCult).Trim()
        }

        # Spell name
        if ($r.PSObject.Properties['Spell']) {
            $spellName = ('' + $r.Spell).Trim()
        } elseif ($r.PSObject.Properties['Provides']) {
            $spellName = ('' + $r.Provides).Trim()
        }

        # Points (new schema only; safe if missing)
        if ($r.PSObject.Properties['Points']) {
            $raw = ('' + $r.Points).Trim()
            if ($raw -ne '') {
                $n = 0
                if ([int]::TryParse($raw, [ref]$n)) {
                    $pointsVal = $n
                } else {
                    $pointsVal = $null
                }
            }
        }

        # Notes
        if ($r.PSObject.Properties['Notes']) {
            $notesVal = ('' + $r.Notes).Trim()
        }

        # skip completely empty rows
        if (-not $assocCult -and -not $spellName) { continue }

        [pscustomobject]@{
            BaseCult = $CultName   # the cult we're querying associations for
            Cult     = $assocCult  # the associated cult providing the spell
            Spell    = $spellName
            Points   = $pointsVal
            Notes    = $notesVal
        }
    }

    $out
}


Export-ModuleMember -Function Get-RQGCultsRoot,Get-CultNames,Get-CultMagic,Get-CultRoles,Get-CultAssociations
