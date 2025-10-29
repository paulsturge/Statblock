#requires -Modules ImportExcel
param(
  [string]$InPath  = 'Y:\Stat_blocks\Data\Cults.xlsx',
  [string]$OutPath = 'Y:\Stat_blocks\Data\Cults.cleaned.xlsx'
)

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel -ErrorAction Stop

function Normalize-Header {
  param([string]$h)
  if ([string]::IsNullOrWhiteSpace($h)) { return '' }
  ($h -replace '[^\p{L}\p{Nd}]+',' ') -replace '\s+',' ' |
    ForEach-Object { $_.Trim() } |
    ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) } |
    ForEach-Object {
      $_ -replace ' ',''
    }
}

function Normalize-Table {
  param([object[]]$Rows)
  if (-not $Rows) { return @() }
  $first = $Rows | Select-Object -First 1
  $map = @{}
  foreach ($p in $first.PSObject.Properties) {
    $nk = Normalize-Header $p.Name
    if (-not $nk) { continue }
    if ($map.ContainsKey($nk)) { continue }
    $map[$p.Name] = $nk
  }
  foreach ($r in $Rows) {
    $o = [ordered]@{}
    foreach ($p in $r.PSObject.Properties) {
      if (-not $map.ContainsKey($p.Name)) { continue }
      $nk = $map[$p.Name]
      $v = $p.Value
      if ($v -is [string]) { $v = $v.Trim() }
      $o[$nk] = $v
    }
    [pscustomobject]$o
  }
}
# Ensure a clean workbook start
if (Test-Path $OutPath) { Remove-Item $OutPath -Force }

$script:firstSheet = $true
function Write-ExcelSheet {
  param(
    [Parameter(Mandatory)] $Data,
    [Parameter(Mandatory)] [string] $SheetName
  )
  if ($null -eq $Data) { return }
  $rows =
    if ($Data -is [System.Collections.IEnumerable] -and -not ($Data -is [string])) {
      [System.Linq.Enumerable]::Count([System.Collections.Generic.List[object]]@($Data))
    } else { 1 }

  if ($rows -eq 0) { return }  # skip empty

  $common = @{
    Path          = $OutPath
    WorksheetName = $SheetName
    TableName     = ($SheetName -replace '[^\w]','_')
    AutoSize      = $true
    BoldTopRow    = $true
  }

  if ($script:firstSheet) {
    $null = $Data | Export-Excel @common
    $script:firstSheet = $false
  } else {
    $null = $Data | Export-Excel @common -Append
  }
}

function Get-Sheet {
  param([string]$Path,[string]$Name)
  if (-not (Test-Path $Path)) { return @() }
  try {
    $rows = Import-Excel -Path $Path -WorksheetName $Name -ErrorAction Stop
    if (-not $rows) { return @() }
    Normalize-Table -Rows $rows
  } catch { @() }
}

# Parse a big text blob into rune/spirit spell rows:
# Matches lines like:  "Carry (Disease)  2 Points"  or  "Create Skeleton 1 Point"
function Parse-SpellsFromText {
  param(
    [string]$Cult,
    [string]$Text,
    [ValidateSet('Rune','Spirit')][string]$MagicType
  )
  if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

  # Regex: Start of line, capture spell name (letters, spaces, punctuation), then a number, then "Point" or "Points"
  $regex = "^\s*([A-Za-z][A-Za-z0-9'()\-/ ]+?)\s+(\d+)\s*Point[s]?\b"
  $list = New-Object System.Collections.Generic.List[object]
  $idx  = 0
  foreach ($line in ($Text -split "\r?\n")) {
    $idx++
    $m = [regex]::Match($line, $regex)
    if ($m.Success) {
      $spell  = $m.Groups[1].Value.Trim()
      $points = [int]$m.Groups[2].Value
      if ($spell -and $points -ge 0) {
        $list.Add([pscustomobject]@{
          Cult      = $Cult
          MagicType = $MagicType
          Spell     = $spell
          Points    = $points
          Access    = ''
          Notes     = ''
          SourceRow = $idx
        })
      }
    }
  }
  $list
}

# Load sheets we expect / gracefully handle if missing
$spellsSheet = Get-Sheet -Path $InPath -Name 'Spells'
$rolesSheet  = Get-Sheet -Path $InPath -Name 'Roles'
$assocSheet  = Get-Sheet -Path $InPath -Name 'Associations'
$tablesSheet = Get-Sheet -Path $InPath -Name 'Tables'

# Canonicalize key column names we rely on
function CoalesceCol {
  param($row, [string[]]$candidates)
  foreach ($c in $candidates) {
    if ($row.PSObject.Properties.Name -contains $c -and $row.$c) { return $row.$c }
  }
  return $null
}

# --- Build clean Spells ---
$cleanSpells = New-Object System.Collections.Generic.List[object]

# 1) Start with any good rows as-is (Spell not equal to Cult, non-empty)
if ($spellsSheet) {
  foreach ($r in $spellsSheet) {
    $cult  = CoalesceCol $r @('Cult')
    $spell = CoalesceCol $r @('Spell','Name','SpellName')
    $mtype = CoalesceCol $r @('MagicType','Type')
    $pts   = CoalesceCol $r @('Points','Cost')
    $acc   = CoalesceCol $r @('Access')
    $notes = CoalesceCol $r @('Notes','Detail')
    if ($spell -and $cult -and ($spell -ne $cult)) {
      $row = [ordered]@{
        Cult      = ''+$cult
        MagicType = ''+$mtype
        Spell     = ''+$spell
        Points    = if ($pts -ne $null -and "$pts" -match '^\d+$') { [int]$pts } else { "$pts" }
        Access    = ''+$acc
        Notes     = ''+$notes
      }
      $cleanSpells.Add([pscustomobject]$row) | Out-Null
    }
  }
}

# 2) For cults that look broken (Spell==Cult), rebuild from Tables text
$cults = @()
if ($spellsSheet) { $cults += ($spellsSheet | ForEach-Object { $_.Cult } | Where-Object { $_ } | Select-Object -Unique) }
if ($tablesSheet) { $cults += ($tablesSheet | ForEach-Object { $_.Cult } | Where-Object { $_ } | Select-Object -Unique) }
$cults = $cults | Sort-Object -Unique

foreach ($cult in $cults) {
  $existing = $cleanSpells | Where-Object { $_.Cult -eq $cult }
  # if we already have usable spells for this cult, keep them and also try to add any missing parsed ones.
  # If we *don’t* have any, try to fully parse.
  $tablesForCult = @()
  if ($tablesSheet) {
    $tablesForCult = $tablesSheet | Where-Object { $_.Cult -eq $cult }
  }

  if ($tablesForCult.Count -gt 0) {
    foreach ($t in $tablesForCult) {
      $tname = CoalesceCol $t @('Table','Name','Title','Section')
      $body  = CoalesceCol $t @('Body','Text','Notes','Content')
      if (-not $body) { continue }

      $isRune   = $false
      $isSpirit = $false
      if ($tname) {
        $n = (''+$tname).ToLower()
        if ($n -match 'rune')   { $isRune = $true }
        if ($n -match 'spirit') { $isSpirit = $true }
      }
      # If table name doesn't hint, try detecting “Points” pattern
      if (-not ($isRune -or $isSpirit)) {
        if ($body -match '(?m)^\s*[A-Za-z].+\s\d+\s*Point') { $isRune = $true } # heuristic
      }

      $mtype = if ($isSpirit) { 'Spirit' } elseif ($isRune) { 'Rune' } else { $null }
      if ($mtype) {
        $parsed = Parse-SpellsFromText -Cult $cult -Text $body -MagicType $mtype
        foreach ($p in $parsed) {
          # Skip if we already have this exact spell entry
          $dup = $cleanSpells | Where-Object {
            $_.Cult -eq $cult -and $_.Spell -eq $p.Spell -and $_.MagicType -eq $p.MagicType
          } | Select-Object -First 1
          if (-not $dup) { $cleanSpells.Add($p) | Out-Null }
        }
      }
    }
  }
}

# Final tidy of Spells
$cleanSpells = $cleanSpells |
  Sort-Object Cult, MagicType, Spell |
  Select-Object Cult, MagicType, Spell,
                @{n='Points';e={ $_.Points }},
                @{n='Access';e={ $_.Access }},
                @{n='Notes';e={ $_.Notes }}

# --- Clean Roles (light pass) ---
$cleanRoles = @()
if ($rolesSheet) {
  $cleanRoles = $rolesSheet | ForEach-Object {
    [pscustomobject]@{
      Cult         = CoalesceCol $_ @('Cult')
      Role         = CoalesceCol $_ @('Role','Rank')
      Requirements = CoalesceCol $_ @('Requirements','Reqs')
      Duties       = CoalesceCol $_ @('Duties','Obligations')
      Benefits     = CoalesceCol $_ @('Benefits','Perks')
      Notes        = CoalesceCol $_ @('Notes','Detail')
    }
  } | Where-Object { $_.Cult -and $_.Role }
}

# --- Clean Associations (light pass) ---
$cleanAssoc = @()
if ($assocSheet) {
  $cleanAssoc = $assocSheet | ForEach-Object {
    [pscustomobject]@{
      Cult   = CoalesceCol $_ @('Cult')
      From   = CoalesceCol $_ @('From','SourceCult','Associate')
      Gives  = CoalesceCol $_ @('Gives','Type','Relation')
      Notes  = CoalesceCol $_ @('Notes','Detail')
      What_Provided = CoalesceCol $_ @('WhatProvided','What_Provided','Provides')
    }
  } | Where-Object { $_.Cult -and $_.From }
}

# --- Tables (rename and trim only) ---
$cleanTables = @()
if ($tablesSheet) {
  $cleanTables = $tablesSheet | ForEach-Object {
    [pscustomobject]@{
      Cult   = CoalesceCol $_ @('Cult')
      Table  = CoalesceCol $_ @('Table','Name','Title','Section')
      Body   = CoalesceCol $_ @('Body','Text','Notes','Content')
    }
  } | Where-Object { $_.Cult -and ($_.Table -or $_.Body) }
}

# Write out the cleaned workbook
if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
"Writing: $OutPath"
$pkg = Export-Excel -Path $OutPath -WorksheetName 'Spells'       -InputObject $cleanSpells -AutoSize -PassThru
    $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Roles'        -InputObject $cleanRoles  -AutoSize -ClearSheet
    $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Associations' -InputObject $cleanAssoc  -AutoSize -ClearSheet
    $null = Export-Excel -ExcelPackage $pkg -WorksheetName 'Tables'       -InputObject $cleanTables -AutoSize -ClearSheet
Close-ExcelPackage $pkg

"Done."
