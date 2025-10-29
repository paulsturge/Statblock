Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$InputPath  = 'Y:\Stat_blocks\Data\Cults.xlsx'
$OutputPath = 'Y:\Stat_blocks\Data\Cults.cleaned.xlsx'

function N([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return '' }
  $s = $s -replace "`r`n|`r|`n",' '
  $s = $s -replace "[\u00A0\u2000-\u200B]",' '
  $s = $s -replace '[\u2018\u2019\u02BC]',"'"
  $s = $s -replace '[\u201C\u201D]','"'
  $s = $s -replace '[\u2013\u2014\u2212]', '-'
  ($s -replace '\s+',' ').Trim()
}
function IsBlank($v){ [string]::IsNullOrWhiteSpace((N (''+$v))) }
function DedupHeaders($arr){
  $seen=@{}
  for($i=0;$i -lt $arr.Count;$i++){
    $h = N $arr[$i]; if(-not $h){ $h = "Column$($i+1)" }
    $base=$h; $k=$h.ToLowerInvariant(); $n=1
    while($seen.ContainsKey($k)){ $n++; $h="$base`_$n"; $k=$h.ToLowerInvariant() }
    $seen[$k]=$true; $arr[$i]=$h
  }
  ,$arr
}

# Robustly read a cell from UsedRange.Value2 no matter how Excel shapes it.
# r, c are zero-based positions within UsedRange.
function Get-Cell($vals, $used, [int]$r, [int]$c) {
  if ($null -eq $vals) { return $null }

  # Scalar (single cell) – Excel sometimes returns the value directly
  if ($vals -isnot [Array]) {
    return $vals
  }

  # True 2-D COM array (most common for multi-cell ranges)
  if ($vals.Rank -eq 2) {
    $rLo = $vals.GetLowerBound(0)
    $cLo = $vals.GetLowerBound(1)
    return $vals.GetValue($rLo + $r, $cLo + $c)
  }

  # 1-D array – Excel does this for 1×N or N×1 ranges
  if ($vals.Rank -eq 1) {
    $rows = $used.Rows.Count
    $cols = $used.Columns.Count
    # If one of the dimensions is 1, index along the other.
    if ($rows -eq 1 -and $cols -ge 1) {
      return $vals[$c]
    } elseif ($cols -eq 1 -and $rows -ge 1) {
      return $vals[$r]
    } else {
      # Rare: Excel flattened a bigger block – assume row-major
      $idx = $r * $cols + $c
      return $vals[$idx]
    }
  }

  # Extreme fallback: read the single cell via COM (slower, but safe)
  return $used.Cells.Item($r + 1, $c + 1).Value2
}



Write-Host "Input : $InputPath"
Write-Host "Output: $OutputPath"

if(-not (Test-Path $InputPath)){ throw "File not found: $InputPath" }
if(Test-Path $OutputPath){ Remove-Item $OutputPath -Force }

$excel = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($InputPath)
  if(-not $wb){ throw "Failed to open workbook: $InputPath" }

  $wbOut = $excel.Workbooks.Add()
  while($wbOut.Worksheets.Count -gt 1){ ($wbOut.Worksheets.Item(1)).Delete() | Out-Null }

  $first = $true
  foreach($ws in $wb.Worksheets){
    $name = $ws.Name
    $used = $ws.UsedRange
    if(-not $used -or $used.Rows.Count -eq 0 -or $used.Columns.Count -eq 0){
      Write-Host "Skip empty: '$name'"; continue
    }

    $rows = [int]$used.Rows.Count
    $cols = [int]$used.Columns.Count
    $vals = $used.Value2

    # Build normalized 2D array of cleaned strings
    $arr = New-Object 'object[,]' $rows, $cols
    for($r=0;$r -lt $rows;$r++){
      for($c=0;$c -lt $cols;$c++){
      $cell = Get-Cell $vals $used $r $c

        $arr[$r,$c] = N (''+$cell)
      }
    }

    # Keep non-empty rows
    $keepR = New-Object System.Collections.Generic.List[int]
    for($r=0;$r -lt $rows;$r++){
      $has=$false
      for($c=0;$c -lt $cols;$c++){ if(-not (IsBlank $arr[$r,$c])){ $has=$true; break } }
      if($has){ $keepR.Add($r) }
    }
    if($keepR.Count -eq 0){ Write-Host " -> all blank after clean: '$name'"; continue }

    # Keep non-empty columns
    $keepC = New-Object System.Collections.Generic.List[int]
    for($c=0;$c -lt $cols;$c++){
      $has=$false
      foreach($r in $keepR){ if(-not (IsBlank $arr[$r,$c])){ $has=$true; break } }
      if($has){ $keepC.Add($c) }
    }
    if($keepC.Count -eq 0){ Write-Host " -> no non-empty cols: '$name'"; continue }

    $rows2=$keepR.Count; $cols2=$keepC.Count
    $out = New-Object 'object[,]' $rows2, $cols2
    for($ri=0;$ri -lt $rows2;$ri++){
      for($ci=0;$ci -lt $cols2;$ci++){
        $out[$ri,$ci] = $arr[$keepR[$ri], $keepC[$ci]]
      }
    }

    # Header row = first row; dedupe/patch blanks
    $hdr = @(); for($c=0;$c -lt $cols2;$c++){ $hdr += (''+$out[0,$c]) }
    $hdr = DedupHeaders $hdr
    for($c=0;$c -lt $cols2;$c++){ $out[0,$c]=$hdr[$c] }

    # Write this sheet
    $wso = if($first){ $first=$false; $wbOut.Worksheets.Item(1) } else { $wbOut.Worksheets.Add() }
    $wso.Name = ($name.Substring(0,[Math]::Min(31,$name.Length)))
    $rng = $wso.Range($wso.Cells(1,1), $wso.Cells($rows2,$cols2))
    $rng.Value2 = $out
    $wso.Range($wso.Cells(1,1), $wso.Cells(1,$cols2)).Font.Bold = $true
    $wso.UsedRange.Columns.AutoFit() | Out-Null

    Write-Host ("Wrote '{0}'  rows:{1} cols:{2}" -f $wso.Name, $rows2, $cols2)
  }

  $wbOut.SaveAs($OutputPath) | Out-Null
  $wbOut.Close($true) | Out-Null
  $wb.Close($false) | Out-Null
  Write-Host "Cleaned workbook saved: $OutputPath"
}
finally {
  if($excel){ $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
