Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Path = 'Y:\Stat_blocks\Data\Cults.xlsx'

if(-not (Test-Path $Path)){ throw "File not found: $Path" }

$excel = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($Path)
  if(-not $wb){ throw "Failed to open workbook: $Path" }

  Write-Host "Opened: $Path"
  foreach($ws in $wb.Worksheets){
    $used = $ws.UsedRange
    $rows = if($used){ [int]$used.Rows.Count } else { 0 }
    $cols = if($used){ [int]$used.Columns.Count } else { 0 }
    Write-Host (" - Sheet: '{0}'  ({1} x {2})" -f $ws.Name, $rows, $cols)
  }
  $wb.Close($false) | Out-Null
} finally {
  if($excel){ $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
