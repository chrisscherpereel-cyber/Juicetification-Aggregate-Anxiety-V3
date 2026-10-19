param([string]$src, [string]$dst)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open((Resolve-Path $src).Path)
  $xl.CalculateFull()
  $full = Join-Path (Resolve-Path (Split-Path $dst)).Path (Split-Path $dst -Leaf)
  if (Test-Path $full) { Remove-Item $full }
  $wb.SaveAs($full, 51)
  $wb.Close($false)
  "saved $full"
} finally { $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null }
