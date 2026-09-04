param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)
$ErrorActionPreference = "Stop"
$src = [System.IO.Path]::GetFullPath($InputPath)
$dst = [System.IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $src)) { throw "source missing" }
if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }

$ppt = New-Object -ComObject PowerPoint.Application
try {
  try { $ppt.DisplayAlerts = 1 } catch {}
  $pres = $ppt.Presentations.Open($src, $true, $false, $false)
  try {
    $pres.SaveAs($dst, 32)
  } finally {
    $pres.Close()
  }
} finally {
  try { $ppt.Quit() } catch {}
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
if (-not (Test-Path -LiteralPath $dst)) { throw "pdf not created" }
