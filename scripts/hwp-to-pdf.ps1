param(
  [Parameter(Mandatory = $true)][string]$InputPath,
  [Parameter(Mandatory = $true)][string]$OutputPath
)
$ErrorActionPreference = "Stop"
$src = [System.IO.Path]::GetFullPath($InputPath)
$dst = [System.IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $src)) { throw "source missing" }
if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }

$hwp = New-Object -ComObject HWPFrame.HwpObject
try {
  try { [void]$hwp.SetMessageBoxMode(0x00020000) } catch {}
  foreach ($name in @("FilePathCheckerModuleExample", "FilePathCheckerModule", "SecurityModule")) {
    try { [void]$hwp.RegisterModule("FilePathCheckDLL", $name) } catch {}
  }
  $opened = $false
  try { $opened = [bool]$hwp.Open($src, "", "lock:false;versionwarning:false;forceopen:true") } catch {}
  if (-not $opened) {
    try { $opened = [bool]$hwp.Open($src, "", "") } catch {}
  }
  if (-not $opened) { throw "hwp open failed" }
  $saved = $false
  try { $saved = [bool]$hwp.SaveAs($dst, "PDF", "") } catch {}
  if (-not $saved) { $hwp.SaveAs($dst, "PDF") }
} finally {
  try { [void]$hwp.Clear(1) } catch {}
  try { $hwp.Quit() } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($hwp) } catch {}
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
if (-not (Test-Path -LiteralPath $dst)) { throw "pdf not created" }
