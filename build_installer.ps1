$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\JBB\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$MakeNsis = "C:\Program Files (x86)\NSIS\makensis.exe"

Set-Location $Root

$ResourceDir = Join-Path $env:LOCALAPPDATA "ResearchReportAutomation\resources"
foreach ($Name in @("report_prompt.md", "report_template.docx")) {
  $Source = Join-Path $ResourceDir $Name
  if (Test-Path $Source) {
    Copy-Item -LiteralPath $Source -Destination (Join-Path $Root $Name) -Force
  }
}

& $Python -m pip install --upgrade googlenewsdecoder
if ($LASTEXITCODE -ne 0) {
  throw "Failed to install googlenewsdecoder with exit code $LASTEXITCODE"
}

& $Python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --noconsole `
  --name ResearchReportHelper `
  --hidden-import googlenewsdecoder `
  --add-data "report_prompt.md;." `
  --add-data "report_template.docx;." `
  app.py
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

& $MakeNsis installer.nsi
if ($LASTEXITCODE -ne 0) {
  throw "NSIS failed with exit code $LASTEXITCODE"
}

Write-Host "Installer: $Root\dist\ResearchReportHelperSetup.exe"
