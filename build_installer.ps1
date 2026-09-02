$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = "C:\Users\JBB\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$MakeNsis = "C:\Program Files (x86)\NSIS\makensis.exe"

Set-Location $Root

# 버전의 단일 원본은 updater.py 의 APP_VERSION 이다. 여기서 읽어 NSIS 로 넘긴다 —
# 앱이 스스로 보고하는 버전과 제어판 '프로그램 추가/제거' 의 버전이 어긋나면,
# 어느 쪽이 맞는지 확인할 방법이 없어진다.
$VersionLine = Select-String -LiteralPath (Join-Path $Root "updater.py") -Pattern '^APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionLine) { throw "updater.py 에서 APP_VERSION 을 찾지 못했습니다" }
$AppVersion = $VersionLine.Matches[0].Groups[1].Value
Write-Host "APP_VERSION: $AppVersion"

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
  --add-data "updater\apply.ps1;updater" `
  app.py
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller failed with exit code $LASTEXITCODE"
}

& $MakeNsis "/DAPP_VERSION=$AppVersion" installer.nsi
if ($LASTEXITCODE -ne 0) {
  throw "NSIS failed with exit code $LASTEXITCODE"
}

Write-Host "Installer: $Root\dist\ResearchReportHelperSetup.exe"
