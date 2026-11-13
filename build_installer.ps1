$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
# 정품 CPython 3.12 로 빌드한다. Codex 런타임 Python 은 Codex 가 Lib\ssl.py 를 패치해 두어
# (2.0.7 장애 원인) 빌드에 쓰면 안 된다.
$Python = "C:\Users\JBB\AppData\Local\Programs\Python\Python312\python.exe"
$MakeNsis = "C:\Program Files (x86)\NSIS\makensis.exe"

# 빌드에 쓰는 Python 의 표준 ssl 모듈이 변조됐는지 검사한다.
# 정품 CPython 의 SSLContext.__new__ 는 load_verify_locations 를 호출하지 않는다.
# 변조본(예: Codex 런타임의 codex-rejected-cas.pem 주입)을 얼리면 exe 의 HTTPS 가
# 전부 [Errno 2] 로 죽는다 — 2.0.7 이 그렇게 나갔다.
$SslSrc = & $Python -c "import ssl,inspect,sys; sys.stdout.write(inspect.getsource(ssl.SSLContext.__new__))"
if ($LASTEXITCODE -ne 0) { throw "ssl 점검 실패: $Python 을 실행할 수 없습니다" }
if (($SslSrc -join "`n") -match "load_verify_locations") {
  throw "빌드 중단: $Python 의 Lib\ssl.py 가 변조돼 있습니다. 정품 CPython 3.12 로 `$Python 을 바꾸세요.`r`n$($SslSrc -join "`n")"
}
Write-Host "ssl.py 점검: 정상"

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

# 산출 exe 를 실제로 띄워 HTTPS 가 되는지 확인한다. 2.0.7 장애는 소스가 아니라
# 번들에서 생겼으므로, 소스 테스트로는 잡히지 않는다.
# /update/check 는 설치본(exe 옆에 Uninstall.exe)일 때만 GitHub 에 접속하므로
# (updater.install_root), 임시 폴더에 exe 와 빈 Uninstall.exe 표식을 두고 띄운다.
# 8765 포트를 쓰므로 앱이 떠 있으면 먼저 닫아야 한다.
$Busy = Get-NetTCPConnection -State Listen -LocalPort 8765 -ErrorAction SilentlyContinue
if ($Busy) { throw "스모크 테스트 불가: 8765 포트가 사용 중입니다 — 조사연구 도우미를 닫고 다시 빌드하세요." }
$SmokeDir = Join-Path ([System.IO.Path]::GetTempPath()) "rrh-smoke-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Force -Path $SmokeDir | Out-Null
Copy-Item -LiteralPath (Join-Path $Root "dist\ResearchReportHelper.exe") -Destination $SmokeDir
New-Item -ItemType File -Path (Join-Path $SmokeDir "Uninstall.exe") | Out-Null
$ExeOut = Join-Path $SmokeDir "ResearchReportHelper.exe"
$env:RRA_NO_BROWSER = "1"
$Smoke = Start-Process -FilePath $ExeOut -PassThru
try {
  $Ok = $false
  for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Seconds 2
    try {
      $Res = Invoke-WebRequest -Uri "http://127.0.0.1:8765/update/check" -UseBasicParsing -TimeoutSec 20
    } catch { continue }   # 아직 서버가 안 떴을 뿐
    $Body = $Res.Content | ConvertFrom-Json
    if ($Body.note -and $Body.note -match "Errno 2") {
      throw "스모크 테스트 실패: 빌드된 exe 의 HTTPS 가 죽어 있습니다 — $($Body.note)"
    }
    if ($Body.latest) { $Ok = $true; break }   # GitHub 까지 실제로 다녀왔다
    if ($Body.note) { Write-Host "update/check note: $($Body.note)" }
  }
  if (-not $Ok) { throw "스모크 테스트 실패: /update/check 가 매니페스트를 받지 못했습니다" }
  Write-Host "스모크 테스트: HTTPS 정상 (latest=$($Body.latest))"
} finally {
  Get-CimInstance Win32_Process -Filter "Name='ResearchReportHelper.exe'" |
    Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($SmokeDir, [System.StringComparison]::OrdinalIgnoreCase) } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
  Start-Sleep -Seconds 2
  Remove-Item -LiteralPath $SmokeDir -Recurse -Force -ErrorAction SilentlyContinue
}
