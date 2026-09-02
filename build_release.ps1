# ============================================================================
# 릴리스 산출물 만들기 — 자동 업데이트가 먹을 **업데이트 팩**과 **latest.json** 을 만든다.
#
# 자동 업데이트의 규격(updater.py 참조):
#   * 앱은 릴리스의 latest.json 을 읽어 version 을 비교한다.
#   * 새 버전이면 pack.url 에서 zip 을 받아 pack.sha256 과 대조한다. 안 맞으면 중단한다.
#   * 팩 안에는 ResearchReportHelper.exe 하나만 들어간다.
#
# 쓰는 법:
#   1) .\build_installer.ps1        (exe + Setup.exe 를 만든다)
#   2) .\build_release.ps1          (팩 + latest.json 을 dist\release\ 에 만든다)
#   3) GitHub 릴리스를 v<버전> 태그로 만들고, 아래 셋을 **에셋으로** 올린다.
#        - latest.json
#        - ResearchReportHelper-Update-<버전>.zip
#        - ResearchReportHelperSetup.exe   (신규 설치용)
#      latest.json 이 최신 릴리스의 에셋으로 붙어 있어야
#      /releases/latest/download/latest.json 이 그것을 가리킨다.
#
# ⚠ latest.json 은 **BOM 없이** 쓴다. 앱은 BOM 을 걷어내고 읽지만(updater.fetch_manifest),
#   다른 도구가 이 파일을 읽을 때 걸린다.
# ============================================================================
[CmdletBinding()]
param(
    # 릴리스 노트 — 배너와 확인 창에 그대로 보인다. 첫 줄이 배너 한 줄 요약이 된다.
    [string]$Notes = "",
    # 팩을 내려받을 주소. 기본값은 이 저장소의 v<버전> 릴리스 에셋.
    [string]$RepoUrl = "https://github.com/lunanalyze/JBResearchReportHelper"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

# 버전의 단일 원본 — updater.py 의 APP_VERSION.
$VersionLine = Select-String -LiteralPath (Join-Path $Root "updater.py") -Pattern '^APP_VERSION\s*=\s*"([^"]+)"'
if (-not $VersionLine) { throw "updater.py 에서 APP_VERSION 을 찾지 못했습니다" }
$Version = $VersionLine.Matches[0].Groups[1].Value

$ExePath = Join-Path $Root "dist\ResearchReportHelper.exe"
if (-not (Test-Path -LiteralPath $ExePath)) {
    throw "빌드된 실행 파일이 없습니다: $ExePath — 먼저 .\build_installer.ps1 을 돌리세요."
}

$OutDir = Join-Path $Root "dist\release"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$PackName = "ResearchReportHelper-Update-$Version.zip"
$PackPath = Join-Path $OutDir $PackName
if (Test-Path -LiteralPath $PackPath) { Remove-Item -LiteralPath $PackPath -Force }

# 팩은 exe 하나만 담는다. 압축 폴더를 만들지 않도록 파일을 직접 넣는다 —
# apply.ps1 은 zip 항목 이름으로 exe 를 찾으므로 폴더가 껴 있어도 동작하지만, 평평한 편이 낫다.
Compress-Archive -LiteralPath $ExePath -DestinationPath $PackPath -CompressionLevel Optimal
$Sha = (Get-FileHash -LiteralPath $PackPath -Algorithm SHA256).Hash.ToLower()
$Size = (Get-Item -LiteralPath $PackPath).Length

$Manifest = [ordered]@{
    version          = $Version
    released_at      = (Get-Date -Format "yyyy-MM-dd")
    notes            = $Notes
    restart_required = $true
    pack             = [ordered]@{
        name   = $PackName
        url    = "$RepoUrl/releases/download/v$Version/$PackName"
        size   = $Size
        sha256 = $Sha
    }
}

$Json = $Manifest | ConvertTo-Json -Depth 5
$JsonPath = Join-Path $OutDir "latest.json"
# Set-Content -Encoding UTF8 은 BOM 을 붙인다. BOM 없이 쓰려면 .NET 으로 직접 쓴다.
[System.IO.File]::WriteAllText($JsonPath, $Json, [System.Text.UTF8Encoding]::new($false))

Write-Host ""
Write-Host "버전    : $Version"
Write-Host "팩      : $PackPath"
Write-Host "크기    : $([math]::Round($Size / 1MB, 1)) MB"
Write-Host "SHA-256 : $Sha"
Write-Host "매니페스트: $JsonPath"
Write-Host ""
Write-Host "다음: GitHub 에 v$Version 태그로 릴리스를 만들고 latest.json / $PackName / ResearchReportHelperSetup.exe 를 에셋으로 올리세요."
