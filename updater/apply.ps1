# ============================================================================
# 조사연구 도우미 업데이터 — 앱이 죽은 뒤 exe 를 갈아끼우고 다시 띄운다.
#
# ⚠ 이 파일은 **UTF-8 + BOM 으로 저장**한다. Windows PowerShell 5.1 은 BOM 없는 파일을
#   ANSI(한글 Windows 면 CP949)로 읽는다. 그러면 한글이 깨지는 데서 끝나지 않고 — 3바이트 UTF-8
#   문자가 CP949 2바이트 경계와 어긋나면서 뒤따르는 따옴표·중괄호까지 삼켜 **스크립트가 아예
#   파싱되지 않는다**(실측: "The string is missing the terminator"). updater.py 의 spawn() 이
#   임시 폴더로 복사할 때 BOM 을 다시 붙여 주지만, 원본도 BOM 을 유지해야 직접 실행·편집이 된다.
#
# 이 스크립트는 **설치 폴더 밖(%TEMP%)에서 실행**된다. 설치 폴더 안에서 돌면 자기가 교체하는
# 대상과 같은 폴더에 갇혀 잠금이 풀리지 않는다. 앱의 Python 은 교체 대상인 exe 안에 갇혀 있어
# 쓸 수 없으므로, Windows 에 항상 있는 PowerShell 로 돈다.
#
# 되돌릴 수 있게 만든다: 기존 exe 를 .bak 으로 옮겨두고, 새 앱이 헬스체크를 통과하지 못하면
# 원래대로 복원한 뒤 다시 띄운다. 업데이트 한 번 잘못돼서 앱이 안 켜지는 상황을 막는 유일한 장치다.
#
# 교체 전에 **파일 잠금이 실제로 풀렸는지** 확인한다. 포트가 닫힌 것은 프로세스가 끝난 증거가
# 아니다 — HTTP 서버가 먼저 닫히고 프로세스는 잠깐 더 살아 있다. 게다가 onefile PyInstaller 는
# 부트로더(부모)와 실제 앱(자식) 두 프로세스라, 자식이 끝난 뒤에도 부모가 임시 전개 폴더를
# 치우는 동안 exe 를 계속 쥐고 있다. 그 사이에 교체를 시작하면 실패한다.
# ============================================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Root,
    [Parameter(Mandatory = $true)][string]$Pack,
    [Parameter(Mandatory = $true)][string]$Work,
    [int]$Port = 8765,
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ExeName = "ResearchReportHelper.exe"
$ExePath = Join-Path $Root $ExeName
$BakPath = Join-Path $Root "$ExeName.bak"
$LogPath = Join-Path $Work "update.log"

$Host.UI.RawUI.WindowTitle = "조사연구 도우미 업데이트"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    Write-Host $line
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
}

# 서버가 정말 내려갔는지는 포트로 본다 — PID 는 재사용될 수 있다.
function Wait-PortFree {
    param([int]$TimeoutSeconds = 90)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $client = New-Object System.Net.Sockets.TcpClient
        try {
            $async = $client.BeginConnect("127.0.0.1", $Port, $null, $null)
            $open = $async.AsyncWaitHandle.WaitOne(1000) -and $client.Connected
        } catch {
            $open = $false
        } finally {
            $client.Close()
        }
        if (-not $open) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}

# exe 를 다른 프로세스가 쥐고 있는가. 공유 금지(FileShare::None)로 열어본다 — 쥐고 있으면
# 예외가 나고, 열리면 곧바로 닫으므로 파일은 그대로다.
function Test-FileUnlocked {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    try {
        $stream = [System.IO.File]::Open(
            $Path, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        return $true
    } catch {
        return $false
    }
}

function Wait-Unlocked {
    param([string]$Path, [int]$TimeoutSeconds = 120)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastNotice = Get-Date
    while ($true) {
        if (Test-FileUnlocked -Path $Path) { return $true }
        if ((Get-Date) -ge $deadline) { return $false }
        if (((Get-Date) - $lastNotice).TotalSeconds -gt 15) {
            $lastNotice = Get-Date
            Write-Log "아직 사용 중입니다 — 기다립니다: $Path"
        }
        Start-Sleep -Seconds 1
    }
}

# 무엇이 exe 를 쥐고 있는지 로그에 남긴다 — 무엇을 닫아야 하는지 알려주려고.
function Write-RunningProcesses {
    try {
        $found = Get-CimInstance Win32_Process |
            Where-Object { $_.ExecutablePath -and $_.ExecutablePath.StartsWith($Root, [System.StringComparison]::OrdinalIgnoreCase) } |
            ForEach-Object { "$($_.ProcessId) $($_.Name)" }
        $text = if ($found) { $found -join ", " } else { "없음" }
        Write-Log "설치 폴더에서 실행 중인 프로세스: $text"
    } catch { }
}

# 팩에서 exe 하나만 꺼낸다. 이름을 대조해 고르므로 zip 안의 `..\` 같은 경로로 설치 폴더 밖을
# 덮어쓸 수 없다(zip slip 방지).
function Expand-PackExe {
    param([string]$PackPath, [string]$Target)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($PackPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.Name -ieq $ExeName } | Select-Object -First 1
        if (-not $entry) { throw "팩에 $ExeName 이 없습니다" }
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $Target, $true)
        return $entry.Length
    } finally {
        $zip.Dispose()
    }
}

# 새 콘솔로 띄운다 — 이 스크립트가 끝나도 앱이 살아 있어야 한다.
# 브라우저는 열지 않는다(RRA_NO_BROWSER=1). 업데이트를 시작한 탭이 이미 /heartbeat 를
# 두드리며 기다리다 스스로 새로고침하므로, 여기서 또 열면 탭이 두 개가 된다.
function Start-App {
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ExePath
        $psi.WorkingDirectory = $Root
        $psi.UseShellExecute = $false
        $psi.EnvironmentVariables["RRA_NO_BROWSER"] = "1"
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    } catch {
        Write-Log "앱을 다시 띄우지 못했습니다: $($_.Exception.Message)"
    }
}

function Test-Healthy {
    param([int]$TimeoutSeconds = 120)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $url = "http://127.0.0.1:$Port/heartbeat"
    while ((Get-Date) -lt $deadline) {
        try {
            $res = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 3
            if ($res.StatusCode -eq 200) { return $true }
        } catch {
            # 아직 안 떴을 뿐
        }
        Start-Sleep -Seconds 2
    }
    return $false
}

function Restore-Backup {
    if (-not (Test-Path -LiteralPath $BakPath)) { return $false }
    try {
        if (Test-Path -LiteralPath $ExePath) { Remove-Item -LiteralPath $ExePath -Force }
        Move-Item -LiteralPath $BakPath -Destination $ExePath -Force
        Write-Log "이전 버전으로 복원 완료"
        return $true
    } catch {
        Write-Log "복원 실패: $($_.Exception.Message) — `"$BakPath`" 를 `"$ExePath`" 로 직접 옮겨야 합니다"
        return $false
    }
}

function Stop-WithFailure {
    param([string]$Message)
    Write-Log $Message
    Write-Host ""
    Write-Host "업데이트에 실패했습니다. 이전 버전으로 되돌렸습니다."
    Write-Host "로그: $LogPath"
    Write-Host ""
    Read-Host "확인하셨으면 Enter 를 누르세요"
    exit 1
}

# ── 본 작업 ────────────────────────────────────────────────────────────────
Write-Log "업데이트 시작 — 대상 $Root, 버전 $Version"

if (-not (Wait-PortFree)) {
    Write-Log "포트 $Port 가 계속 열려 있습니다 — 앱이 안 내려갔습니다."
    Write-RunningProcesses
    Stop-WithFailure "아무것도 건드리지 않고 중단합니다."
}
Write-Log "앱 종료 확인"

# 포트가 닫힌 것 != 프로세스가 끝난 것. exe 를 쥔 프로세스가 남아 있으면 교체가 막힌다.
if (-not (Wait-Unlocked -Path $ExePath -TimeoutSeconds 180)) {
    Write-Log "$ExeName 이 아직 사용 중입니다."
    Write-RunningProcesses
    Write-Log "업데이트를 중단합니다 — 열려 있는 '조사연구 도우미' 를 모두 닫고 다시 시도하세요."
    Start-App   # 아무것도 건드리지 않았으므로 원래 앱을 그대로 띄운다
    Stop-WithFailure "아무것도 교체하지 않았습니다."
}
Write-Log "파일 잠금 해제 확인"

# 지난 업데이트가 중간에 죽어 남은 백업은 치운다 — 새 백업 자리를 비워야 한다.
if (Test-Path -LiteralPath $BakPath) {
    try { Remove-Item -LiteralPath $BakPath -Force } catch { }
}

try {
    Move-Item -LiteralPath $ExePath -Destination $BakPath -Force
    Write-Log "이전 $ExeName → $ExeName.bak 보관"
} catch {
    Write-Log "이전 실행 파일을 옮기지 못했습니다: $($_.Exception.Message)"
    Start-App
    Stop-WithFailure "아무것도 교체하지 않았습니다."
}

try {
    $size = Expand-PackExe -PackPath $Pack -Target $ExePath
    if (-not (Test-Path -LiteralPath $ExePath) -or (Get-Item -LiteralPath $ExePath).Length -le 0) {
        throw "전개된 파일이 비어 있습니다"
    }
    Write-Log "새 $ExeName 전개 완료 ($size 바이트)"
} catch {
    Write-Log "전개 실패: $($_.Exception.Message) — 되돌립니다"
    if (Test-Path -LiteralPath $ExePath) { try { Remove-Item -LiteralPath $ExePath -Force } catch { } }
    Restore-Backup | Out-Null
    Start-App
    Stop-WithFailure "이전 버전으로 되돌렸습니다."
}

Write-Log "앱 재시작 …"
Start-App
if (Test-Healthy) {
    Write-Log "헬스체크 통과 — 업데이트 완료 (버전 $Version)"
    # 제어판 '프로그램 추가/제거' 의 버전도 맞춰 둔다. 인스톨러가 아니라 우리가 exe 를 바꿨으니
    # 여기서 고치지 않으면 제어판은 영영 옛 버전을 보여주고, 어느 쪽이 맞는지 알 수 없게 된다.
    if ($Version) {
        try {
            $key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Jeonbuk Bank AI Innovation Department"
            if (Test-Path $key) { Set-ItemProperty -Path $key -Name "DisplayVersion" -Value $Version }
        } catch {
            Write-Log "제어판 버전 표기를 갱신하지 못했습니다: $($_.Exception.Message)"
        }
    }
    try { Remove-Item -LiteralPath $BakPath -Force } catch { }
    try { Remove-Item -LiteralPath $Pack -Force } catch { }
    exit 0
}

Write-Log "새 버전이 뜨지 않습니다 — 되돌립니다"
if (Wait-Unlocked -Path $ExePath -TimeoutSeconds 60) {
    try { Remove-Item -LiteralPath $ExePath -Force } catch { }
    Restore-Backup | Out-Null
    Start-App
} else {
    Write-Log "새 exe 가 사용 중이라 되돌리지 못했습니다 — `"$BakPath`" 가 이전 버전입니다"
}
Stop-WithFailure "새 버전이 실행되지 않았습니다."
