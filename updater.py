"""원격 업데이트 — GitHub Releases 의 ``latest.json`` 을 보고, 새 버전이면 팩을 받아 적용한다.

AI_IB_Agent_Converged 의 자동 업데이트(Spring Boot ``service/update/*``)와 **같은 규격**이다.
manifest 형식·SHA-256 검증·별도 프로세스 교체·실패 시 자동 복구가 모두 동일하고, 아래 두 가지만
이 앱의 배포 형태(PyInstaller onefile + NSIS)에 맞춰 바뀌었다.

1. **팩에 담기는 것**
   Converged 는 ``<root>\\app\\`` 폴더(jar + web 정적 파일)를 통째로 담는다. 이 앱은 실행 파일이
   ``ResearchReportHelper.exe`` **하나뿐**이라 팩도 그 exe 하나만 담는다. 프롬프트·보고서 서식은
   exe 안에 들어 있고 실행할 때마다 ``paths.copy_default_resource(..., overwrite=True)`` 가
   ``%LOCALAPPDATA%`` 로 다시 펼치므로, exe 만 갈아끼우면 자원도 함께 새것이 된다.

2. **업데이터를 무엇으로 돌리는가**
   Converged 는 설치본에 번들된 Python(``runtime\\python\\python.exe``)으로 ``apply.py`` 를 돌린다.
   업데이트가 ``runtime\\`` 을 건드리지 않으니 교체 중에도 인터프리터가 안전하기 때문이다.
   이 앱의 Python 은 **교체 대상인 exe 안에 갇혀 있어** 그렇게 쓸 수 없다. 그래서 교체 스크립트는
   Windows 에 항상 있는 **PowerShell**(``updater/apply.ps1``)로 쓴다.

그대로 가져온 설계 판단(이유는 Converged 쪽 주석에 자세하다):

* **manifest 는 우리 형식**(``latest.json``). GitHub API 대신
  ``/releases/latest/download/latest.json`` 을 쓴다 — 최신 릴리스 에셋으로 리다이렉트해 주므로
  토큰도, 레이트리밋도, 릴리스 목록 파싱도 없다. 대신 **공개 저장소**여야 한다.
* **교체는 별도 프로세스**가 한다. 실행 중인 exe 는 Windows 가 잠그고 있어 스스로 못 바꾼다 —
  스크립트를 임시 폴더에 풀어 띄우고 이 프로세스는 종료한다.
* **SHA-256 을 반드시 검증**한다. 업데이터는 사용자 PC 로 코드가 들어오는 통로라, 받은 파일이
  릴리스가 공표한 그것인지 확인하지 않으면 통로가 그대로 구멍이 된다.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

import paths


#: 이 빌드의 버전. **여기가 단일 원본**이다 — ``build_installer.ps1`` 이 이 값을 읽어
#: NSIS 의 ``DisplayVersion`` 으로 넘기고, ``build_release.ps1`` 이 ``latest.json`` 에 적는다.
APP_VERSION = "2.0.4"

#: 업데이트 manifest 주소. 배포처를 바꿀 수 있게 환경변수로 덮을 수 있다(사내 파일서버 등).
#: 소스가 아니라 **릴리스 배포용 공개 저장소**를 가리킨다 — 클라이언트에 토큰을 심지 않으려면
#: 인증 없이 읽혀야 한다.
DEFAULT_FEED_URL = (
    "https://github.com/lunanalyze/JBResearchReportHelper"
    "/releases/latest/download/latest.json"
)

APP_EXE_NAME = "ResearchReportHelper.exe"
UPDATER_SCRIPT = "apply.ps1"
WORK_DIR_NAME = "research-report-helper-update"

_USER_AGENT = "research-report-helper-updater"
_FALSY = {"0", "false", "off", "no"}

# Windows CreateProcess 플래그 — 자식을 **새 콘솔**로 띄운다. 이 프로세스는 곧 죽는데
# (자기 exe 를 잠그고 있어 교체가 안 되므로) 자식이 같이 죽으면 아무것도 갈아끼우지 못한 채
# 앱만 닫힌다. 창이 보이는 것도 의도다 — 업데이트가 조용히 실패하는 게 최악이다.
_CREATE_NEW_CONSOLE = 0x00000010


class UpdateError(Exception):
    """업데이트를 진행할 수 없는 사유. ``status`` 는 화면에 그대로 내려보낼 HTTP 코드."""

    def __init__(self, status: int, message: str) -> None:
        super().__init__(message)
        self.status = status
        self.message = message


def feed_url() -> str:
    return (os.environ.get("RRH_UPDATE_FEED") or DEFAULT_FEED_URL).strip()


def enabled() -> bool:
    return (os.environ.get("RRH_UPDATE_ENABLED") or "1").strip().lower() not in _FALSY


def install_root() -> Path | None:
    """설치 루트. 설치본이 아니면 ``None`` — 그때는 업데이트를 아예 막는다.

    개발 환경(``python app.py``)에서 업데이트를 걸면 소스 트리를 덮어쓰는 사고가 난다.
    경로를 설정으로 받지 않는 이유도 같다 — 설정이 실제 위치와 어긋나면 엉뚱한 폴더를 갈아엎는다.
    지금 실행 중인 exe 의 위치에서 되짚는다.

    인스톨러가 만드는 구조::

        <root>\\ResearchReportHelper.exe   ← 지금 실행 중인 이 파일(교체 대상)
        <root>\\Uninstall.exe              ← NSIS 가 쓴다. 설치본이라는 표식
    """
    if not getattr(sys, "frozen", False):
        return None  # 개발 실행 — exe 가 아니다
    exe = Path(sys.executable).resolve()
    if exe.name.lower() != APP_EXE_NAME.lower():
        return None
    root = exe.parent
    return root if (root / "Uninstall.exe").is_file() else None


def installed() -> bool:
    """설치본으로 실행 중인가(= 업데이트를 걸 수 있는가)."""
    return install_root() is not None


def is_newer(latest: str, current: str) -> bool:
    """``2.0.9`` vs ``2.0.10`` 처럼 자리수가 다른 경우를 문자열 비교로 틀리지 않게 숫자 단위로 본다.

    숫자가 아닌 꼬리(``-rc1`` 등)는 무시한다.
    """
    if not (latest or "").strip():
        return False
    if not current or current == "dev":
        return False  # 개발본에 업데이트를 권하지 않는다
    a = _parse_version(latest)
    b = _parse_version(current)
    for i in range(max(len(a), len(b))):
        x = a[i] if i < len(a) else 0
        y = b[i] if i < len(b) else 0
        if x != y:
            return x > y
    return False


def _parse_version(value: str) -> list[int]:
    text = value.strip()
    if text[:1] in {"v", "V"}:
        text = text[1:]
    out = []
    for part in text.split("."):
        digits = ""
        for ch in part:
            if not ch.isdigit():
                break
            digits += ch
        out.append(int(digits) if digits else 0)
    return out


# ── manifest ────────────────────────────────────────────────────────────────
def _open(url: str, timeout: float):
    """GitHub 릴리스 에셋은 302 로 CDN 으로 넘긴다 — urllib 은 리다이렉트를 따라간다."""
    request = urllib.request.Request(
        url, headers={"Accept": "*/*", "User-Agent": _USER_AGENT}
    )
    return urllib.request.urlopen(request, timeout=timeout)


def fetch_manifest(timeout: float = 15.0) -> dict:
    try:
        with _open(feed_url(), timeout) as response:
            raw = response.read()
    except urllib.error.HTTPError as exc:  # noqa: PERF203
        if exc.code == 404:
            raise RuntimeError(
                "HTTP 404 (릴리스에 latest.json 이 없거나 저장소가 비공개입니다)"
            ) from exc
        raise RuntimeError(f"HTTP {exc.code}") from exc
    # BOM 제거 — latest.json 은 PowerShell 이 만드는데 Set-Content -Encoding UTF8 이 BOM 을
    # 붙인다. json.loads 는 선두 U+FEFF 에서 통째로 실패하고, 그러면 업데이트 알림이
    # 아무에게도 안 뜬다. 빌드 쪽에서도 BOM 없이 쓰지만 여기서 한 번 더 막는다.
    text = raw.decode("utf-8-sig", errors="replace")
    data = json.loads(text or "{}")
    if not isinstance(data, dict):
        raise RuntimeError("latest.json 형식이 올바르지 않습니다")
    return data


# ── 확인 ────────────────────────────────────────────────────────────────────
def check() -> dict:
    """새 버전이 있는지 확인.

    **실패는 오류가 아니다** — 망이 막힌 PC 도 있으므로 사유만 담아 돌려준다. 화면이 배너를
    안 띄우면 그만이다. 여기서 예외를 던지면 앱 첫 화면에 에러가 뜬다.
    """
    root = install_root()
    out: dict = {
        "ok": True,
        "current": APP_VERSION,
        "installed": root is not None,
        "available": False,
    }
    if not enabled():
        out["note"] = "자동 업데이트가 꺼져 있습니다."
        return out
    if root is None:
        out["note"] = "설치본이 아닙니다(개발 실행) — 업데이트를 적용할 수 없습니다."
        return out
    try:
        manifest = fetch_manifest()
    except Exception as exc:  # noqa: BLE001 — 망 문제까지 전부 '확인 못 함'으로 접는다
        out["note"] = f"업데이트 정보를 가져오지 못했습니다: {exc}"
        return out

    latest = str(manifest.get("version") or "")
    pack = manifest.get("pack") or {}
    out["latest"] = latest
    out["notes"] = str(manifest.get("notes") or "")
    out["released_at"] = str(manifest.get("released_at") or "")
    out["size"] = int(pack.get("size") or 0)
    out["restart_required"] = bool(manifest.get("restart_required", True))
    out["available"] = is_newer(latest, APP_VERSION)
    if out["available"] and not (pack.get("url") and pack.get("sha256")):
        # 해시가 없으면 적용을 막는다 — 검증 못 하는 코드를 받아 실행할 수는 없다.
        out["available"] = False
        out["note"] = "릴리스에 팩 주소나 SHA-256이 없어 적용할 수 없습니다."
    return out


# ── 적용 ────────────────────────────────────────────────────────────────────
def apply(port: int, busy_reason=None) -> dict:
    """업데이트 적용 — 내려받아 검증하고 업데이터를 띄운다.

    돌아온 뒤 **호출자가 이 프로세스를 종료해야 한다**. 실행 중인 exe 를 잠근 채로는 교체가 안 된다.

    진행 중인 작업이 있으면 거부한다(409). 수집·보고서 생성 도중에 프로세스를 죽이면 사용자는
    몇 분치 작업과 LLM 비용을 잃는다.
    """
    if not enabled():
        raise UpdateError(403, "자동 업데이트가 꺼져 있습니다.")
    root = install_root()
    if root is None:
        raise UpdateError(412, "설치본이 아닙니다(개발 실행) — 업데이트를 적용할 수 없습니다.")
    busy = busy_reason() if busy_reason else None
    if busy:
        raise UpdateError(409, busy)

    try:
        manifest = fetch_manifest()
    except Exception as exc:  # noqa: BLE001
        raise UpdateError(502, f"업데이트 정보를 가져오지 못했습니다: {exc}") from exc

    latest = str(manifest.get("version") or "")
    if not is_newer(latest, APP_VERSION):
        raise UpdateError(409, "이미 최신 버전입니다.")
    pack = manifest.get("pack") or {}
    url = str(pack.get("url") or "")
    expected = str(pack.get("sha256") or "").lower()
    if not url or not expected:
        raise UpdateError(502, "릴리스에 팩 주소나 SHA-256이 없습니다.")

    work = Path(tempfile.gettempdir()) / WORK_DIR_NAME
    work.mkdir(parents=True, exist_ok=True)
    pack_path = work / f"update-{latest}.zip"
    try:
        _download(url, pack_path)
        actual = _sha256(pack_path)
    except Exception as exc:  # noqa: BLE001
        raise UpdateError(502, f"업데이트 파일을 받지 못했습니다: {exc}") from exc
    if actual != expected:
        pack_path.unlink(missing_ok=True)
        raise UpdateError(
            502,
            "받은 파일의 검증값이 릴리스와 다릅니다 — 적용을 중단했습니다."
            f" (기대 {expected[:12]}…, 실제 {actual[:12]}…)",
        )

    try:
        spawn(root, work, pack_path, port, latest)
    except Exception as exc:  # noqa: BLE001
        raise UpdateError(500, f"업데이터를 시작하지 못했습니다: {exc}") from exc

    return {
        "ok": True,
        "version": latest,
        "message": "업데이트를 적용합니다. 앱이 자동으로 다시 시작됩니다.",
    }


def _download(url: str, target: Path) -> None:
    with _open(url, 20.0) as response, open(target, "wb") as out:
        while True:
            chunk = response.read(1 << 16)
            if not chunk:
                break
            out.write(chunk)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            chunk = handle.read(1 << 16)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def spawn(root: Path, work: Path, pack: Path, port: int, version: str) -> None:
    """업데이터 기동 — exe 안의 스크립트를 임시 폴더에 풀고 **떼어낸 프로세스**로 띄운다.

    임시 폴더에 푸는 이유가 두 가지다. (1) onefile exe 안의 파일은 경로로 직접 줄 수 없다.
    (2) PyInstaller 는 종료할 때 전개 폴더(``sys._MEIPASS``)를 지우므로, 거기서 스크립트를
    돌리면 앱이 죽는 순간 스크립트도 발밑이 사라진다.
    """
    source = Path(paths.APP_DIR) / "updater" / UPDATER_SCRIPT
    if not source.is_file():
        raise RuntimeError(f"업데이터 스크립트가 없습니다: {source}")
    script = work / UPDATER_SCRIPT
    # **BOM 을 붙여 쓴다.** Windows PowerShell 5.1 은 BOM 없는 파일을 ANSI(한글 Windows 면
    # CP949)로 읽어서, UTF-8 한글이 깨진 채 파싱된다. 저장소의 원본이 어떤 인코딩이든
    # 여기서 한 번에 맞춰 둔다.
    script.write_text(source.read_text(encoding="utf-8-sig"), encoding="utf-8-sig")

    subprocess.Popen(  # noqa: S603 — 인자는 전부 우리가 만든 경로다
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-Root",
            str(root),
            "-Pack",
            str(pack),
            "-Work",
            str(work),
            "-Port",
            str(port),
            "-Version",
            version or "",
        ],
        cwd=str(work),
        creationflags=_CREATE_NEW_CONSOLE,
        close_fds=True,
    )
