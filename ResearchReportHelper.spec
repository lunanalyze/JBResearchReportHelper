# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    # updater/apply.ps1 은 자동 업데이트가 exe 를 갈아끼울 때 쓰는 교체 스크립트다.
    # exe 안의 Python 은 교체 대상 자신이라 쓸 수 없어, Windows 기본 PowerShell 로 돈다.
    datas=[
        ('report_prompt.md', '.'),
        ('report_template.docx', '.'),
        ('updater/apply.ps1', 'updater'),
    ],
    hiddenimports=['googlenewsdecoder'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ResearchReportHelper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
