# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

_python_root = Path(sys.base_prefix)
_tcl_dir = _python_root / "tcl"
_tkinter_dir = _python_root / "Lib" / "tkinter"
_dll_dir = _python_root / "DLLs"

_datas = []
if _tcl_dir.exists():
    _datas.append((str(_tcl_dir), "tcl"))
if _tkinter_dir.exists():
    _datas.append((str(_tkinter_dir), "tkinter"))

_binaries = [
    (str(_dll_dir / "tcl86t.dll"), "."),
    (str(_dll_dir / "tk86t.dll"), "."),
]
_binaries = [item for item in _binaries if Path(item[0]).exists()]

a = Analysis(
    ['PASTAS_DT.py'],
    pathex=[],
    binaries=_binaries,
    datas=_datas,
    hiddenimports=['tkinter', 'tkinter.messagebox', '_tkinter'],
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
    name='PASTAS_DT',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
