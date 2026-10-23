from pathlib import Path
import sys

root = Path(SPECPATH)
python_root = Path(sys.base_prefix)
tk_datas = [(str(python_root / "tcl" / directory), destination)
            for directory, destination in (("tcl8.6", "_tcl_data"), ("tk8.6", "_tk_data"))
            if (python_root / "tcl" / directory).is_dir()]
tk_binaries = [(str(python_root / "DLLs" / name), ".")
               for name in ("tcl86t.dll", "tk86t.dll")
               if (python_root / "DLLs" / name).is_file()]

analysis = Analysis(
    [str(root / "launch.py")],
    pathex=[str(root)],
    binaries=tk_binaries,
    datas=tk_datas + [(str(root / "assets" / filename), "assets") for filename in ("gawa.ico", "gawa.png")],
    hiddenimports=["tkinter", "tkinter.ttk", "_tkinter"],
    hookspath=[str(root / "hooks")],
    runtime_hooks=[str(root / "scripts" / "pyi_rth_tkinter.py")],
    excludes=["numpy", "pandas", "matplotlib", "scipy", "IPython", "pytest"],
    noarchive=False,
)
archive = PYZ(analysis.pure)
executable = EXE(
    archive,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="加华加油与发票核对助手",
    icon=str(root / "assets" / "gawa_exe.ico"),
    debug=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
