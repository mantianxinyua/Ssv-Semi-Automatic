# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec：收集 Python 全部标准库（含全部子模块与 C 扩展），
第三方库（pandas/numpy/cv2 等）保持外置（excludes），运行时从 exe 同级的 libs\ 加载。
"""
import os
import sys

# 1) 顶层标准库模块（sys.stdlib_module_names 为 Python 3.10+ 内置 API）
hiddenimports = list(sys.stdlib_module_names)

# 2) 遍历 Lib\ 目录，把全部标准库子模块也加入（tkinter.ttk、concurrent.futures 等）
_LIB = os.path.join(os.path.dirname(sys.executable), 'Lib')
for _root, _dirs, _files in os.walk(_LIB):
    _dirs[:] = [d for d in _dirs if d not in ('site-packages', 'test', '__pycache__')]
    _rel = os.path.relpath(_root, _LIB)
    for _f in _files:
        if _f.endswith('.py') and not _f.startswith('__'):
            _mod = os.path.splitext(_f)[0]
            hiddenimports.append(_mod if _rel == '.' else _rel.replace(os.sep, '.') + '.' + _mod)

hiddenimports = sorted(set(hiddenimports))

a = Analysis(
    ['main.py'],
    pathex=[],
    # python3.dll：cv2.pyd 等稳定 ABI 扩展依赖的转发层，必须与 python313.dll 同目录
    binaries=[(r'C:\Users\Whj\AppData\Local\Programs\Python\Python313\python3.dll', '.')],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 第三方库全部排除：它们由 main.py 注入的 libs\ 目录提供
    excludes=['pandas', 'numpy', 'openpyxl', 'PIL', 'cv2',
              'dateutil', 'tzdata', 'et_xmlfile'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='报告生成工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='报告生成工具',
)
