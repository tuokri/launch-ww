# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

_crt_dlls_path = 'C:\\Program Files (x86)\\Windows Kits\\10\\Redist\\10.0.18362.0\\ucrt\\DLLs\\x64'
binaries = [(str(p), ".") for p in Path(_crt_dlls_path).iterdir()]

_icon_path = Path('resources\\ww_icon.ico')
icon = str(_icon_path) if _icon_path.exists() else None

block_cipher = None

a = Analysis(['src\\launch_wwserver.py'],
             pathex=[],
             binaries=binaries,
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data,
          cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='LaunchWinterWarServer',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          icon=icon)
