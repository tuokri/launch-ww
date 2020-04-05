# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None

icon_path = Path('resources\\ww_icon.ico')
icon = str(icon_path) if icon_path.exists() else None

crt_dlls_path = 'C:\\Program Files (x86)\\Windows Kits\\10\\Redist\\10.0.18362.0\\ucrt\\DLLs\\x64'
binaries = [crt_dlls_path]

a = Analysis(['src\\launch_winterwar.py'],
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
          name='LaunchWinterWar',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon=icon)
