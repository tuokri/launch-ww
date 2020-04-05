import sys
from pathlib import Path

_crt_dlls_path = 'C:\\Program Files (x86)\\Windows Kits\\10\\Redist\\10.0.18362.0\\ucrt\\DLLs\\x64'
binaries = []
try:
    binaries = [str(p) for p in Path(_crt_dlls_path).iterdir()]
except Exception as e:
    sys.stderr.write(f"error finding UCRT x64 DLLs: {e}\n")

_icon_path = Path('resources\\ww_icon.ico')
icon = str(_icon_path) if _icon_path.exists() else None
