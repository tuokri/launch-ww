"""
Purge cached Winter War mod files in Rising Storm 2
workshop cache to avoid conflicts and run Rising Storm 2.
"""
import ctypes.wintypes
import os
import sys
from pathlib import Path
from typing import List

import logbook
from logbook import Logger
from logbook import RotatingFileHandler
from logbook import StreamHandler

VNGAME_EXE = "VNGame.exe"
WW_PACKAGE = "WinterWar.u"
WW_WORKSHOP_ID = 1758494341
AUDIO_DIR = "WwiseAudio"

# Windows constants.
CSIDL_PERSONAL = 5  # My Documents.
SHGFP_TYPE_CURRENT = 0  # Get current, not default value.


def user_documents_dir() -> Path:
    """Return the user's Windows documents directory path."""
    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(
        None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    return Path(buf.value)


USER_DOCS_DIR = user_documents_dir()
CACHE_DIR = USER_DOCS_DIR / Path("My Games\\Rising Storm 2\\ROGame\\Cache")
PUBLISHED_DIR = USER_DOCS_DIR / Path("My Games\\Rising Storm 2\\ROGame\\Published")
LOGS_DIR = USER_DOCS_DIR / Path("My Games\\Rising Storm 2\\ROGame\\Logs")

logbook.set_datetime_format("local")
_sh = StreamHandler(sys.stdout, level="INFO")
_rfh = RotatingFileHandler(
    LOGS_DIR / Path("LaunchWinterWar.log"), level="INFO", bubble=True)
_rfh.format_string = (
    "[{record.time}] {record.level_name}: {record.channel}: "
    "{record.func_name}(): {record.message}"
)
_rfh.push_application()
logger = Logger(__name__)
logger.handlers.append(_rfh)

# No console window in frozen mode.
if not hasattr(sys, "frozen"):
    _sh.push_application()
    logger.handlers.append(_sh)


def find_ww_cache_dirs() -> List[Path]:
    """Find all cache directories containing Winter War files."""
    cache_dirs = [p for p in CACHE_DIR.rglob(WW_PACKAGE)]
    cache_dirs = [
        CACHE_DIR / Path(str(p).lstrip(str(CACHE_DIR)).split(os.path.sep)[0])
        for p in cache_dirs
    ]

    # Defensively add WW_WORKSHOP_ID-directory even if WW_PACKAGE was not found
    # and the directory exists.
    ww_cache_dir = CACHE_DIR / Path(str(WW_WORKSHOP_ID))
    if ww_cache_dir.exists():
        cache_dirs.append(ww_cache_dir)

    return list(set(cache_dirs))


def main():
    logger.info("user documents directory: '{uh}'", uh=USER_DOCS_DIR)
    cache_dirs = find_ww_cache_dirs()
    count = len(cache_dirs)
    logger.info(
        "found {count} Winter War cache directo{p}", count=count,
        p="ry" if count == 1 else "ries")


if __name__ == "__main__":
    main()
