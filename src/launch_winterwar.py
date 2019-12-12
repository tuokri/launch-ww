"""
Purge cached Winter War mod files in Rising Storm 2
workshop cache to avoid conflicts and run Rising Storm 2.
"""
import argparse
import ctypes.wintypes
import os
import shutil
import subprocess
import sys
from argparse import Namespace
from pathlib import Path
from typing import List

import logbook
from logbook import Logger
from logbook import RotatingFileHandler
from logbook import StreamHandler

VNGAME_EXE = "VNGame.exe"
VNGAME_EXE_PATH = Path(
    "steamapps\\common\\Rising Storm 2\\Binaries\\Win64") / Path(VNGAME_EXE)
WW_PACKAGE = "WinterWar.u"
WW_WORKSHOP_ID = 1758494341
AUDIO_DIR = Path("CookedPC\\WwiseAudio")

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
SCRIPT_LOG_PATH = LOGS_DIR / Path("LaunchWinterWar.log")

logbook.set_datetime_format("local")
_rfh_bubble = False if hasattr(sys, "frozen") else True
_rfh = RotatingFileHandler(SCRIPT_LOG_PATH, level="INFO", bubble=_rfh_bubble)
_rfh.format_string = (
    "[{record.time}] {record.level_name}: {record.channel}: "
    "{record.func_name}(): {record.message}"
)
_rfh.push_application()
logger = Logger(__name__)
logger.handlers.append(_rfh)

# No console window in frozen mode.
if hasattr(sys, "frozen"):
    logger.info("not adding stdout logging handler in frozen mode")
else:
    _sh = StreamHandler(sys.stdout, level="INFO")
    logger.info("adding logging handler: {h}", h=_sh)
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


def error_handler(function, path, exc_info):
    logger.error("function={f} path={p}",
                 f=function, p=path, exc_info=exc_info)


def parse_args() -> Namespace:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--dry-run", dest="dry_run",
        action="store_true", default=False,
        help="run without deleting files or launching Rising Storm 2",
    )
    return ap.parse_args()


def main():
    args = parse_args()

    logger.info("user documents directory: '{uh}'", uh=USER_DOCS_DIR)
    cache_dirs = find_ww_cache_dirs()
    count = len(cache_dirs)
    logger.info(
        "found {count} Winter War cache directo{p}", count=count,
        p="ry" if count == 1 else "ries")

    audio_dir = PUBLISHED_DIR / AUDIO_DIR
    if audio_dir.exists():
        logger.info("found audio directory: '{ad}'", ad=audio_dir)
        cache_dirs.append(audio_dir)

    for cache_dir in cache_dirs:
        if args.dry_run:
            logger.info("dry run, not removing: {cd}", cd=cache_dir)
        else:
            logger.info("removing: {cd}", cd=cache_dir)
            shutil.rmtree(
                cache_dir,
                onerror=error_handler,
            )

    if not args.dry_run:
        logger.info("launching Rising Storm 2")
        subprocess.run(["start", "steam://rungameid/418460"], shell=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as _e:
        extra = (f"Check '{SCRIPT_LOG_PATH}' for more information."
                 if SCRIPT_LOG_PATH.exists() else "")
        ctypes.windll.user32.MessageBoxW(
            0,
            f"Error launching Winter War!\r\n"
            f"{type(_e).__name__}: {_e}.\r\n"
            f"{extra}\r\n",
            "Error",
            0
        )
        # noinspection PyBroadException
        try:
            logger.exception("error running script")
        except Exception:
            pass
