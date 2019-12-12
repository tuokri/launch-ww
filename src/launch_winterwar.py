"""
Purge cached Winter War mod files in Rising Storm 2
workshop cache to avoid conflicts and run Rising Storm 2.
"""
import argparse
import ctypes.wintypes
import errno
import os
import shutil
import subprocess
import sys
import winreg
from argparse import Namespace
from pathlib import Path
from typing import List
from winreg import HKEY_LOCAL_MACHINE

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
STEAM_REG_32 = "SOFTWARE\\Valve\\Steam"
STEAM_REG_64 = "SOFTWARE\\Wow6432Node\\Valve\\Steam"


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


def find_steam_install_dir() -> Path:
    """Read Steam installation directory from Windows registry."""
    steam_reg_keys = [STEAM_REG_32, STEAM_REG_64]

    for steam_reg_key in steam_reg_keys:
        logger.info("trying to read key: {k}", k=steam_reg_key)
        try:
            with winreg.OpenKey(HKEY_LOCAL_MACHINE, steam_reg_key) as key:
                try:
                    return Path(winreg.QueryValueEx(key, "InstallPath")[0])
                except OSError as ose:
                    if ose.errno != errno.ENOENT:
                        raise
        except FileNotFoundError:
            logger.info("key not found: {k}", k=steam_reg_key)

    raise RuntimeError("Could not find Steam directory")


def find_vngame_exe() -> Path:
    steam_dir = find_steam_install_dir()
    vngame_exe_path = steam_dir / VNGAME_EXE_PATH

    logger.info("looking for '{exe}'", exe=vngame_exe_path)
    if not vngame_exe_path.exists():
        raise RuntimeError(f"Could not find '{VNGAME_EXE}'")

    return vngame_exe_path


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

    vngame_exe = find_vngame_exe()
    logger.info("found Rising Storm 2 executable: '{exe}'", exe=vngame_exe)

    if not args.dry_run:
        logger.info("launching Rising Storm 2")
        subprocess.Popen(str(vngame_exe))


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
