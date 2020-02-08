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

RS2_APP_ID = 418460
VNGAME_EXE = "VNGame.exe"
VNGAME_EXE_PATH = Path("Binaries\\Win64") / Path(VNGAME_EXE)
WW_PACKAGE = "WinterWar.u"
WW_WORKSHOP_ID = 1758494341
AUDIO_DIR = Path("CookedPC\\WwiseAudio")

# Windows constants.
CSIDL_PERSONAL = 5  # My Documents.
SHGFP_TYPE_CURRENT = 0  # Get current, not default value.

CMD_BAT_FILE = "Binaries\\Win64\\LaunchWinterWarCommand.bat"
CMD_BAT_FILE_PATH = Path(CMD_BAT_FILE)
VBS_QUIET_PROXY_FILE = "Binaries\\Win64\\LaunchWinterWarCommand.vbs"
VBS_QUIET_PROXY_FILE_PATH = Path(VBS_QUIET_PROXY_FILE)


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
WW_INT_PATH = USER_DOCS_DIR / Path("My Games\\Rising Storm 2\\ROGame\\Localization\\INT\\WinterWar.int")
WW_INI_PATH = USER_DOCS_DIR / Path("My Games\\Rising Storm 2\\ROGame\\Config\\ROGame_WinterWar.ini")

logger = Logger(__name__)
if LOGS_DIR.exists():
    logbook.set_datetime_format("local")
    _rfh_bubble = False if hasattr(sys, "frozen") else True
    _rfh = RotatingFileHandler(SCRIPT_LOG_PATH, level="INFO", bubble=_rfh_bubble)
    _rfh.format_string = (
        "[{record.time}] {record.level_name}: {record.channel}: "
        "{record.func_name}(): {record.message}"
    )
    _rfh.push_application()
    logger.handlers.append(_rfh)

# Check if running as PyInstaller frozen executable.
FROZEN = True if hasattr(sys, "frozen") else False

# No console window in frozen mode.
if FROZEN:
    logger.info("not adding stdout logging handler in frozen mode")
else:
    _sh = StreamHandler(sys.stdout, level="INFO")
    logger.info("adding logging handler: {h}", h=_sh)
    _sh.push_application()
    logger.handlers.append(_sh)


def find_ww_cache_dirs() -> List[Path]:
    """Find all cache/config directories containing Winter War files."""
    cache_dirs = [p for p in CACHE_DIR.rglob(WW_PACKAGE)]
    cache_dirs = [
        CACHE_DIR / Path(str(p).lstrip(str(CACHE_DIR)).split(os.path.sep)[0])
        for p in cache_dirs
    ]

    audio_dir = PUBLISHED_DIR / AUDIO_DIR
    if audio_dir.exists():
        logger.info("found audio directory: '{ad}'", ad=audio_dir)
        cache_dirs.append(audio_dir)

    if WW_INT_PATH.exists():
        cache_dirs.append(WW_INT_PATH)
        logger.info("found '{int}'", int=WW_INT_PATH)

    if WW_INI_PATH.exists():
        cache_dirs.append(WW_INI_PATH)
        logger.info("found '{ini}'", ini=WW_INI_PATH)

    # Defensively add WW_WORKSHOP_ID-directory even if WW_PACKAGE was not found
    # and the directory exists.
    ww_cache_dir = CACHE_DIR / Path(str(WW_WORKSHOP_ID))
    if ww_cache_dir.exists():
        cache_dirs.append(ww_cache_dir)

    return list(set(cache_dirs))


def error_handler(function, path, exc_info):
    """Error handler for shutil.rmtree."""
    logger.error("function={f} path={p}",
                 f=function, p=path, exc_info=exc_info)


def parse_args() -> Namespace:
    ap = argparse.ArgumentParser()
    ap.description = "Launcher for Winter War mod for Rising Storm 2: Vietnam."

    ap.add_argument(
        "--dry-run", dest="dry_run",
        action="store_true", default=False,
        help="run without deleting files or launching Rising Storm 2",
    )
    ap.add_argument(
        "--launch-options", nargs="*",
        help=f"launch options passed to {VNGAME_EXE}",
    )

    args, unknown = ap.parse_known_args()
    if unknown:
        logger.info("interpreting {a} unknown argument(s) as Steam "
                    "launch options: {u}", a=len(unknown), u=unknown)
        if args.launch_options:
            args.launch_options.extend(unknown)
        else:
            args.launch_options = unknown

    if args.launch_options:
        lo = args.launch_options

        # List of 1 string.
        if len(lo) == 1:
            lo = lo[0].split(" ")

        args.launch_options = lo

    return args


def main():
    args = parse_args()

    logger.info("user documents directory: '{uh}'", uh=USER_DOCS_DIR)
    cache_dirs = find_ww_cache_dirs()
    count = len(cache_dirs)
    logger.info(
        "found {count} Winter War cache directo{p}", count=count,
        p="ry" if count == 1 else "ries")

    for cache_dir in cache_dirs:
        if args.dry_run:
            logger.info("dry run, not removing: {cd}", cd=cache_dir)
        else:
            logger.info("removing: {cd}", cd=cache_dir)
            shutil.rmtree(
                cache_dir,
                onerror=error_handler,
            )

    lo = args.launch_options
    if not lo:
        lo = []

    logger.info("Steam launch options: {lo}", lo=lo)
    steam_proto_cmd = f'"steam://run/{RS2_APP_ID}//{" ".join(lo)}"'

    # START has a peculiarity involving double quotes around the first parameter.
    # If the first parameter has double quotes it uses that as the optional TITLE for the new window.
    command = ["START", '""', steam_proto_cmd]

    logger.info("launch arguments: {cmd}", cmd=command)
    command_str = " ".join(command)
    logger.info("launch command string: '{s}'", s=command_str)

    # Saving the command in a file is a workaround for subprocess.Popen failing
    # when calling the Windows START command directly due to some weird error with our
    # custom arguments and Steam protocol URL.
    with open(CMD_BAT_FILE, "w") as f:
        logger.info("writing start command to file: '{f}'", f=CMD_BAT_FILE)
        f.write(command_str)

    popen_args = {}
    # Redirecting stdout/stderr when frozen causes OSError for "invalid handle".
    if not FROZEN:
        popen_args["stdout"] = subprocess.PIPE
        popen_args["stderr"] = subprocess.PIPE

    if not CMD_BAT_FILE_PATH.exists():
        logger.error("unable to locate '{cmd_file}'", cmd_file=CMD_BAT_FILE_PATH)

    if not VBS_QUIET_PROXY_FILE_PATH.exists():
        logger.error("unable to locate '{vbs_file}'", vbs_file=VBS_QUIET_PROXY_FILE_PATH)

    if not VNGAME_EXE_PATH.exists():
        logger.error("unable to locate '{vngame_file}'", vngame_file=VNGAME_EXE_PATH)
        # Show pop-up to user explaining installation directory requirements?

    if not args.dry_run:
        logger.info("launching Rising Storm 2")
        p = subprocess.Popen([VBS_QUIET_PROXY_FILE, CMD_BAT_FILE], **popen_args)
        out, err = p.communicate()
        if out:
            logger.info("command stdout: {o}", o=out.decode("cp850"))
        if err:
            logger.error("command stderr: {o}", o=err.decode("cp850"))


if __name__ == "__main__":
    try:
        main()
    except Exception as _e:
        extra = (f"Check '{SCRIPT_LOG_PATH}' for more information."
                 if SCRIPT_LOG_PATH.exists() else "")
        ctypes.windll.user32.MessageBoxW(
            0,
            f"Error launching Winter War!\r\n"
            f"{type(_e).__name__}: {_e}\r\n"
            f"{extra}\r\n",
            "Error",
            0
        )
        # noinspection PyBroadException
        try:
            logger.exception("error running script")
        except Exception:
            pass
