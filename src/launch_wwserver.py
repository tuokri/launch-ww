import os
import shutil
import sys
from pathlib import Path
from typing import List

import logbook
from logbook import Logger
from logbook import StreamHandler

VNGAME_EXE = "VNGame.exe"
WW_PACKAGE = "WinterWar.u"
WW_WORKSHOP_ID = 1758494341

logger = Logger(__name__)
logbook.set_datetime_format("local")
_sh = StreamHandler(stream=sys.stdout, level="INFO")
_sh.format_string = (
    "[{record.time}] {record.level_name}: {record.channel}: "
    "{record.func_name}(): {record.message}"
)
_sh.push_application()
logger.handlers.append(_sh)


def resolve_server_root_dir() -> Path:
    fpath = Path(__file__)
    vngame_path = fpath.parent / Path(VNGAME_EXE)
    if not vngame_path.exists():
        logger.warn("'{vg} does not exist'", vg=vngame_path.absolute())
    root_dir = (fpath.parent / Path("..\\..\\")).resolve(strict=True)
    logger.info("resolved server root directory: '{rd}'", rd=root_dir)
    return root_dir


ROOT_DIR = resolve_server_root_dir()
CACHE_DIR = ROOT_DIR / Path("ROGame\\Cache")
PUBLISHED_DIR = ROOT_DIR / Path("ROGame\\Published")
WW_INT_PATH = ROOT_DIR / Path("ROGame\\Localization\\INT\\WinterWar.int")
WW_INI_PATH = ROOT_DIR / Path("ROGame\\Config\\ROGame_WinterWar.ini")


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
        logger.info("found WW cache directory: {cd}", cd=ww_cache_dir)

    return list(set(cache_dirs))


def find_ww_config_files() -> List[Path]:
    config_files = []

    if WW_INT_PATH.exists():
        config_files.append(WW_INT_PATH)
        logger.info("found config file: '{int}'", int=WW_INT_PATH)

    if WW_INI_PATH.exists():
        config_files.append(WW_INI_PATH)
        logger.info("found localization file: '{ini}'", ini=WW_INI_PATH)

    return config_files


def error_handler(function, path, exc_info):
    """Error handler for shutil.rmtree."""
    logger.error("function={f} path={p}",
                 f=function, p=path, exc_info=exc_info)


def main():
    cache_dirs = find_ww_cache_dirs()
    config_files = find_ww_config_files()
    count = len(cache_dirs)
    logger.info(
        "found {count} Winter War cache directo{p}", count=count,
        p="ry" if count == 1 else "ries")

    for cache_dir in cache_dirs:
        logger.info("removing: {cd}", cd=cache_dir.absolute())
        shutil.rmtree(
            cache_dir,
            onerror=error_handler,
        )

    for config_file in config_files:
        logger.info("removing: {cf}", cf=config_file.absolute())
        config_file.unlink()

    logger.info("done")


if __name__ == "__main__":
    main()
