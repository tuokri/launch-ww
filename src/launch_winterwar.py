"""
Launcher for Talvisota - Winter War community expansion
for Rising Storm 2: Vietnam.

Purges cached Talvisota - Winter War mod files in user documents
directory to avoid conflicts and runs Rising Storm 2.
"""
import argparse
import ctypes.wintypes
import math
import os
import shutil
import subprocess
import sys
import time
from argparse import Namespace
from pathlib import Path
from typing import List

import logbook
import psutil
import resources
from PyQt5 import QtCore
from PyQt5.QtCore import QObject
from PyQt5.QtCore import QThread
from PyQt5.QtGui import QIcon
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QWidget
from logbook import Logger
from logbook import RotatingFileHandler
from logbook import StreamHandler

RS2_APP_ID = 418460
WIN64_BINARIES_PATH = Path("Binaries\\Win64\\")
VNGAME_EXE = "VNGame.exe"
VNGAME_EXE_PATH = WIN64_BINARIES_PATH / Path(VNGAME_EXE)
WW_PACKAGE = "WinterWar.u"
WW_WORKSHOP_ID = 1758494341
AUDIO_DIR = Path("CookedPC\\WwiseAudio")

# Windows constants.
CSIDL_PERSONAL = 5  # My Documents.
SHGFP_TYPE_CURRENT = 0  # Get current, not default value.
S_OK = 0


def user_documents_dir() -> Path:
    """Return the user's Windows documents directory path."""
    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ret = ctypes.windll.shell32.SHGetFolderPathW(
        None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    if ret != S_OK:
        raise EnvironmentError(
            f"user documents directory not found with error code: {ret}")
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

# Check if running as PyInstaller generated frozen executable.
FROZEN = True if hasattr(sys, "frozen") else False

# No console window in frozen mode.
if FROZEN:
    logger.info("not adding stdout logging handler in frozen mode")
else:
    _sh = StreamHandler(sys.stdout, level="INFO")
    logger.info("adding logging handler: {h}", h=_sh)
    _sh.push_application()
    logger.handlers.append(_sh)


class VNGameExeNotFoundError(Exception):
    pass


def find_ww_cache_dirs() -> List[Path]:
    """Find all cache directories containing Winter War files."""
    cache_dirs = [p for p in CACHE_DIR.rglob(WW_PACKAGE)]
    cache_dirs = [
        CACHE_DIR / Path(str(p).lstrip(str(CACHE_DIR)).split(os.path.sep)[0])
        for p in cache_dirs
    ]

    audio_dir = PUBLISHED_DIR / AUDIO_DIR
    if audio_dir.exists():
        logger.info("found audio directory: '{ad}'", ad=audio_dir)
        cache_dirs.append(audio_dir)

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


def parse_args() -> Namespace:
    ap = argparse.ArgumentParser()
    ap.description = "Launcher for Winter War mod for Rising Storm 2: Vietnam."

    ap.add_argument(
        "--dry-run", dest="dry_run",
        action="store_true", default=False,
        help="run without deleting files or launching Rising Storm 2",
    )

    return ap.parse_args()


def resolve_binary_paths():
    """
    TODO: Refactor (remove) global usage.
    """
    global VNGAME_EXE_PATH

    if not VNGAME_EXE_PATH.exists():
        if Path(VNGAME_EXE).exists():
            VNGAME_EXE_PATH = Path(VNGAME_EXE)
        else:
            raise VNGameExeNotFoundError(
                f"{VNGAME_EXE} not found, please make sure Winter War mod "
                f"and Rising Storm 2: Vietnam are installed on the same drive.")
    logger.info("VNGame.exe found in '{p}'",
                p=VNGAME_EXE_PATH.absolute())


def main():
    args = parse_args()

    logger.info("user documents directory: '{uh}'", uh=USER_DOCS_DIR)
    cache_dirs = find_ww_cache_dirs()
    config_files = find_ww_config_files()
    count = len(cache_dirs)
    logger.info(
        "found {count} Winter War cache directo{p}", count=count,
        p="ry" if count == 1 else "ries")

    for cache_dir in cache_dirs:
        if args.dry_run:
            logger.info("dry run, not removing: {cd}", cd=cache_dir.absolute())
        else:
            logger.info("removing: {cd}", cd=cache_dir.absolute())
            shutil.rmtree(
                cache_dir,
                onerror=error_handler,
            )

    for config_file in config_files:
        if args.dry_run:
            logger.info("dry run, not removing: {cf}", cf=config_file.absolute())
        else:
            logger.info("removing: {cf}", cf=config_file.absolute())
            config_file.unlink()

    steam_proto_cmd = f"steam://run/{RS2_APP_ID}"

    resolve_binary_paths()

    popen_kwargs = {"shell": True}
    # Redirecting stdout/stderr when frozen causes OSError for "invalid handle".
    if not FROZEN:
        popen_kwargs["stdout"] = subprocess.PIPE
        popen_kwargs["stderr"] = subprocess.PIPE

    popen_args = [steam_proto_cmd]
    if not args.dry_run:
        logger.info("launching Rising Storm 2, Popen args={a} kwargs={kw}",
                    a=popen_args, kw=popen_kwargs)
        p = subprocess.Popen(popen_args, **popen_kwargs)
        out, err = p.communicate()
        if out:
            logger.info("command stdout: {o}", o=out.decode("cp850"))
        if err:
            logger.error("command stderr: {o}", o=err.decode("cp850"))


class VNGameProcessListener(QObject):
    vngame_exe_finished = QtCore.pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)

    @staticmethod
    def is_alive() -> bool:
        return "VNGame.exe" in [p.name() for p in psutil.process_iter()]

    @QtCore.pyqtSlot()
    def listen(self):
        logger.info("{c}: waiting for VNGame.exe to start",
                    c=self.__class__.__name__)
        while not self.is_alive():
            time.sleep(1)
        logger.info("{c}: VNGame.exe started",
                    c=self.__class__.__name__)

        logger.info("{c}: waiting for VNGame.exe to finish",
                    c=self.__class__.__name__)
        while self.is_alive():
            time.sleep(1)
        logger.info("{c}: VNGame.exe finished",
                    c=self.__class__.__name__)

        self.vngame_exe_finished.emit(0)


class Gui(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.layout = QVBoxLayout()

        self.img_label = QLabel()
        self.image = QPixmap(":/ww_banner.png")
        size = self.image.size()
        width = math.floor(size.width() * 0.33)
        height = math.floor(size.height() * 0.33)
        self.image = self.image.scaled(width, height, QtCore.Qt.KeepAspectRatio)
        self.img_label.setPixmap(self.image)

        self.text_label = QLabel(
            "Talvisota - Winter War launcher. This launcher is ran minimized "
            "to track Steam play time hours. You may close this window or leave it open."
        )
        self.text_label.setWordWrap(True)
        self.text_label.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)

        self.layout.addWidget(self.img_label)
        self.layout.addWidget(self.text_label)

        self.setLayout(self.layout)
        self.setWindowFlags(QtCore.Qt.MSWindowsFixedSizeDialogHint)
        self.setWindowTitle("Talvisota - Winter War")
        self.setWindowIcon(QIcon(":/ww_icon.ico"))
        self.setWindowState(QtCore.Qt.WindowMinimized)

    def warn(self, title: str, msg: str):
        QMessageBox.warning(self, title, msg)


if __name__ == "__main__":
    logger.info("resources loaded: {r}", r=resources)

    _app = QApplication(sys.argv)
    _gui = Gui()
    _thread = QThread()
    _vpl = VNGameProcessListener()

    _vpl.vngame_exe_finished.connect(_app.exit)
    _thread.started.connect(_vpl.listen)

    _vpl.moveToThread(_thread)
    _thread.start()

    _gui.show()
    logger.info("GUI init done")

    try:
        logger.info("cwd='{cwd}'", cwd=os.getcwd())
        main()
    except Exception as _e:
        _extra = (f"Check '{SCRIPT_LOG_PATH}' for more information."
                  if SCRIPT_LOG_PATH.exists() else "")
        _msg = (f"Error launching Winter War!\r\n"
                f"{type(_e).__name__}: {_e}\r\n"
                f"{_extra}\r\n")
        _gui.warn("Error", _msg)
        # noinspection PyBroadException
        try:
            logger.exception("error running script")
        except Exception:
            pass

    sys.exit(_app.exec_())
