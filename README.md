# Talvisota - Winter War launcher

Launcher for Talvisota - Winter war mod for Rising Storm 2: Vietnam.

## Building distributable standalone executable (on Windows 10)

1. Install Python version 3.6 or greater. 
Newest Python version available at the time is recommended.

2. Install Windows 10 SDK.

3. Clone this repository. (Requires Git).

    `git clone git@github.com:tuokri/launch-ww.git`

4. Create virtualenv for the project.

    `python -m venv venv`
    
5. Activate the virtualenv.

    `.\venv\Scripts\activate.bat`
    
6. Install dependencies.

    `pip install -r requirements.txt`

7. In `build_utils.py` update the `_crt_dlls_path` variable to point to the correct location
for x64 UCRT DLLs. This path will vary depending on where Windows 10 SDK was installed.

8. Generate resources file.

    1. Place `ww_icon.ico` and `ww_banner.png` files in `resources` directory.
    Note that this repository does provide the image files.

    2. `pyrcc5 resources\resources.qrc > src\resources.py`

9. Run PyInstaller.

    `pyinstaller launch_winterwar.spec`
    
    The generated executable will be placed in `dist` directory.
    
10. Copy the executable to RS2 `Binaries\Win64` directory and verify that it works.
