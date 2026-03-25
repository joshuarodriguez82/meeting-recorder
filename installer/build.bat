@echo off
echo ============================================
echo  Meeting Recorder Installer Builder
echo ============================================
echo.

cd /d C:\meeting_recorder

echo [1/4] Activating virtual environment...
call .venv\Scripts\activate
if errorlevel 1 (
    echo ERROR: Could not activate venv.
    pause & exit /b 1
)

echo [2/4] Installing PyInstaller...
pip install pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: Could not install PyInstaller.
    pause & exit /b 1
)

echo [3/4] Bundling app files into installer...
python installer\bundle.py
if errorlevel 1 (
    echo ERROR: Bundling failed.
    pause & exit /b 1
)

echo [4/4] Compiling MeetingRecorderSetup.exe...
if exist "meeting_recorder.ico" (
    pyinstaller --onefile --windowed ^
        --name "MeetingRecorderSetup" ^
        --icon "meeting_recorder.ico" ^
        installer\installer_bundled.py
) else (
    pyinstaller --onefile --windowed ^
        --name "MeetingRecorderSetup" ^
        installer\installer_bundled.py
)

if errorlevel 1 (
    echo ERROR: PyInstaller failed.
    pause & exit /b 1
)

echo.
echo ============================================
echo  SUCCESS!
echo.
echo  Installer ready at:
echo  C:\meeting_recorder\dist\MeetingRecorderSetup.exe
echo.
echo  Share this one file with your team.
echo.
echo  NOTE: Windows Defender may flag the .exe
echo  as unknown. Tell your team to click
echo  "More info" then "Run anyway" when
echo  they see the SmartScreen warning.
echo ============================================
echo.
pause
