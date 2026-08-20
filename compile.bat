@echo off
chcp 1251>nul

:: Обычный exe
python -m PyInstaller --noconsole --onefile --name "WinGrid" --hidden-import=win32gui --hidden-import=win32con --hidden-import=win32event --hidden-import=win32api --hidden-import=winerror "WinGrid.py"

:: CLI режим
python  -m PyInstaller --console --onefile --name "WinGridCLI" --hidden-import=win32gui --hidden-import=win32con --hidden-import=win32event --hidden-import=win32api --hidden-import=winerror "WinGrid.py"

pause