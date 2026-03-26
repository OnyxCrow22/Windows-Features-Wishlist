@echo off
echo Building WindowsAutoSwitch…
python -m PyInstaller --noconsole --onefile --add-data "WASIcon.png;." WindowsAutoSwitch.pyw

if errorlevel 1 (
	echo.
	echo Build failure! Check the output logs for error cause!
	pause
	exit
)

echo.
echo Cleaning up build files...
rmdir /s /q build
del WindowsAutoSwitch.spec

echo.
echo Moving .exe to scripts folder...
move dist\WindowsAutoSwitch.exe .\WindowsAutoSwitch.exe
rmdir /s /q dist

echo.
echo Build complete
exit