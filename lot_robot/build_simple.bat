@echo off
echo Installing PyInstaller...
pip install pyinstaller

echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building application (this may take a few minutes)...
pyinstaller --name="Поиск закупок" --onefile --windowed --add-data "config.py;." main.py

echo.
echo Done! Check the 'dist' folder for the executable.
pause

