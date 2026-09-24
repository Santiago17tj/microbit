@echo off
cd /d "%~dp0"
echo Instalando pyserial y pywin32...
python -m pip install -r requirements.txt
echo.
python verificar_instalacion.py
pause
