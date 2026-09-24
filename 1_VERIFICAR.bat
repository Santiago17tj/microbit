@echo off
setlocal
cd /d "%~dp0"
title Verificar monitor de compostaje

rem "python" puede ser el acceso directo de Microsoft Store, que no hace nada.
set "PY="
python -c "import sys" >nul 2>&1
if %errorlevel%==0 (
    set "PY=python"
    goto :hay_python
)
py -3 -c "import sys" >nul 2>&1
if %errorlevel%==0 (
    set "PY=py -3"
    goto :hay_python
)
echo ERROR: no se encontro Python.
echo  - Instala Python desde https://www.python.org/downloads/
echo    y marca la casilla "Add python.exe to PATH".
echo  - O desactiva "python.exe" en Configuracion ^> Aplicaciones ^>
echo    Alias de ejecucion de aplicaciones (el acceso de Microsoft Store).
goto :fin

:hay_python
echo Instalando pyserial y pywin32...
%PY% -m pip install -r requirements.txt
echo.
%PY% -u verificar_instalacion.py

:fin
echo.
pause
