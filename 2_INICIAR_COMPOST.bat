@echo off
setlocal
cd /d "%~dp0"
title Monitor de compostaje
echo ==================================================
echo   Monitor de compostaje  micro:bit -^> Excel
echo ==================================================
echo Carpeta: %CD%
echo.

if not exist "DataStream_compost.py" (
    echo ERROR: no se encuentra DataStream_compost.py en esta carpeta.
    echo Si abriste el .bat desde dentro de un .zip, extrae primero la carpeta.
    goto :fin
)

rem --- Buscar un Python que funcione de verdad -------------------------------
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
%PY% -c "import sys; print('Python', sys.version.split()[0], '-', sys.executable)"
echo.

rem --- Instalar pyserial / pywin32 si faltan ---------------------------------
%PY% -c "import serial, win32com.client" >nul 2>&1
if not %errorlevel%==0 (
    echo Instalando pyserial y pywin32...
    %PY% -m pip install -r requirements.txt
    echo.
)

%PY% -u DataStream_compost.py %*
echo.
echo El programa termino (codigo %errorlevel%).

:fin
echo.
pause
