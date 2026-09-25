# -*- coding: utf-8 -*-
"""
verificar_instalacion.py
Comprueba (y corrige cuando puede) toda la cadena del proyecto de compostaje:
Python, pip, pyserial, pywin32, micro:bit en Windows, puerto COM, datos serie
y Excel.  Ademas hace una copia de seguridad del archivo del profesor.

    python verificar_instalacion.py
"""

import datetime
import glob
import importlib
import os
import shutil
import subprocess
import sys
import time

BAUDIOS = 115200
ESPERA_DATOS = 45          # segundos: la micro:bit envia cada 10 s (los programas antiguos, cada ~35 s)
ARCHIVO_PROFESOR = "DataStream_V2_var2.py"
AQUI = os.path.dirname(os.path.abspath(__file__))

resultados = []


def marcar(ok, texto, detalle=""):
    resultados.append((ok, texto))
    print("[{}] {}{}".format("OK" if ok else "X ", texto, "  -> " + detalle if detalle else ""))


def pip_instalar(paquete):
    print("   Instalando {} ...".format(paquete))
    return subprocess.call([sys.executable, "-m", "pip", "install", "--upgrade", paquete]) == 0


def asegurar_modulo(modulo, paquete):
    try:
        importlib.import_module(modulo)
        return True
    except ImportError:
        if pip_instalar(paquete):
            importlib.invalidate_caches()
            try:
                importlib.import_module(modulo)
                return True
            except ImportError:
                pass
    return False


def respaldar_profesor():
    candidatos = glob.glob(os.path.join(AQUI, ARCHIVO_PROFESOR)) + \
        glob.glob(os.path.join(os.path.expanduser("~"), "*", ARCHIVO_PROFESOR)) + \
        glob.glob(os.path.join(os.path.expanduser("~"), "*", "*", ARCHIVO_PROFESOR))
    if not candidatos:
        marcar(True, "Archivo del profesor no esta en esta carpeta (no se toca nada)")
        return
    original = candidatos[0]
    carpeta = os.path.join(AQUI, "respaldo")
    os.makedirs(carpeta, exist_ok=True)
    ya = glob.glob(os.path.join(carpeta, "DataStream_V2_var2_ORIGINAL_*.py"))
    if ya:
        marcar(True, "Copia de seguridad del archivo del profesor ya existe", ya[0])
        return
    destino = os.path.join(carpeta, "DataStream_V2_var2_ORIGINAL_{}.py".format(
        datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
    shutil.copy2(original, destino)
    marcar(True, "Copia de seguridad del archivo del profesor", destino)


def main():
    print("=== Verificacion del monitor de compostaje ===\n")

    marcar(sys.version_info >= (3, 7), "Python " + sys.version.split()[0], sys.executable)
    pip_ok = subprocess.call([sys.executable, "-m", "pip", "--version"]) == 0
    marcar(pip_ok, "pip disponible")

    marcar(asegurar_modulo("serial", "pyserial"), "pyserial instalado")
    if os.name == "nt":
        marcar(asegurar_modulo("win32com.client", "pywin32"), "pywin32 instalado")

    respaldar_profesor()

    import serial
    import serial.tools.list_ports
    sys.path.insert(0, AQUI)
    from DataStream_compost import es_microbit, parsear_linea

    print("\nPuertos serie en el sistema:")
    puertos = list(serial.tools.list_ports.comports())
    for p in puertos:
        print("   {:8s} {}  (VID:PID {}:{})".format(
            p.device, p.description,
            "%04X" % p.vid if p.vid else "----", "%04X" % p.pid if p.pid else "----"))
    mb = [p for p in puertos if es_microbit(p)]
    marcar(bool(mb), "micro:bit reconocida por Windows")
    if not mb:
        print("\n   Conecta la micro:bit con un cable USB de DATOS. Si aparece la unidad MICROBIT pero")
        print("   no un COM 'mbed Serial Port', actualiza el firmware DAPLink:")
        print("   https://microbit.org/get-started/user-guide/firmware/")
        resumen()
        return
    puerto = mb[0].device
    marcar(True, "COM correcto identificado", puerto)

    try:
        s = serial.Serial(puerto, BAUDIOS, timeout=1)
    except serial.SerialException as e:
        marcar(False, "Python abre el puerto a 115200", str(e))
        print("   Cierra MakeCode (pestana conectada por WebUSB), Tera Term, monitor serie")
        print("   u otro Python que este usando {}.".format(puerto))
        resumen()
        return
    marcar(True, "Python abre el puerto a 115200")

    print("\nEscuchando {} hasta {} s ...".format(puerto, ESPERA_DATOS))
    validas, otras = [], []
    fin = time.time() + ESPERA_DATOS
    while time.time() < fin and len(validas) < 2:
        linea = s.readline().decode("utf-8", errors="ignore").strip()
        if not linea:
            continue
        datos = parsear_linea(linea)
        print("   <- {!r}  {}".format(linea, "valida" if datos else "formato no reconocido"))
        (validas if datos else otras).append(datos or linea)
    s.close()
    marcar(bool(validas or otras), "micro:bit envia datos")
    marcar(bool(validas), "Formato TEMP:x,HUM:y,RAW:z correcto")
    if validas:
        t, h, r = validas[-1]
        marcar(t is not None, "Temperatura DS18B20 valida",
               "{} °C".format(t) if t is not None else "revisa cableado P0 y resistencia 4.7k")
        marcar(True, "Humedad y RAW visibles", "HUM={}%  RAW={}".format(h, r))
        if h in (0, 100):
            print("   Nota: HUM={}% esta en el limite de la escala. Anota el RAW ({}) para calibrar.".format(h, r))
    elif not otras:
        print("   No llego nada: carga el .hex en la micro:bit (python cargar_hex.py archivo.hex).")

    if os.name == "nt":
        try:
            import win32com.client
            xl = win32com.client.DispatchEx("Excel.Application")
            libro = xl.Workbooks.Add()
            libro.Close(False)
            version = str(xl.Version)
            xl.Quit()   # no dejar un Excel invisible abierto en segundo plano
            marcar(True, "Excel se puede abrir desde Python", "version " + version)
        except Exception as e:
            marcar(False, "Excel se puede abrir desde Python", str(e))
    resumen()


def resumen():
    fallos = [t for ok, t in resultados if not ok]
    print("\n=== {} comprobaciones, {} fallos ===".format(len(resultados), len(fallos)))
    for t in fallos:
        print("   FALLA: " + t)
    if not fallos:
        print("Todo listo. Ejecuta:  python DataStream_compost.py")


if __name__ == "__main__":
    main()
    if os.name == "nt" and not os.environ.get("PROMPT"):
        input("\nPulsa Enter para cerrar...")
