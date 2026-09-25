# -*- coding: utf-8 -*-
"""
cargar_hex.py
Carga un .hex en la micro:bit copiandolo a la unidad MICROBIT (sin WebUSB)
y comprueba que la carga termino bien.

    python cargar_hex.py                      (usa el .hex mas reciente de esta carpeta o de Descargas)
    python cargar_hex.py compostaje-microbit.hex

DataStream_compost.py usa grabar_hex() para cargar el programa solo al detectar
la micro:bit.
"""

import glob
import os
import shutil
import string
import sys
import time

AQUI = os.path.dirname(os.path.abspath(__file__))


def unidad_microbit():
    if os.name == "nt":
        raices = ["{}:\\".format(l) for l in string.ascii_uppercase]
    else:
        raices = glob.glob("/media/*/*") + glob.glob("/run/media/*/*") + glob.glob("/Volumes/*")
    for r in raices:
        if os.path.isfile(os.path.join(r, "DETAILS.TXT")) and os.path.isfile(os.path.join(r, "MICROBIT.HTM")):
            return r
    return None


def hex_reciente():
    carpetas = [AQUI, os.path.join(os.path.expanduser("~"), "Downloads"),
                os.path.join(os.path.expanduser("~"), "Descargas")]
    hexes = [h for c in carpetas for h in glob.glob(os.path.join(c, "*.hex"))]
    return max(hexes, key=os.path.getmtime) if hexes else None


def grabar_hex(archivo, unidad=None, avisar=print):
    """Copia el .hex a la unidad MICROBIT y espera a que la micro:bit lo grabe.
    Devuelve (ok, mensaje).  avisar(texto) se usa para ir contando el progreso."""
    unidad = unidad or unidad_microbit()
    if not unidad:
        return False, "No encuentro la unidad MICROBIT. Conecta la micro:bit por USB (cable de datos)."

    avisar("Copiando {} -> {}".format(os.path.basename(archivo), unidad))
    destino = os.path.join(unidad, os.path.basename(archivo))
    try:
        with open(archivo, "rb") as origen, open(destino, "wb") as dst:
            shutil.copyfileobj(origen, dst)
            dst.flush()
            os.fsync(dst.fileno())
    except OSError as e:
        # La unidad se desmonta al terminar de grabar; si falla justo al cerrar
        # el archivo no es necesariamente un error: se comprueba FAIL.TXT abajo.
        if unidad_microbit():
            return False, "No se pudo copiar el .hex a {}: {}".format(unidad, e)

    avisar("Grabando (la luz amarilla parpadea; la unidad se desmonta y vuelve)...")
    # 1) esperar a que la unidad desaparezca (la micro:bit se reinicia al grabar)
    fin = time.time() + 20
    while time.time() < fin and unidad_microbit():
        time.sleep(0.5)
    # 2) esperar a que vuelva y mirar si dejo FAIL.TXT
    fin = time.time() + 30
    while time.time() < fin:
        nueva = unidad_microbit()
        if nueva:
            time.sleep(1)   # dar tiempo a que aparezcan los archivos de la unidad
            fallo = os.path.join(nueva, "FAIL.TXT")
            if os.path.isfile(fallo):
                with open(fallo, errors="ignore") as f:
                    return False, "La micro:bit rechazo el .hex: " + f.read().strip()
            return True, "Programa cargado en la micro:bit (sin FAIL.TXT)."
        time.sleep(1)
    return False, "La unidad MICROBIT no volvio a aparecer en 30 s; revisa la micro:bit."


def main():
    archivo = sys.argv[1] if len(sys.argv) > 1 else hex_reciente()
    if not archivo or not os.path.isfile(archivo):
        sys.exit("No encuentro ningun .hex. Descargalo de MakeCode ('Descargar como archivo').")
    ok, mensaje = grabar_hex(archivo)
    if not ok:
        sys.exit(mensaje)
    print("OK: " + mensaje)


if __name__ == "__main__":
    main()
