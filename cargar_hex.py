# -*- coding: utf-8 -*-
"""
cargar_hex.py
Carga un .hex en la micro:bit copiandolo a la unidad MICROBIT (sin WebUSB)
y comprueba que la carga termino bien.

    python cargar_hex.py                      (usa el .hex mas reciente de esta carpeta o de Descargas)
    python cargar_hex.py compostaje-microbit.hex
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


def main():
    archivo = sys.argv[1] if len(sys.argv) > 1 else hex_reciente()
    if not archivo or not os.path.isfile(archivo):
        sys.exit("No encuentro ningun .hex. Descargalo de MakeCode ('Descargar como archivo').")
    unidad = unidad_microbit()
    if not unidad:
        sys.exit("No encuentro la unidad MICROBIT. Conecta la micro:bit por USB (cable de datos).")

    print("Copiando {} -> {}".format(archivo, unidad))
    destino = os.path.join(unidad, os.path.basename(archivo))
    with open(archivo, "rb") as origen, open(destino, "wb") as dst:
        shutil.copyfileobj(origen, dst)
        dst.flush()
        os.fsync(dst.fileno())

    print("Grabando (la luz amarilla parpadea; la unidad se desmonta y vuelve)...")
    time.sleep(3)
    for _ in range(30):
        nueva = unidad_microbit()
        if nueva:
            fallo = os.path.join(nueva, "FAIL.TXT")
            if os.path.isfile(fallo):
                with open(fallo, errors="ignore") as f:
                    sys.exit("La micro:bit rechazo el .hex:\n" + f.read())
            print("OK: programa cargado en la micro:bit (sin FAIL.TXT).")
            return
        time.sleep(1)
    print("La unidad no volvio a aparecer en 30 s; revisa la micro:bit (deberia mostrar 'T:').")


if __name__ == "__main__":
    main()
