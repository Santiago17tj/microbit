# -*- coding: utf-8 -*-
"""
DataStream_compost.py
Adaptacion de DataStream_V2_var2.py (programa base del profesor) para el
monitor de compostaje.  El archivo original del profesor NO se modifica.

    DS18B20 -> micro:bit P0      (temperatura)
    FC-28   -> micro:bit P1      (humedad del suelo, analogica)
    micro:bit -> USB/Serial 115200 -> Python -> Excel

La micro:bit envia lineas con este formato exacto:
    TEMP:33,HUM:58,RAW:430

Uso:
    python DataStream_compost.py                 (detecta el COM automaticamente)
    python DataStream_compost.py --puerto COM5   (fuerza un COM concreto)
    python DataStream_compost.py --simular       (prueba Excel sin micro:bit)
    python DataStream_compost.py --sin-excel     (solo consola + CSV)
    python DataStream_compost.py --sin-panel     (no abrir el panel web)
    python DataStream_compost.py --sin-cargar    (no revisar ni cargar el programa)
    python DataStream_compost.py --hex compostaje-microbit.hex   (usar otro programa)

Al detectar la micro:bit le pregunta que programa tiene ("ID?").  Si no es
HEX_A_CARGAR (ver CONFIGURACION) se lo carga y comprueba que responde bien.
Si la micro:bit no esta conectada, la espera.

Ademas de Excel abre un panel web en tiempo real en http://localhost:8765
(ver panel_web.py y panel/index.html).
"""

import argparse
import csv
import datetime
import glob
import os
import random
import re
import shutil
import subprocess
import sys
import threading
import time

from cargar_hex import grabar_hex, unidad_microbit

print("Monitor de compostaje: iniciando...", flush=True)
try:
    import serial
    import serial.tools.list_ports
except ImportError:
    print("Falta la libreria pyserial. Ejecuta 1_VERIFICAR.bat (o: python -m pip install pyserial).")
    sys.exit(1)

# --------------------------------------------------------------------------
# CONFIGURACION
# --------------------------------------------------------------------------
# Alternativa manual: si la deteccion automatica falla, escribe aqui el COM
# que aparece en el Administrador de dispositivos, por ejemplo "COM5".
PUERTO_MANUAL = None
BAUDIOS = 115200
GUARDAR_CADA = 10          # guarda el .xlsx cada N filas (y al salir)
CARPETA_DATOS = "datos_compost"
PUERTO_WEB = 8765          # panel web: http://localhost:8765

# Programa que debe tener la micro:bit.  Al detectarla se comprueba y, si tiene
# otro, se le carga este.  PROVISIONAL mientras se reemplaza el DS18B20 (usa el
# sensor interno del chip).  Con el DS18B20 nuevo: "compostaje-microbit.hex".
HEX_A_CARGAR = "PROVISIONAL-temp-interna.hex"
FUENTES_HEX = {            # .hex -> carpeta con su main.ts (para compilarlo e identificarlo)
    "compostaje-microbit.hex": ".",
    "PROVISIONAL-temp-interna.hex": "provisional_temp_interna",
    "diagnostico-ds18b20.hex": "diagnostico_ds18b20",
}
AQUI = os.path.dirname(os.path.abspath(__file__))

MICROBIT_VID = 0x0D28      # ARM / DAPLink (interfaz USB de la micro:bit)

# Constantes de Excel (equivalentes a las de la libreria de tipos)
XL_SRC_RANGE = 1
XL_YES = 1
XL_LINE_MARKERS = 65
XL_VALUE = 2
XL_CATEGORY = 1
XL_PRIMARY = 1
XL_SECONDARY = 2
XL_OPEN_XML_WORKBOOK = 51
XL_LEGEND_BOTTOM = -4107
XL_NORMAL = -4143

ANCHO_GRAFICA, ALTO_GRAFICA = 470, 300     # grafica de Excel, en puntos (cabe en media pantalla)

ENCABEZADOS = ["Hora", "Temperatura (°C)", "Humedad (%)", "Humedad RAW"]

NUMERO = r"(-?(?:\d+(?:\.\d+)?|Infinity)|NaN)"
PATRON = re.compile(r"^TEMP:" + NUMERO + r",HUM:" + NUMERO + r",RAW:(\d+)$")

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

panel = None               # PanelWeb, si el panel web esta activo


def avisar(texto, nivel="info"):
    """Imprime en consola y lo muestra en el registro del panel web."""
    print(texto)
    if panel is not None:
        panel.evento(" ".join(texto.split()), nivel)


def fase_compost(temp):
    """Mismos umbrales que main.ts de la micro:bit."""
    if temp is None:
        return "ERROR SENSOR"
    if temp > 65:
        return "PELIGRO CALOR"
    if temp > 60:
        return "TERMOFILA ALTA"
    if temp >= 40:
        return "TERMOFILA"
    return "MESOFILA"


# --------------------------------------------------------------------------
# SERIAL
# --------------------------------------------------------------------------
def es_microbit(p):
    texto = " ".join(str(x) for x in (p.description, p.manufacturer, p.product) if x).lower()
    return p.vid == MICROBIT_VID or "micro:bit" in texto or "mbed" in texto or "daplink" in texto


def detectar_puerto(silencioso=False):
    if not silencioso:
        print("Buscando la micro:bit en los puertos USB...")
    puertos = list(serial.tools.list_ports.comports())
    candidatos = [p for p in puertos if es_microbit(p)]
    if candidatos:
        p = candidatos[0]
        print("micro:bit detectada en {} ({})".format(p.device, p.description))
        return p.device
    if silencioso:
        return None
    print("No se encontro ninguna micro:bit. Puertos serie disponibles:")
    for p in puertos:
        print("   {:8s} {}".format(p.device, p.description))
    if not puertos:
        print("   (ninguno) - revisa el cable USB (debe ser de datos, no solo de carga)")
    print("Si tu micro:bit es uno de ellos, ejecuta:  python DataStream_compost.py --puerto COMx")
    print("o escribe el COM en PUERTO_MANUAL al principio de este archivo.")
    return None


def abrir_serial(puerto):
    s = serial.Serial(puerto, BAUDIOS, timeout=1)
    s.reset_input_buffer()
    return s


def abrir_con_reintentos(puerto, segundos=10):
    """Tras grabar o reconectar, Windows tarda un poco en soltar el COM."""
    fin = time.time() + segundos
    while True:
        try:
            return abrir_serial(puerto)
        except serial.SerialException:
            if time.time() > fin:
                raise
            time.sleep(1)


def esperar_puerto(args):
    """Devuelve el COM de la micro:bit; si no esta conectada, la espera."""
    fijo = args.puerto or PUERTO_MANUAL
    if fijo:
        return fijo
    puerto = detectar_puerto()
    if puerto:
        return puerto
    avisar("Esperando la micro:bit: conectala por USB (cable de datos)...", "aviso")
    if panel is not None:
        panel.conexion("esperando")
    inicio = ultimo = time.time()
    while True:
        time.sleep(1)
        puerto = detectar_puerto(silencioso=True)
        if puerto:
            return puerto
        if time.time() - ultimo >= 15:
            ultimo = time.time()
            print("   ... sigo esperando la micro:bit ({} s). Ctrl+C para salir.".format(int(ultimo - inicio)))


# --------------------------------------------------------------------------
# COMPROBAR Y CARGAR EL PROGRAMA DE LA MICRO:BIT
# --------------------------------------------------------------------------
def id_esperado(nombre_hex):
    """Lee PROGRAMA_ID del main.ts del programa (lo que la micro:bit responde a "ID?")."""
    carpeta = FUENTES_HEX.get(os.path.basename(nombre_hex))
    if carpeta is None:
        return None
    try:
        with open(os.path.join(AQUI, carpeta, "main.ts"), encoding="utf-8") as f:
            m = re.search(r'const\s+PROGRAMA_ID\s*=\s*"([^"]+)"', f.read())
    except OSError:
        return None
    return m.group(1) if m else None


def identificar(s, intentos=3, espera=2.0):
    """Pregunta "ID?" a la micro:bit.  Devuelve (id o None, lineas de datos recibidas
    mientras tanto, para no perderlas)."""
    sobrantes = []
    for _ in range(intentos):
        try:
            s.write(b"ID?\n")
            fin = time.time() + espera
            while time.time() < fin:
                linea = s.readline().decode("utf-8", errors="ignore").strip()
                if linea.startswith("ID:"):
                    return linea[3:].strip(), sobrantes
                if linea:
                    sobrantes.append(linea)
        except serial.SerialException:
            break
    return None, sobrantes


def pedir_dato(s):
    """Pide una lectura ya mismo ("DATO?") para no esperar un ciclo entero.
    Los programas antiguos de la micro:bit simplemente lo ignoran."""
    try:
        s.write(b"DATO?\n")
    except serial.SerialException:
        pass


def compilar_hex(nombre):
    """Compila el .hex con MakeCode (npx makecode build, necesita Node).
    Devuelve la ruta del .hex o None."""
    carpeta = FUENTES_HEX.get(nombre)
    npx = shutil.which("npx")
    if carpeta is None or npx is None:
        return None
    carpeta = os.path.join(AQUI, carpeta)
    avisar("Compilando {} con MakeCode (la primera vez tarda un poco)...".format(nombre))
    try:
        r = subprocess.run([npx, "--yes", "makecode", "build"], cwd=carpeta,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           text=True, errors="replace", timeout=600)
    except (OSError, subprocess.TimeoutExpired) as e:
        avisar("No se pudo compilar {} ({}).".format(nombre, e), "aviso")
        return None
    compilado = os.path.join(carpeta, "built", "binary.hex")
    if r.returncode != 0 or not os.path.isfile(compilado):
        ultimas = "\n".join(r.stdout.strip().splitlines()[-5:])
        avisar("MakeCode no pudo compilar {}:\n{}".format(nombre, ultimas), "aviso")
        return None
    destino = os.path.join(AQUI, nombre)
    shutil.copyfile(compilado, destino)
    avisar("Compilado {}.".format(nombre), "ok")
    return destino


def fuentes_cambiadas(nombre, ruta_hex):
    """True si main.ts / pxt.json son mas nuevos que el .hex (p. ej. tras calibrar el FC-28)."""
    carpeta = FUENTES_HEX.get(nombre)
    if carpeta is None:
        return False
    t_hex = os.path.getmtime(ruta_hex) + 2   # margen: git crea los archivos casi a la vez
    fuentes = [os.path.join(AQUI, carpeta, f) for f in ("main.ts", "pxt.json")]
    return any(os.path.isfile(f) and os.path.getmtime(f) > t_hex for f in fuentes)


def buscar_hex(nombre):
    """Devuelve (ruta del .hex o None, recompilado).  Busca el de esta carpeta
    (recompilandolo si cambio main.ts); si no esta, lo compila con MakeCode; y
    si no hay Node, busca el descargado de GitHub en Descargas."""
    if os.path.isfile(nombre) and os.path.basename(nombre) not in FUENTES_HEX:
        return os.path.abspath(nombre), False     # un .hex cualquiera indicado con --hex
    nombre = os.path.basename(nombre)
    ruta = os.path.join(AQUI, nombre)
    if os.path.isfile(ruta):
        if fuentes_cambiadas(nombre, ruta):
            avisar("El main.ts cambio despues de crear {}; se recompila.".format(nombre))
            nuevo = compilar_hex(nombre)
            return (nuevo, True) if nuevo else (ruta, False)
        return ruta, False
    compilado = compilar_hex(nombre)
    if compilado:
        return compilado, True
    descargas = []
    for d in ("Downloads", "Descargas"):
        base = os.path.join(os.path.expanduser("~"), d)
        descargas += glob.glob(os.path.join(base, nombre)) + glob.glob(os.path.join(base, "*", nombre))
    return (max(descargas, key=os.path.getmtime), False) if descargas else (None, False)


def esperar_unidad(segundos=15):
    fin = time.time() + segundos
    while time.time() < fin:
        unidad = unidad_microbit()
        if unidad:
            return unidad
        time.sleep(1)
    return None


def preparar_microbit(args, puerto):
    """Abre el serial y se asegura de que la micro:bit tenga el programa correcto.
    Devuelve (serial abierto, puerto, lineas de datos ya recibidas)."""
    s = abrir_con_reintentos(puerto)
    if args.sin_cargar:
        return s, puerto, []

    nombre = os.path.basename(args.hex)
    esperado = id_esperado(nombre)
    ruta, recompilado = buscar_hex(args.hex)
    sobrantes = []
    if panel is not None:
        panel.conexion("verificando", puerto)
    if esperado and not recompilado:
        print("Comprobando que programa tiene la micro:bit...")
        actual, sobrantes = identificar(s)
        if actual == esperado:
            avisar("La micro:bit ya tiene el programa correcto ({}). No hace falta cargarlo.".format(actual), "ok")
            return s, puerto, sobrantes
        avisar("La micro:bit tiene {}; se esperaba {}. Se carga el programa.".format(
            actual or "otro programa (o una version anterior)", esperado), "aviso")
    if not ruta:
        avisar("No encuentro {} ni puedo compilarlo (falta Node). Descargalo del artefacto\n"
               "   'compostaje-microbit-hex' en GitHub Actions y dejalo en esta carpeta.\n"
               "   Se sigue con el programa que ya tenga la micro:bit.".format(nombre), "aviso")
        return s, puerto, sobrantes

    unidad = esperar_unidad()
    if not unidad:
        avisar("No aparece la unidad MICROBIT, asi que no se puede cargar el programa.\n"
               "   Se sigue con el que ya tenga la micro:bit.", "aviso")
        return s, puerto, sobrantes
    s.close()
    if panel is not None:
        panel.conexion("cargando", puerto)
    avisar("Cargando {} en la micro:bit ({})...".format(os.path.basename(ruta), unidad))
    ok, mensaje = grabar_hex(ruta, unidad, avisar=print)
    avisar(mensaje, "ok" if ok else "error")

    puerto = esperar_puerto(args)
    s = abrir_con_reintentos(puerto)
    if ok and esperado:
        time.sleep(1)    # que arranque el programa recien cargado
        actual, sobrantes = identificar(s)
        if actual == esperado:
            avisar("Verificado: la micro:bit responde como {}.".format(actual), "ok")
        else:
            avisar("Se cargo el programa pero la micro:bit no respondio a la identificacion.", "aviso")
    return s, puerto, sobrantes


def a_numero(texto):
    """Convierte el texto recibido a numero. Devuelve None si el sensor dio error."""
    if texto in ("NaN", "Infinity", "-Infinity"):
        return None
    valor = float(texto)
    return int(valor) if valor.is_integer() else valor


def parsear_linea(linea):
    """'TEMP:33,HUM:58,RAW:430' -> (33, 58, 430).  Otra cosa -> None."""
    m = PATRON.match(linea.strip())
    if not m:
        return None
    temp = a_numero(m.group(1))
    hum = a_numero(m.group(2))
    raw = int(m.group(3))
    # El DS18B20 mide de -55 a 125 °C; fuera de eso es un error de lectura
    if temp is not None and not -55 <= temp <= 125:
        temp = None
    return temp, hum, raw


def generador_simulado():
    """Lineas falsas con el mismo formato que la micro:bit (modo --simular)."""
    temp, raw = 33.0, 430
    while True:
        temp = min(70, max(20, temp + random.uniform(-0.8, 1.2)))
        raw = min(1023, max(150, raw + random.randint(-25, 25)))
        hum = min(100, max(0, round((raw - 1023) * (100 - 0) / (250 - 1023))))
        time.sleep(2)
        yield "TEMP:{},HUM:{},RAW:{}".format(round(temp), hum, raw)


# --------------------------------------------------------------------------
# EXCEL
# --------------------------------------------------------------------------
# Excel rechaza las ordenes mientras esta ocupado (una celda en edicion, un aviso
# abierto, la ventana de activacion...).  Esos errores se reintentan.
EXCEL_OCUPADO = (-2147418111, -2147417846)   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def excel_ocupado(e):
    codigo = getattr(e, "hresult", None)
    if codigo is None and getattr(e, "args", None):
        codigo = e.args[0]
    return codigo in EXCEL_OCUPADO


def con_reintentos(funcion, segundos=10):
    """Llama a funcion(); si Excel esta ocupado lo reintenta durante unos segundos."""
    fin = time.time() + segundos
    while True:
        try:
            return funcion()
        except Exception as e:
            if not excel_ocupado(e) or time.time() > fin:
                raise
            time.sleep(0.3)


class HojaExcel:
    def __init__(self, ruta_xlsx):
        import win32com.client
        self.ruta = os.path.abspath(ruta_xlsx)
        # DispatchEx abre un Excel nuevo en vez de engancharse a uno oculto o bloqueado
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.fila = 1            # ultima fila ya escrita en Excel
        self.pendientes = []     # lecturas que Excel aun no ha aceptado
        self.tabla = None
        self.grafica = None
        self.ejes_listos = False
        self.listo = False       # hoja creada (Excel puede estar ocupado al arrancar)
        self.al_estar_listo = None
        try:
            self.asegurar_listo(segundos=20)
        except Exception as e:
            if not excel_ocupado(e):
                raise
            # Excel sigue ocupado (aviso de activacion...): se reintenta con cada lectura

    def asegurar_listo(self, segundos=10):
        if self.listo:
            return
        con_reintentos(self._preparar, segundos)
        self.listo = True
        if self.al_estar_listo is not None:
            try:
                self.al_estar_listo()
            except Exception:
                pass

    def _preparar(self):
        self.excel.Visible = True
        if not hasattr(self, "libro"):
            self.libro = self.excel.Workbooks.Add()
        self.hoja = self.libro.Worksheets(1)
        self.hoja.Name = "Compostaje"
        self.hoja.Range("A1:D1").Value = tuple(ENCABEZADOS)
        self.hoja.Range("A1:D1").Font.Bold = True
        self.hoja.Columns("A").NumberFormat = "hh:mm:ss"
        self.hoja.Columns("A").ColumnWidth = 10
        self.hoja.Columns("B:D").ColumnWidth = 12
        self.hoja.Columns("E").ColumnWidth = 2          # separacion con la grafica
        self.hoja.Range("A1:D1").WrapText = True       # encabezados en dos lineas
        self.hoja.Rows(1).AutoFit()

    def agregar(self, hora, temp, hum, raw):
        """Anade una lectura.  Si Excel sigue ocupado tras reintentar, lanza la
        excepcion pero la lectura queda pendiente y se escribe con la siguiente."""
        segundos = hora.hour * 3600 + hora.minute * 60 + hora.second
        self.pendientes.append((segundos / 86400.0,      # hora como valor de Excel
                                "" if temp is None else temp,
                                "" if hum is None else hum,
                                raw))
        con_reintentos(self.volcar)

    def volcar(self):
        """Escribe las lecturas pendientes y actualiza tabla y grafica."""
        self.asegurar_listo()
        if self.pendientes:
            desde, hasta = self.fila + 1, self.fila + len(self.pendientes)
            # una sola orden para todas las filas: o entran todas o ninguna
            self.hoja.Range("A{}:D{}".format(desde, hasta)).Value = tuple(self.pendientes)
            self.fila, self.pendientes = hasta, []
        if self.fila < 2:
            return

        rango = self.hoja.Range("A1:D{}".format(self.fila))
        if self.tabla is None:
            if self.hoja.ListObjects.Count > 0:      # creada en un intento anterior
                self.tabla = self.hoja.ListObjects(1)
            else:
                self.tabla = self.hoja.ListObjects.Add(SourceType=XL_SRC_RANGE, Source=rango,
                                                       XlListObjectHasHeaders=XL_YES)
            self.tabla.Name = "TablaCompost"
            self.tabla.TableStyle = "TableStyleMedium2"
            self.hoja.Range("A1:D1").WrapText = True   # la tabla lo quita: encabezados en dos lineas
            self.hoja.Rows(1).RowHeight = 30
        else:
            self.tabla.Resize(rango)
        if self.grafica is None:
            self._crear_grafica()
        self._actualizar_grafica()
        try:
            self.excel.ActiveWindow.ScrollRow = max(1, self.fila - 20)   # mantiene visible la ultima fila
        except Exception:
            pass   # el usuario esta usando otra ventana de Excel; no pasa nada

    def _crear_grafica(self):
        while self.hoja.ChartObjects().Count > 0:   # restos de un intento anterior
            self.hoja.ChartObjects(1).Delete()
        obj = self.hoja.ChartObjects().Add(self.hoja.Range("F1").Left, self.hoja.Range("F1").Top, ANCHO_GRAFICA, ALTO_GRAFICA)
        g = obj.Chart
        g.ChartType = XL_LINE_MARKERS
        while g.SeriesCollection().Count > 0:
            g.SeriesCollection(1).Delete()
        s_temp = g.SeriesCollection().NewSeries()
        s_temp.Name = "=Compostaje!$B$1"
        s_hum = g.SeriesCollection().NewSeries()
        s_hum.Name = "=Compostaje!$C$1"
        s_hum.AxisGroup = XL_SECONDARY
        g.HasTitle = True
        g.ChartTitle.Text = "Compostaje: temperatura y humedad en tiempo real"
        g.HasLegend = True
        g.Legend.Position = XL_LEGEND_BOTTOM
        self.s_temp, self.s_hum = s_temp, s_hum
        self.grafica = g

    def _actualizar_grafica(self):
        f = self.fila
        horas = self.hoja.Range("A2:A{}".format(f))
        self.s_temp.XValues = horas
        self.s_temp.Values = self.hoja.Range("B2:B{}".format(f))
        self.s_hum.XValues = horas
        self.s_hum.Values = self.hoja.Range("C2:C{}".format(f))
        if not self.ejes_listos:   # los ejes existen cuando ya hay datos
            g = self.grafica
            ejes = ((XL_CATEGORY, XL_PRIMARY, "Hora"),
                    (XL_VALUE, XL_PRIMARY, "Temperatura (°C)"),
                    (XL_VALUE, XL_SECONDARY, "Humedad (%)"))
            for tipo, grupo, titulo in ejes:
                eje = g.Axes(tipo, grupo)
                eje.HasTitle = True
                eje.AxisTitle.Text = titulo
            g.Axes(XL_VALUE, XL_SECONDARY).MinimumScale = 0
            g.Axes(XL_VALUE, XL_SECONDARY).MaximumScale = 100
            self.ejes_listos = True

    def ventana(self):
        """Hwnd de la ventana de Excel, ya sin maximizar (para colocarla)."""
        def _ventana():
            self.excel.WindowState = XL_NORMAL
            return self.excel.Hwnd
        return con_reintentos(_ventana)

    def ajustar_zoom(self):
        """Zoom para que la tabla y la grafica quepan a lo ancho de la ventana."""
        def _ajustar():
            v = self.excel.ActiveWindow
            # hasta el borde de la grafica + numeros de fila y barra de desplazamiento
            necesario = self.hoja.Range("F1").Left + ANCHO_GRAFICA + 55
            v.Zoom = max(40, min(100, int(v.UsableWidth / necesario * 100)))
        con_reintentos(_ajustar)

    def guardar(self):
        def _guardar():
            self.excel.DisplayAlerts = False
            try:
                self.libro.SaveAs(self.ruta, XL_OPEN_XML_WORKBOOK)
            finally:
                self.excel.DisplayAlerts = True
        con_reintentos(_guardar)


# --------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# --------------------------------------------------------------------------
def iniciar_panel(args, ruta_csv):
    global panel
    try:
        from panel_web import PanelWeb
        p = PanelWeb(args.puerto_web, "simulacion" if args.simular else "real", ruta_csv)
        url = p.iniciar(abrir_navegador=not args.no_abrir_navegador)
    except OSError as e:
        print("No se pudo iniciar el panel web en el puerto {} ({}).".format(args.puerto_web, e))
        print("   Quiza ya hay otro DataStream_compost.py abierto. Se continua sin panel.")
        return
    except Exception as e:
        print("No se pudo iniciar el panel web ({}). Se continua sin panel.".format(e))
        return
    panel = p
    print("Panel web en vivo: {}".format(url))


def acomodar_ventanas(args, excel):
    """Panel web a la izquierda y Excel a la derecha (ver ventanas.py)."""
    try:
        import ventanas
    except Exception:
        return
    panel_visible = panel is not None and not args.no_abrir_navegador
    if excel is None:
        if panel_visible:
            ventanas.lado_a_lado_en_segundo_plano()      # sin Excel: panel a pantalla completa
        return
    if not excel.listo:
        # Excel aun ocupado: se colocan las ventanas en cuanto responda
        excel.al_estar_listo = lambda: acomodar_ventanas(args, excel)
        return
    try:
        hwnd = excel.ventana()
        ventanas.lado_a_lado(hwnd, espera_panel=15 if panel_visible else 0)
        excel.ajustar_zoom()
    except Exception as e:
        print("   (no se pudieron colocar las ventanas: {})".format(e))


def main():
    ap = argparse.ArgumentParser(description="Registro del compostaje micro:bit -> Excel")
    ap.add_argument("--puerto", help="Puerto COM manual, p. ej. COM5")
    ap.add_argument("--simular", action="store_true", help="Genera datos falsos (sin micro:bit)")
    ap.add_argument("--sin-excel", action="store_true", help="No abrir Excel (solo consola y CSV)")
    ap.add_argument("--max-lecturas", type=int, default=0, help="Termina tras N lecturas (0 = infinito)")
    ap.add_argument("--sin-panel", action="store_true", help="No iniciar el panel web")
    ap.add_argument("--puerto-web", type=int, default=PUERTO_WEB, help="Puerto del panel web (por defecto 8765)")
    ap.add_argument("--no-abrir-navegador", action="store_true", help="Inicia el panel pero no abre el navegador")
    ap.add_argument("--sin-acomodar", action="store_true", help="No colocar panel y Excel lado a lado")
    ap.add_argument("--sin-cargar", action="store_true", help="No comprobar ni cargar el programa de la micro:bit")
    ap.add_argument("--hex", default=HEX_A_CARGAR, help="Programa que debe tener la micro:bit (por defecto {})".format(HEX_A_CARGAR))
    args = ap.parse_args()

    os.makedirs(CARPETA_DATOS, exist_ok=True)
    marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_csv = os.path.join(CARPETA_DATOS, "compost_{}.csv".format(marca))
    ruta_xlsx = os.path.join(CARPETA_DATOS, "compost_{}.xlsx".format(marca))

    if not args.sin_panel:
        iniciar_panel(args, ruta_csv)

    # ---- Excel ----
    excel = None
    if not args.sin_excel:
        print("Abriendo Excel (puede tardar unos segundos)...")
        aviso_lento = threading.Timer(20, lambda: print(
            "   Excel esta tardando. Si no se abre: cierra todos los Excel (tambien en el\n"
            "   Administrador de tareas, EXCEL.EXE) o ejecuta con --sin-excel."))
        aviso_lento.daemon = True
        aviso_lento.start()
        try:
            excel = HojaExcel(ruta_xlsx)
            if excel.listo:
                print("Excel abierto. Hoja 'Compostaje' creada.")
            else:
                avisar("Excel esta ocupado (un aviso abierto, p. ej. 'Error de activacion'?).\n"
                       "   Cierra el aviso en Excel: la hoja se creara sola con la proxima lectura.\n"
                       "   Mientras tanto los datos se guardan en el CSV.", "aviso")
        except Exception as e:
            print("No se pudo abrir Excel ({}). Se continua guardando en CSV.".format(e))
        finally:
            aviso_lento.cancel()

    if not args.sin_acomodar:
        acomodar_ventanas(args, excel)

    # ---- origen de datos ----
    microbit = None
    if args.simular:
        avisar("MODO SIMULACION: no se usa la micro:bit.")
        if panel is not None:
            panel.conexion("conectado", "datos simulados")
        lineas = generador_simulado()
    else:
        puerto = esperar_puerto(args)
        try:
            microbit, puerto, pendientes = preparar_microbit(args, puerto)
        except serial.SerialException as e:
            avisar("No se pudo abrir {}: {}".format(puerto, e), "error")
            print("Cierra MakeCode, el monitor serie u otro programa que este usando el puerto.")
            if panel is not None:
                panel.detener()
            sys.exit(1)
        avisar("Puerto {} abierto a {} baudios.".format(puerto, BAUDIOS), "ok")
        pedir_dato(microbit)
        if panel is not None:
            panel.conexion("esperando", puerto)

        def lineas_serial():
            nonlocal microbit
            ultimo_dato = ultimo_latido = time.time()
            avisado = False
            primer_dato = True
            for linea in pendientes:        # recibidas mientras se identificaba la micro:bit
                yield linea
            while True:
                ahora = time.time()
                if primer_dato and ahora - ultimo_latido >= 10:
                    ultimo_latido = ahora
                    print("   ... esperando la primera linea de la micro:bit ({} s)".format(int(ahora - ultimo_dato)))
                if not avisado and ahora - ultimo_dato > 30:
                    avisado = True
                    avisar("AVISO: 30 s sin recibir nada de la micro:bit. Si su pantalla no muestra\n"
                           "       'T:', 'MESOFILA'..., cierra y vuelve a abrir este programa para\n"
                           "       que le cargue {}.\n"
                           "       Este programa sigue esperando y se reconecta solo.".format(os.path.basename(args.hex)), "aviso")
                try:
                    dato = microbit.readline()
                except serial.SerialException:
                    avisar("Se perdio la conexion con la micro:bit. Reintentando...", "error")
                    if panel is not None:
                        panel.conexion("reconectando")
                    microbit.close()
                    while True:
                        time.sleep(2)
                        try:
                            nuevo = args.puerto or PUERTO_MANUAL or detectar_puerto(silencioso=True) or puerto
                            microbit = abrir_serial(nuevo)
                            pedir_dato(microbit)
                            avisar("Reconectado.", "ok")
                            if panel is not None:
                                panel.conexion("esperando", nuevo)
                            primer_dato = True
                            break
                        except serial.SerialException:
                            pass
                    continue
                if dato:
                    if primer_dato:
                        primer_dato = False
                        avisar("Conexion OK: llegan datos de la micro:bit.", "ok")
                        if panel is not None:
                            panel.conexion("conectado")
                    ultimo_dato, avisado = time.time(), False
                    yield dato.decode("utf-8", errors="ignore")
        lineas = lineas_serial()

    archivo_csv = open(ruta_csv, "w", newline="", encoding="utf-8-sig")
    escritor = csv.writer(archivo_csv, delimiter=";")
    escritor.writerow(ENCABEZADOS)

    print("Esperando datos (la micro:bit envia una linea cada 10 s)... Ctrl+C para terminar.\n")
    raws, n = [], 0
    fallos_excel = 0
    fase_anterior = None
    try:
        for linea in lineas:
            linea = linea.strip()
            if not linea or linea.startswith("ID:"):
                continue
            datos = parsear_linea(linea)
            if datos is None:
                if linea.startswith("ERROR_DS18B20"):
                    avisar("   DS18B20 -> {}  (revisa cableado P0 y resistencia 4.7k)".format(linea), "error")
                else:
                    print("   (linea ignorada: {!r})".format(linea))
                continue
            temp, hum, raw = datos
            hora = datetime.datetime.now()
            n += 1
            raws.append(raw)

            if panel is not None:
                panel.lectura(hora, temp, hum, raw)
                fase = fase_compost(temp)
                if fase != fase_anterior:
                    nivel = {"TERMOFILA": "ok", "MESOFILA": "info"}.get(fase, "aviso" if fase == "TERMOFILA ALTA" else "error")
                    panel.evento("Fase: {}".format(fase) if fase_anterior is None
                                 else "Cambio de fase: {} -> {}".format(fase_anterior, fase), nivel)
                    fase_anterior = fase

            t_txt = "ERROR SENSOR ({})".format(linea.split(",")[0]) if temp is None else "{} °C".format(temp)
            print("{} | Temperatura: {} | Humedad: {}% | RAW: {}".format(
                hora.strftime("%H:%M:%S"), t_txt, hum, raw))

            escritor.writerow([hora.strftime("%H:%M:%S"), "" if temp is None else temp, hum, raw])
            archivo_csv.flush()

            if excel is not None:
                try:
                    excel.agregar(hora, temp, hum, raw)
                    fallos_excel = 0
                    if n % GUARDAR_CADA == 0:
                        excel.guardar()
                except Exception as e:
                    if excel_ocupado(e):
                        avisar("   Excel esta ocupado (una celda en edicion o un aviso abierto?).\n"
                               "   Pulsa Esc o cierra el aviso en Excel; las filas pendientes ({}) se\n"
                               "   escribiran con la proxima lectura. El CSV ya las tiene guardadas.".format(
                                   len(excel.pendientes)), "aviso")
                    else:
                        fallos_excel += 1
                        avisar("Error escribiendo en Excel ({}).".format(e), "aviso")
                        if fallos_excel >= 3:
                            avisar("Excel no responde (se cerro?). Se sigue guardando solo en CSV.", "error")
                            excel = None

            if args.max_lecturas and n >= args.max_lecturas:
                break
    except KeyboardInterrupt:
        print("\nDetenido por el usuario.")
    finally:
        archivo_csv.close()
        if microbit is not None:
            microbit.close()
        if excel is not None:
            try:
                con_reintentos(excel.volcar)
                excel.guardar()
                print("Excel guardado en: {}".format(excel.ruta))
            except Exception as e:
                print("No se pudo guardar el Excel: {}".format(e))
        print("CSV guardado en: {}".format(os.path.abspath(ruta_csv)))
        if raws:
            print("\nResumen FC-28 (para calibrar): {} lecturas | RAW min={} max={} media={:.0f}".format(
                len(raws), min(raws), max(raws), sum(raws) / len(raws)))
        if panel is not None:
            panel.evento("Registro detenido. Datos guardados.", "info")
            panel.detener()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:     # Ctrl+C mientras se espera la micro:bit
        print("\nDetenido por el usuario.")
        if panel is not None:
            panel.detener()
