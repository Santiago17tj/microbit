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
"""

import argparse
import csv
import datetime
import os
import random
import re
import sys
import threading
import time

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

ENCABEZADOS = ["Hora", "Temperatura (°C)", "Humedad (%)", "Humedad RAW"]

NUMERO = r"(-?(?:\d+(?:\.\d+)?|Infinity)|NaN)"
PATRON = re.compile(r"^TEMP:" + NUMERO + r",HUM:" + NUMERO + r",RAW:(\d+)$")

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass


# --------------------------------------------------------------------------
# SERIAL
# --------------------------------------------------------------------------
def es_microbit(p):
    texto = " ".join(str(x) for x in (p.description, p.manufacturer, p.product) if x).lower()
    return p.vid == MICROBIT_VID or "micro:bit" in texto or "mbed" in texto or "daplink" in texto


def detectar_puerto():
    print("Buscando la micro:bit en los puertos USB...")
    puertos = list(serial.tools.list_ports.comports())
    candidatos = [p for p in puertos if es_microbit(p)]
    if candidatos:
        p = candidatos[0]
        print("micro:bit detectada en {} ({})".format(p.device, p.description))
        return p.device
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
class HojaExcel:
    def __init__(self, ruta_xlsx):
        import win32com.client
        self.ruta = os.path.abspath(ruta_xlsx)
        # DispatchEx abre un Excel nuevo en vez de engancharse a uno oculto o bloqueado
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = True
        self.libro = self.excel.Workbooks.Add()
        self.hoja = self.libro.Worksheets(1)
        self.hoja.Name = "Compostaje"
        for col, texto in enumerate(ENCABEZADOS, start=1):
            self.hoja.Cells(1, col).Value = texto
        self.hoja.Range("A1:D1").Font.Bold = True
        self.hoja.Columns("A").NumberFormat = "hh:mm:ss"
        self.hoja.Columns("A:D").ColumnWidth = 18
        self.fila = 1
        self.tabla = None
        self.grafica = None

    def agregar(self, hora, temp, hum, raw):
        self.fila += 1
        f = self.fila
        segundos = hora.hour * 3600 + hora.minute * 60 + hora.second
        self.hoja.Cells(f, 1).Value = segundos / 86400.0      # hora como valor de Excel
        self.hoja.Cells(f, 2).Value = "" if temp is None else temp
        self.hoja.Cells(f, 3).Value = "" if hum is None else hum
        self.hoja.Cells(f, 4).Value = raw

        rango = self.hoja.Range("A1:D{}".format(f))
        if self.tabla is None:
            self.tabla = self.hoja.ListObjects.Add(SourceType=XL_SRC_RANGE, Source=rango,
                                                   XlListObjectHasHeaders=XL_YES)
            self.tabla.Name = "TablaCompost"
            self.tabla.TableStyle = "TableStyleMedium2"
            self._crear_grafica()
        else:
            self.tabla.Resize(rango)
        self._actualizar_grafica()
        try:
            self.excel.ActiveWindow.ScrollRow = max(1, f - 20)   # mantiene visible la ultima fila
        except Exception:
            pass   # el usuario esta usando otra ventana de Excel; no pasa nada

    def _crear_grafica(self):
        obj = self.hoja.ChartObjects().Add(self.hoja.Range("F2").Left, self.hoja.Range("F2").Top, 620, 340)
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
        self.grafica = g
        self.s_temp, self.s_hum = s_temp, s_hum

    def _actualizar_grafica(self):
        f = self.fila
        horas = self.hoja.Range("A2:A{}".format(f))
        self.s_temp.XValues = horas
        self.s_temp.Values = self.hoja.Range("B2:B{}".format(f))
        self.s_hum.XValues = horas
        self.s_hum.Values = self.hoja.Range("C2:C{}".format(f))
        if f == 2:   # los ejes existen cuando ya hay datos
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

    def guardar(self):
        self.excel.DisplayAlerts = False
        try:
            self.libro.SaveAs(self.ruta, XL_OPEN_XML_WORKBOOK)
        finally:
            self.excel.DisplayAlerts = True


# --------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Registro del compostaje micro:bit -> Excel")
    ap.add_argument("--puerto", help="Puerto COM manual, p. ej. COM5")
    ap.add_argument("--simular", action="store_true", help="Genera datos falsos (sin micro:bit)")
    ap.add_argument("--sin-excel", action="store_true", help="No abrir Excel (solo consola y CSV)")
    ap.add_argument("--max-lecturas", type=int, default=0, help="Termina tras N lecturas (0 = infinito)")
    args = ap.parse_args()

    os.makedirs(CARPETA_DATOS, exist_ok=True)
    marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_csv = os.path.join(CARPETA_DATOS, "compost_{}.csv".format(marca))
    ruta_xlsx = os.path.join(CARPETA_DATOS, "compost_{}.xlsx".format(marca))

    # ---- origen de datos ----
    microbit = None
    if args.simular:
        print("MODO SIMULACION: no se usa la micro:bit.")
        lineas = generador_simulado()
    else:
        puerto = args.puerto or PUERTO_MANUAL or detectar_puerto()
        if not puerto:
            sys.exit(1)
        try:
            microbit = abrir_serial(puerto)
        except serial.SerialException as e:
            print("No se pudo abrir {}: {}".format(puerto, e))
            print("Cierra MakeCode, el monitor serie u otro programa que este usando el puerto.")
            sys.exit(1)
        print("Puerto {} abierto a {} baudios.".format(puerto, BAUDIOS))

        def lineas_serial():
            nonlocal microbit
            ultimo_dato = ultimo_latido = time.time()
            avisado = False
            primer_dato = True
            while True:
                ahora = time.time()
                if primer_dato and ahora - ultimo_latido >= 10:
                    ultimo_latido = ahora
                    print("   ... esperando la primera linea de la micro:bit ({} s)".format(int(ahora - ultimo_dato)))
                if not avisado and ahora - ultimo_dato > 40:
                    avisado = True
                    print("AVISO: 40 s sin recibir nada de la micro:bit. Si su pantalla no muestra")
                    print("       'T:', 'MESOFILA'..., carga compostaje-microbit.hex (o")
                    print("       PROVISIONAL-temp-interna.hex) en la unidad MICROBIT.")
                    print("       Este programa sigue esperando y se reconecta solo.")
                try:
                    dato = microbit.readline()
                except serial.SerialException:
                    print("Se perdio la conexion con la micro:bit. Reintentando...")
                    microbit.close()
                    while True:
                        time.sleep(2)
                        try:
                            microbit = abrir_serial(args.puerto or PUERTO_MANUAL or detectar_puerto() or puerto)
                            print("Reconectado.")
                            break
                        except serial.SerialException:
                            pass
                    continue
                if dato:
                    if primer_dato:
                        primer_dato = False
                        print("Conexion OK: llegan datos de la micro:bit.")
                    ultimo_dato, avisado = time.time(), False
                    yield dato.decode("utf-8", errors="ignore")
        lineas = lineas_serial()

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
            print("Excel abierto. Hoja 'Compostaje' creada.")
        except Exception as e:
            print("No se pudo abrir Excel ({}). Se continua guardando en CSV.".format(e))
        finally:
            aviso_lento.cancel()

    archivo_csv = open(ruta_csv, "w", newline="", encoding="utf-8-sig")
    escritor = csv.writer(archivo_csv, delimiter=";")
    escritor.writerow(ENCABEZADOS)

    print("Esperando datos (la micro:bit envia una linea cada ~15-20 s)... Ctrl+C para terminar.\n")
    raws, n = [], 0
    try:
        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue
            datos = parsear_linea(linea)
            if datos is None:
                if linea.startswith("ERROR_DS18B20"):
                    print("   DS18B20 -> {}  (revisa cableado P0 y resistencia 4.7k)".format(linea))
                else:
                    print("   (linea ignorada: {!r})".format(linea))
                continue
            temp, hum, raw = datos
            hora = datetime.datetime.now()
            n += 1
            raws.append(raw)

            t_txt = "ERROR SENSOR ({})".format(linea.split(",")[0]) if temp is None else "{} °C".format(temp)
            print("{} | Temperatura: {} | Humedad: {}% | RAW: {}".format(
                hora.strftime("%H:%M:%S"), t_txt, hum, raw))

            escritor.writerow([hora.strftime("%H:%M:%S"), "" if temp is None else temp, hum, raw])
            archivo_csv.flush()

            if excel is not None:
                try:
                    excel.agregar(hora, temp, hum, raw)
                    if n % GUARDAR_CADA == 0:
                        excel.guardar()
                except Exception as e:
                    print("Error escribiendo en Excel ({}). Se sigue guardando en CSV.".format(e))
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
                excel.guardar()
                print("Excel guardado en: {}".format(excel.ruta))
            except Exception as e:
                print("No se pudo guardar el Excel: {}".format(e))
        print("CSV guardado en: {}".format(os.path.abspath(ruta_csv)))
        if raws:
            print("\nResumen FC-28 (para calibrar): {} lecturas | RAW min={} max={} media={:.0f}".format(
                len(raws), min(raws), max(raws), sum(raws) / len(raws)))


if __name__ == "__main__":
    main()
