# -*- coding: utf-8 -*-
"""
ventanas.py
Coloca el panel web y Excel lado a lado en la pantalla (solo Windows):

    +-------------------+-------------------+
    |    Panel web      |      Excel        |
    |  (Edge / Chrome)  |  tabla + grafica  |
    +-------------------+-------------------+

El panel se abre como ventana de aplicacion (sin pestanas ni barra de
direcciones) con Edge o Chrome.  Si no hay ninguno, se usa el navegador
por defecto.  Solo usa ctypes: no hace falta instalar nada.
"""

import ctypes
import os
import subprocess
import sys
import threading
import time
import webbrowser

TITULO_PANEL = "Monitor de compostaje"     # <title> de panel/index.html
ES_WINDOWS = sys.platform == "win32"

if ES_WINDOWS:
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    _WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = [_WNDENUMPROC, wintypes.LPARAM]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.MoveWindow.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.BOOL]


def _dpi_real():
    """Trabajar en pixeles reales aunque Windows tenga el escalado al 125 % / 150 %."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass


if ES_WINDOWS:
    _dpi_real()


# --------------------------------------------------------------------------
# Abrir el panel
# --------------------------------------------------------------------------
def _navegador_app():
    """Ruta de Edge o Chrome (los que tienen modo --app).  Prefiere el navegador
    por defecto si es uno de ellos."""
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    local = os.environ.get("LocalAppData", "")
    chrome = [os.path.join(pf, r"Google\Chrome\Application\chrome.exe"),
              os.path.join(pf86, r"Google\Chrome\Application\chrome.exe"),
              os.path.join(local, r"Google\Chrome\Application\chrome.exe")]
    edge = [os.path.join(pf86, r"Microsoft\Edge\Application\msedge.exe"),
            os.path.join(pf, r"Microsoft\Edge\Application\msedge.exe")]
    por_defecto = ""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as k:
            por_defecto = winreg.QueryValueEx(k, "ProgId")[0].lower()
    except Exception:
        pass
    orden = chrome + edge if "chrome" in por_defecto else edge + chrome
    for ruta in orden:
        if os.path.isfile(ruta):
            return ruta
    return None


def abrir_panel(url):
    """Abre el panel en una ventana propia.  Si ya hay una abierta (de una
    ejecucion anterior) la reutiliza: la pagina se reconecta sola."""
    if ES_WINDOWS:
        ya_abierta = buscar_ventana(TITULO_PANEL, CLASES_NAVEGADOR)
        if ya_abierta:
            user32.ShowWindow(ya_abierta, 9)            # SW_RESTORE
            user32.SetForegroundWindow(ya_abierta)
            return True
    exe = _navegador_app() if ES_WINDOWS else None
    if exe:
        try:
            subprocess.Popen([exe, "--app=" + url], close_fds=True,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except OSError:
            pass
    webbrowser.open(url)
    return False


# --------------------------------------------------------------------------
# Colocar ventanas
# --------------------------------------------------------------------------
class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def area_trabajo():
    """(x, y, ancho, alto) de la pantalla principal sin la barra de tareas."""
    r = _RECT()
    user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(r), 0)    # SPI_GETWORKAREA
    return r.left, r.top, r.right - r.left, r.bottom - r.top


CLASES_NAVEGADOR = ("Chrome_WidgetWin_1", "MozillaWindowClass")   # Edge/Chrome, Firefox


def buscar_ventana(titulo, clases=None):
    """Primera ventana visible cuyo titulo empieza por `titulo` (y cuya clase
    esta en `clases`, si se indica)."""
    if not ES_WINDOWS:
        return None
    encontradas = []

    def cada(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, buf, 512)
            if buf.value.startswith(titulo):
                clase = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, clase, 256)
                if clases is None or clase.value in clases:
                    encontradas.append(hwnd)
        return True

    user32.EnumWindows(_WNDENUMPROC(cada), 0)
    return encontradas[0] if encontradas else None


def colocar(hwnd, x, y, ancho, alto):
    """Mueve la ventana al rectangulo visible indicado.  Windows 10/11 dibuja
    un borde invisible de ~7 px alrededor de cada ventana: se compensa para que
    las dos mitades queden pegadas, sin huecos."""
    user32.ShowWindow(hwnd, 9)                          # SW_RESTORE (por si esta maximizada)
    ventana, visible = _RECT(), _RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(ventana))
    try:
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, 9, ctypes.byref(visible), ctypes.sizeof(visible))   # DWMWA_EXTENDED_FRAME_BOUNDS
    except Exception:
        visible = ventana
    izq, arriba = visible.left - ventana.left, visible.top - ventana.top
    der, abajo = ventana.right - visible.right, ventana.bottom - visible.bottom
    user32.MoveWindow(hwnd, x - izq, y - arriba, ancho + izq + der, alto + arriba + abajo, True)


def lado_a_lado(hwnd_excel=None, espera_panel=15):
    """Panel a la izquierda y Excel a la derecha.  Si no hay Excel, el panel
    ocupa toda la pantalla.  Espera unos segundos a que aparezca el panel."""
    if not ES_WINDOWS:
        return False
    x, y, ancho, alto = area_trabajo()
    mitad = ancho // 2
    fin = time.time() + espera_panel
    panel = buscar_ventana(TITULO_PANEL, CLASES_NAVEGADOR)
    while panel is None and time.time() < fin:
        time.sleep(0.5)
        panel = buscar_ventana(TITULO_PANEL, CLASES_NAVEGADOR)
    if hwnd_excel:
        colocar(hwnd_excel, x + mitad, y, ancho - mitad, alto)
    if panel:
        if hwnd_excel:
            colocar(panel, x, y, mitad, alto)
        else:
            user32.ShowWindow(panel, 3)                 # SW_MAXIMIZE
    return panel is not None


def lado_a_lado_en_segundo_plano(hwnd_excel=None):
    """Igual que lado_a_lado() pero sin bloquear el programa."""
    hilo = threading.Thread(target=lado_a_lado, args=(hwnd_excel,), daemon=True)
    hilo.start()
    return hilo
