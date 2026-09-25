# -*- coding: utf-8 -*-
"""
panel_web.py
Panel web local en tiempo real para el monitor de compostaje.

Solo usa la libreria estandar de Python (no hace falta instalar nada ni tener
internet).  DataStream_compost.py lo arranca solo y abre el navegador en
http://localhost:8765

    /                -> panel/index.html
    /api/estado      -> JSON con todo el historial (al abrir la pagina)
    /eventos         -> Server-Sent Events con cada lectura nueva
    /descargar.csv   -> el CSV de la sesion actual
"""

import json
import os
import queue
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

AQUI = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_HTML = os.path.join(AQUI, "panel", "index.html")
MAX_LECTURAS = 5000        # historial que se guarda en memoria para la grafica
MAX_EVENTOS = 60


class _Servidor(ThreadingHTTPServer):
    # En Windows SO_REUSEADDR deja que dos programas usen el mismo puerto a la vez
    allow_reuse_address = sys.platform != "win32"
    daemon_threads = True

    def handle_error(self, request, client_address):
        # El navegador corta conexiones al recargar o cerrar la pestana: no es un error
        if not isinstance(sys.exc_info()[1], (ConnectionError, TimeoutError)):
            super().handle_error(request, client_address)


class PanelWeb:
    def __init__(self, puerto_web=8765, modo="real", ruta_csv=None):
        self.puerto_web = puerto_web
        self.ruta_csv = ruta_csv
        self._cerrojo = threading.Lock()
        self._clientes = set()
        self._estado = {
            "modo": modo,              # "real" o "simulacion"
            "conexion": "iniciando",   # iniciando / esperando / conectado / reconectando / detenido
            "puerto": None,
            "lecturas": [],
            "eventos": [],
        }
        self._servidor = None

    # ---- API que usa DataStream_compost.py ---------------------------------
    def iniciar(self, abrir_navegador=True):
        panel = self

        class Manejador(_Manejador):
            pass
        Manejador.panel = panel

        self._servidor = _Servidor(("127.0.0.1", self.puerto_web), Manejador)
        hilo = threading.Thread(target=self._servidor.serve_forever, daemon=True)
        hilo.start()
        url = "http://localhost:{}".format(self.puerto_web)
        if abrir_navegador:
            try:
                from ventanas import abrir_panel      # ventana propia (Edge/Chrome en modo app)
                abrir_panel(url)
            except Exception:
                webbrowser.open(url)
        return url

    def detener(self):
        self.conexion("detenido")
        if self._servidor is not None:
            self._servidor.shutdown()

    def lectura(self, hora, temp, hum, raw):
        dato = {"t": int(hora.timestamp() * 1000), "temp": temp, "hum": hum, "raw": raw}
        with self._cerrojo:
            lecturas = self._estado["lecturas"]
            lecturas.append(dato)
            if len(lecturas) > MAX_LECTURAS:
                del lecturas[0]
        self._publicar("lectura", dato)

    def evento(self, texto, nivel="info"):
        """nivel: info / ok / aviso / error"""
        dato = {"t": int(time.time() * 1000), "texto": texto, "nivel": nivel}
        with self._cerrojo:
            eventos = self._estado["eventos"]
            eventos.append(dato)
            if len(eventos) > MAX_EVENTOS:
                del eventos[0]
        self._publicar("evento", dato)

    def conexion(self, estado, puerto=None):
        with self._cerrojo:
            self._estado["conexion"] = estado
            if puerto:
                self._estado["puerto"] = puerto
            dato = {"conexion": estado, "puerto": self._estado["puerto"]}
        self._publicar("conexion", dato)

    # ---- interno -------------------------------------------------------------
    def _instantanea(self):
        with self._cerrojo:
            copia = dict(self._estado)
            copia["lecturas"] = list(self._estado["lecturas"])
            copia["eventos"] = list(self._estado["eventos"])
        copia["csv"] = os.path.basename(self.ruta_csv) if self.ruta_csv else None
        return copia

    def _publicar(self, tipo, dato):
        mensaje = "event: {}\ndata: {}\n\n".format(tipo, json.dumps(dato, ensure_ascii=False))
        with self._cerrojo:
            clientes = list(self._clientes)
        for q in clientes:
            q.put(mensaje)


class _Manejador(BaseHTTPRequestHandler):
    panel = None

    def log_message(self, *args):
        pass   # no ensuciar la consola con cada peticion

    def _enviar(self, codigo, tipo, cuerpo, extra=None):
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(cuerpo)

    def do_GET(self):
        ruta = self.path.split("?")[0]
        if ruta in ("/", "/index.html"):
            try:
                with open(ARCHIVO_HTML, "rb") as f:
                    self._enviar(200, "text/html; charset=utf-8", f.read())
            except OSError:
                self._enviar(500, "text/plain; charset=utf-8", "Falta panel/index.html".encode("utf-8"))
        elif ruta == "/api/estado":
            cuerpo = json.dumps(self.panel._instantanea(), ensure_ascii=False).encode("utf-8")
            self._enviar(200, "application/json; charset=utf-8", cuerpo)
        elif ruta == "/eventos":
            self._stream()
        elif ruta == "/descargar.csv" and self.panel.ruta_csv and os.path.exists(self.panel.ruta_csv):
            with open(self.panel.ruta_csv, "rb") as f:
                nombre = os.path.basename(self.panel.ruta_csv)
                self._enviar(200, "text/csv; charset=utf-8", f.read(),
                             {"Content-Disposition": 'attachment; filename="{}"'.format(nombre)})
        else:
            self._enviar(404, "text/plain; charset=utf-8", b"No encontrado")

    def _stream(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Connection", "keep-alive")
        self.end_headers()
        q = queue.Queue()
        with self.panel._cerrojo:
            self.panel._clientes.add(q)
        try:
            self.wfile.write(b"retry: 2000\n\n")
            self.wfile.flush()
            while True:
                try:
                    mensaje = q.get(timeout=15)
                    self.wfile.write(mensaje.encode("utf-8"))
                except queue.Empty:
                    self.wfile.write(b": latido\n\n")   # mantiene viva la conexion
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
            pass
        finally:
            with self.panel._cerrojo:
                self.panel._clientes.discard(q)
