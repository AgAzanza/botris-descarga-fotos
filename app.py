#!/usr/bin/env python3
"""
Interfaz grafica para la descarga de fotos.

Para quien no quiere usar la terminal: se eligen los excels con un boton,
se elige la carpeta de destino y se aprieta Descargar.

IMPORTANTE PARA QUIEN TOQUE ESTE ARCHIVO
========================================
Aca arriba SOLO se importa lo que trae Python de fabrica. Nada de pandas,
requests ni Pillow.

El motivo: este programa tiene que poder abrirse ANTES de que las
dependencias esten instaladas, porque una de sus funciones es justamente
instalarlas. Si importara pandas arriba, no abriria nunca en una maquina
recien configurada y la persona veria un error en vez de un boton.

Las marcas se importan recien cuando se aprieta Descargar, dentro del hilo
de trabajo, no al abrir la ventana.
"""

import os
import queue
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# Dentro del .exe no hay archivos sueltos: todo esta empaquetado y
# sys.executable apunta al propio ejecutable, no a un Python.
CONGELADO = getattr(sys, "frozen", False)
AQUI = Path(sys.executable).parent if CONGELADO else Path(__file__).parent

# nombre para importar -> nombre para instalar
DEPENDENCIAS = {
    "pandas": "pandas",
    "xlrd": "xlrd",
    "requests": "requests",
    "PIL": "Pillow",
    "curl_cffi": "curl_cffi",
}


def faltantes():
    """Que dependencias no estan instaladas."""
    import importlib.util
    return [paquete for modulo, paquete in DEPENDENCIAS.items()
            if importlib.util.find_spec(modulo) is None]


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Descarga de fotos de producto")
        self.geometry("880x640")
        self.minsize(720, 520)

        self.excels = []
        self.salida = tk.StringVar(value=str(Path.home() / "fotos"))
        self.limite = tk.StringVar(value="")
        self.forzar = tk.BooleanVar(value=False)
        self.cola = queue.Queue()
        self.corriendo = False

        self._construir()
        self._revisar_dependencias()
        self.after(100, self._vaciar_cola)

    # -------------------------------------------------------------- interfaz

    def _construir(self):
        pad = {"padx": 12, "pady": 6}

        # --- aviso de dependencias (se oculta si esta todo bien) ---
        self.barra_dep = tk.Frame(self, bg="#f6e6c8")
        self.txt_dep = tk.Label(self.barra_dep, bg="#f6e6c8", anchor="w",
                                justify="left", wraplength=620)
        self.txt_dep.pack(side="left", **pad)
        self.btn_dep = ttk.Button(self.barra_dep, text="Instalar ahora",
                                  command=self._instalar)
        self.btn_dep.pack(side="right", **pad)

        # --- 1. archivos ---
        f1 = ttk.LabelFrame(self, text="1. Excels de estructura")
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="Elegir archivos...",
                   command=self._elegir_excels).pack(side="left", **pad)
        self.lbl_excels = ttk.Label(f1, text="ningun archivo elegido")
        self.lbl_excels.pack(side="left", **pad)

        # --- 2. destino ---
        f2 = ttk.LabelFrame(self, text="2. Dónde guardar las fotos")
        f2.pack(fill="x", **pad)
        ttk.Entry(f2, textvariable=self.salida).pack(
            side="left", fill="x", expand=True, **pad)
        ttk.Button(f2, text="Cambiar...",
                   command=self._elegir_salida).pack(side="left", **pad)

        # --- 3. opciones ---
        f3 = ttk.LabelFrame(self, text="3. Opciones")
        f3.pack(fill="x", **pad)
        ttk.Label(f3, text="Probar solo los primeros").pack(side="left", padx=(12, 4))
        ttk.Entry(f3, textvariable=self.limite, width=5).pack(side="left")
        ttk.Label(f3, text="productos  (vacío = todos)").pack(side="left", padx=4)
        ttk.Checkbutton(f3, text="Rehacer los ya bajados",
                        variable=self.forzar).pack(side="left", padx=20)

        # --- accion ---
        f4 = tk.Frame(self)
        f4.pack(fill="x", **pad)
        self.btn_run = ttk.Button(f4, text="Descargar fotos",
                                  command=self._arrancar)
        self.btn_run.pack(side="left", padx=12)
        self.btn_carpeta = ttk.Button(f4, text="Abrir carpeta",
                                      command=self._abrir_carpeta,
                                      state="disabled")
        self.btn_carpeta.pack(side="left")
        self.btn_revision = ttk.Button(f4, text="Abrir hojas de revisión",
                                       command=self._abrir_revisiones,
                                       state="disabled")
        self.btn_revision.pack(side="left", padx=8)
        self.progreso = ttk.Progressbar(f4, mode="indeterminate", length=160)
        self.progreso.pack(side="right", padx=12)

        # --- consola ---
        f5 = ttk.LabelFrame(self, text="Progreso")
        f5.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(f5, wrap="none", height=16, bg="#1c1f22",
                           fg="#e6e8ea", insertbackground="#e6e8ea",
                           font=("Consolas" if os.name == "nt" else "monospace", 10))
        sb = ttk.Scrollbar(f5, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set, state="disabled")
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    # --------------------------------------------------------- dependencias

    def _revisar_dependencias(self):
        # En el .exe las dependencias van adentro: no hay nada que instalar.
        falta = [] if CONGELADO else faltantes()
        if falta:
            self.barra_dep.pack(fill="x", before=self.winfo_children()[1])
            self.txt_dep.config(
                text=("Falta instalar: " + ", ".join(falta) +
                      ".  Sin esto el programa no puede funcionar. "
                      "Hace falta conexión a internet una sola vez."))
            self.btn_run.config(state="disabled")
        else:
            self.barra_dep.pack_forget()
            self.btn_run.config(state="normal")

    def _instalar(self):
        if self.corriendo:
            return
        self.btn_dep.config(state="disabled")
        self._escribir("Instalando dependencias. Puede tardar un par de minutos.\n")
        self._lanzar(
            [sys.executable, "-m", "pip", "install", "-r",
             str(AQUI / "requirements.txt")],
            al_terminar=self._fin_instalacion)

    def _fin_instalacion(self, codigo):
        if codigo == 0 and not faltantes():
            self._escribir("\nListo, dependencias instaladas.\n\n")
        else:
            self._escribir(
                "\nNo se pudieron instalar. Probá desde una terminal:\n"
                f"  {sys.executable} -m pip install -r requirements.txt\n\n")
        self.btn_dep.config(state="normal")
        self._revisar_dependencias()

    # -------------------------------------------------------------- acciones

    def _elegir_excels(self):
        rutas = filedialog.askopenfilenames(
            title="Elegir los excels de estructura",
            filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
        if rutas:
            self.excels = [Path(r) for r in rutas]
            self.lbl_excels.config(
                text=(f"{len(self.excels)} archivo(s): " +
                      ", ".join(p.name for p in self.excels)[:70]))

    def _elegir_salida(self):
        d = filedialog.askdirectory(title="Dónde guardar las fotos")
        if d:
            self.salida.set(d)

    def _arrancar(self):
        if self.corriendo:
            return
        if not self.excels:
            messagebox.showwarning("Falta algo", "Elegí al menos un excel.")
            return

        self.btn_run.config(state="disabled")
        self.btn_carpeta.config(state="disabled")
        self.btn_revision.config(state="disabled")
        self.progreso.start(12)
        self._descargar()

    def _fin_descarga(self, codigo):
        self.progreso.stop()
        self.btn_run.config(state="normal")
        self.btn_carpeta.config(state="normal")
        if self._revisiones():
            self.btn_revision.config(state="normal")
        self._escribir(
            "\nTerminado. Abrí las hojas de revisión para controlar las fotos.\n"
            if codigo == 0 else f"\nTerminó con error {codigo}.\n")

    def _revisiones(self):
        base = Path(self.salida.get())
        return sorted(base.glob("*/revision/revision.html")) if base.exists() else []

    def _abrir_revisiones(self):
        for h in self._revisiones():
            webbrowser.open(h.resolve().as_uri())

    def _abrir_carpeta(self):
        d = Path(self.salida.get())
        d.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(d)                                   # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", str(d)])
        else:
            subprocess.run(["xdg-open", str(d)])

    def _descargar(self):
        """
        Corre las marcas EN ESTE MISMO PROCESO, en un hilo aparte.

        No se usa subprocess a proposito: dentro del .exe, sys.executable
        apunta al propio ejecutable y lanzarlo de nuevo abriria otra ventana
        en bucle en vez de correr el programa.

        Lo que las marcas imprimen se desvia a la consola de la ventana.
        """
        self.corriendo = True
        excels = [str(p) for p in self.excels]
        salida = self.salida.get()
        limite = self.limite.get().strip() or None
        forzar = self.forzar.get()

        class HaciaLaVentana:
            """Hace de sys.stdout: cada linea impresa va a la cola."""
            def __init__(self, cola):
                self.cola = cola
            def write(self, texto):
                if texto:
                    self.cola.put(("log", texto))
            def flush(self):
                pass

        def trabajo():
            import contextlib
            salida_previa = sys.stdout
            try:
                import runner
                with contextlib.redirect_stdout(HaciaLaVentana(self.cola)):
                    runner.ejecutar(excels, salida, limite, forzar)
                self.cola.put(("fin", 0))
            except Exception as e:
                import traceback
                self.cola.put(("log", "\nError:\n" + traceback.format_exc()))
                self.cola.put(("fin", 1))
            finally:
                sys.stdout = salida_previa

        self._al_terminar = self._fin_descarga
        threading.Thread(target=trabajo, daemon=True).start()

    # ------------------------------------------------- ejecutar y mostrar

    def _lanzar(self, cmd, al_terminar):
        """
        Corre un comando externo y manda su salida a la consola.

        Solo se usa para instalar dependencias con pip. La descarga NO pasa
        por aca: ver _descargar().
        """
        self.corriendo = True

        def trabajo():
            entorno = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
            try:
                p = subprocess.Popen(
                    cmd, cwd=str(AQUI), env=entorno,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1,
                    creationflags=(subprocess.CREATE_NO_WINDOW
                                   if os.name == "nt" else 0))
                for linea in p.stdout:
                    self.cola.put(("log", linea))
                p.wait()
                self.cola.put(("fin", p.returncode))
            except Exception as e:
                self.cola.put(("log", f"\nNo se pudo ejecutar: {e}\n"))
                self.cola.put(("fin", 1))

        self._al_terminar = al_terminar
        threading.Thread(target=trabajo, daemon=True).start()

    def _vaciar_cola(self):
        """La consola se actualiza solo desde el hilo de la interfaz."""
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self._escribir(dato)
                else:
                    self.corriendo = False
                    self._al_terminar(dato)
        except queue.Empty:
            pass
        self.after(100, self._vaciar_cola)

    def _escribir(self, texto):
        self.log.config(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
