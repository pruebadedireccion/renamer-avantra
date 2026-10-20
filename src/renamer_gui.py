# -*- coding: utf-8 -*-
"""
Renamer Avantra - Interfaz grafica (v1.2.0).

Novedades v1.2.0:
    - Tabla de resultados con filtros (Todo / Listos / Pendientes / Anomalias)
      y buscador, ademas del registro detallado en pestana aparte
    - Deteccion de mapeos conflictivos en el Excel
    - Validacion de PDFs danados o descargas incompletas
    - Boton Cancelar durante procesos largos
    - Deshacer la ultima ejecucion (solo copias registradas en reporte.csv)
    - Drag and drop de carpetas y Excel (si tkinterdnd2 esta disponible)
"""

from __future__ import annotations

import ctypes
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

sys.path.insert(0, str(Path(__file__).resolve().parent))

import errores as err
import renamer_core as core

try:  # drag and drop opcional
    from tkinterdnd2 import DND_FILES
    from tkinterdnd2.TkinterDnD import _require as _cargar_tkdnd
    TIENE_DND = True
except Exception:  # pragma: no cover - entorno sin tkinterdnd2
    TIENE_DND = False

TITULO = "Renamer Avantra - Renombrado Masivo de Documentos"
VERSION = "1.3.0"

# --- Paleta ------------------------------------------------------------------
AZUL_OSCURO = "#1F3864"
AZUL_MEDIO = "#2E5496"
AZUL_CLARO = "#D9E2F3"
VERDE = "#1E7B34"
VERDE_HOVER = "#156028"
GRIS_SIM = "#5B6E8C"
GRIS_SIM_HOVER = "#44546A"
GRIS_TXT = "#404040"
FONDO = "#F4F6FA"

FUENTE_BASE = ("Segoe UI", 13)
FUENTE_BASE_NEG = ("Segoe UI", 13, "bold")
FUENTE_TITULO = ("Segoe UI Semibold", 22)
FUENTE_SUBTITULO = ("Segoe UI", 13)
FUENTE_SECCION = ("Segoe UI", 14, "bold")
FUENTE_LOG = ("Consolas", 12)
FUENTE_ESTADO = ("Segoe UI", 12)

ARCHIVO_CONFIG = "config.json"

FILTROS = ("Todo", "Listos", "Pendientes", "Anomalias")


def habilitar_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def tamano_inicial() -> str:
    try:
        usuario32 = ctypes.windll.user32
        ancho = usuario32.GetSystemMetrics(0)
        alto = usuario32.GetSystemMetrics(1)
        w = max(900, min(int(ancho * 0.64), 1360))
        h = max(680, min(int(alto * 0.80), 1040))
        return f"{w}x{h}"
    except Exception:
        return "1080x780"


class App:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title(TITULO)
        self.root.geometry(tamano_inicial())
        self.root.minsize(800, 580)
        try:
            self.root.configure(fg_color=FONDO)
        except Exception:
            # Root tkinter plano (sin customtkinter): usar bg clasico
            self.root.configure(bg=FONDO)

        self.var_origen = ctk.StringVar()
        self.var_destino = ctk.StringVar()
        self.var_mapeo = ctk.StringVar()
        self.var_patron = ctk.StringVar(value=core.PATRON_DEFAULT)
        self.var_codigo = ctk.StringVar(value="")
        self.var_filtro = ctk.StringVar(value="Todo")

        self.ruta_config = Path(".") / ARCHIVO_CONFIG
        cfg = core.cargar_config(self.ruta_config)
        if cfg.get("patron"):
            self.var_patron.set(cfg["patron"])
        if cfg.get("codigo") is not None:
            self.var_codigo.set(cfg["codigo"])
        if cfg.get("origen") and Path(cfg["origen"]).is_dir():
            self.var_origen.set(cfg["origen"])
        if cfg.get("destino") and Path(cfg["destino"]).is_dir():
            self.var_destino.set(cfg["destino"])
        if cfg.get("mapeo") and Path(cfg["mapeo"]).is_file():
            self.var_mapeo.set(cfg["mapeo"])

        # Deteccion de identificadores (v1.3.0)
        self.modo = cfg.get("modo", core.MODO_AUTO)
        if self.modo not in (core.MODO_AUTO, core.MODO_RADICADO, core.MODO_CUSTOM):
            self.modo = core.MODO_AUTO
        self.regex_origen_cfg = cfg.get("regex_origen", "")
        self.regex_destino_cfg = cfg.get("regex_destino", "")

        self.plan: list[dict] = []
        self.registros: list[dict] = []
        self.mapeo: dict[str, str] | None = None
        self.mapeo_ruta: str = ""
        self.conflictos: dict[str, list[str]] = {}
        self._cancelar = False

        self._construir_widgets()
        self._centrar()

        self._log("Bienvenido.", "fase")
        self._log("1. Seleccione la carpeta de PDFs y la carpeta destino.", "info")
        self._log("2. Defina el FORMATO DE SALIDA y presione INVENTARIAR.", "info")
        self._log(
            "Con pendientes: EJECUTAR renombra los validos y deja registro "
            "de los demas (pendientes_correccion.csv).", "info",
        )

    # --- Widgets -------------------------------------------------------------

    def _construir_widgets(self) -> None:
        marco = ctk.CTkFrame(self.root, fg_color="transparent")
        marco.pack(fill="both", expand=True)

        # Encabezado
        cabezal = ctk.CTkFrame(
            marco, fg_color=AZUL_OSCURO, corner_radius=0, height=64
        )
        cabezal.pack(fill="x")
        cabezal.pack_propagate(False)
        ctk.CTkLabel(
            cabezal, text="RENAMER AVANTRA", font=FUENTE_TITULO,
            text_color="white",
        ).pack(side="left", padx=(20, 8), pady=10)
        ctk.CTkLabel(
            cabezal,
            text="Renombrado masivo de documentos  |  By: Red Avantra \u00ae",
            font=FUENTE_SUBTITULO,
            text_color=AZUL_CLARO,
        ).pack(side="left", pady=10)
        ctk.CTkLabel(
            cabezal, text=f"v{VERSION}", font=FUENTE_ESTADO,
            text_color=AZUL_CLARO,
        ).pack(side="right", padx=(0, 20))

        cont = ctk.CTkFrame(marco, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=14, pady=8)

        # --- 1. CARPETAS ------------------------------------------------------
        s1 = self._seccion(cont, "1. CARPETAS")
        interno = ctk.CTkFrame(s1, fg_color="white", corner_radius=0)
        interno.pack(fill="x", padx=16, pady=(2, 8))
        self.ent_origen = self._campo_ruta(
            interno, 0, "Carpeta de PDFs (origen):", self.var_origen,
            al_cambiar=self._vista_previa_carpeta,
        )
        self.lbl_vista_carpeta = ctk.CTkLabel(
            interno, text="", font=FUENTE_ESTADO, text_color=AZUL_MEDIO,
            anchor="w",
        )
        self.lbl_vista_carpeta.grid(
            row=3, column=0, columnspan=2, sticky="w", padx=(0, 10),
            pady=(0, 4),
        )
        fila_av = ctk.CTkFrame(interno, fg_color="transparent")
        fila_av.grid(row=3, column=2, sticky="e", pady=(0, 4))
        ctk.CTkButton(
            fila_av, text="Verificar Excel", command=self._verificar_excel,
            fg_color=AZUL_CLARO, hover_color="#C4D2EE", text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO, corner_radius=8, height=28, width=120,
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            fila_av, text="Crear plantilla", command=self._crear_plantilla,
            fg_color=AZUL_CLARO, hover_color="#C4D2EE", text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO, corner_radius=8, height=28, width=118,
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            fila_av, text="Avanzado", command=self._ventana_avanzada,
            fg_color="white", hover_color=AZUL_CLARO, text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO, corner_radius=8, height=28, width=90,
        ).pack(side="left")
        self.ent_destino = self._campo_ruta(
            interno, 1, "Carpeta destino (renombrados):", self.var_destino
        )
        self.ent_mapeo = self._campo_excel(interno, 2)

        self._registrar_dnd(self.ent_origen, "carpeta_origen")
        self._registrar_dnd(self.ent_destino, "carpeta_destino")
        self._registrar_dnd(self.ent_mapeo, "excel")

        if self.var_origen.get():
            self._vista_previa_carpeta()
        if self.var_mapeo.get():
            self._validar_excel_seleccionado()

        # --- 2. FORMATO DE SALIDA ---------------------------------------------
        s2 = self._seccion(cont, "2. FORMATO DE SALIDA", con_accion=True)

        fmt = ctk.CTkFrame(s2, fg_color="white", corner_radius=0)
        fmt.pack(fill="x", padx=16, pady=(2, 6))
        fmt.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            fmt, text="Patron de nombre final:", font=FUENTE_BASE,
            text_color=GRIS_TXT,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.ent_patron = ctk.CTkEntry(
            fmt, textvariable=self.var_patron, font=("Consolas", 13),
            height=34, corner_radius=8, border_color="#C9D2E3",
            fg_color="white", text_color=GRIS_TXT,
        )
        self.ent_patron.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=4)
        self.ent_patron.bind("<KeyRelease>", lambda _e: self._actualizar_ejemplo())

        ctk.CTkLabel(
            fmt, text="Codigo fijo {c}:", font=FUENTE_BASE, text_color=GRIS_TXT
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        self.ent_codigo = ctk.CTkEntry(
            fmt, textvariable=self.var_codigo, font=("Consolas", 13),
            height=34, corner_radius=8, width=180,
            placeholder_text="(opcional, ej. ZREE)",
            fg_color="white", text_color=GRIS_TXT,
        )
        self.ent_codigo.grid(row=1, column=1, sticky="w", pady=4)
        self.ent_codigo.bind("<KeyRelease>", lambda _e: self._actualizar_ejemplo())

        piezas = ctk.CTkFrame(fmt, fg_color="transparent")
        piezas.grid(row=1, column=2, sticky="w", padx=(10, 0), pady=4)
        ctk.CTkLabel(
            piezas, text="Piezas: ", font=FUENTE_ESTADO, text_color=GRIS_TXT
        ).pack(side="left")
        for token in ("{id}", "{n}", "{c}"):
            ctk.CTkButton(
                piezas, text=token, width=52, height=26, corner_radius=6,
                font=("Consolas", 12), fg_color=AZUL_CLARO,
                hover_color="#C4D2EE", text_color=AZUL_OSCURO,
                command=lambda t=token: self._insertar_token(t),
            ).pack(side="left", padx=(0, 6))

        ej = ctk.CTkFrame(s2, fg_color="#F4F6FA", corner_radius=8)
        ej.pack(fill="x", padx=16, pady=(2, 8))
        self.lbl_ejemplo = ctk.CTkLabel(
            ej, text="", font=("Consolas", 14), text_color=GRIS_TXT,
            justify="left", anchor="w",
        )
        self.lbl_ejemplo.pack(anchor="w", padx=14, pady=(8, 4))
        self.lbl_patron_estado = ctk.CTkLabel(
            ej, text="", font=FUENTE_ESTADO, anchor="w"
        )
        self.lbl_patron_estado.pack(anchor="w", padx=14, pady=(0, 10))
        self._actualizar_ejemplo()

        # --- 3. ACCIONES --------------------------------------------------------
        s3 = self._seccion(cont, "3. ACCIONES  (en orden)")
        acciones = ctk.CTkFrame(s3, fg_color="transparent")
        acciones.pack(fill="x", padx=16, pady=(2, 8))
        acciones.columnconfigure((0, 1, 2), weight=1)

        self.btn_inv = ctk.CTkButton(
            acciones, text="INVENTARIAR", command=self._accion_inventariar,
            fg_color=AZUL_MEDIO, hover_color=AZUL_OSCURO, text_color="white",
            font=FUENTE_BASE_NEG, corner_radius=8, height=38,
        )
        self.btn_inv.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.btn_sim = ctk.CTkButton(
            acciones, text="SIMULAR (dry-run)", command=self._accion_simular,
            fg_color=GRIS_SIM, hover_color=GRIS_SIM_HOVER, text_color="white",
            font=FUENTE_BASE_NEG, corner_radius=8, height=38,
            state="disabled",
        )
        self.btn_sim.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self.btn_ejec = ctk.CTkButton(
            acciones, text="EJECUTAR (copiar)", command=self._accion_ejecutar,
            fg_color=VERDE, hover_color=VERDE_HOVER, text_color="white",
            font=FUENTE_BASE_NEG, corner_radius=8, height=38,
            state="disabled",
        )
        self.btn_ejec.grid(row=0, column=2, sticky="ew")

        self.btn_cancelar = ctk.CTkButton(
            acciones, text="Cancelar", command=self._accion_cancelar,
            fg_color="#9C1C1C", hover_color="#7A1616", text_color="white",
            font=FUENTE_BASE, corner_radius=8, height=38, width=100,
            state="disabled",
        )
        self.btn_cancelar.grid(row=0, column=3, padx=(8, 0))

        ctk.CTkButton(
            acciones, text="Salir", command=self.root.destroy,
            fg_color="#8C8C8C", hover_color="#6E6E6E", font=FUENTE_BASE,
            corner_radius=8, height=38, width=90,
        ).grid(row=0, column=4, padx=(8, 0))

        self.progreso = ctk.CTkProgressBar(
            acciones, progress_color=AZUL_MEDIO, height=12
        )
        self.progreso.set(0)
        self.progreso.grid(row=1, column=0, columnspan=5, sticky="ew",
                           pady=(6, 0))

        # --- 4. RESULTADOS ------------------------------------------------------
        s4 = self._seccion(cont, "4. RESULTADOS")
        self.tabs = ctk.CTkTabview(s4, fg_color="white")
        self.tabs.pack(fill="both", expand=True, padx=14, pady=(2, 8))
        tab_plan = self.tabs.add("Plan")
        tab_log = self.tabs.add("Registro")

        # Filtros + buscador + re-verificar
        barra = ctk.CTkFrame(tab_plan, fg_color="transparent")
        barra.pack(fill="x", padx=4, pady=(4, 6))

        self.seg_filtro = ctk.CTkSegmentedButton(
            barra, values=list(FILTROS), variable=self.var_filtro,
            command=lambda _v: self._refrescar_tabla(),
            font=FUENTE_ESTADO, height=30,
        )
        self.seg_filtro.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            barra, text="Re-verificar", command=self._accion_simular,
            fg_color=AZUL_CLARO, hover_color="#C4D2EE", text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO, corner_radius=8, height=30, width=110,
        ).pack(side="right", padx=(10, 0))

        self.ent_buscar = ctk.CTkEntry(
            barra, placeholder_text="Buscar archivo, destino o motivo...",
            font=FUENTE_ESTADO, height=30, corner_radius=8,
            border_color="#C9D2E3", fg_color="white", text_color=GRIS_TXT,
        )
        self.ent_buscar.pack(side="left", fill="x", expand=True, padx=(0, 0))
        self.ent_buscar.bind("<KeyRelease>", lambda _e: self._refrescar_tabla())

        # Tabla de resultados
        estilo = ttk.Style(self.root)
        try:
            estilo.theme_use("clam")
        except Exception:
            pass
        estilo.configure(
            "Renamer.Treeview", font=("Segoe UI", 10), rowheight=24,
            background="white", fieldbackground="white", borderwidth=0,
        )
        estilo.configure(
            "Renamer.Treeview.Heading", font=("Segoe UI", 10, "bold"),
            background=AZUL_CLARO, foreground=AZUL_OSCURO,
        )
        estilo.map(
            "Renamer.Treeview",
            background=[("selected", AZUL_MEDIO)],
            foreground=[("selected", "white")],
        )

        contenedor_tabla = ctk.CTkFrame(tab_plan, fg_color="transparent")
        contenedor_tabla.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        contenedor_tabla.rowconfigure(0, weight=1)
        contenedor_tabla.columnconfigure(0, weight=1)

        columnas = ("original", "destino", "estado", "detalle")
        self.tabla = ttk.Treeview(
            contenedor_tabla, columns=columnas, show="headings",
            style="Renamer.Treeview", selectmode="browse",
        )
        self.tabla.heading("original", text="Archivo original")
        self.tabla.heading("destino", text="Nombre final")
        self.tabla.heading("estado", text="Estado")
        self.tabla.heading("detalle", text="Detalle / accion sugerida")
        self.tabla.column("original", width=330, stretch=True)
        self.tabla.column("destino", width=250, stretch=True)
        self.tabla.column("estado", width=110, stretch=False)
        self.tabla.column("detalle", width=420, stretch=True)
        self.tabla.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(
            contenedor_tabla, orient="vertical", command=self.tabla.yview
        )
        self.tabla.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        self.tabla.tag_configure("ok", foreground=VERDE)
        self.tabla.tag_configure("pend", foreground="#8C6209")
        self.tabla.tag_configure("err", foreground="#9C1C1C")

        # Registro (log)
        self.txt = ctk.CTkTextbox(
            tab_log, font=FUENTE_LOG, fg_color="white", text_color=GRIS_TXT,
            wrap="none",
        )
        self.txt.pack(fill="both", expand=True, padx=4, pady=4)
        self.txt._textbox.tag_configure("fase", foreground=AZUL_OSCURO)
        self.txt._textbox.tag_configure("ok", foreground=VERDE)
        self.txt._textbox.tag_configure("error", foreground="#9C1C1C")
        self.txt._textbox.tag_configure("anomalia", foreground="#9C1C1C")
        self.txt._textbox.tag_configure("sinmapeo", foreground="#8C6209")
        self.txt._textbox.tag_configure("info", foreground=GRIS_TXT)
        self.txt._textbox.tag_configure("ruta", foreground=AZUL_MEDIO)

        # --- Pie: altura automatica, botones sin ancho fijo ----------------------
        pie = ctk.CTkFrame(
            self.root, fg_color="white", corner_radius=0,
            border_width=1, border_color="#D6DCE8",
        )
        pie.pack(fill="x", side="bottom")

        self.var_estado = ctk.StringVar(value="Listo.")
        ctk.CTkLabel(
            pie, textvariable=self.var_estado, font=FUENTE_ESTADO,
            text_color=GRIS_TXT, anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=14, pady=8)

        for texto, comando in (
            ("Copiar lista", self._copiar_lista),
            ("Abrir carpeta destino", self._abrir_destino),
            ("Abrir reporte.csv", self._abrir_reporte),
            ("Deshacer ultima", self._accion_deshacer),
        ):
            ctk.CTkButton(
                pie, text=texto, command=comando,
                fg_color=AZUL_CLARO, hover_color="#C4D2EE",
                text_color=AZUL_OSCURO, font=FUENTE_ESTADO,
                corner_radius=8, height=32, width=130,
            ).pack(side="right", padx=(6, 10))

    # --- Helpers de construccion -----------------------------------------------

    def _seccion(self, contenedor, titulo, con_accion=False) -> ctk.CTkFrame:
        s = ctk.CTkFrame(
            contenedor, fg_color="white", corner_radius=12,
            border_width=1, border_color="#D6DCE8",
        )
        s.pack(fill="x", pady=(0, 6))
        cab = ctk.CTkFrame(s, fg_color="transparent")
        cab.pack(fill="x", padx=16, pady=(10, 2))
        ctk.CTkLabel(
            cab, text=titulo, font=FUENTE_SECCION, text_color=AZUL_OSCURO
        ).pack(side="left")
        if con_accion:
            ctk.CTkButton(
                cab, text="  ?  ", width=30, height=24, corner_radius=12,
                font=FUENTE_BASE_NEG, fg_color=AZUL_CLARO,
                hover_color="#C4D2EE", text_color=AZUL_OSCURO,
                command=self._mostrar_ayuda_formato,
            ).pack(side="right")
        return s

    def _campo_ruta(self, marco, fila, texto, variable, al_cambiar=None):
        ctk.CTkLabel(
            marco, text=texto, font=FUENTE_BASE, text_color=GRIS_TXT
        ).grid(row=fila, column=0, sticky="w", padx=(0, 10), pady=4)
        entrada = ctk.CTkEntry(
            marco, textvariable=variable, font=FUENTE_BASE, height=34,
            corner_radius=8, border_color="#C9D2E3",
            fg_color="white", text_color=GRIS_TXT,
        )
        entrada.grid(row=fila, column=1, sticky="ew", padx=(0, 10), pady=4)
        marco.columnconfigure(1, weight=1)
        ctk.CTkButton(
            marco, text="Examinar...",
            command=lambda: self._examinar_carpeta(variable, al_cambiar),
            fg_color=AZUL_CLARO, hover_color="#C4D2EE", text_color=AZUL_OSCURO,
            font=FUENTE_BASE, corner_radius=8, height=32, width=120,
        ).grid(row=fila, column=2, pady=4)
        return entrada

    def _campo_excel(self, marco, fila):
        ctk.CTkLabel(
            marco, text="Excel de correspondencia (opcional):",
            font=FUENTE_BASE, text_color=GRIS_TXT,
        ).grid(row=fila, column=0, sticky="w", padx=(0, 10), pady=4)
        entrada = ctk.CTkEntry(
            marco, textvariable=self.var_mapeo, font=FUENTE_BASE, height=34,
            corner_radius=8, border_color="#C9D2E3",
            fg_color="white", text_color=GRIS_TXT,
        )
        entrada.grid(row=fila, column=1, sticky="ew", padx=(0, 10), pady=4)
        ctk.CTkButton(
            marco, text="Examinar...", command=self._examinar_excel,
            fg_color=AZUL_CLARO, hover_color="#C4D2EE", text_color=AZUL_OSCURO,
            font=FUENTE_BASE, corner_radius=8, height=32, width=120,
        ).grid(row=fila, column=2, pady=4)
        return entrada

    # --- Drag and drop -----------------------------------------------------------

    def _registrar_dnd(self, widget, tipo: str) -> None:
        if not TIENE_DND:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: self._al_soltar(e, tipo))
        except Exception:
            pass

    def _al_soltar(self, evento, tipo: str) -> None:
        try:
            rutas = list(self.root.tk.splitlist(evento.data))
        except Exception:
            rutas = [evento.data.strip("{}")]
        if not rutas:
            return
        ruta = Path(rutas[0])
        if tipo == "excel":
            if ruta.is_file() and ruta.suffix.lower() in (".xlsx", ".xls"):
                self.var_mapeo.set(str(ruta))
                self._persistir_config()
                self._validar_excel_seleccionado()
            return
        if ruta.is_dir():
            if tipo == "carpeta_origen":
                self.var_origen.set(str(ruta))
                self._vista_previa_carpeta()
            else:
                self.var_destino.set(str(ruta))
            self._persistir_config()

    # --- Quick wins ------------------------------------------------------------

    def _vista_previa_carpeta(self) -> None:
        ruta = self.var_origen.get().strip()
        if not ruta or not Path(ruta).is_dir():
            self.lbl_vista_carpeta.configure(text="")
            return
        try:
            regs = core.inventariar(Path(ruta))
            anom = len(core.listar_anomalias(regs))
            ops = len(core.ids_origen_unicos(regs))
            self.lbl_vista_carpeta.configure(
                text=f"{len(regs)} archivos  |  {ops} operaciones  |  "
                     f"{anom} con problemas"
            )
        except Exception:
            self.lbl_vista_carpeta.configure(text="")

    def _validar_excel_seleccionado(self) -> None:
        ruta = self.var_mapeo.get().strip()
        if not ruta:
            return
        if not Path(ruta).is_file():
            self._estado("El Excel seleccionado no existe en esa ruta.")
            return
        mapeo, mensaje_error = err.leer_mapeo_con_reintentos(
            Path(ruta), intentos=2, espera=1.0,
            modo=self.modo, regex_origen=self.regex_origen_cfg or None,
            regex_destino=self.regex_destino_cfg or None,
        )
        if mensaje_error:
            messagebox.showwarning(TITULO, mensaje_error)
            return
        if not mapeo:
            messagebox.showwarning(TITULO, err.MSJ_EXCEL_SIN_DATOS)
            return
        self._estado(
            f"Excel valido: {len(mapeo)} pares de correspondencia detectados."
        )

    def _mostrar_ayuda_formato(self) -> None:
        mensaje = (
            "FORMATO DE SALIDA - como armar el nombre final\n\n"
            "El nombre se construye con una plantilla de piezas:\n"
            "  {id} = identificador de destino (del Excel de correspondencia)\n"
            "  {n}  = consecutivo automatico (1, 2, 3...) cuando una\n"
            "        operacion recibe varios documentos\n"
            "  {c}  = codigo fijo que usted escribe (ej. ZREE)\n\n"
            "Ejemplos:\n"
            "  {id}-{c}-{n}  con codigo ZREE -> 9900000001-ZREE-1.pdf\n"
            "  {id}-{n}                      -> 9900000001-1.pdf\n"
            "  {id}                          -> 9900000001.pdf\n\n"
            "Para el caso ZREE, copie estos dos valores:\n"
            "  Patron de nombre final:  {id}-{c}-{n}\n"
            "  Codigo fijo:             ZREE\n\n"
            "El ejemplo en vivo muestra la transformacion mientras edita."
        )
        messagebox.showinfo("Ayuda - Formato de salida", mensaje)

    def _guardar_config_patron(self, patron: str, codigo: str) -> None:
        self._persistir_config(patron=patron, codigo=codigo)

    def _persistir_config(self, patron=None, codigo=None) -> None:
        datos = core.cargar_config(self.ruta_config)
        if patron is not None:
            datos["patron"] = patron
        if codigo is not None:
            datos["codigo"] = codigo
        self.regex_origen_cfg = self.regex_origen_cfg  # preservado
        datos["origen"] = self.var_origen.get().strip()
        datos["destino"] = self.var_destino.get().strip()
        datos["mapeo"] = self.var_mapeo.get().strip()
        datos["modo"] = self.modo
        datos["regex_origen"] = self.regex_origen_cfg
        datos["regex_destino"] = self.regex_destino_cfg
        core.guardar_config(self.ruta_config, datos)

    # --- Formato de salida -------------------------------------------------------

    def _insertar_token(self, token: str) -> None:
        self.var_patron.set(self.var_patron.get() + token)
        self._actualizar_ejemplo()

    def _actualizar_ejemplo(self) -> None:
        patron = self.var_patron.get()
        valido, mensaje = core.validar_patron(patron)
        codigo = self.var_codigo.get().strip()
        core.CODIGO_FIJO = codigo
        original = "2026XX840820X1 ACME CORPORACION.pdf"
        if valido:
            try:
                final = core.generar_nombre(patron, "9900000001", 1)
                self.lbl_ejemplo.configure(
                    text=f"Original : {original}\nFinal    : {final}"
                )
                self.lbl_patron_estado.configure(text="", text_color=VERDE)
                self._guardar_config_patron(patron, codigo)
            except Exception as exc:
                self.lbl_ejemplo.configure(text=f"Error generando ejemplo: {exc}")
                self.lbl_patron_estado.configure(text="", text_color=VERDE)
        else:
            self.lbl_ejemplo.configure(
                text=f"Original : {original}\nFinal    : ..."
            )
            self.lbl_patron_estado.configure(
                text=f"Formato invalido: {mensaje}", text_color="#9C1C1C"
            )

    # --- Tabla de resultados ------------------------------------------------------

    def _estado_item(self, item: dict) -> str:
        if item["estado"] == core.OK and item["destino"]:
            return "Listo"
        if item["estado"] == core.SIN_MAPEO:
            return "Sin correspondencia"
        if item["estado"] == core.CONFLICTO:
            return "Mapeo conflictivo"
        if item["estado"] == core.ANOMALIA:
            return "Anomalia"
        return "Pendiente"

    def _refrescar_tabla(self) -> None:
        filtro = self.var_filtro.get()
        busqueda = self.ent_buscar.get().strip().lower()

        self.tabla.delete(*self.tabla.get_children())
        mostrados = 0
        for item in self.plan:
            listo = item["estado"] == core.OK and bool(item["destino"])
            pendiente = item["estado"] in (core.SIN_MAPEO, core.CONFLICTO) or (
                item["estado"] == core.OK and not item["destino"]
            )
            anomalia = item["estado"] == core.ANOMALIA

            if filtro == "Listos" and not listo:
                continue
            if filtro == "Pendientes" and not pendiente:
                continue
            if filtro == "Anomalias" and not anomalia:
                continue

            texto = " ".join(
                str(x) for x in (
                    item["archivo"], item["destino"] or "",
                    self._estado_item(item), item["detalle"] or "",
                )
            ).lower()
            if busqueda and busqueda not in texto:
                continue

            tag = "ok" if listo else ("err" if anomalia else "pend")
            detalle_mostrar = item["detalle"] or ""
            if not listo and detalle_mostrar:
                detalle_mostrar = (
                    f"{detalle_mostrar}  |  Accion: "
                    f"{core.accion_corta(item['detalle'])}"
                )
            self.tabla.insert(
                "", "end",
                values=(
                    item["archivo"],
                    item["destino"] or "",
                    self._estado_item(item),
                    detalle_mostrar,
                ),
                tags=(tag,),
            )
            mostrados += 1

        total = len(self.plan)
        self._estado(
            f"Mostrando {mostrados} de {total} archivo(s)  |  filtro: {filtro}"
        )

    # --- Examina rutas ------------------------------------------------------------

    def _examinar_carpeta(self, variable, al_cambiar=None) -> None:
        ruta = filedialog.askdirectory(title="Seleccione carpeta")
        if ruta:
            variable.set(ruta)
            if al_cambiar:
                al_cambiar()
            self._persistir_config()

    def _examinar_excel(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccione Excel de correspondencia",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if ruta:
            self.var_mapeo.set(ruta)
            self._persistir_config()
            self._validar_excel_seleccionado()

    def _centrar(self) -> None:
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"+{x}+{y}")

    # --- Deteccion de identificadores (v1.3.0) ----------------------------------

    ETIQUETAS_MODO = {
        "auto": "Universal (recomendado)",
        "radicado": "Radicado estricto",
        "custom": "Personalizado (regex)",
    }

    def _regex_cargadas(self):
        """Compila las regex del modo personalizado (o None)."""
        if self.modo != core.MODO_CUSTOM:
            return None, None
        origen = core._compilar_regex(self.regex_origen_cfg or "", "origen")
        destino = None
        if self.regex_destino_cfg:
            destino = core._compilar_regex(self.regex_destino_cfg, "destino")
        return origen, destino

    def _ventana_avanzada(self) -> None:
        ventana = ctk.CTkToplevel(self.root)
        ventana.title("Opciones avanzadas de deteccion")
        ventana.geometry("620x560+120+80")
        ventana.configure(fg_color=FONDO)
        ventana.transient(self.root)
        ventana.grab_set()

        ctk.CTkLabel(
            ventana,
            text="MODO DE DETECCION DE IDENTIFICADORES",
            font=FUENTE_SECCION, text_color=AZUL_OSCURO,
        ).pack(anchor="w", padx=20, pady=(18, 4))
        ctk.CTkLabel(
            ventana,
            text="Universal reconoce radicados, codigos y nombres libres.\n"
                 "Cambie aqui solo si necesita una deteccion exacta.",
            font=FUENTE_ESTADO, text_color=SEC,
            justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        var_modo = ctk.StringVar(value=self.ETIQUETAS_MODO[self.modo])
        menu = ctk.CTkOptionMenu(
            ventana, variable=var_modo,
            values=list(self.ETIQUETAS_MODO.values()),
            font=FUENTE_BASE, height=34, fg_color=AZUL_MEDIO,
            button_color=AZUL_OSCURO, button_hover_color=AZUL_OSCURO,
        )
        menu.pack(fill="x", padx=20, pady=(0, 12))

        marco_reg = ctk.CTkFrame(ventana, fg_color="white", corner_radius=10)
        marco_reg.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        marco_reg.columnconfigure(1, weight=1)
        marco_reg.rowconfigure(0, weight=0)

        lbl = ctk.CTkLabel(
            marco_reg, text="Regex de ORIGEN (coincide con el inicio del nombre):",
            font=FUENTE_ESTADO, text_color=GRIS_TXT,
        )
        lbl.grid(row=0, column=0, sticky="w", padx=14, pady=(12, 3))
        ent_reo = ctk.CTkEntry(
            marco_reg, font=("Consolas", 12), height=32, corner_radius=8,
            fg_color="white", text_color=GRIS_TXT,
            placeholder_text=r"ej. EXP-\d{4}-\d{3}",
        )
        ent_reo.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14,
                     pady=(0, 8))
        ent_reo.insert(0, self.regex_origen_cfg or "")

        ctk.CTkLabel(
            marco_reg, text="Regex de DESTINO (coincide con la celda destino):",
            font=FUENTE_ESTADO, text_color=GRIS_TXT,
        ).grid(row=2, column=0, sticky="w", padx=14, pady=(2, 3))
        ent_red = ctk.CTkEntry(
            marco_reg, font=("Consolas", 12), height=32, corner_radius=8,
            fg_color="white", text_color=GRIS_TXT,
            placeholder_text=r"ej. (CT|EXP)-\d{4}-[0-9A-Z]+",
        )
        ent_red.grid(row=3, column=0, columnspan=2, sticky="ew", padx=14,
                     pady=(0, 8))
        ent_red.insert(0, self.regex_destino_cfg or "")

        resultado = ctk.CTkLabel(
            marco_reg, text="", font=("Consolas", 11), text_color=GRIS_TXT,
            justify="left", anchor="w",
        )
        resultado.grid(row=4, column=0, columnspan=2, sticky="ew", padx=14,
                       pady=(2, 12))

        def actualizar_estado_reg():
            es_custom = (
                var_modo.get() == self.ETIQUETAS_MODO[core.MODO_CUSTOM]
            )
            estado = "normal" if es_custom else "disabled"
            gris = "#8C6209" if not es_custom else SEC
            ctk.CTkLabel(
                marco_reg,
                text=("Solo se usan en el modo Personalizado"
                      if not es_custom else ""),
                font=FUENTE_ESTADO, text_color=gris,
            ).grid(row=0, column=1, sticky="e", pady=(12, 3))

        def probar():
            modo_sel = {v: k for k, v in self.ETIQUETAS_MODO.items()}.get(
                var_modo.get(), core.MODO_AUTO
            )
            nombre = ent_prueba.get().strip() or "2026EE963171O1 1.pdf"
            lineas = [f"Archivo de prueba : {nombre}"]
            try:
                regex_or = (
                    core._compilar_regex(ent_reo.get().strip(), "origen")
                    if modo_sel == core.MODO_CUSTOM else None
                )
                id_detectado = core.extraer_id_origen_flexible(
                    nombre, modo_sel, regex_or
                )
                lineas.append(f"Origen detectado  : {id_detectado or '(no detectado)'}")
            except ValueError as exc:
                lineas.append(f"Error             : {exc}")

            ruta_mapeo = self.var_mapeo.get().strip()
            if ruta_mapeo and Path(ruta_mapeo).is_file():
                try:
                    regex_de = (
                        core._compilar_regex(ent_red.get().strip(), "destino")
                        if modo_sel == core.MODO_CUSTOM and ent_red.get().strip()
                        else None
                    )
                    mapeo, conflictos, informe = core.cargar_mapeo_detallado(
                        Path(ruta_mapeo), modo_sel, regex_or, regex_de
                    )
                    lineas.append(
                        f"Excel             : {informe['pares']} pares, "
                        f"{informe['conflictos']} conflictos, "
                        f"{len(informe['saltadas'])} filas saltadas"
                    )
                except ValueError as exc:
                    lineas.append(f"Excel             : error: {exc}")
                except Exception:
                    lineas.append(
                        "Excel             : no se pudo leer (abierto o bloqueado)"
                    )
            resultado.configure(text="\n".join(lineas))

        ctk.CTkLabel(
            marco_reg, text="Archivo de prueba (nombre tal como llega):",
            font=FUENTE_ESTADO, text_color=GRIS_TXT,
        ).grid(row=5, column=0, sticky="w", padx=14, pady=(6, 3))
        ent_prueba = ctk.CTkEntry(
            marco_reg, font=("Consolas", 12), height=32, corner_radius=8,
            fg_color="white", text_color=GRIS_TXT,
        )
        ent_prueba.grid(row=6, column=0, columnspan=2, sticky="ew", padx=14,
                        pady=(0, 8))
        ent_prueba.insert(0, "2026EE963171O1 1.pdf")
        ctk.CTkButton(
            marco_reg, text="Probar deteccion", command=probar,
            fg_color=AZUL_MEDIO, hover_color=AZUL_OSCURO, text_color="white",
            font=FUENTE_BASE, corner_radius=8, height=32, width=150,
        ).grid(row=7, column=0, sticky="w", padx=14, pady=(0, 12))

        actualizar_estado_reg()
        menu.configure(command=lambda _v: actualizar_estado_reg())

        def guardar():
            modo_sel = {v: k for k, v in self.ETIQUETAS_MODO.items()}.get(
                var_modo.get(), core.MODO_AUTO
            )
            self.modo = modo_sel
            self.regex_origen_cfg = ent_reo.get().strip()
            self.regex_destino_cfg = ent_red.get().strip()
            self._persistir_config()
            ventana.destroy()
            self._estado(f"Modo de deteccion: {self.ETIQUETAS_MODO[modo_sel]}")

        ctk.CTkButton(
            ventana, text="Guardar", command=guardar,
            fg_color=VERDE, hover_color=VERDE_HOVER, text_color="white",
            font=FUENTE_BASE_NEG, corner_radius=8, height=36, width=140,
        ).pack(pady=(0, 14))

    def _verificar_excel(self) -> None:
        """Dictamen estructural del Excel de correspondencia seleccionado."""
        ruta = self.var_mapeo.get().strip()
        if not ruta:
            messagebox.showinfo(
                TITULO, "Primero seleccione el Excel de correspondencia."
            )
            return
        if not Path(ruta).is_file():
            messagebox.showwarning(
                TITULO,
                "El Excel seleccionado no existe en esa ruta.\n\n"
                "Vuelva a seleccionarlo con 'Examinar...'.",
            )
            return

        try:
            regex_or, regex_de = self._regex_cargadas()
        except ValueError as exc:
            messagebox.showwarning(TITULO, str(exc))
            return

        mapeo, conflictos, informe = None, None, None
        try:
            rutas = self._rutas_validas()
            _src, _dst = rutas if rutas else (None, None)
            mapeo, conflictos, informe = core.cargar_mapeo_detallado(
                Path(ruta), self.modo, regex_or, regex_de
            )
        except ValueError as exc:
            messagebox.showwarning(TITULO, str(exc))
            return
        except Exception as exc:
            messagebox.showerror(TITULO, err.traducir_error_excel(exc))
            return

        lineas = [
            "VERIFICACION DEL EXCEL",
            f"Archivo        : {Path(ruta).name}",
            f"Modo           : {self.ETIQUETAS_MODO[self.modo]}",
            f"Filas leidas   : {informe['filas']}",
            f"Pares validos  : {informe['pares']}",
            f"Conflictos     : {informe['conflictos']}",
        ]
        if informe["saltadas"]:
            lineas.append(f"Filas saltadas : {len(informe['saltadas'])}")
            for num, motivo in informe["saltadas"][:6]:
                lineas.append(f"   fila {num}: {motivo}")
            if len(informe["saltadas"]) > 6:
                lineas.append(
                    f"   ...y {len(informe['saltadas']) - 6} mas"
                )
        for origen_id in list(conflictos)[:3]:
            lineas.append(
                f"   [CONFLICTO] {origen_id} -> "
                f"{', '.join(conflictos[origen_id])}"
            )
        messagebox.showinfo(TITULO, "\n".join(lineas))
        self._log("")
        self._log("VERIFICACION DEL EXCEL:", "fase")
        for linea in lineas[2:]:
            tag = "sinmapeo" if "CONFLICTO" in linea else "info"
            self._log(f"  {linea}", tag)

    def _crear_plantilla(self) -> None:
        ruta = filedialog.asksaveasfilename(
            title="Guardar plantilla de correspondencia",
            defaultextension=".xlsx",
            initialfile="plantilla_correspondencia.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            core.crear_plantilla_excel(Path(ruta))
        except Exception as exc:
            messagebox.showerror(
                TITULO, f"No se pudo crear la plantilla: {exc}"
            )
            return
        self._log("")
        self._log(f"Plantilla creada: {ruta}", "ruta")
        self._estado("Plantilla de correspondencia creada")

    # --- Utilidades de interfaz -------------------------------------------------------

    def _log(self, msg: str, tag: str = "info") -> None:
        self.txt._textbox.insert("end", msg + "\n", tag)
        self.txt._textbox.see("end")

    def _linea_fase(self, titulo: str) -> None:
        self._log("", "info")
        self._log("=" * 72, "fase")
        self._log(titulo, "fase")
        self._log("=" * 72, "fase")

    def _estado(self, msg: str) -> None:
        self.var_estado.set(msg)

    def _progreso(self, actual: int, total: int) -> None:
        frac = (actual / total) if total else 0
        self.progreso.set(frac)
        self._estado(f"Procesando... {actual}/{total} ({int(frac * 100)}%)")

    def _rutas_validas(self, con_mapeo: bool = False):
        patron = self.var_patron.get()
        valido, mensaje = core.validar_patron(patron)
        if not valido:
            self._estado(f"Formato de salida invalido: {mensaje}")
            return None
        core.CODIGO_FIJO = self.var_codigo.get().strip()

        origen = self.var_origen.get().strip()
        destino = self.var_destino.get().strip()
        if not origen or not Path(origen).is_dir():
            messagebox.showwarning(
                TITULO, "Seleccione una carpeta ORIGEN valida."
            )
            return None
        if not destino:
            messagebox.showwarning(TITULO, "Seleccione la carpeta DESTINO.")
            return None
        if con_mapeo and not self.var_mapeo.get().strip():
            messagebox.showwarning(
                TITULO,
                "Para simular o ejecutar necesita el Excel de correspondencia "
                "(origen -> destino).",
            )
            return None
        return Path(origen).resolve(), Path(destino).resolve()

    def _bloquear(self, ocupado: bool) -> None:
        estado = "disabled" if ocupado else "normal"
        self.btn_inv.configure(state=estado)
        self.btn_sim.configure(state=estado)
        self.btn_ejec.configure(state=estado)
        self.btn_cancelar.configure(
            state="normal" if ocupado else "disabled",
            text="Cancelar",
        )
        if ocupado:
            self._cancelar = False

    def _accion_cancelar(self) -> None:
        self._cancelar = True
        self.btn_cancelar.configure(text="Cancelando...", state="disabled")
        self._estado("Cancelando: se detendra al terminar el archivo actual...")

    # --- Acciones -------------------------------------------------------------------------

    def _accion_inventariar(self) -> None:
        try:
            regex_or, _regex_de = self._regex_cargadas()
        except ValueError as exc:
            messagebox.showwarning(TITULO, str(exc))
            return
        rutas = self._rutas_validas()
        if not rutas:
            return
        origen, destino = rutas
        self._bloquear(True)
        self.txt.delete("1.0", "end")
        threading.Thread(
            target=self._hilo_inventariar, args=(origen, destino, regex_or),
            daemon=True,
        ).start()

    def _hilo_inventariar(self, origen: Path, destino: Path,
                          regex_or=None) -> None:
        try:
            self._linea_fase(f"INVENTARIO  |  {origen}")
            self.registros = core.inventariar(
                origen, progreso=self._progreso,
                cancelado=lambda: self._cancelar,
                modo=self.modo, regex_origen=regex_or,
            )
            anom = core.listar_anomalias(self.registros)
            validos = len(self.registros) - len(anom)
            n_ids = len(core.ids_origen_unicos(self.registros))

            self._log("")
            self._log(f"Archivos leidos      : {len(self.registros)}", "info")
            self._log(f"Validos              : {validos}", "ok")
            self._log(f"Con anomalia         : {len(anom)}",
                      "anomalia" if anom else "ok")
            self._log(f"Ids de origen unicos : {n_ids}", "info")

            for r in anom[:20]:
                self._log(
                    f"  [ANOMALIA] {r['archivo']} -> {r['anomalia']}",
                    "anomalia",
                )
            if len(anom) > 20:
                self._log(f"  ... y {len(anom) - 20} mas", "anomalia")

            if n_ids:
                ruta_lista = destino / "ids_para_consulta.txt"
                core.generar_lista_ids(self.registros, ruta_lista)
                self._log("")
                self._log(
                    f"Lista para consulta (lotes de {core.LOTE_CONSULTA}): "
                    f"{ruta_lista}", "ruta",
                )

            self.plan = core.construir_plan(
                self.registros, None, self.var_patron.get()
            )
            self._refrescar_tabla()
            self._log("")
            self._log(
                "Siguiente paso: obtenga el Excel de correspondencia, "
                "luego SIMULAR.", "fase",
            )
            self._estado(
                f"Inventario listo: {validos} validos | {len(anom)} anomalias "
                f"| {n_ids} ids"
            )
            self.btn_sim.configure(state="normal")
            self.btn_ejec.configure(state="disabled")
            self._bloquear(False)
        except core.ProcesoCancelado as exc:
            self._log(f"CANCELADO: {exc}", "sinmapeo")
            self._estado("Inventario cancelado")
            self._bloquear(False)
            self.btn_sim.configure(state="disabled")
        except Exception as exc:
            self._log(f"ERROR en inventario: {type(exc).__name__}: {exc}",
                      "error")
            self._estado("Error en inventario")
            self._bloquear(False)

    def _accion_simular(self) -> None:
        if not self.registros:
            messagebox.showinfo(
                TITULO, "Primero ejecute INVENTARIAR con la carpeta de PDFs."
            )
            return
        rutas = self._rutas_validas(con_mapeo=True)
        if not rutas:
            return
        origen, destino = rutas
        ruta_excel = Path(self.var_mapeo.get().strip())
        if not ruta_excel.is_file():
            messagebox.showwarning(TITULO, "El Excel de mapeo no existe.")
            return

        try:
            regex_or, regex_de = self._regex_cargadas()
        except ValueError as exc:
            messagebox.showwarning(TITULO, str(exc))
            return
        self._bloquear(True)
        self._estado("Leyendo Excel de correspondencia...")
        threading.Thread(
            target=self._hilo_simular,
            args=(origen, destino, ruta_excel, regex_or, regex_de),
            daemon=True,
        ).start()

    def _hilo_simular(self, origen: Path, destino: Path,
                      ruta_excel: Path, regex_or=None, regex_de=None) -> None:
        try:
            patron = self.var_patron.get()
            core.CODIGO_FIJO = self.var_codigo.get().strip()

            if self.mapeo is None or self.mapeo_ruta != str(ruta_excel):
                try:
                    mapeo, conflictos, informe = core.cargar_mapeo_detallado(
                        ruta_excel, self.modo, regex_or, regex_de
                    )
                except ValueError as exc:
                    self._log(f"CONFIGURACION DE DETECCION INVALIDA: {exc}",
                              "error")
                    self._estado("Deteccion mal configurada")
                    self.btn_sim.configure(state="normal")
                    self.btn_inv.configure(state="normal")
                    self._bloquear(False)
                    messagebox.showwarning(TITULO, str(exc))
                    return
                self.mapeo = mapeo
                self.conflictos = conflictos
                self.mapeo_ruta = str(ruta_excel)

            self._log("")
            self._log(
                f"Mapeo cargado: {len(self.mapeo)} pares origen -> destino",
                "ok",
            )
            if self.conflictos:
                self._log(
                    f"Mapeos conflictivos detectados: {len(self.conflictos)} "
                    "origen(es) con varios destinos", "anomalia",
                )
                for origen_id, destinos in list(self.conflictos.items())[:10]:
                    self._log(
                        f"  [CONFLICTO] {origen_id} -> {', '.join(destinos)}",
                        "anomalia",
                    )

            if not self.mapeo and not self.conflictos:
                self._log("")
                self._log("EL EXCEL NO CONTIENE CORRESPONDENCIA:", "error")
                for linea in err.MSJ_EXCEL_SIN_DATOS.splitlines():
                    self._log(f"  {linea}", "error")
                self._estado("Excel sin datos de correspondencia")
                self.btn_sim.configure(state="normal")
                self.btn_inv.configure(state="normal")
                messagebox.showwarning(TITULO, err.MSJ_EXCEL_SIN_DATOS)
                return

            self.plan = core.construir_plan(
                self.registros, self.mapeo, patron, self.conflictos
            )
            res = core.resumen_plan(self.plan)

            self._log("")
            self._log(
                f"PLAN: {res['ok']} listos | {res['sin_mapeo']} sin "
                f"correspondencia | {res['conflictos']} con conflicto | "
                f"{res['anomalias']} anomalias", "info",
            )
            self._log("")

            permitido, avisos = core.validar_plan(self.plan, destino)
            for aviso in avisos[:20]:
                tag = "anomalia" if aviso.startswith("COLISION") else "sinmapeo"
                self._log(f"  {aviso}", tag)

            if not permitido:
                self._log("")
                self._log(
                    "NO SE PUEDE EJECUTAR: hay COLISIONES de nombre destino. "
                    "Corrija los datos y vuelva a simular.", "error",
                )
                self._estado("Simulacion con colisiones - no ejecutable")
                self.btn_ejec.configure(state="disabled")
            else:
                pendientes = (
                    res["sin_mapeo"] + res["conflictos"] + res["anomalias"]
                )
                self._log("")
                if pendientes:
                    self._log(
                        f"SIMULACION OK: {res['ok']} se pueden renombrar. "
                        f"{pendientes} quedaran pendientes con registro y "
                        "sugerencia (pendientes_correccion.csv).", "ok",
                    )
                else:
                    self._log(
                        "SIMULACION CORRECTA: todos los archivos se pueden "
                        "renombrar.", "ok",
                    )
                self._estado("Simulacion correcta - listo para EJECUTAR")
                self.btn_ejec.configure(state="normal")

            self._refrescar_tabla()
            self.btn_sim.configure(state="normal")
            self.btn_inv.configure(state="normal")
            self._bloquear(False)
            # El boton EJECUTAR debe reflejar la validacion, no el reset
            if not permitido:
                self.btn_ejec.configure(state="disabled")
        except Exception as exc:
            self._log(
                f"ERROR inesperado en simulacion: {type(exc).__name__}: {exc}",
                "error",
            )
            self._estado("Error en simulacion")
            self.btn_sim.configure(state="normal")
            self.btn_inv.configure(state="normal")
            self._bloquear(False)

    def _accion_ejecutar(self) -> None:
        rutas = self._rutas_validas(con_mapeo=True)
        if not rutas:
            return
        origen, destino = rutas

        res = core.resumen_plan(self.plan)
        pendientes = res["sin_mapeo"] + res["conflictos"] + res["anomalias"]

        if pendientes:
            mensaje = (
                f"{res['ok']} de {res['total']} archivos se pueden renombrar.\n\n"
                f"{pendientes} tienen problemas y NO se renombraran: quedaran "
                f"registrados en 'pendientes_correccion.csv' con la accion "
                f"sugerida para corregirlos.\n\n"
                f"Desea renombrar los {res['ok']} validos?"
            )
            if not messagebox.askyesno(TITULO, mensaje):
                return
        else:
            if not messagebox.askyesno(
                TITULO,
                f"Se copiaran {res['ok']} archivo(s) renombrados a:\n\n"
                f"{destino}\n\n"
                "Los originales NO se tocan. Desea continuar?",
            ):
                return

        self._bloquear(True)
        self._linea_fase(f"EJECUTANDO copia  |  {destino}")
        threading.Thread(
            target=self._hilo_ejecutar, args=(origen, destino), daemon=True
        ).start()

    def _hilo_ejecutar(self, origen: Path, destino: Path) -> None:
        try:
            copiados, omitidos = core.ejecutar_copias(
                self.plan, origen, destino, progreso=self._progreso,
                cancelado=lambda: self._cancelar,
            )
            core.escribir_reporte(self.plan, destino / "reporte.csv")

            n_pend, n_ids = core.escribir_pendientes(
                self.plan, destino / "pendientes_correccion.csv",
                destino / "ids_pendientes.txt",
            )
            core.generar_lista_ids(
                self.registros, destino / "ids_para_consulta.txt"
            )

            self._log("")
            self._log(
                f"LISTO: {copiados} copiado(s), {omitidos} omitido(s) "
                f"(ya existian).", "ok",
            )
            if n_pend:
                self._log(
                    f"Pendientes con seguimiento: {n_pend} "
                    f"(ids para re-consulta: {n_ids})", "sinmapeo",
                )
                self._log(
                    f"Registro: {destino / 'pendientes_correccion.csv'}",
                    "ruta",
                )
            self._log(f"Reporte: {destino / 'reporte.csv'}", "ruta")
            self._log("")
            self._log(
                "Siguiente paso: cargue la carpeta destino en el sistema. "
                "Si hubo pendientes, corrigalos y re-ejecute el mismo lote: "
                "los ya hechos se omiten solos.", "fase",
            )
            resumen = f"Completado: {copiados} copiados, {omitidos} omitidos"
            if n_pend:
                resumen += f", {n_pend} pendientes"
            self._estado(resumen)
        except core.ProcesoCancelado as exc:
            self._log(f"CANCELADO: {exc}", "sinmapeo")
            self._log(
                "Los archivos copiados hasta este punto se conservan; el "
                "reporte reflejara el estado final.", "info",
            )
            self._estado("Ejecucion cancelada")
        except Exception as exc:
            self._log(
                f"ERROR en ejecucion: {type(exc).__name__}: {exc}", "error"
            )
            self._estado("Error en ejecucion")
        finally:
            self._bloquear(False)

    # --- Acciones auxiliares ---------------------------------------------------------

    def _accion_deshacer(self) -> None:
        destino = self.var_destino.get().strip()
        origen = self.var_origen.get().strip()
        if not destino or not Path(destino).is_dir():
            messagebox.showinfo(
                TITULO, "Seleccione la carpeta destino para deshacer."
            )
            return
        ruta_reporte = Path(destino) / "reporte.csv"
        if not ruta_reporte.is_file():
            messagebox.showinfo(
                TITULO,
                "No hay reporte.csv en la carpeta destino: no hay nada que "
                "deshacer de una ejecucion anterior.",
            )
            return
        if not messagebox.askyesno(
            TITULO,
            "Se eliminaran SOLO las copias registradas en reporte.csv de esta "
            "carpeta destino (los originales no se tocan).\n\n"
            "Desea continuar?",
        ):
            return
        eliminados, errores = core.deshacer_desde_reporte(
            ruta_reporte, Path(origen) if origen else Path("."), Path(destino)
        )
        self._log("")
        self._log(f"DESHACER: {eliminados} copia(s) eliminada(s).", "ok")
        for e in errores[:10]:
            self._log(f"  {e}", "anomalia")
        self._estado(f"Deshacer: {eliminados} eliminados, {len(errores)} avisos")

    def _copiar_lista(self) -> None:
        ruta = self.var_destino.get().strip() or self.var_origen.get().strip()
        archivo = Path(ruta) / "ids_para_consulta.txt" if ruta else None
        if archivo and archivo.is_file():
            texto = archivo.read_text(encoding="utf-8")
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self._estado("Lista de ids copiada al portapapeles.")
        else:
            messagebox.showinfo(
                TITULO,
                "Aun no existe la lista. Ejecute INVENTARIAR primero.",
            )

    def _abrir_destino(self) -> None:
        ruta = self.var_destino.get().strip()
        if ruta and Path(ruta).is_dir():
            subprocess.Popen(["explorer", ruta])

    def _abrir_reporte(self) -> None:
        ruta = self.var_destino.get().strip()
        archivo = Path(ruta) / "reporte.csv" if ruta else None
        if archivo and archivo.is_file():
            subprocess.Popen(["notepad", str(archivo)])
        else:
            messagebox.showinfo(
                TITULO,
                "Aun no existe reporte.csv. Ejecute EJECUTAR primero.",
            )


def main() -> None:
    global TIENE_DND
    habilitar_dpi_awareness()
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # La raiz siempre es customtkinter (soporta fg_color y temas).
    # El soporte de arrastrar y soltar se carga sobre esa misma raiz.
    root = ctk.CTk()
    if TIENE_DND:
        try:
            _cargar_tkdnd(root)
        except Exception:
            TIENE_DND = False

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()