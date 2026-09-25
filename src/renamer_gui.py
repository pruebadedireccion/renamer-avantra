# -*- coding: utf-8 -*-
"""
Renamer Avantra - Interfaz grafica (v1.1.0).

Novedades v1.1.0:
    - Layout responsivo (pie adaptable, log con scroll nativo, DPI aware)
    - Errores de Excel traducidos a instrucciones, con reintentos
    - Ejecucion parcial con seguimiento (pendientes_correccion.csv)
    - Quick wins: vista previa de carpeta, validacion de Excel al elegirlo,
      recordatorio de sesion y ayuda del formato de salida
"""

from __future__ import annotations

import ctypes
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

sys.path.insert(0, str(Path(__file__).resolve().parent))

import renamer_core as core
import errores as err

TITULO = "Renamer Avantra - Renombrado Masivo de Documentos"

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


def habilitar_dpi_awareness() -> None:
    """Nitidez en pantallas con escalado (Windows 8.1+)."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def tamano_inicial() -> str:
    """Ventana inicial ~60% del area util del monitor, con limites."""
    try:
        usuario32 = ctypes.windll.user32
        ancho = usuario32.GetSystemMetrics(0)
        alto = usuario32.GetSystemMetrics(1)
        w = max(860, min(int(ancho * 0.62), 1280))
        h = max(640, min(int(alto * 0.68), 900))
        return f"{w}x{h}"
    except Exception:
        return "1020x760"


class App:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title(TITULO)
        self.root.geometry(tamano_inicial())
        self.root.minsize(760, 560)
        self.root.configure(fg_color=FONDO)

        self.var_origen = ctk.StringVar()
        self.var_destino = ctk.StringVar()
        self.var_mapeo = ctk.StringVar()
        self.var_patron = ctk.StringVar(value=core.PATRON_DEFAULT)
        self.var_codigo = ctk.StringVar(value="")

        # Configuracion persistente (patron, codigo y rutas de la sesion)
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

        self.plan: list[dict] = []
        self.registros: list[dict] = []
        self.mapeo: dict[str, str] | None = None
        self.mapeo_ruta: str = ""

        self._construir_widgets()
        self._centrar()

        self._log("Bienvenido.", "fase")
        self._log("1. Seleccione la carpeta de PDFs y la carpeta destino.", "info")
        self._log("2. Defina el FORMATO DE SALIDA y presione INVENTARIAR.", "info")
        self._log(
            "Con pendientes: EJECUTAR renombra los validos y deja registro "
            "de los demas (pendientes_correccion.csv).",
            "info",
        )

    # --- Widgets -------------------------------------------------------------

    def _construir_widgets(self) -> None:
        marco = ctk.CTkFrame(self.root, fg_color="transparent")
        marco.pack(fill="both", expand=True)

        # Encabezado
        cabezal = ctk.CTkFrame(marco, fg_color=AZUL_OSCURO, corner_radius=0, height=64)
        cabezal.pack(fill="x")
        cabezal.pack_propagate(False)
        ctk.CTkLabel(
            cabezal, text="RENAMER AVANTRA", font=FUENTE_TITULO, text_color="white"
        ).pack(side="left", padx=(20, 8), pady=10)
        ctk.CTkLabel(
            cabezal,
            text="Renombrado masivo de documentos  |  By: Red Avantra \u00ae",
            font=FUENTE_SUBTITULO,
            text_color=AZUL_CLARO,
        ).pack(side="left", pady=10)

        cont = ctk.CTkFrame(marco, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=14, pady=8)

        # --- 1. CARPETAS ------------------------------------------------------
        s1 = self._seccion(cont, "1. CARPETAS")
        interno = ctk.CTkFrame(s1, fg_color="white", corner_radius=0)
        interno.pack(fill="x", padx=16, pady=(2, 10))
        self._campo_ruta(
            interno,
            0,
            "Carpeta de PDFs (origen):",
            self.var_origen,
            al_cambiar=self._vista_previa_carpeta,
        )
        self.lbl_vista_carpeta = ctk.CTkLabel(
            interno, text="", font=FUENTE_ESTADO, text_color=AZUL_MEDIO, anchor="w"
        )
        self.lbl_vista_carpeta.grid(
            row=3, column=0, columnspan=3, sticky="w", padx=(0, 10), pady=(0, 4)
        )
        self._campo_ruta(interno, 1, "Carpeta destino (renombrados):", self.var_destino)
        self._campo_excel(interno, 2)
        if self.var_origen.get():
            self._vista_previa_carpeta()
        if self.var_mapeo.get():
            self._validar_excel_seleccionado()

        # --- 2. FORMATO DE SALIDA ---------------------------------------------
        s2 = self._seccion(cont, "2. FORMATO DE SALIDA")
        fila_titulo = ctk.CTkFrame(s2, fg_color="transparent")
        fila_titulo.pack(fill="x", padx=16, pady=(8, 0))
        ctk.CTkButton(
            fila_titulo,
            text="  ?  ",
            width=30,
            height=24,
            corner_radius=12,
            font=FUENTE_BASE_NEG,
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            command=self._mostrar_ayuda_formato,
        ).pack(side="right")

        fmt = ctk.CTkFrame(s2, fg_color="white", corner_radius=0)
        fmt.pack(fill="x", padx=16, pady=(2, 10))
        fmt.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            fmt, text="Patron de nombre final:", font=FUENTE_BASE, text_color=GRIS_TXT
        ).grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.ent_patron = ctk.CTkEntry(
            fmt,
            textvariable=self.var_patron,
            font=("Consolas", 13),
            height=34,
            corner_radius=8,
            border_color="#C9D2E3",
        )
        self.ent_patron.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=4)
        self.ent_patron.bind("<KeyRelease>", lambda _e: self._actualizar_ejemplo())

        ctk.CTkLabel(
            fmt, text="Codigo fijo {c}:", font=FUENTE_BASE, text_color=GRIS_TXT
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        self.ent_codigo = ctk.CTkEntry(
            fmt,
            textvariable=self.var_codigo,
            font=("Consolas", 13),
            height=34,
            corner_radius=8,
            width=180,
            placeholder_text="(opcional, ej. ZREE)",
        )
        self.ent_codigo.grid(row=1, column=1, sticky="w", pady=4)
        self.ent_codigo.bind("<KeyRelease>", lambda _e: self._actualizar_ejemplo())

        piezas = ctk.CTkFrame(fmt, fg_color="transparent")
        piezas.grid(row=2, column=1, sticky="w", pady=(2, 4))
        ctk.CTkLabel(
            piezas, text="Piezas: ", font=FUENTE_ESTADO, text_color=GRIS_TXT
        ).pack(side="left")
        for token in ("{id}", "{n}", "{c}"):
            ctk.CTkButton(
                piezas,
                text=token,
                width=52,
                height=26,
                corner_radius=6,
                font=("Consolas", 12),
                fg_color=AZUL_CLARO,
                hover_color="#C4D2EE",
                text_color=AZUL_OSCURO,
                command=lambda t=token: self._insertar_token(t),
            ).pack(side="left", padx=(0, 6))

        ej = ctk.CTkFrame(s2, fg_color="#F4F6FA", corner_radius=8)
        ej.pack(fill="x", padx=16, pady=(2, 12))
        self.lbl_ejemplo = ctk.CTkLabel(
            ej,
            text="",
            font=("Consolas", 14),
            text_color=GRIS_TXT,
            justify="left",
            anchor="w",
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
        acciones.pack(fill="x", padx=16, pady=(2, 12))
        acciones.columnconfigure((0, 1, 2), weight=1)

        self.btn_inv = ctk.CTkButton(
            acciones,
            text="INVENTARIAR",
            command=self._accion_inventariar,
            fg_color=AZUL_MEDIO,
            hover_color=AZUL_OSCURO,
            text_color="white",
            font=FUENTE_BASE_NEG,
            corner_radius=8,
            height=38,
        )
        self.btn_inv.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.btn_sim = ctk.CTkButton(
            acciones,
            text="SIMULAR (dry-run)",
            command=self._accion_simular,
            fg_color=GRIS_SIM,
            hover_color=GRIS_SIM_HOVER,
            text_color="white",
            font=FUENTE_BASE_NEG,
            corner_radius=8,
            height=38,
            state="disabled",
        )
        self.btn_sim.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self.btn_ejec = ctk.CTkButton(
            acciones,
            text="EJECUTAR (copiar)",
            command=self._accion_ejecutar,
            fg_color=VERDE,
            hover_color=VERDE_HOVER,
            text_color="white",
            font=FUENTE_BASE_NEG,
            corner_radius=8,
            height=38,
            state="disabled",
        )
        self.btn_ejec.grid(row=0, column=2, sticky="ew")

        ctk.CTkButton(
            acciones,
            text="Salir",
            command=self.root.destroy,
            fg_color="#8C8C8C",
            hover_color="#6E6E6E",
            font=FUENTE_BASE,
            corner_radius=8,
            height=38,
            width=110,
        ).grid(row=0, column=3, padx=(8, 0))

        self.progreso = ctk.CTkProgressBar(
            acciones, progress_color=AZUL_MEDIO, height=12
        )
        self.progreso.set(0)
        self.progreso.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))

        # --- 4. RESULTADOS ------------------------------------------------------
        s4 = self._seccion(cont, "4. RESULTADOS")
        self.txt = ctk.CTkTextbox(
            s4,
            font=FUENTE_LOG,
            fg_color="white",
            text_color=GRIS_TXT,
            wrap="none",
        )
        self.txt.pack(fill="both", expand=True, padx=14, pady=(2, 12))
        self.txt._textbox.tag_configure("fase", foreground=AZUL_OSCURO)
        self.txt._textbox.tag_configure("ok", foreground=VERDE)
        self.txt._textbox.tag_configure("error", foreground="#9C1C1C")
        self.txt._textbox.tag_configure("anomalia", foreground="#9C1C1C")
        self.txt._textbox.tag_configure("sinmapeo", foreground="#8C6209")
        self.txt._textbox.tag_configure("info", foreground=GRIS_TXT)
        self.txt._textbox.tag_configure("ruta", foreground=AZUL_MEDIO)

        # --- Pie: altura automatica, botones sin ancho fijo ----------------------
        pie = ctk.CTkFrame(
            self.root,
            fg_color="white",
            corner_radius=0,
            border_width=1,
            border_color="#D6DCE8",
        )
        pie.pack(fill="x", side="bottom")

        self.var_estado = ctk.StringVar(value="Listo.")
        ctk.CTkLabel(
            pie,
            textvariable=self.var_estado,
            font=FUENTE_ESTADO,
            text_color=GRIS_TXT,
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=14, pady=8)

        for texto, comando in (
            ("Copiar lista", self._copiar_lista),
            ("Abrir carpeta destino", self._abrir_destino),
            ("Abrir reporte.csv", self._abrir_reporte),
        ):
            ctk.CTkButton(
                pie,
                text=texto,
                command=comando,
                fg_color=AZUL_CLARO,
                hover_color="#C4D2EE",
                text_color=AZUL_OSCURO,
                font=FUENTE_ESTADO,
                corner_radius=8,
                height=32,
                width=120,
            ).pack(side="right", padx=(6, 12))

    # --- Helpers de construccion -----------------------------------------------

    def _seccion(self, contenedor, titulo) -> ctk.CTkFrame:
        s = ctk.CTkFrame(
            contenedor,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color="#D6DCE8",
        )
        s.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(s, text=titulo, font=FUENTE_SECCION, text_color=AZUL_OSCURO).pack(
            anchor="w", padx=16, pady=(10, 2)
        )
        return s

    def _campo_ruta(self, marco, fila, texto, variable, al_cambiar=None):
        ctk.CTkLabel(marco, text=texto, font=FUENTE_BASE, text_color=GRIS_TXT).grid(
            row=fila, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ctk.CTkEntry(
            marco,
            textvariable=variable,
            font=FUENTE_BASE,
            height=34,
            corner_radius=8,
            border_color="#C9D2E3",
        ).grid(row=fila, column=1, sticky="ew", padx=(0, 10), pady=4)
        marco.columnconfigure(1, weight=1)
        ctk.CTkButton(
            marco,
            text="Examinar...",
            command=lambda: self._examinar_carpeta(variable, al_cambiar),
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_BASE,
            corner_radius=8,
            height=32,
            width=120,
        ).grid(row=fila, column=2, pady=4)

    def _campo_excel(self, marco, fila):
        ctk.CTkLabel(
            marco,
            text="Excel de correspondencia (opcional):",
            font=FUENTE_BASE,
            text_color=GRIS_TXT,
        ).grid(row=fila, column=0, sticky="w", padx=(0, 10), pady=4)
        ctk.CTkEntry(
            marco,
            textvariable=self.var_mapeo,
            font=FUENTE_BASE,
            height=34,
            corner_radius=8,
            border_color="#C9D2E3",
        ).grid(row=fila, column=1, sticky="ew", padx=(0, 10), pady=4)
        ctk.CTkButton(
            marco,
            text="Examinar...",
            command=self._examinar_excel,
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_BASE,
            corner_radius=8,
            height=32,
            width=120,
        ).grid(row=fila, column=2, pady=4)

    # --- Quick wins ------------------------------------------------------------

    def _vista_previa_carpeta(self) -> None:
        """Al elegir la carpeta origen, muestra conteo instantaneo."""
        ruta = self.var_origen.get().strip()
        if not ruta or not Path(ruta).is_dir():
            self.lbl_vista_carpeta.configure(text="")
            return
        try:
            pdfs = [
                p for p in Path(ruta).iterdir() if p.is_file() and core.es_pdf(p.name)
            ]
            regs = core.inventariar(Path(ruta))
            anom = len(core.listar_anomalias(regs))
            ops = len(core.ids_origen_unicos(regs))
            self.lbl_vista_carpeta.configure(
                text=f"{len(pdfs)} PDFs encontrados  |  {ops} operaciones  |  "
                f"{anom} con problemas"
            )
        except Exception:
            self.lbl_vista_carpeta.configure(text="")

    def _validar_excel_seleccionado(self) -> None:
        """Al elegir el Excel: valida al instante que traiga correspondencia."""
        ruta = self.var_mapeo.get().strip()
        if not ruta:
            return
        if not Path(ruta).is_file():
            self._estado("El Excel seleccionado no existe en esa ruta.")
            return
        try:
            mapeo, mensaje_error = err.leer_mapeo_con_reintentos(
                Path(ruta), intentos=2, espera=1.0
            )
        except Exception as exc:
            mapeo, mensaje_error = None, err.traducir_error_excel(exc)
        if mensaje_error:
            messagebox.showwarning(TITULO, mensaje_error)
            return
        if not mapeo:
            messagebox.showwarning(TITULO, err.MSJ_EXCEL_SIN_DATOS)
            return
        self._estado(f"Excel valido: {len(mapeo)} pares de correspondencia detectados.")

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
        datos["origen"] = self.var_origen.get().strip()
        datos["destino"] = self.var_destino.get().strip()
        datos["mapeo"] = self.var_mapeo.get().strip()
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
        original = "2026XX840820X1 INTERCONDI SA.pdf"
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
            self.lbl_ejemplo.configure(text=f"Original : {original}\nFinal    : ...")
            self.lbl_patron_estado.configure(
                text=f"Formato invalido: {mensaje}", text_color="#9C1C1C"
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
        self._estado(f"Copiando... {actual}/{total} ({int(frac * 100)}%)")

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
            messagebox.showwarning(TITULO, "Seleccione una carpeta ORIGEN valida.")
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

    def _bloquear(self, valor: bool) -> None:
        estado = "disabled" if valor else "normal"
        self.btn_inv.configure(state=estado)
        self.btn_sim.configure(state=estado)
        self.btn_ejec.configure(state=estado)

    # --- Acciones -------------------------------------------------------------------------

    def _accion_inventariar(self) -> None:
        rutas = self._rutas_validas()
        if not rutas:
            return
        origen, destino = rutas
        self._bloquear(True)
        self.txt.delete("1.0", "end")
        threading.Thread(
            target=self._hilo_inventariar, args=(origen, destino), daemon=True
        ).start()

    def _hilo_inventariar(self, origen: Path, destino: Path) -> None:
        try:
            self._linea_fase(f"INVENTARIO  |  {origen}")
            self.registros = core.inventariar(origen, progreso=self._progreso)
            anom = core.listar_anomalias(self.registros)
            validos = len(self.registros) - len(anom)
            n_ids = len(core.ids_origen_unicos(self.registros))

            self._log("")
            self._log(f"Archivos leidos      : {len(self.registros)}", "info")
            self._log(f"Validos              : {validos}", "ok")
            self._log(
                f"Anomalias            : {len(anom)}", "anomalia" if anom else "ok"
            )
            self._log(f"Ids de origen unicos : {n_ids}", "info")

            for r in anom[:20]:
                self._log(f"  [ANOMALIA] {r['archivo']} -> {r['anomalia']}", "anomalia")
            if len(anom) > 20:
                self._log(f"  ... y {len(anom) - 20} mas", "anomalia")

            if n_ids:
                ruta_lista = destino / "ids_para_consulta.txt"
                core.generar_lista_ids(self.registros, ruta_lista)
                self._log("")
                self._log(
                    f"Lista para consulta (lotes de "
                    f"{core.LOTE_CONSULTA}): {ruta_lista}",
                    "ruta",
                )

            self.plan = core.construir_plan(self.registros, None, self.var_patron.get())
            self._log("")
            self._log(
                "Siguiente paso: obtenga el Excel de correspondencia, luego SIMULAR.",
                "fase",
            )
            self._estado(
                f"Inventario listo: {validos} validos | "
                f"{len(anom)} anomalias | {n_ids} ids"
            )
            self.btn_sim.configure(state="normal")
            self._bloquear(False)
        except Exception as exc:
            self._log(f"ERROR en inventario: {exc}", "error")
            self._estado("Error en inventario")
            self._bloquear(False)

    def _accion_simular(self) -> None:
        rutas = self._rutas_validas(con_mapeo=True)
        if not rutas:
            return
        origen, destino = rutas
        ruta_excel = Path(self.var_mapeo.get().strip())
        if not ruta_excel.is_file():
            messagebox.showwarning(TITULO, "El Excel de mapeo no existe.")
            return

        self._bloquear(True)
        self._estado("Leyendo Excel de correspondencia...")
        threading.Thread(
            target=self._hilo_simular, args=(origen, destino, ruta_excel), daemon=True
        ).start()

    def _hilo_simular(self, origen: Path, destino: Path, ruta_excel: Path) -> None:
        try:
            patron = self.var_patron.get()
            core.CODIGO_FIJO = self.var_codigo.get().strip()

            # Lectura con reintentos + traduccion de errores
            if self.mapeo is None or self.mapeo_ruta != str(ruta_excel):
                mapeo, mensaje_error = err.leer_mapeo_con_reintentos(
                    ruta_excel, log=self._log
                )
                if mapeo is None:
                    self._log("")
                    self._log("NO SE PUDO LEER EL EXCEL:", "error")
                    for linea in mensaje_error.splitlines():
                        self._log(f"  {linea}", "error")
                    self._estado("Error: revise las instrucciones arriba")
                    self.btn_sim.configure(state="normal")
                    self.btn_inv.configure(state="normal")
                    messagebox.showerror(TITULO, mensaje_error)
                    return
                self.mapeo = mapeo
                self.mapeo_ruta = str(ruta_excel)

            self._log("")
            self._log(f"Mapeo cargado: {len(self.mapeo)} pares origen -> destino", "ok")

            if not self.mapeo:
                # Excel vacio: detener aqui con explicacion (no muro de SIN MAPEO)
                self._log("")
                self._log("EL EXCEL NO CONTIENE CORRESPONDENCIA:", "error")
                for linea in err.MSJ_EXCEL_SIN_DATOS.splitlines():
                    self._log(f"  {linea}", "error")
                self._estado("Excel sin datos de correspondencia")
                self.btn_sim.configure(state="normal")
                self.btn_inv.configure(state="normal")
                messagebox.showwarning(TITULO, err.MSJ_EXCEL_SIN_DATOS)
                return

            self.plan = core.construir_plan(self.registros, self.mapeo, patron)
            res = core.resumen_plan(self.plan)

            self._log("")
            self._log(
                f"PLAN: {res['ok']} listos | "
                f"{res['sin_mapeo']} con problema | "
                f"{res['anomalias']} anomalias",
                "info",
            )
            self._log("")

            permitido, avisos = core.validar_plan(self.plan, destino)
            ok_listado = [
                p for p in self.plan if p["estado"] == core.OK and p["destino"]
            ]
            for p in ok_listado[:50]:
                self._log(f"  {p['archivo']}", "info")
                self._log(f"      ->  {p['destino']}", "ok")
            if len(ok_listado) > 50:
                self._log(f"  ... y {len(ok_listado) - 50} mas", "info")

            if avisos:
                self._log("")
                self._log("AVISOS:", "sinmapeo")
                for pr in avisos[:20]:
                    tag = "anomalia" if pr.startswith("COLISION") else "sinmapeo"
                    self._log(f"  {pr}", tag)
                if len(avisos) > 20:
                    self._log(f"  ... y {len(avisos) - 20} mas", "anomalia")

            if not permitido:
                self._log("")
                self._log(
                    "NO SE PUEDE EJECUTAR: hay COLISIONES de nombre "
                    "destino. Corrija los datos y vuelva a simular.",
                    "error",
                )
                self._estado("Simulacion con colisiones - no ejecutable")
                self.btn_ejec.configure(state="disabled")
            else:
                pendientes = res["sin_mapeo"] + res["anomalias"]
                self._log("")
                if pendientes:
                    self._log(
                        f"SIMULACION OK: {res['ok']} se pueden renombrar. "
                        f"{pendientes} quedaran pendientes con registro y "
                        "sugerencia (pendientes_correccion.csv).",
                        "ok",
                    )
                else:
                    self._log(
                        "SIMULACION CORRECTA: todos los archivos se pueden renombrar.",
                        "ok",
                    )
                self._estado("Simulacion correcta - listo para EJECUTAR")
                self.btn_ejec.configure(state="normal")

            self.btn_sim.configure(state="normal")
            self.btn_inv.configure(state="normal")
        except Exception as exc:
            self._log(
                f"ERROR inesperado en simulacion: {type(exc).__name__}: {exc}", "error"
            )
            self._estado("Error en simulacion")
            self.btn_sim.configure(state="normal")
            self.btn_inv.configure(state="normal")

    def _accion_ejecutar(self) -> None:
        rutas = self._rutas_validas(con_mapeo=True)
        if not rutas:
            return
        origen, destino = rutas

        res = core.resumen_plan(self.plan)
        pendientes = res["sin_mapeo"] + res["anomalias"]

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
        self.btn_ejec.configure(state="disabled")
        self._linea_fase(f"EJECUTANDO copia  |  {destino}")
        threading.Thread(
            target=self._hilo_ejecutar, args=(origen, destino), daemon=True
        ).start()

    def _hilo_ejecutar(self, origen: Path, destino: Path) -> None:
        try:
            copiados, omitidos = core.ejecutar_copias(
                self.plan, origen, destino, progreso=self._progreso
            )
            core.escribir_reporte(self.plan, destino / "reporte.csv")

            n_pend, n_ids = core.escribir_pendientes(
                self.plan,
                destino / "pendientes_correccion.csv",
                destino / "ids_pendientes.txt",
            )
            core.generar_lista_ids(self.registros, destino / "ids_para_consulta.txt")

            self._log("")
            self._log(
                f"LISTO: {copiados} copiado(s), {omitidos} omitido(s) (ya existian).",
                "ok",
            )
            if n_pend:
                self._log(
                    f"Pendientes con seguimiento: {n_pend} "
                    f"(ids para re-consulta: {n_ids})",
                    "sinmapeo",
                )
                self._log(f"Registro: {destino / 'pendientes_correccion.csv'}", "ruta")
            self._log(f"Reporte: {destino / 'reporte.csv'}", "ruta")
            self._log("")
            self._log(
                "Siguiente paso: cargue la carpeta destino en el "
                "sistema. Si hubo pendientes, corrigalos y re-ejecute "
                "el mismo lote: los ya hechos se omiten solos.",
                "fase",
            )
            resumen = f"Completado: {copiados} copiados, {omitidos} omitidos"
            if n_pend:
                resumen += f", {n_pend} pendientes"
            self._estado(resumen)
        except Exception as exc:
            self._log(f"ERROR en ejecucion: {type(exc).__name__}: {exc}", "error")
            self._estado("Error en ejecucion")
        finally:
            self.btn_inv.configure(state="normal")
            self.btn_sim.configure(state="normal")
            self.btn_ejec.configure(state="normal")

    # --- Acciones auxiliares ---------------------------------------------------------

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
                TITULO, "Aun no existe la lista. Ejecute INVENTARIAR primero."
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
                TITULO, "Aun no existe reporte.csv. Ejecute EJECUTAR primero."
            )


def main() -> None:
    habilitar_dpi_awareness()
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
