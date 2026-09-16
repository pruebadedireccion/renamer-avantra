# -*- coding: utf-8 -*-
"""
Renamer Avantra - Interfaz grafica.

Ventana "puente" para el renombrado masivo de PDFs:
    - Carpeta origen / carpeta destino / Excel de correspondencia
    - FORMATO DE SALIDA configurable con ejemplo en vivo
    - Flujo: INVENTARIAR -> SIMULAR -> EJECUTAR
"""

from __future__ import annotations

import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

sys.path.insert(0, str(Path(__file__).resolve().parent))

import renamer_core as core

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
FUENTE_MONO = ("Consolas", 14)

ARCHIVO_CONFIG = "config.json"


class App:
    def __init__(self, root) -> None:
        self.root = root
        self.root.title(TITULO)
        self.root.geometry("980x780")
        self.root.minsize(860, 660)
        self.root.configure(fg_color=FONDO)

        self.var_origen = ctk.StringVar()
        self.var_destino = ctk.StringVar()
        self.var_mapeo = ctk.StringVar()
        self.var_patron = ctk.StringVar(value=core.PATRON_DEFAULT)
        self.var_codigo = ctk.StringVar(value="")

        # Configuracion persistente
        ruta_config = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
        self.ruta_config = Path(".") / "config.json"
        cfg = core.cargar_config(self.ruta_config)
        if cfg.get("patron"):
            self.var_patron.set(cfg["patron"])
        if cfg.get("codigo"):
            self.var_codigo.set(cfg["codigo"])

        self.plan: list[dict] = []
        self.registros: list[dict] = []
        self.mapeo: dict[str, str] | None = None
        self.mapeo_ruta: str = ""

        self._construir_widgets()
        self._centrar()

        self._log("Bienvenido.", "fase")
        self._log("1. Seleccione la carpeta de PDFs y la carpeta destino.", "info")
        self._log("2. Defina el FORMATO DE SALIDA y presione INVENTARIAR.", "info")

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
        cont.pack(fill="both", expand=True, padx=16, pady=10)

        # --- 1. CARPETAS ------------------------------------------------------
        s1 = ctk.CTkFrame(
            cont,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color="#D6DCE8",
        )
        s1.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            s1, text="1. CARPETAS", font=FUENTE_SECCION, text_color=AZUL_OSCURO
        ).pack(anchor="w", padx=16, pady=(10, 2))
        interno = ctk.CTkFrame(s1, fg_color="white", corner_radius=0)
        interno.pack(fill="x", padx=16, pady=(2, 10))
        self._campo_ruta(interno, 0, "Carpeta de PDFs (origen):", self.var_origen)
        self._campo_ruta(interno, 1, "Carpeta destino (renombrados):", self.var_destino)
        self._campo_excel(interno, 2)

        # --- 2. FORMATO DE SALIDA ---------------------------------------------
        s2 = ctk.CTkFrame(
            cont,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color="#D6DCE8",
        )
        s2.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            s2, text="2. FORMATO DE SALIDA", font=FUENTE_SECCION, text_color=AZUL_OSCURO
        ).pack(anchor="w", padx=16, pady=(10, 2))

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
            width=160,
            placeholder_text="(opcional, ej. ZREE)",
        )
        self.ent_codigo.grid(row=1, column=1, sticky="w", pady=4)
        self.ent_codigo.bind("<KeyRelease>", lambda _e: self._actualizar_ejemplo())

        # Piezas disponibles
        piezas = ctk.CTkFrame(fmt, fg_color="transparent")
        piezas.grid(row=2, column=1, sticky="w", pady=(2, 4))
        ctk.CTkLabel(
            piezas, text="Piezas: ", font=FUENTE_ESTADO, text_color=GRIS_TXT
        ).pack(side="left")
        for token in ("{id}", "{n}", "{c}"):
            b = ctk.CTkButton(
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
            )
            b.pack(side="left", padx=(0, 6))

        # Ejemplo en vivo
        ej = ctk.CTkFrame(s2, fg_color="#F4F6FA", corner_radius=8)
        ej.pack(fill="x", padx=16, pady=(2, 12))
        ctk.CTkLabel(
            ej, text="EJEMPLO EN VIVO", font=FUENTE_ESTADO, text_color=AZUL_MEDIO
        ).pack(anchor="w", padx=14, pady=(8, 2))
        self.lbl_ejemplo = ctk.CTkLabel(
            ej,
            text="",
            font=("Consolas", 14),
            text_color=GRIS_TXT,
            justify="left",
            anchor="w",
        )
        self.lbl_ejemplo.pack(anchor="w", padx=14, pady=(0, 10))
        self.lbl_patron_estado = ctk.CTkLabel(
            ej, text="", font=FUENTE_ESTADO, anchor="w"
        )
        self.lbl_patron_estado.pack(anchor="w", padx=14, pady=(0, 10))
        self._actualizar_ejemplo()

        # --- 3. ACCIONES --------------------------------------------------------
        s3 = ctk.CTkFrame(
            cont,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color="#D6DCE8",
        )
        s3.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            s3,
            text="3. ACCIONES  (en orden)",
            font=FUENTE_SECCION,
            text_color=AZUL_OSCURO,
        ).pack(anchor="w", padx=16, pady=(10, 2))
        acciones = ctk.CTkFrame(s3, fg_color="transparent")
        acciones.pack(fill="x", padx=16, pady=(2, 12))

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
            width=170,
        )
        self.btn_inv.pack(side="left", padx=(0, 10))

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
            width=190,
            state="disabled",
        )
        self.btn_sim.pack(side="left", padx=(0, 10))

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
            width=190,
            state="disabled",
        )
        self.btn_ejec.pack(side="left", padx=(0, 10))

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
        ).pack(side="right", padx=(8, 0))

        self.progreso = ctk.CTkProgressBar(
            acciones, progress_color=AZUL_MEDIO, height=14, width=280
        )
        self.progreso.set(0)
        self.progreso.pack(side="right", padx=(0, 14))

        # --- 4. RESULTADOS ------------------------------------------------------
        s4 = ctk.CTkFrame(
            cont,
            fg_color="white",
            corner_radius=12,
            border_width=1,
            border_color="#D6DCE8",
        )
        s4.pack(fill="both", expand=True, pady=(0, 8))
        ctk.CTkLabel(
            s4, text="4. RESULTADOS", font=FUENTE_SECCION, text_color=AZUL_OSCURO
        ).pack(anchor="w", padx=16, pady=(10, 2))

        self.txt = tk.Text(
            s4,
            height=12,
            font=FUENTE_LOG,
            bg="white",
            fg=GRIS_TXT,
            relief="flat",
            padx=10,
            pady=8,
            bd=0,
            wrap="none",
        )
        self.txt.pack(fill="both", expand=True, padx=14, pady=(2, 10))
        barra = tk.Scrollbar(s4, command=self.txt.yview, width=10)
        self.txt.configure(yscrollcommand=barra.set)
        barra.place(relx=1.0, rely=0.02, relheight=0.95, anchor="ne", relwidth=0.012)

        self.txt.tag_configure("fase", foreground=AZUL_OSCURO)
        self.txt.tag_configure("ok", foreground=VERDE)
        self.txt.tag_configure("error", foreground="#9C1C1C")
        self.txt.tag_configure("anomalia", foreground="#9C1C1C")
        self.txt.tag_configure("sinmapeo", foreground="#8C6209")
        self.txt.tag_configure("info", foreground=GRIS_TXT)
        self.txt.tag_configure("ruta", foreground=AZUL_MEDIO)

        # --- Pie -----------------------------------------------------------------
        pie = ctk.CTkFrame(
            self.root,
            fg_color="white",
            corner_radius=0,
            height=48,
            border_width=1,
            border_color="#D6DCE8",
        )
        pie.pack(fill="x", side="bottom")
        pie.pack_propagate(False)

        self.var_estado = ctk.StringVar(value="Listo.")
        ctk.CTkLabel(
            pie,
            textvariable=self.var_estado,
            font=FUENTE_ESTADO,
            text_color=GRIS_TXT,
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=16)

        ctk.CTkButton(
            pie,
            text="Copiar lista",
            command=self._copiar_lista,
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO,
            corner_radius=8,
            height=30,
            width=130,
        ).pack(side="right", padx=6, pady=8)
        ctk.CTkButton(
            pie,
            text="Abrir carpeta destino",
            command=self._abrir_destino,
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO,
            corner_radius=8,
            height=30,
            width=180,
        ).pack(side="right", padx=6)
        ctk.CTkButton(
            pie,
            text="Abrir reporte.csv",
            command=self._abrir_reporte,
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_ESTADO,
            corner_radius=8,
            height=30,
            width=150,
        ).pack(side="right", padx=(6, 12))

    def _campo_ruta(self, marco, fila, texto, variable):
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
            command=lambda: self._examinar_carpeta(variable),
            fg_color=AZUL_CLARO,
            hover_color="#C4D2EE",
            text_color=AZUL_OSCURO,
            font=FUENTE_BASE,
            corner_radius=8,
            height=32,
            width=110,
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
            width=110,
        ).grid(row=fila, column=2, pady=4)

    # --- Formato de salida -----------------------------------------------------

    def _insertar_token(self, token: str) -> None:
        actual = self.var_patron.get()
        self.var_patron.set(actual + token)
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

    def _guardar_config_patron(self, patron: str, codigo: str) -> None:
        core.guardar_config(self.ruta_config, {"patron": patron, "codigo": codigo})

    # --- Examina rutas -----------------------------------------------------------

    def _examinar_carpeta(self, variable) -> None:
        ruta = filedialog.askdirectory(title="Seleccione carpeta")
        if ruta:
            variable.set(ruta)

    def _examinar_excel(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccione Excel de correspondencia",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if ruta:
            self.var_mapeo.set(ruta)

    def _centrar(self) -> None:
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 3
        self.root.geometry(f"+{x}+{y}")

    # --- Utilidades de interfaz ---------------------------------------------------

    def _log(self, msg: str, tag: str = "info") -> None:
        self.txt.insert("end", msg + "\n", tag)
        self.txt.see("end")
        self.root.update_idletasks()

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
            from tkinter import messagebox

            messagebox.showwarning(TITULO, "Seleccione una carpeta ORIGEN valida.")
            return None
        if not destino:
            from tkinter import messagebox

            messagebox.showwarning(TITULO, "Seleccione la carpeta DESTINO.")
            return None
        if con_mapeo and not self.var_mapeo.get().strip():
            from tkinter import messagebox

            messagebox.showwarning(
                TITULO,
                "Para simular o ejecutar necesita el Excel de "
                "correspondencia (origen -> destino).",
            )
            return None
        return Path(origen).resolve(), Path(destino).resolve()

    def _bloquear(self, valor: bool) -> None:
        estado = "disabled" if valor else "normal"
        self.btn_inv.configure(state=estado)
        self.btn_sim.configure(state=estado)
        self.btn_ejec.configure(state=estado)

    # --- Acciones -------------------------------------------------------------------

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

            if anom:
                self._log("")
                self._log("ANOMALIAS DETECTADAS:", "anomalia")
                for r in anom[:20]:
                    self._log(
                        f"  [ANOMALIA] {r['archivo']} -> {r['anomalia']}", "anomalia"
                    )
                if len(anom) > 20:
                    self._log(f"  ... y {len(anom) - 20} mas", "anomalia")

            if n_ids:
                ruta_lista = destino / "ids_para_consulta.txt"
                core.generar_lista_ids(self.registros, ruta_lista)
                self._log("")
                self._log(
                    f"Lista para consulta (lotes de {core.LOTE_CONSULTA}): "
                    f"{ruta_lista}",
                    "ruta",
                )

            self.plan = core.construir_plan(self.registros, None, self.var_patron.get())
            self._log("")
            self._log(
                "Siguiente paso: obtenga el Excel de correspondencia, "
                "seleccione SIMULAR.",
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
            from tkinter import messagebox

            messagebox.showwarning(TITULO, "El Excel de mapeo no existe.")
            return

        self._bloquear(True)
        self._linea_fase("SIMULACION (dry-run)  |  no se modificara nada")
        threading.Thread(
            target=self._hilo_simular, args=(origen, destino, ruta_excel), daemon=True
        ).start()

    def _hilo_simular(self, origen: Path, destino: Path, ruta_excel: Path) -> None:
        try:
            patron = self.var_patron.get()
            core.CODIGO_FIJO = self.var_codigo.get().strip()

            if self.mapeo is None or self.mapeo_ruta != str(ruta_excel):
                self.mapeo = core.cargar_mapeo(ruta_excel)
                self.mapeo_ruta = str(ruta_excel)
            self._log(f"Mapeo cargado: {len(self.mapeo)} pares origen -> destino", "ok")

            self.plan = core.construir_plan(self.registros, self.mapeo, patron)
            res = core.resumen_plan(self.plan)

            self._log("")
            self._log(
                f"PLAN: {res['ok']} a renombrar | "
                f"{res['sin_mapeo']} sin mapeo | "
                f"{res['anomalias']} anomalias",
                "info",
            )
            self._log("")

            permitido, problemas = core.validar_plan(self.plan, destino)
            ok_listado = [
                p for p in self.plan if p["estado"] == core.OK and p["destino"]
            ]
            for p in ok_listado[:50]:
                self._log(f"  {p['archivo']}", "info")
                self._log(f"      ->  {p['destino']}", "ok")
            if len(ok_listado) > 50:
                self._log(f"  ... y {len(ok_listado) - 50} mas", "info")

            if problemas:
                self._log("")
                self._log("PROBLEMAS DE VALIDACION:", "anomalia")
                for pr in problemas[:20]:
                    tag = "sinmapeo" if "SIN MAPEO" in pr else "anomalia"
                    self._log(f"  {pr}", tag)
                if len(problemas) > 20:
                    self._log(f"  ... y {len(problemas) - 20} mas", "anomalia")

            if not permitido:
                self._log("")
                self._log(
                    "NO SE PUEDE EJECUTAR: corrija los problemas y vuelva a simular.",
                    "error",
                )
                self._estado("Simulacion con problemas - no ejecutable")
                self.btn_ejec.configure(state="disabled")
            else:
                self._log("")
                self._log(
                    "SIMULACION CORRECTA. Presione EJECUTAR cuando este conforme.", "ok"
                )
                self._estado("Simulacion correcta - listo para EJECUTAR")
                self.btn_ejec.configure(state="normal")

            self.btn_sim.configure(state="normal")
        except Exception as exc:
            self._log(f"ERROR en simulacion: {exc}", "error")
            self._estado("Error en simulacion")
            self.btn_sim.configure(state="normal")

    def _accion_ejecutar(self) -> None:
        rutas = self._rutas_validas(con_mapeo=True)
        if not rutas:
            return
        origen, destino = rutas

        res = core.resumen_plan(self.plan)
        from tkinter import messagebox

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
            core.generar_lista_ids(self.registros, destino / "ids_para_consulta.txt")

            self._log("")
            self._log(
                f"LISTO: {copiados} archivo(s) copiados, {omitidos} "
                f"omitidos (ya existian).",
                "ok",
            )
            self._log(f"Reporte: {destino / 'reporte.csv'}", "ruta")
            self._estado(f"Completado: {copiados} copiados, {omitidos} omitidos")
        except Exception as exc:
            self._log(f"ERROR en ejecucion: {exc}", "error")
            self._estado("Error en ejecucion")
        finally:
            self.btn_inv.configure(state="normal")
            self.btn_sim.configure(state="normal")
            self.btn_ejec.configure(state="normal")

    # --- Acciones auxiliares -------------------------------------------------------

    def _copiar_lista(self) -> None:
        ruta = self.var_destino.get().strip() or self.var_origen.get().strip()
        archivo = Path(ruta) / "ids_para_consulta.txt" if ruta else None
        if archivo and archivo.is_file():
            texto = archivo.read_text(encoding="utf-8")
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self._estado("Lista de ids copiada al portapapeles.")
        else:
            from tkinter import messagebox

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
            from tkinter import messagebox

            messagebox.showinfo(
                TITULO, "Aun no existe reporte.csv. Ejecute EJECUTAR primero."
            )


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
