# -*- coding: utf-8 -*-
"""
Nucleo de Renamer Avantra: renombrado masivo de archivos PDF con
mapeo por tabla de correspondencia y formato de salida configurable.

Concepto de formato de salida:
    El nombre final se construye con una plantilla parametrizable:

        {id}   -> identificador de destino (del Excel de mapeo)
        {n}    -> consecutivo cuando la operacion recibe varios archivos
        {c}    -> codigo fijo configurable (ej. ZREE)

    Ejemplos de patron:
        "{id}"          -> 9900000001.pdf
        "{id}-{n}"      -> 9900000001-1.pdf
        "{id}-{c}-{n}"  -> 9900000001-ZREE-1.pdf

Este modulo NO tiene interfaz grafica: es la logica pura usada por
renamer_gui.py (ventana) y renamer_cli.py (consola).
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path

try:
    import openpyxl
except ImportError as e:
    raise SystemExit("Falta la libreria openpyxl (pip install openpyxl)") from e

# --- Configuracion por defecto ------------------------------------------------

# Identificador de origen en el nombre del archivo (ej. ficticio: 2026XX840820X1)
RE_ID_ORIGEN = re.compile(r"^(?P<id>\d{4}[A-Z]{2}\d+[A-Z]\d)", re.IGNORECASE)
RE_DESCARGA_DUP = re.compile(r"\s\(\d+\)")
# Identificador de destino en el Excel de correspondencia (6-12 digitos)
RE_ID_DESTINO = re.compile(r"^\d{6,12}$")

PATRON_DEFAULT = "{id}-{c}-{n}"
CODIGO_FIJO = ""  # codigo fijo configurable (ej. ZREE)
LOTE_CONSULTA = 200

ARCHIVOS_PROPIOS = {"ids_para_consulta.txt", "reporte.csv", "config.json"}

TOKENS = ("{id}", "{n}", "{c}")

# --- Estados ----------------------------------------------------------------

OK = "OK"
SIN_MAPEO = "SIN MAPEO"
ANOMALIA = "ANOMALIA"


# --- Validacion del patron -----------------------------------------------------


def validar_patron(patron: str) -> tuple[bool, str]:
    """
    Valida el patron de salida. Devuelve (valido, mensaje).
    Reglas:
        - Debe existir y no estar vacio
        - Debe incluir {id}
        - Si incluye {n}, debe ir despues de {id}
        - Solo tokens conocidos
    """
    if not patron or not patron.strip():
        return False, "El patron no puede estar vacio"

    p = patron.strip()
    if "{id}" not in p:
        return False, "El patron debe incluir {id} (identificador de destino)"

    for tok in re.findall(r"\{[^}]*\}", p):
        if tok not in TOKENS:
            return False, f"Token no reconocido: {tok} (validos: {', '.join(TOKENS)})"

    if "{n}" in p and p.index("{n}") < p.index("{id}"):
        return False, "El consecutivo {n} debe ir despues de {id}"

    return True, ""


def _limpiar(nombre: str) -> str:
    """Quita guiones dobles y sobrantes en bordes."""
    nombre = re.sub(r"-{2,}", "-", nombre)
    return nombre.strip("-")


def generar_nombre(patron: str, id_destino: str, n: int) -> str:
    """Aplica el patron y devuelve el nombre final con extension .pdf."""
    nombre = patron.replace("{id}", id_destino)
    nombre = nombre.replace("{n}", str(n))
    if "{c}" in nombre:
        nombre = nombre.replace("{c}", CODIGO_FIJO or "").replace("--", "-")
    return _limpiar(nombre) + ".pdf"


# --- Utilidades ---------------------------------------------------------------


def es_pdf(nombre: str) -> bool:
    return nombre.lower().endswith(".pdf")


def extraer_id_origen(nombre: str) -> str | None:
    m = RE_ID_ORIGEN.match(nombre)
    return m.group("id").upper() if m else None


def es_duplicado_descarga(nombre: str) -> bool:
    return bool(RE_DESCARGA_DUP.search(Path(nombre).stem))


# --- Configuracion persistente --------------------------------------------------


def cargar_config(ruta: Path) -> dict:
    """Lee config.json junto a la aplicacion (si existe)."""
    if ruta.is_file():
        try:
            return json.loads(ruta.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def guardar_config(ruta: Path, datos: dict) -> None:
    try:
        ruta.write_text(
            json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


# --- Inventario ---------------------------------------------------------------


def inventariar(carpeta: Path, progreso=None) -> list[dict]:
    """
    Escanea la carpeta plana y clasifica cada archivo.

    Devuelve registros: {archivo, id_origen, anomalia}
    """
    registros: list[dict] = []
    archivos = sorted(
        p.name
        for p in carpeta.iterdir()
        if p.is_file()
        and p.name not in ARCHIVOS_PROPIOS
        and p.suffix.lower() != ".xlsx"
    )

    total = len(archivos)
    for i, nombre in enumerate(archivos, 1):
        if progreso and i % 50 == 0:
            progreso(i, total)

        reg = {"archivo": nombre, "id_origen": None, "anomalia": None}

        if not es_pdf(nombre):
            reg["anomalia"] = "No es PDF"
            registros.append(reg)
            continue

        if es_duplicado_descarga(nombre):
            reg["anomalia"] = "Duplicado de descarga ' (N)'"
            registros.append(reg)
            continue

        id_origen = extraer_id_origen(nombre)
        if not id_origen:
            reg["anomalia"] = "No inicia con el patron de identificador"
            registros.append(reg)
            continue
        reg["id_origen"] = id_origen

        registros.append(reg)

    if progreso:
        progreso(total, total)

    return registros


def listar_anomalias(registros: list[dict]) -> list[dict]:
    return [r for r in registros if r["anomalia"] is not None]


def ids_origen_unicos(registros: list[dict]) -> list[str]:
    return sorted(
        {r["id_origen"] for r in registros if r["anomalia"] is None and r["id_origen"]}
    )


def generar_lista_ids(registros: list[dict], ruta: Path) -> int:
    """Exporta ids de origen unicos, LOTE_CONSULTA por linea."""
    ids = ids_origen_unicos(registros)
    lineas = [
        ", ".join(ids[i : i + LOTE_CONSULTA]) for i in range(0, len(ids), LOTE_CONSULTA)
    ]
    ruta.write_text("\n".join(lineas), encoding="utf-8")
    return len(ids)


# --- Mapeo --------------------------------------------------------------------


def cargar_mapeo(ruta_excel: Path) -> dict[str, str]:
    """
    Lee el Excel de correspondencia y auto-detecta columnas:
        - origen: celda que calza con el patron de id de origen
        - destino: celda numerica de 6-12 digitos
    Devuelve {id_origen: id_destino}.
    """
    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    try:
        ws = wb.active
        if ws is None:
            raise ValueError("El Excel no tiene hoja activa")
        mapeo: dict[str, str] = {}
        for fila in ws.iter_rows(values_only=True):
            if not fila:
                continue
            origen = destino = None
            for celda in fila:
                if celda is None:
                    continue
                texto = str(celda).strip()
                if origen is None and RE_ID_ORIGEN.fullmatch(texto):
                    origen = texto.upper()
                elif destino is None and RE_ID_DESTINO.fullmatch(texto):
                    destino = texto
            if origen and destino:
                mapeo[origen] = destino
        return mapeo
    finally:
        wb.close()


# --- Plan de renombrado ---------------------------------------------------------


def construir_plan(
    registros: list[dict],
    mapeo: dict[str, str] | None,
    patron: str = PATRON_DEFAULT,
) -> list[dict]:
    """
    Construye el plan por archivo con el patron de salida dado:
        OK         -> tiene id de origen y mapeo
        SIN MAPEO  -> id valido pero sin correspondencia en el Excel
        ANOMALIA   -> problema de estructura (no se renombra)

    Numeracion: si una misma operacion de destino recibe varios archivos,
    {n} toma 1, 2, 3... (orden alfabetico). Si la operacion tiene un solo
    archivo, {n} vale 1 (o no aparece si el patron no lo usa).
    """
    plan = []
    for r in registros:
        item = {
            "archivo": r["archivo"],
            "id_origen": r["id_origen"],
            "destino": None,
            "estado": None,
            "detalle": r["anomalia"] or "",
        }

        if r["anomalia"] is not None:
            item["estado"] = ANOMALIA
            plan.append(item)
            continue

        if mapeo is None:
            item["estado"] = OK
            item["detalle"] = "sin mapeo (pendiente Excel de correspondencia)"
            plan.append(item)
            continue

        destino_id = mapeo.get(r["id_origen"])
        if not destino_id:
            item["estado"] = SIN_MAPEO
            item["detalle"] = (
                f"id origen {r['id_origen']} sin correspondencia en el Excel"
            )
            plan.append(item)
            continue

        item["estado"] = OK
        plan.append(item)

    if mapeo is None:
        return plan

    # Numerar por operacion de destino
    por_destino: dict[str, list[dict]] = {}
    for item in plan:
        if item["estado"] == OK and item["id_origen"]:
            destino_id = mapeo.get(item["id_origen"])
            if destino_id:
                por_destino.setdefault(destino_id, []).append(item)

    for destino_id, items in por_destino.items():
        items_orden = sorted(items, key=lambda x: x["archivo"])
        for n, item in enumerate(items_orden, 1):
            item["destino"] = generar_nombre(
                patron, destino_id, n if len(items) > 1 else 1
            )
            if len(items) > 1:
                item["detalle"] = (
                    f"operacion {destino_id} con {len(items)} archivos "
                    f"(numerado {n} de {len(items)})"
                )

    return plan


def resumen_plan(plan: list[dict]) -> dict:
    return {
        "total": len(plan),
        "ok": sum(1 for p in plan if p["estado"] == OK and p["destino"]),
        "sin_mapeo": sum(1 for p in plan if p["estado"] == SIN_MAPEO),
        "anomalias": sum(1 for p in plan if p["estado"] == ANOMALIA),
        "pendientes_mapeo": sum(
            1 for p in plan if p["estado"] == OK and not p["destino"]
        ),
    }


def validar_plan(
    plan: list[dict], destino_dir: Path, requerir_mapeo: bool = True
) -> tuple[bool, list[str]]:
    """
    Valida el plan antes de ejecutar. Devuelve (permitido, problemas).
    """
    problemas: list[str] = []

    ok = [p for p in plan if p["estado"] == OK and p["destino"]]

    if requerir_mapeo:
        sin_mapeo = [p for p in plan if p["estado"] == SIN_MAPEO]
        if sin_mapeo:
            problemas.append(
                f"{len(sin_mapeo)} id(s) de origen SIN correspondencia en el Excel"
            )
            for p in sin_mapeo[:10]:
                problemas.append(f"  SIN MAPEO: {p['archivo']} ({p['detalle']})")
            if len(sin_mapeo) > 10:
                problemas.append(f"  ... y {len(sin_mapeo) - 10} mas")

    # Colisiones de nombre destino
    vistos: dict[str, list[str]] = {}
    for p in ok:
        vistos.setdefault(p["destino"], []).append(p["archivo"])
    for d, origs in sorted(vistos.items()):
        if len(origs) > 1:
            problemas.append(f"COLISION: {d} <- {', '.join(origs)}")

    # Destinos ya existentes de una corrida anterior
    ya_existen = [p for p in ok if (destino_dir / p["destino"]).exists()]
    for p in ya_existen:
        problemas.append(f"YA EXISTE en destino: {p['destino']}")

    return (len(problemas) == 0, problemas)


def ejecutar_copias(
    plan: list[dict], carpeta: Path, destino_dir: Path, progreso=None
) -> tuple[int, int]:
    """
    Copia los archivos con estado OK+destino a destino_dir.
    Devuelve (copiados, omitidos).
    """
    destino_dir.mkdir(parents=True, exist_ok=True)
    pendientes = [p for p in plan if p["estado"] == OK and p["destino"]]
    total = len(pendientes)
    copiados = omitidos = 0

    for i, p in enumerate(pendientes, 1):
        origen = carpeta / p["archivo"]
        destino = destino_dir / p["destino"]
        if destino.exists():
            omitidos += 1
        else:
            destino.write_bytes(origen.read_bytes())
            copiados += 1
        if progreso and (i % 10 == 0 or i == total):
            progreso(i, total)

    return copiados, omitidos


def escribir_reporte(plan: list[dict], ruta_reporte: Path) -> None:
    with open(ruta_reporte, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["archivo_original", "archivo_nuevo", "estado", "detalle"])
        for p in plan:
            w.writerow([p["archivo"], p["destino"] or "", p["estado"], p["detalle"]])
