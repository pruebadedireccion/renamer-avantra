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
CONFLICTO = "CONFLICTO"


class ProcesoCancelado(Exception):
    """Cancelacion solicitada por el usuario durante un proceso largo."""


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


def inventariar(carpeta: Path, progreso=None, cancelado=None) -> list[dict]:
    """
    Escanea la carpeta plana y clasifica cada archivo.

    Verifica la extension y el encabezado %PDF del contenido (detecta
    descargas incompletas o archivos danados).

    `cancelado`: callable opcional que devuelve True para abortar.
    Lanza ProcesoCancelado si se solicita la cancelacion.

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
        if cancelado and cancelado():
            raise ProcesoCancelado(
                f"Cancelado durante el inventario ({i - 1} de {total} leidos)"
            )
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

        if not _pdf_valido(carpeta / nombre):
            reg["anomalia"] = "PDF danado o descarga incompleta"
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


def _pdf_valido(ruta: Path) -> bool:
    """True si el archivo abre y su encabezado es %PDF-."""
    try:
        with open(ruta, "rb") as f:
            cabecera = f.read(5)
        return cabecera == b"%PDF-"
    except OSError:
        return False


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
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text("\n".join(lineas), encoding="utf-8")
    return len(ids)


# --- Mapeo --------------------------------------------------------------------


def cargar_mapeo(ruta_excel: Path) -> dict[str, str]:
    """
    Lee el Excel de correspondencia y auto-detecta columnas.
    Compatibilidad: devuelve solo {id_origen: id_destino}.
    Para detectar conflictos use cargar_mapeo_detallado().
    """
    mapeo, _conflictos = cargar_mapeo_detallado(ruta_excel)
    return mapeo


def cargar_mapeo_detallado(
    ruta_excel: Path,
) -> tuple[dict[str, str], dict[str, list[str]]]:
    """
    Lee el Excel de correspondencia y auto-detecta columnas:
        - origen: celda que calza con el patron de id de origen
        - destino: celda numerica de 6-12 digitos

    Detecta mapeos conflictivos: un mismo id de origen con dos o mas
    destinos distintos. Esos origenes se excluyen del mapeo y se
    devuelven aparte en `conflictos`.

    Devuelve (mapeo, conflictos):
        mapeo      : {id_origen: id_destino}
        conflictos : {id_origen: [destinos distintos detectados]}
    """
    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    try:
        ws = wb.active
        if ws is None:
            raise ValueError("El Excel no tiene hoja activa")
        vistos: dict[str, set[str]] = {}
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
                vistos.setdefault(origen, set()).add(destino)

        mapeo: dict[str, str] = {}
        conflictos: dict[str, list[str]] = {}
        for origen, destinos in vistos.items():
            if len(destinos) == 1:
                mapeo[origen] = next(iter(destinos))
            else:
                conflictos[origen] = sorted(destinos)
        return mapeo, conflictos
    finally:
        wb.close()


# --- Plan de renombrado ---------------------------------------------------------


def construir_plan(
    registros: list[dict],
    mapeo: dict[str, str] | None,
    patron: str = PATRON_DEFAULT,
    conflictos: dict[str, list[str]] | None = None,
) -> list[dict]:
    """
    Construye el plan por archivo con el patron de salida dado:
        OK         -> tiene id de origen y mapeo
        SIN MAPEO  -> id valido pero sin correspondencia en el Excel
        CONFLICTO  -> el Excel asigna dos o mas destinos al mismo origen
        ANOMALIA   -> problema de estructura (no se renombra)

    Numeracion: si una misma operacion de destino recibe varios archivos,
    {n} toma 1, 2, 3... (orden alfabetico). Si la operacion tiene un solo
    archivo, {n} vale 1 (o no aparece si el patron no lo usa).
    """
    conflictos = conflictos or {}
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

        if r["id_origen"] in conflictos:
            destinos = ", ".join(conflictos[r["id_origen"]])
            item["estado"] = CONFLICTO
            item["detalle"] = (
                f"conflicto: el Excel asigna varios destinos a este origen "
                f"({destinos})"
            )
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
        "conflictos": sum(1 for p in plan if p["estado"] == CONFLICTO),
        "pendientes_mapeo": sum(
            1 for p in plan if p["estado"] == OK and not p["destino"]
        ),
    }


def validar_plan(
    plan: list[dict], destino_dir: Path, requerir_mapeo: bool = True
) -> tuple[bool, list[str]]:
    """
    Valida el plan antes de ejecutar. Devuelve (permitido, problemas).

    Semantica v1.1.0 (ejecucion parcial):
        - SIN MAPEO no bloquea: esos archivos quedan en el registro de
          pendientes para seguimiento.
        - "YA EXISTE en destino" no bloquea: en re-ejecuciones los archivos
          ya copiados se omiten (ejecutar_copias ya los salta).
        - COLISION sigue bloqueando: dos archivos hacia el mismo nombre
          en la misma corrida indican datos inconsistentes.
    """
    problemas: list[str] = []
    ok = [p for p in plan if p["estado"] == OK and p["destino"]]

    # Colisiones de nombre destino (unico bloqueo duro)
    vistos: dict[str, list[str]] = {}
    for p in ok:
        vistos.setdefault(p["destino"], []).append(p["archivo"])
    for d, origs in sorted(vistos.items()):
        if len(origs) > 1:
            problemas.append(f"COLISION: {d} <- {', '.join(origs)}")

    # Avisos informativos (no bloquean)
    sin_mapeo = [p for p in plan if p["estado"] == SIN_MAPEO]
    if sin_mapeo:
        problemas.append(
            f"AVISO: {len(sin_mapeo)} archivo(s) sin correspondencia quedaran "
            "pendientes (se registran en pendientes_correccion.csv)"
        )
    conflictos = [p for p in plan if p["estado"] == CONFLICTO]
    if conflictos:
        problemas.append(
            f"AVISO: {len(conflictos)} archivo(s) con mapeo conflictivo "
            "quedaran pendientes (el Excel asigna varios destinos al mismo origen)"
        )
    ya_existen = [p for p in ok if (destino_dir / p["destino"]).exists()]
    if ya_existen:
        problemas.append(
            f"AVISO: {len(ya_existen)} archivo(s) ya existen en el destino "
            "y se omitiran (re-ejecucion)"
        )

    return (not any(p.startswith("COLISION") for p in problemas), problemas)


def accion_sugerida(detalle: str) -> str:
    """Sugerencia de correccion segun el motivo del pendiente."""
    d = detalle.lower()
    if "sin correspondencia" in d:
        return (
            "Obtenga el identificador de destino de este radicado en el "
            "sistema, agreguelo al Excel de correspondencia y vuelva a "
            "ejecutar (los ya renombrados se omiten solos)."
        )
    if "conflicto" in d:
        return (
            "El Excel asigna dos o mas destinos al mismo radicado. "
            "Verifique cual operacion es la correcta, deje una sola fila "
            "para ese radicado en el Excel y vuelva a ejecutar."
        )
    if "danado" in d or "incompleta" in d:
        return (
            "El archivo no tiene un PDF valido (descarga incompleta o "
            "archivo corrupto). Vuelva a descargarlo, reemplacelo en la "
            "carpeta origen y ejecute de nuevo."
        )
    if "duplicado" in d:
        return (
            "Este archivo es una descarga repetida. Verifique cual copia "
            "es la correcta, elimine la sobrante de la carpeta origen y "
            "vuelva a ejecutar."
        )
    if "no es pdf" in d:
        return "No es un PDF: excluyalo de la carpeta origen."
    if "no inicia" in d:
        return (
            "El nombre no inicia con el identificador de origen. Verifique "
            "el nombre del archivo; si corresponde, renombrelo anteponiendo "
            "el radicado y vuelva a ejecutar."
        )
    return "Revise el motivo y corrija el archivo en la carpeta origen."


def accion_corta(detalle: str) -> str:
    """Version breve de la accion sugerida para la tabla de la interfaz."""
    d = detalle.lower()
    if "sin correspondencia" in d:
        return "Complete el Excel y re-ejecute"
    if "conflicto" in d:
        return "Verifique el destino correcto en el Excel"
    if "danado" in d or "incompleta" in d:
        return "Vuelva a descargar el PDF"
    if "duplicado" in d:
        return "Elimine la copia repetida"
    if "no es pdf" in d:
        return "Excluya el archivo de la carpeta"
    if "no inicia" in d:
        return "Renombre el archivo con el radicado"
    return "Revise el archivo en la carpeta origen"


def escribir_pendientes(
    plan: list[dict], ruta_pendientes: Path, ids_pendientes_ruta: Path | None = None
) -> tuple[int, int]:
    """
    Genera el registro de seguimiento de pendientes:
        - pendientes_correccion.csv: archivo, id origen, motivo, accion sugerida
        - ids_pendientes.txt (opcional): ids de origen sin correspondencia,
          en lotes, listos para re-consultar en el sistema

    Devuelve (total_pendientes, ids_pendientes_unicos).
    """
    pendientes = [
        p
        for p in plan
        if p["estado"] in (SIN_MAPEO, ANOMALIA, CONFLICTO)
        or (p["estado"] == OK and not p["destino"])
    ]

    ruta_pendientes.parent.mkdir(parents=True, exist_ok=True)
    with open(ruta_pendientes, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["archivo_original", "id_origen", "motivo", "accion_sugerida"])
        for p in pendientes:
            motivo = p["detalle"] or p["estado"]
            w.writerow(
                [
                    p["archivo"],
                    p["id_origen"] or "",
                    motivo,
                    accion_sugerida(motivo),
                ]
            )

    # Lista de ids sin correspondencia (o con conflicto) para re-consulta
    ids_sin = sorted(
        {
            p["id_origen"]
            for p in pendientes
            if p["estado"] in (SIN_MAPEO, CONFLICTO) and p["id_origen"]
        }
    )
    if ids_pendientes_ruta is not None:
        lineas = [
            ", ".join(ids_sin[i : i + LOTE_CONSULTA])
            for i in range(0, len(ids_sin), LOTE_CONSULTA)
        ]
        ids_pendientes_ruta.write_text("\n".join(lineas), encoding="utf-8")

    return len(pendientes), len(ids_sin)


def ejecutar_copias(
    plan: list[dict], carpeta: Path, destino_dir: Path, progreso=None,
    cancelado=None,
) -> tuple[int, int]:
    """
    Copia los archivos con estado OK+destino a destino_dir.
    Devuelve (copiados, omitidos).

    `cancelado`: callable opcional que devuelve True para abortar.
    Lanza ProcesoCancelado si se solicita la cancelacion (lo ya copiado
    se conserva y queda registrado en el reporte).
    """
    destino_dir.mkdir(parents=True, exist_ok=True)
    pendientes = [p for p in plan if p["estado"] == OK and p["destino"]]
    total = len(pendientes)
    copiados = omitidos = 0

    for i, p in enumerate(pendientes, 1):
        if cancelado and cancelado():
            raise ProcesoCancelado(
                f"Cancelado durante la copia ({copiados} copiados de {total})"
            )
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


def deshacer_desde_reporte(
    ruta_reporte: Path, carpeta_origen: Path, destino_dir: Path
) -> tuple[int, list[str]]:
    """
    Deshace la ultima ejecucion: elimina SOLO las copias registradas en el
    reporte que cumplan todas estas condiciones:
        1. El archivo destino existe.
        2. Corresponde a una fila con estado OK del reporte.
        3. Si el original existe, el tamano coincide con la copia.

    Devuelve (eliminados, errores).
    """
    if not ruta_reporte.is_file():
        return 0, ["No existe reporte.csv en la carpeta destino."]

    eliminados = 0
    errores: list[str] = []
    with open(ruta_reporte, encoding="utf-8-sig") as f:
        for fila in csv.DictReader(f, delimiter=";"):
            if fila.get("estado") != OK or not fila.get("archivo_nuevo"):
                continue
            destino = destino_dir / fila["archivo_nuevo"]
            origen = carpeta_origen / fila["archivo_original"]
            if not destino.is_file():
                continue
            try:
                if origen.is_file() and destino.stat().st_size != origen.stat().st_size:
                    errores.append(
                        f"NO se borro (tamano distinto al original): "
                        f"{fila['archivo_nuevo']}"
                    )
                    continue
                destino.unlink()
                eliminados += 1
            except OSError as exc:
                errores.append(f"No se pudo borrar {fila['archivo_nuevo']}: {exc}")

    return eliminados, errores


def escribir_reporte(plan: list[dict], ruta_reporte: Path) -> None:
    with open(ruta_reporte, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["archivo_original", "archivo_nuevo", "estado", "detalle"])
        for p in plan:
            w.writerow([p["archivo"], p["destino"] or "", p["estado"], p["detalle"]])
