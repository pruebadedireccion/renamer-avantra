# -*- coding: utf-8 -*-
"""
Traductor de errores tecnicos a mensajes accionables para el usuario final.

Cada funcion devuelve un texto en lenguaje de oficina que explica:
    que paso, por que probablemente paso, y que hacer para resolverlo.
"""

from __future__ import annotations

import time
from pathlib import Path

import openpyxl

from openpyxl.utils.exceptions import InvalidFileException


# --- Mensajes accionables ------------------------------------------------------

MSJ_EXCEL_BLOQUEADO = (
    "No se pudo leer el Excel: esta abierto en Excel o bloqueado por OneDrive.\n\n"
    "Que hacer:\n"
    "  1. Cierre el archivo si lo tiene abierto en Excel.\n"
    "  2. Si esta en OneDrive, clic derecho sobre el archivo y elija\n"
    "     'Descargar' o 'Siempre conservar en este dispositivo'.\n"
    "  3. Vuelva a presionar SIMULAR."
)

MSJ_NO_EXCEL = (
    "El archivo seleccionado no es un Excel valido.\n\n"
    "Que hacer:\n"
    "  Verifique que el archivo sea .xlsx y que se abra correctamente\n"
    "  en Excel, luego seleccione de nuevo con 'Examinar...'."
)

MSJ_NO_EXISTE = (
    "El archivo Excel ya no existe en la ruta seleccionada.\n\n"
    "Que hacer:\n"
    "  Vuelva a seleccionarlo con el boton 'Examinar...'."
)

MSJ_EXCEL_SIN_DATOS = (
    "El Excel se abrio pero no contiene filas de correspondencia validas.\n\n"
    "Causas posibles:\n"
    "  - Es la plantilla vacia, no el reporte descargado del sistema.\n"
    "  - No tiene las columnas con el identificador de origen y el de destino.\n\n"
    "Que hacer:\n"
    "  Descargue el reporte del sistema con los identificadores y vuelva\n"
    "  a seleccionarlo con 'Examinar...'."
)


def traducir_error_excel(exc: Exception) -> str:
    """Convierte una excepcion de lectura de Excel en mensaje accionable."""
    nombre = type(exc).__name__
    texto = str(exc).lower()

    if isinstance(exc, PermissionError) or "errno 13" in texto or "permission" in texto:
        return MSJ_EXCEL_BLOQUEADO
    if isinstance(exc, InvalidFileException) or "not support" in texto:
        return MSJ_NO_EXCEL
    if isinstance(exc, FileNotFoundError) or "no such file" in texto:
        return MSJ_NO_EXISTE
    # Genérico legible
    return (
        f"No se pudo leer el Excel.\n\nDetalle tecnico: {nombre}: {exc}\n\n"
        "Que hacer:\n"
        "  Verifique que el archivo sea un Excel valido, que no este abierto\n"
        "  en Excel y que no este bloqueado por OneDrive, e intente de nuevo."
    )


# --- Lectura con reintentos ------------------------------------------------------


def leer_mapeo_con_reintentos(
    ruta_excel: Path, intentos: int = 3, espera: float = 1.5, log=None
):
    """
    Intenta leer el Excel hasta `intentos` veces con espera entre intentos.
    Resuelve bloqueos momentaneos (OneDrive/antivirus).

    Devuelve (mapeo, mensaje_error).
        - mapeo: dict si la lectura fue exitosa, None si fallo
        - mensaje_error: '' si fue exitosa, el mensaje accionable si fallo
    """
    ultimo_error: Exception | None = None
    for intento in range(1, intentos + 1):
        try:
            mapeo = _cargar_mapeo_interno(ruta_excel)
            return mapeo, ""
        except (PermissionError, OSError) as exc:
            ultimo_error = exc
            if intento < intentos:
                if hasattr(log, "__call__"):
                    log(
                        f"El Excel esta ocupado, reintentando... "
                        f"(intento {intento + 1} de {intentos})",
                        "sinmapeo",
                    )
                time.sleep(espera)
            else:
                return None, traducir_error_excel(ultimo_error)
        except (InvalidFileException, ValueError) as exc:
            return None, traducir_error_excel(exc)
        except Exception as exc:  # noqa: BLE001 - traducir cualquier fallo
            return None, traducir_error_excel(exc)
    return None, MSJ_EXCEL_BLOQUEADO


def _cargar_mapeo_interno(ruta_excel: Path) -> dict[str, str]:
    """
    Lectura tolerante del Excel de correspondencia.

    A diferencia de la version estricta del core, esta:
        - ignora hojas ilegibles (recorre todas las hojas)
        - tolera filas corruptas
        - solo falla si el archivo no se puede abrir en absoluto
    """
    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    try:
        mapeo: dict[str, str] = {}
        from renamer_core import RE_ID_ORIGEN, RE_ID_DESTINO

        for ws in wb.worksheets:
            try:
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
            except Exception:  # noqa: BLE001 - hoja ilegible: pasar a la siguiente
                continue
        return mapeo
    finally:
        wb.close()
