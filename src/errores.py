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
    ruta_excel: Path,
    intentos: int = 3,
    espera: float = 1.5,
    log=None,
    modo: str = None,
    regex_origen=None,
    regex_destino=None,
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
            mapeo = _cargar_mapeo_interno(
                ruta_excel, modo, regex_origen, regex_destino
            )
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


def _cargar_mapeo_interno(
    ruta_excel: Path,
    modo: str = None,
    regex_origen=None,
    regex_destino=None,
) -> dict[str, str]:
    """
    Lectura del Excel de correspondencia con la deteccion por capas del
    nucleo (radicado+numerico o regla general, segun el modo).
    """
    import renamer_core as core

    modo = modo or core.MODO_AUTO
    mapeo, _conflictos, _informe = core.cargar_mapeo_detallado(
        ruta_excel, modo, regex_origen, regex_destino
    )
    return mapeo
