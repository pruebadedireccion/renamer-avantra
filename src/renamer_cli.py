# -*- coding: utf-8 -*-
"""
Renamer Avantra - Consola.

Renombrado masivo de PDFs con mapeo por tabla de correspondencia
y formato de salida configurable.

Uso:
    1) Inventario (no requiere mapeo, no modifica nada):
        python renamer_cli.py --carpeta RUTA

    2) Simulacion (dry-run, no modifica nada):
        python renamer_cli.py --carpeta RUTA --mapeo mapeo.xlsx

    3) Renombrado real (copia a carpeta destino, originales intactos):
        python renamer_cli.py --carpeta RUTA --mapeo mapeo.xlsx \\
               --destino SALIDA --ejecutar

    Patron de salida configurable:
        --patron "{id}"              -> 9900000001.pdf
        --patron "{id}-{n}"          -> 9900000001-1.pdf
        --patron "{id}-{c}-{n}"      -> 9900000001-ZREE-1.pdf  (con --codigo ZREE)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import renamer_core as core


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="Renamer Avantra",
        description="Renombrado masivo de PDFs con mapeo por tabla de "
        "correspondencia y patron configurable ({id} / {n} / {c})",
    )
    ap.add_argument("--carpeta", required=True, help="Carpeta con los PDFs de origen")
    ap.add_argument("--mapeo", help="Excel de correspondencia origen -> destino")
    ap.add_argument("--destino", help="Carpeta destino para los renombrados")
    ap.add_argument(
        "--patron",
        default=core.PATRON_DEFAULT,
        help="Plantilla del nombre final: {id} {n} {c}. "
        f"Default: {core.PATRON_DEFAULT}",
    )
    ap.add_argument(
        "--codigo", default="", help="Codigo fijo para el token {c} (ej. ZREE)"
    )
    ap.add_argument(
        "--ejecutar",
        action="store_true",
        help="Ejecuta de verdad (copia). Sin esta bandera solo simula.",
    )
    args = ap.parse_args()

    # Patron y codigo fijos
    valido, mensaje = core.validar_patron(args.patron)
    if not valido:
        print(f"ERROR en el patron: {mensaje}")
        sys.exit(1)
    core.CODIGO_FIJO = args.codigo

    carpeta = Path(args.carpeta).resolve()
    if not carpeta.is_dir():
        print(f"ERROR: la carpeta no existe: {carpeta}")
        sys.exit(1)

    destino_dir = (
        Path(args.destino).resolve() if args.destino else carpeta / "RENOMBRADOS"
    )

    # FASE 1: inventario
    registros = core.inventariar(carpeta)
    anom = core.listar_anomalias(registros)
    n_ids = len(core.ids_origen_unicos(registros))

    print("=" * 62)
    print(f"INVENTARIO  |  {carpeta}")
    print("=" * 62)
    print(f"Archivos leidos  : {len(registros)}")
    print(f"Validos          : {len(registros) - len(anom)}")
    print(f"Anomalias        : {len(anom)}")
    print(f"Ids de origen    : {n_ids}")
    print(f"Patron de salida : {args.patron}")
    if args.codigo:
        print(f"Codigo fijo      : {args.codigo}")

    for r in anom[:20]:
        print(f"  [ANOMALIA] {r['archivo']} -> {r['anomalia']}")
    if len(anom) > 20:
        print(f"  ... y {len(anom) - 20} mas")

    if anom:
        print()
        print("PROCESO DETENIDO: hay anomalias que requieren revision manual.")
        sys.exit(2)

    if not registros:
        print("Nada que procesar.")
        sys.exit(0)

    # Lista de ids para la consulta de correspondencia
    ruta_lista = destino_dir / "ids_para_consulta.txt"
    core.generar_lista_ids(registros, ruta_lista)
    print(f"Lista para consulta (lotes de {core.LOTE_CONSULTA}): {ruta_lista}")

    # FASE 2: mapeo
    if not args.mapeo:
        print()
        print("Sin mapeo: solo se genero el inventario y la lista de ids.")
        print(
            "Siguiente paso: obtener el Excel de correspondencia y ejecutar "
            "de nuevo con --mapeo."
        )
        sys.exit(0)

    mapeo = core.cargar_mapeo(Path(args.mapeo).resolve())
    if not mapeo:
        print("ERROR: el Excel de mapeo no produjo pares origen -> destino.")
        sys.exit(1)

    faltantes = set(core.ids_origen_unicos(registros)) - set(mapeo)
    print()
    print(f"Mapeo cargado: {len(mapeo)} pares | sin correspondencia: {len(faltantes)}")
    if faltantes:
        print("PROCESO DETENIDO: ids sin correspondencia en el Excel.")
        for rad in sorted(faltantes):
            print(f"  [SIN MAPEO] {rad}")
        sys.exit(2)

    # FASE 4: plan + validacion
    plan = core.construir_plan(registros, mapeo, args.patron)
    res = core.resumen_plan(plan)

    print()
    print("=" * 62)
    print("PLAN DE RENOMBRADO (dry-run)")
    print("=" * 62)
    for p in plan:
        if p["estado"] == core.OK and p["destino"]:
            print(f"  {p['archivo']}")
            print(f"      ->  {p['destino']}")
    print()
    print(
        f"A renombrar: {res['ok']} | sin mapeo: {res['sin_mapeo']} | "
        f"anomalias: {res['anomalias']}"
    )

    permitido, problemas = core.validar_plan(plan, destino_dir)
    if problemas:
        print()
        print("PROBLEMAS DE VALIDACION:")
        for pr in problemas:
            print(f"  {pr}")

    if not permitido:
        print()
        print("PROCESO DETENIDO: corrija los problemas y vuelva a simular.")
        sys.exit(2)

    if not args.ejecutar:
        print()
        print("DRY-RUN: no se modifico nada. Agregue --ejecutar para copiar.")
        sys.exit(0)

    # Ejecucion real
    print()
    print("EJECUTANDO...")
    copiados, omitidos = core.ejecutar_copias(plan, carpeta, destino_dir)
    core.escribir_reporte(plan, destino_dir / "reporte.csv")
    core.generar_lista_ids(registros, destino_dir / "ids_para_consulta.txt")
    print()
    print(f"LISTO: {copiados} archivo(s) copiados, {omitidos} omitidos.")
    print(f"Carpeta destino: {destino_dir}")
    print(f"Reporte: {destino_dir / 'reporte.csv'}")


if __name__ == "__main__":
    main()
