# -*- coding: utf-8 -*-
"""Pruebas del nucleo: inventario, mapeo, plan, ejecucion, deshacer."""

import shutil

import openpyxl
import pytest

import renamer_core as core

PDF_OK = b"%PDF-1.4 contenido de prueba"


@pytest.fixture(autouse=True)
def codigo_limpio():
    core.CODIGO_FIJO = ""
    yield
    core.CODIGO_FIJO = ""


@pytest.fixture()
def carpeta(tmp_path):
    origen = tmp_path / "origen"
    destino = tmp_path / "destino"
    origen.mkdir()
    return origen, destino


def crear_pdf(carpeta, nombre, contenido=PDF_OK):
    (carpeta / nombre).write_bytes(contenido)


def crear_excel(ruta, filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID Origen", "ID Destino"])
    for origen, destino in filas:
        ws.append([origen, destino])
    wb.save(ruta)


# --- Inventario ---------------------------------------------------------------


def test_inventario_basico(carpeta):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    crear_pdf(origen, "2026XX840823X1 BETA.pdf")
    (origen / "nota.txt").write_text("x", encoding="utf-8")
    (origen / "2026XX840824X1 GAMMA (1).pdf").write_bytes(PDF_OK)

    regs = core.inventariar(origen)

    assert len(regs) == 4
    anom = core.listar_anomalias(regs)
    assert len(anom) == 2
    motivos = {r["anomalia"] for r in anom}
    assert "No es PDF" in motivos
    assert any("Duplicado" in m for m in motivos)
    assert core.ids_origen_unicos(regs) == [
        "2026XX840820X1",
        "2026XX840823X1",
    ]


def test_inventario_detecta_pdf_danado(carpeta):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840825X1 DANADO.pdf", b"esto no es un pdf")
    crear_pdf(origen, "2026XX840826X1 VACIO.pdf", b"")

    regs = core.inventariar(origen)

    assert all(r["anomalia"] is not None for r in regs)
    assert all("danado" in (r["anomalia"] or "").lower() for r in regs)


def test_inventario_sin_extension_pdf(carpeta):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840827X1 SINEXT")
    regs = core.inventariar(origen)
    assert regs[0]["anomalia"] == "No es PDF"


def test_inventario_cancelable(carpeta):
    origen, _ = carpeta
    for i in range(5):
        crear_pdf(origen, f"2026XX84082{i}X1 DOC{i}.pdf")

    with pytest.raises(core.ProcesoCancelado):
        core.inventariar(origen, cancelado=lambda: True)


# --- Mapeo --------------------------------------------------------------------


def test_mapeo_columnas_invertidas(tmp_path):
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("9900000001", "2026XX840820X1")])
    mapeo = core.cargar_mapeo(ruta)
    assert mapeo == {"2026XX840820X1": "9900000001"}


def test_mapeo_vacio(tmp_path):
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [])
    mapeo, conflictos, informe = core.cargar_mapeo_detallado(ruta)
    assert mapeo == {} and conflictos == {}
    assert informe["pares"] == 0


def test_mapeo_conflicto_detectado(tmp_path):
    ruta = tmp_path / "m.xlsx"
    crear_excel(
        ruta,
        [
            ("2026XX840820X1", "9900000001"),
            ("2026XX840820X1", "9900000099"),
            ("2026XX840823X1", "9900000002"),
        ],
    )
    mapeo, conflictos, informe = core.cargar_mapeo_detallado(ruta)

    assert informe["conflictos"] == 1
    assert "2026XX840820X1" in conflictos
    assert conflictos["2026XX840820X1"] == ["9900000001", "9900000099"]
    assert "2026XX840820X1" not in mapeo
    assert mapeo["2026XX840823X1"] == "9900000002"


# --- Plan ---------------------------------------------------------------------


def test_plan_nombre_con_codigo(carpeta, tmp_path):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    core.CODIGO_FIJO = "ZREE"
    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}-{c}-{n}"
    )

    assert [p["destino"] for p in plan] == ["9900000001-ZREE-1.pdf"]


def test_plan_numeracion_multiarchivo(carpeta, tmp_path):
    origen, _ = carpeta
    for nombre in (
        "2026XX840824X1 CERTIFICADO.pdf",
        "2026XX840824X1 ANEXO.pdf",
        "2026XX840824X1 OFICIO.pdf",
    ):
        crear_pdf(origen, nombre)
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840824X1", "9900000003")])

    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}-{n}"
    )
    destinos = {p["archivo"]: p["destino"] for p in plan}

    assert destinos["2026XX840824X1 ANEXO.pdf"] == "9900000003-1.pdf"
    assert destinos["2026XX840824X1 CERTIFICADO.pdf"] == "9900000003-2.pdf"
    assert destinos["2026XX840824X1 OFICIO.pdf"] == "9900000003-3.pdf"


def test_plan_conflicto_queda_pendiente(carpeta, tmp_path):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(
        ruta,
        [
            ("2026XX840820X1", "9900000001"),
            ("2026XX840820X1", "9900000099"),
        ],
    )
    mapeo, conflictos, _informe = core.cargar_mapeo_detallado(ruta)
    plan = core.construir_plan(
        core.inventariar(origen), mapeo, "{id}", conflictos
    )

    assert plan[0]["estado"] == core.CONFLICTO
    assert plan[0]["destino"] is None
    assert "conflicto" in plan[0]["detalle"].lower()


def test_rechazo_de_patron_sin_id():
    valido, mensaje = core.validar_patron("{n}-{c}")
    assert not valido and "{id}" in mensaje


# --- Validacion y ejecucion -----------------------------------------------------


def test_validacion_permite_ejecucion_parcial(carpeta, tmp_path):
    origen, destino = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    crear_pdf(origen, "2026XX840823X1 BETA.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )
    permitido, avisos = core.validar_plan(plan, destino)

    assert permitido
    assert any("AVISO" in a for a in avisos)


def test_colision_bloquea(carpeta, tmp_path):
    origen, destino = carpeta
    crear_pdf(origen, "2026XX840820X1 A.pdf")
    crear_pdf(origen, "2026XX840820X1 B.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )
    permitido, avisos = core.validar_plan(plan, destino)

    assert not permitido
    assert any(a.startswith("COLISION") for a in avisos)


def test_ejecucion_parcial_y_reejecucion(carpeta, tmp_path):
    origen, destino = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    crear_pdf(origen, "2026XX840823X1 BETA.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )
    copiados, omitidos = core.ejecutar_copias(plan, origen, destino)
    assert (copiados, omitidos) == (1, 0)
    assert (destino / "9900000001.pdf").is_file()

    # Se corrige el Excel y se re-ejecuta: omite el ya hecho
    crear_excel(
        ruta,
        [
            ("2026XX840820X1", "9900000001"),
            ("2026XX840823X1", "9900000002"),
        ],
    )
    plan2 = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )
    copiados2, omitidos2 = core.ejecutar_copias(plan2, origen, destino)
    assert (copiados2, omitidos2) == (1, 1)


def test_pendientes_con_sugerencias(carpeta, tmp_path):
    origen, destino = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    crear_pdf(origen, "2026XX840823X1 BETA.pdf")
    (origen / "sin_radicado.pdf").write_bytes(PDF_OK)
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    regs = core.inventariar(origen)
    plan = core.construir_plan(regs, core.cargar_mapeo(ruta), "{id}")
    n_pend, n_ids = core.escribir_pendientes(
        plan, destino / "pendientes_correccion.csv", destino / "ids.txt"
    )

    # modo universal: BETA sin correspondencia y "sin_radicado" usa el
    # nombre completo como origen; ambos quedan pendientes
    assert n_pend == 2
    assert n_ids == 2
    contenido = (destino / "pendientes_correccion.csv").read_text(
        encoding="utf-8-sig"
    )
    assert "accion_sugerida" in contenido
    assert "Obtenga el identificador" in contenido
    assert "sin correspondencia" in contenido.lower()


def test_modo_radicado_estricto_rechaza_otros_nombres(carpeta, tmp_path):
    origen, _ = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    crear_pdf(origen, "acta_de_reunion.pdf")
    regs = core.inventariar(origen, modo=core.MODO_RADICADO)

    por_estado = {r["archivo"]: r["anomalia"] for r in regs}
    assert por_estado["2026XX840820X1 ACME.pdf"] is None
    assert "No inicia" in por_estado["acta_de_reunion.pdf"]


def test_cancelacion_copia(carpeta, tmp_path):
    origen, destino = carpeta
    for i in range(4):
        crear_pdf(origen, f"2026XX84082{i}X1 DOC{i}.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(
        ruta,
        [(f"2026XX84082{i}X1", f"99000000{i:02d}") for i in range(4)],
    )
    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )

    with pytest.raises(core.ProcesoCancelado):
        core.ejecutar_copias(plan, origen, destino, cancelado=lambda: True)


# --- Deshacer -------------------------------------------------------------------


def test_deshacer_borra_solo_copias_del_reporte(carpeta, tmp_path):
    origen, destino = carpeta
    crear_pdf(origen, "2026XX840820X1 ACME.pdf")
    ruta = tmp_path / "m.xlsx"
    crear_excel(ruta, [("2026XX840820X1", "9900000001")])

    plan = core.construir_plan(
        core.inventariar(origen), core.cargar_mapeo(ruta), "{id}"
    )
    core.ejecutar_copias(plan, origen, destino)
    core.escribir_reporte(plan, destino / "reporte.csv")

    # Archivo ajeno que NO debe borrarse
    (destino / "archivo_ajeno.pdf").write_bytes(PDF_OK)

    eliminados, errores = core.deshacer_desde_reporte(
        destino / "reporte.csv", origen, destino
    )

    assert eliminados == 1 and errores == []
    assert not (destino / "9900000001.pdf").exists()
    assert (destino / "archivo_ajeno.pdf").exists()


def test_deshacer_sin_reporte(carpeta):
    origen, destino = carpeta
    eliminados, errores = core.deshacer_desde_reporte(
        destino / "reporte.csv", origen, destino
    )
    assert eliminados == 0 and errores