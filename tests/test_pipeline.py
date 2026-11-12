import io

import pandas as pd

from src.pipeline import generar

CFG = {
    "hojas_base": ["BASE FINAL", "BASE"], "mes_actual": "2026-10", "fecha_cierre": "2026-11-05",
    "tabla1_hasta_dia": 9, "anio_inicio_columnas": 2022, "periodo_anteriores": 2019,
    "etiqueta_tabla": "x", "asesoras": [{"nombre": "Ana, A", "hoja": "Ana"}, {"nombre": "Beto, B", "hoja": "Beto"}],
}


def _excel(filas):
    df = pd.DataFrame(filas, columns=["PERIODO", "Fecha Vencimiento Actual", "Fecha Pago",
                                      "Real Pago Dolar", "Coordinadora Cobranza"])
    df["DNI"] = "12345678"            # dato personal: el programa no debe cargarlo
    df["Email"] = "alguien@correo.com"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as xw:
        df.to_excel(xw, sheet_name="BASE FINAL", index=False)
    buf.seek(0)
    return buf


def test_totales_y_adelanto_cierre():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 100.0, "Ana, A"),   # oct-26 (no es adelanto)
        (2026, d("2026-11-05"), d("2026-10-06"), 50.0, "Ana, A"),    # vence 01-05/11, pago oct -> Cierre
        (2026, d("2026-11-06"), d("2026-10-07"), 30.0, "Ana, A"),    # vence tras el cierre -> Cierre
        (2026, d("2026-11-03"), d("2026-11-02"), 20.0, "Ana, A"),    # vence 01-05/11, pago nov -> NO Cierre
        (2026, d("2026-11-20"), d("2026-11-03"), 40.0, "Ana, A"),    # vence tras el cierre, pago nov -> Cierre
        (2019, d("2021-12-31"), d("2026-10-07"), 10.0, "Beto, B"),   # Anteriores
        (2026, d("2026-10-20"), d("2026-10-07"), 0.0, "Beto, B"),    # pago 0: no cuenta
    ])
    r = generar(base, CFG)
    assert r.n_filas == 6
    assert r.res_mes.loc["Total general", "Total general"] == 250.0
    m = r.vistas["Ana"]["monto"]
    assert m.loc["2026", "Adelanto"] == 140.0           # 50 + 30 + 20 + 40
    assert m.loc["2026", "Adelanto Cierre"] == 120.0    # 50 + 30 + 40 (sin los 20 pagados en nov)
    assert m.loc["2026", "Total general"] == 240.0      # 100 + 140; Cierre no se duplica
    c = r.vistas["Ana"]["cuotas"]
    assert c.loc["2026", "Adelanto"] == 4 and c.loc["2026", "Adelanto Cierre"] == 3
    assert r.vistas["Beto"]["cuotas"].loc["Anteriores", "Anteriores"] == 1


def test_coordinadora_nueva_avisa():
    d = pd.Timestamp
    r = generar(_excel([(2026, d("2026-10-20"), d("2026-10-06"), 5.0, "Perez, Nora")]), CFG)
    assert any("Perez, Nora" in a for a in r.advertencias)
    assert "Nora" in r.libro.sheetnames


def test_periodo_texto_anteriores_no_se_descarta():
    d = pd.Timestamp
    df = pd.DataFrame([
        ("Anteriores", d("2020-01-10"), d("2026-10-06"), 44.78, "Ana, A"),
        (2026, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A"),
    ], columns=["PERIODO", "Fecha Vencimiento Actual", "Fecha Pago", "Real Pago Dolar", "Coordinadora Cobranza"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as xw:
        df.to_excel(xw, sheet_name="BASE", index=False)
    buf.seek(0)
    r = generar(buf, CFG)
    assert r.n_filas == 2
    assert r.vistas["Ana"]["monto"].loc["Anteriores", "Total general"] == 44.78


def test_hoja_nueva_no_choca_y_asesora_sin_pagos_se_omite():
    d = pd.Timestamp
    cfg = {**CFG, "asesoras": [{"nombre": "Ana, A", "hoja": "Ana"}, {"nombre": "Beto, B", "hoja": "Beto"}]}
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 5.0, "Ana, A"),
        (2026, d("2026-10-20"), d("2026-10-06"), 7.0, "Lopez, Ana Maria"),   # primer nombre 'Ana' ya existe
    ])
    r = generar(base, cfg)
    assert r.libro.sheetnames == ["RES MES", "RESUMEN ADELANTOS", "VENCIDOS 1-5 NOV", "VENCIDOS 6 NOV EN ADELANTE", "Ana", "Ana L", "DETALLE"]
    assert any("SIN pagos" in a and "Beto, B" in a for a in r.advertencias)


def test_aviso_si_el_mes_actual_no_corresponde_a_la_base():
    d = pd.Timestamp
    base = _excel([(2026, d("2026-09-20"), d("2026-08-20"), 5.0, "Ana, A")])   # pagos de agosto
    r = generar(base, CFG)                                                     # config en octubre
    assert "2026-08" in r.advertencias[0] and "REVISA EL CICLO" in r.advertencias[0]


def test_detalle_audita_los_totales():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 100.0, "Ana, A"),
        (2026, d("2026-11-05"), d("2026-10-06"), 50.0, "Ana, A"),    # Cierre
        (2026, d("2026-11-03"), d("2026-11-02"), 20.0, "Ana, A"),    # Adelanto, NO cierre
        (2019, d("2021-12-31"), d("2026-10-07"), 10.0, "Beto, B"),
    ])
    r = generar(base, CFG)
    assert r.libro.sheetnames[-1] == "DETALLE"
    ws = r.libro["DETALLE"]
    filas = list(ws.iter_rows(values_only=True))
    enc = list(filas[0])
    assert len(filas) - 1 == r.n_filas
    i_monto, i_cierre, i_col, i_fila = (enc.index(x) for x in
                                        ("Real Pago Dolar", "Adelanto Cierre", "Columna de vencimiento", "Fila en BASE"))
    assert sum(f[i_monto] for f in filas[1:]) == 180.0                      # = total de RES MES
    assert sum(f[i_monto] for f in filas[1:] if f[i_cierre] == "SI") == 50.0
    assert sum(f[i_monto] for f in filas[1:] if f[i_col] == "Adelanto") == 70.0
    assert sorted(f[i_fila] for f in filas[1:]) == [2, 3, 4, 5]              # filas de la hoja original
    assert ws.tables["Detalle"].ref == f"A1:{chr(64 + len(enc))}{len(filas)}"


def test_fusionar_suma_la_cartera_a_otra_asesora():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 100.0, "Beto, B"),
        (2026, d("2026-10-20"), d("2026-10-07"), 30.0, "Nueva, N"),
        (2026, d("2026-10-20"), d("2026-10-08"), 20.0, "Otra, O"),
    ])
    cfg = {**CFG, "fusionar": {"Nueva, N": "Ana, A", "Otra, O": "Ana, A", "Fantasma, F": "Ana, A"}}
    r = generar(base, cfg)
    assert r.res_mes.loc["Total general", "Ana, A"] == 50.0           # 30 + 20 sumados a Ana
    assert r.libro.sheetnames == ["RES MES", "RESUMEN ADELANTOS", "VENCIDOS 1-5 NOV", "VENCIDOS 6 NOV EN ADELANTE", "Ana", "Beto", "DETALLE"]  # no se crean hojas para las unidas
    assert any("Unidas a la cartera de 'Ana, A'" in a and "Nueva, N" in a for a in r.advertencias)
    assert any("Fantasma, F" in a and "no aparece" in a for a in r.advertencias)
    ws = r.libro["DETALLE"]
    enc = [c.value for c in ws[1]]
    orig = {row[enc.index("Coordinadora original")] for row in ws.iter_rows(min_row=2, values_only=True)
            if row[enc.index("Hoja")] == "Ana"}
    assert orig == {"Nueva, N", "Otra, O"}                              # se conserva el nombre original


def test_resumen_adelantos_y_clasificacion_en_detalle():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 100.0, "Ana, A"),   # oct-26: no es adelanto
        (2026, d("2026-11-05"), d("2026-10-06"), 50.0, "Ana, A"),    # vence 1-5, pagado oct   -> Adelanto, Cierre, V1-5
        (2026, d("2026-11-06"), d("2026-10-07"), 30.0, "Ana, A"),    # vence desde el 6        -> Adelanto, Cierre, V6+
        (2026, d("2026-11-03"), d("2026-11-02"), 20.0, "Ana, A"),    # vence 1-5, pagado nov   -> Adelanto, V1-5 (no Cierre)
        (2026, d("2026-11-20"), d("2026-11-03"), 40.0, "Ana, A"),    # vence desde el 6        -> Adelanto, Cierre, V6+
        (2026, d("2026-12-15"), d("2026-10-30"), 10.0, "Beto, B"),   # vence desde el 6
    ])
    r = generar(base, CFG)
    assert r.libro.sheetnames[:2] == ["RES MES", "RESUMEN ADELANTOS"]
    rs = r.resumen
    assert list(rs.columns) == ["Adelanto mes natural", "Adelanto considerado meta en reportes anteriores",
                                "Pagos vencidos del 1 al 5 de nov", "Pagos vencidos a partir del 6 de Noviembre"]
    assert list(rs.loc["Ana"]) == [140.0, 120.0, 70.0, 70.0]
    assert list(rs.loc["Beto"]) == [10.0, 10.0, 0.0, 10.0]
    assert list(rs.loc["Total"]) == [150.0, 130.0, 70.0, 80.0]
    assert (rs.iloc[:, 2] + rs.iloc[:, 3]).round(2).tolist() == rs.iloc[:, 0].round(2).tolist()   # 1-5 + 6+ = Adelanto
    ws = r.libro["RESUMEN ADELANTOS"]
    assert [c.value for c in ws[1]] == ["Asesora", *rs.columns]
    assert [c.value for c in ws[ws.max_row]] == ["Total", 150.0, 130.0, 70.0, 80.0]
    # DETALLE trae la clasificación de cada pago
    det = r.libro["DETALLE"]
    enc = [c.value for c in det[1]]
    for col in ("Adelanto", "Adelanto Cierre", "Vencido del 1 al 5 de nov", "Vencido desde el 6 de nov"):
        assert col in enc
    filas = {row[enc.index("Real Pago Dolar")]: row for row in det.iter_rows(min_row=2, values_only=True)}
    f20 = filas[20.0]
    assert (f20[enc.index("Adelanto")], f20[enc.index("Adelanto Cierre")],
            f20[enc.index("Vencido del 1 al 5 de nov")], f20[enc.index("Vencido desde el 6 de nov")]) == ("SI", "NO", "SI", "NO")
    f100 = filas[100.0]
    assert (f100[enc.index("Adelanto")], f100[enc.index("Adelanto Cierre")]) == ("NO", "NO")


def test_etiquetas_de_diciembre():
    from src.buckets import etiquetas_adelantos, params_desde_config
    p = params_desde_config({"mes_actual": "2026-11", "fecha_cierre": "2026-12-05",
                             "anio_inicio_columnas": 2022, "periodo_anteriores": 2019})
    et = etiquetas_adelantos(p)
    assert et["v15"] == "Pagos vencidos del 1 al 5 de dic"
    assert et["v6"] == "Pagos vencidos a partir del 6 de Diciembre"


def test_periodo_vacio_cuenta_en_totales_pero_no_agrega_fila():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A"),
        (None, d("2026-10-30"), d("2026-10-08"), 87.13, "Ana, A"),         # PERIODO vacío
    ])
    r = generar(base, CFG)
    assert r.n_filas == 2
    assert r.res_mes.loc["Total general", "Total general"] == 97.13          # cuenta en RES MES
    m = r.vistas["Ana"]["monto"]
    assert list(m.index) == ["Anteriores", "2023", "2024", "2025", "2026"]   # filas fijas: no se agrega "Sin PERIODO"
    assert m.loc["2026", "Total general"] == 10.0                            # el pago sin PERIODO no está en las tablas 2 y 3
    assert any("PERIODO vacío" in a and "87.13" in a for a in r.advertencias)


def test_filas_de_periodo_fijas_aunque_la_base_no_tenga_esos_anios():
    d = pd.Timestamp
    base = _excel([(2025, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A"),
                   (2026, d("2026-10-20"), d("2026-10-07"), 20.0, "Ana, A")])   # no hay 2023 ni 2024 ni Anteriores
    r = generar(base, CFG)
    assert list(r.vistas["Ana"]["monto"].index) == ["Anteriores", "2023", "2024", "2025", "2026"]
    ws = r.libro["Ana"]
    etiquetas = [ws.cell(i, 2).value for i in range(1, ws.max_row + 1)]
    for k in ("Anteriores", "2023", "2024", "2025", "2026"):
        assert etiquetas.count(k) == 2                                       # una vez en la tabla 2 y otra en la 3


def test_periodos_antes_del_primer_anio_van_a_anteriores_y_fuera_de_lista_avisa():
    d = pd.Timestamp
    base = _excel([(2021, d("2026-10-20"), d("2026-10-06"), 5.0, "Ana, A"),      # 2021 -> Anteriores
                   (2027, d("2026-10-20"), d("2026-10-07"), 7.0, "Ana, A")])     # fuera de la lista
    r = generar(base, CFG)
    assert r.vistas["Ana"]["monto"].loc["Anteriores", "Total general"] == 5.0
    assert any("2027" in a and "periodos_filas" in a for a in r.advertencias)


def test_fila_sin_fecha_se_excluye_con_aviso_sin_error():
    d = pd.Timestamp
    base = _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A"),
        (2026, pd.NaT, d("2026-10-08"), 20.0, "Ana, A"),                     # sin fecha de vencimiento
    ])
    r = generar(base, CFG)
    assert r.n_filas == 1
    assert any("datos vacíos" in a and "Fecha Vencimiento Actual" in a for a in r.advertencias)


def test_periodo_que_no_es_un_anio_no_cae_en_anteriores_y_codigo_de_respaldo():
    d = pd.Timestamp
    df = pd.DataFrame([
        (2026, "10751748A00001", "10751748A00001", d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A"),
        (282.02, d("2026-11-11"), "10751748A32378", d("2026-10-30"), d("2026-10-08"), 20.0, "Ana, A"),   # fila desalineada
    ], columns=["PERIODO", "Código Matricula", "Código Matricula", "Fecha Vencimiento Actual", "Fecha Pago",
                "Real Pago Dolar", "Coordinadora Cobranza"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as xw:
        df.to_excel(xw, sheet_name="BASE", index=False)
    buf.seek(0)
    r = generar(buf, CFG)
    m = r.vistas["Ana"]["monto"]
    assert m.loc["Anteriores", "Total general"] == 0.0                       # 282.02 NO se toma como 'Anteriores'
    assert m.loc["2026", "Total general"] == 10.0
    assert r.res_mes.loc["Total general", "Total general"] == 30.0           # pero el pago sí cuenta en RES MES
    assert any("282.02" in a and "PERIODO" in a for a in r.advertencias)
    assert any("Código Matricula' inválido" in a for a in r.advertencias)
    det = r.libro["DETALLE"]; enc = [c.value for c in det[1]]
    codigos = {row[enc.index("Código Matricula")] for row in det.iter_rows(min_row=2, values_only=True)}
    assert codigos == {"10751748A00001", "10751748A32378"}                    # el código válido, no la fecha


def _base_vencidos():
    d = pd.Timestamp
    return _excel([
        (2026, d("2026-10-20"), d("2026-10-06"), 100.0, "Ana, A"),   # oct-26: no es adelanto
        (2026, d("2026-11-05"), d("2026-10-06"), 50.0, "Ana, A"),    # vence 1-5, pagado oct  -> Adelanto Cierre
        (2026, d("2026-11-03"), d("2026-11-02"), 20.0, "Ana, A"),    # vence 1-5, pagado nov  -> Adelanto (sin cierre)
        (2026, d("2026-11-06"), d("2026-10-07"), 30.0, "Beto, B"),   # vence desde el 6        -> Adelanto Cierre
        (2026, d("2026-12-15"), d("2026-10-07"), 10.0, "Beto, B"),   # vence desde el 6
    ])


def test_hojas_de_vencidos_1_al_cierre_y_desde_el_dia_siguiente():
    r = generar(_base_vencidos(), CFG)
    assert r.libro.sheetnames[:4] == ["RES MES", "RESUMEN ADELANTOS", "VENCIDOS 1-5 NOV", "VENCIDOS 6 NOV EN ADELANTE"]
    p15, p6 = r.vencidos["1-5"], r.vencidos["6+"]
    assert p15.loc["Total general", "Total general"] == 70.0            # 50 + 20
    assert p15.loc["Total general", "Ana, A"] == 70.0 and pd.isna(p15.loc["Total general", "Beto, B"])
    assert p6.loc["Total general", "Total general"] == 40.0             # 30 + 10
    assert p6.loc["Total general", "Beto, B"] == 40.0
    # las dos hojas suman el Adelanto mes natural y usan los mismos días que RES MES
    assert list(p15.index) == list(r.res_mes.index)
    assert p15.loc["Total general", "Total general"] + p6.loc["Total general", "Total general"] == 110.0

    ws = r.libro["VENCIDOS 1-5 NOV"]
    texto = [[c for c in fila] for fila in ws.iter_rows(values_only=True)]
    assert any(f[0] == "Numero de cuotas" and f[1] == 2 for f in texto)               # 2 pagos de Ana
    enc = next(f for f in texto if f[0] == "Fecha Pago" and f[1] == "Asesora")
    assert "Adelanto Cierre" in enc and "Fila en BASE" in enc
    filas = [f for f in texto[texto.index(enc) + 1:] if f[0] is not None]
    assert sorted(f[enc.index("Real Pago Dolar")] for f in filas) == [20.0, 50.0]
    cierre = {f[enc.index("Real Pago Dolar")]: f[enc.index("Adelanto Cierre")] for f in filas}
    assert cierre == {50.0: "SI", 20.0: "NO"}


def test_hojas_de_vencidos_sin_pagos_en_el_rango_no_fallan():
    d = pd.Timestamp
    r = generar(_excel([(2026, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A")]), CFG)       # nada vence desde el 1/11
    ws = r.libro["VENCIDOS 1-5 NOV"]
    assert any("no hay pagos" in str(c) for fila in ws.iter_rows(values_only=True) for c in fila)
    assert r.vencidos["1-5"].loc["Total general", "Total general"] != r.vencidos["1-5"].loc["Total general", "Total general"]  # NaN


def test_detalle_trae_categoria_y_motivo_de_cada_pago():
    r = generar(_base_vencidos(), CFG)
    det = r.libro["DETALLE"]
    enc = [c.value for c in det[1]]
    assert "Categoría" in enc and "Motivo" in enc
    por_monto = {row[enc.index("Real Pago Dolar")]: row for row in det.iter_rows(min_row=2, values_only=True)}
    cat = lambda m: por_monto[m][enc.index("Categoría")]
    mot = lambda m: por_monto[m][enc.index("Motivo")]
    assert cat(100.0) == "Otros (no es adelanto)"
    assert cat(50.0) == "Adelanto Cierre" and "mes anterior" in mot(50.0)
    assert cat(20.0) == "Adelanto (sin cierre)" and "mismo mes" in mot(20.0)
    assert cat(30.0) == "Adelanto Cierre" and "después del cierre" in mot(30.0)
    assert cat(10.0) == "Adelanto Cierre"


def test_nombre_de_asesora_no_puede_chocar_con_hojas_de_vencidos():
    import pytest
    d = pd.Timestamp
    cfg = {**CFG, "asesoras": [{"nombre": "Ana, A", "hoja": "VENCIDOS 1-5 NOV"}]}
    with pytest.raises(ValueError):
        generar(_excel([(2026, d("2026-10-20"), d("2026-10-06"), 10.0, "Ana, A")]), cfg)
