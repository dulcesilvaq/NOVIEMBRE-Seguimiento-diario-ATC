from datetime import date

import pytest

from src.buckets import ADEL, ANT, clave_periodo, clave_vencimiento, es_cierre, params_desde_config

CFG = {"mes_actual": "2026-10", "fecha_cierre": "2026-11-05",
       "anio_inicio_columnas": 2022, "periodo_anteriores": 2019}
P = params_desde_config(CFG)


def test_limites_adelanto():
    assert clave_vencimiento(date(2026, 10, 31), P) == "2026-10"
    assert clave_vencimiento(date(2026, 11, 1), P) == ADEL


def test_cierre_vence_despues_del_cierre_entra_siempre():
    assert es_cierre(date(2026, 11, 6), date(2026, 10, 28), P)    # regla de antes
    assert es_cierre(date(2026, 11, 20), date(2026, 10, 28), P)
    assert es_cierre(date(2026, 11, 20), date(2026, 11, 3), P)    # paga y vence en nov, pero vence tras el cierre
    assert es_cierre(date(2026, 12, 15), date(2026, 11, 3), P)


def test_cierre_vence_del_1_al_5_entra_si_se_pago_en_mes_anterior():
    # filas verdes/amarillas de la captura (trasladadas a octubre): pago en oct, vence 01-05/11
    for venc in (date(2026, 11, 1), date(2026, 11, 3), date(2026, 11, 5)):
        assert es_cierre(venc, date(2026, 10, 28), P)
    assert es_cierre(date(2026, 11, 5), date(2026, 9, 30), P)     # pagado aún antes


def test_cierre_vence_del_1_al_5_y_pagado_en_noviembre_no_entra():
    # filas blancas de la captura: pago y vencimiento en el mismo mes
    assert not es_cierre(date(2026, 11, 3), date(2026, 11, 1), P)
    assert not es_cierre(date(2026, 11, 5), date(2026, 11, 5), P)
    assert not es_cierre(date(2026, 11, 3), date(2026, 11, 4), P)


def test_cierre_no_aplica_fuera_de_adelanto():
    assert not es_cierre(date(2026, 10, 20), date(2026, 10, 6), P)   # vence en el mes actual
    assert not es_cierre(date(2026, 9, 10), date(2026, 10, 6), P)    # pago atrasado


def test_cierre_editable():
    p = params_desde_config({**CFG, "fecha_cierre": "07/11/2026"})
    assert p.cierre_desde == date(2026, 11, 8)
    assert es_cierre(date(2026, 11, 7), date(2026, 10, 30), p)       # 07/11 ya es "del 1 al cierre"
    assert not es_cierre(date(2026, 11, 7), date(2026, 11, 2), p)
    assert es_cierre(date(2026, 11, 8), date(2026, 11, 2), p)


def test_anteriores_vencimiento():
    assert clave_vencimiento(date(2021, 12, 31), P) == ANT
    assert clave_vencimiento(date(2022, 1, 1), P) == "2022"
    assert clave_vencimiento(date(2025, 12, 31), P) == "2025"


def test_periodo_anteriores():
    assert clave_periodo(2018, P) == ANT and clave_periodo(2019, P) == ANT
    assert clave_periodo(2023, P) == 2023


def test_diciembre_pasa_a_enero():
    p = params_desde_config({**CFG, "mes_actual": "2026-12", "fecha_cierre": "2027-01-05"})
    assert p.adelanto_desde == date(2027, 1, 1) and p.cierre_desde == date(2027, 1, 6)
    assert es_cierre(date(2027, 1, 3), date(2026, 12, 30), p)
    assert not es_cierre(date(2027, 1, 3), date(2027, 1, 2), p)
    assert es_cierre(date(2027, 1, 20), date(2027, 1, 3), p)


def test_fecha_cierre_invalida_o_faltante():
    with pytest.raises(ValueError):
        params_desde_config({**CFG, "fecha_cierre": "2026-10-05"})   # no es del mes siguiente
    with pytest.raises(ValueError):
        params_desde_config({k: v for k, v in CFG.items() if k != "fecha_cierre"})
    with pytest.raises(ValueError):
        params_desde_config({**CFG, "mes_actual": "2026-13"})


def test_rango_tabla1():
    from src.tables import dias_tabla1
    d = dias_tabla1(P, 9)
    assert d[0].date() == date(2026, 10, 1) and d[-1].date() == date(2026, 11, 9) and len(d) == 40
