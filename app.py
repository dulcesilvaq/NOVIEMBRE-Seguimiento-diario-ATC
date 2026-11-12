"""App web (Streamlit): sube la base, ajusta mes / cierre / asesoras en pantalla y descarga el Excel.

Ejecutar en local:   streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
import yaml

from src import buckets
from src.loader import listar_coordinadoras
from src.pipeline import generar

MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
CFG = yaml.safe_load(Path("config.yaml").read_text(encoding="utf-8"))   # valores iniciales
NOMBRE_COL, HOJA_COL, ORDEN_COL = "Coordinadora Cobranza (como viene en la base)", "Hoja", "Orden"

st.set_page_config(page_title="Reporte de cobranza", page_icon="📊", layout="wide")
st.title("📊 Reporte de cobranza")
st.caption("Sube la base, ajusta el ciclo y descarga el Excel (RES MES + una hoja por asesora). "
           "El archivo se procesa en memoria y solo se leen las columnas necesarias "
           "(no se leen DNI, correos ni teléfonos).")

# ------------------------------------------------------------------ 1. Archivo
archivo = st.file_uploader("1 · Excel con la hoja BASE FINAL", type=["xlsx", "xlsm"])

# ------------------------------------------------------------------ 2. Ciclo
st.subheader("2 · Ciclo")
a0, m0 = (int(x) for x in str(CFG["mes_actual"]).split("-"))
c1, c2, c3, c4 = st.columns(4)
anio = c1.number_input("Año actual", 2020, 2100, a0, step=1)
mes = c2.selectbox("Mes actual", range(1, 13), index=m0 - 1, format_func=lambda m: MESES_ES[m - 1])

# Cierre por defecto: el de la config si es el mismo mes; si no, el mismo día del mes siguiente
a_sig, m_sig = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
cierre_cfg = buckets._leer_fecha(CFG.get("fecha_cierre") or f"{a0 + (m0 == 12)}-{m0 % 12 + 1:02d}-05", "fecha_cierre")
cierre_def = cierre_cfg if (anio, mes) == (a0, m0) else date(a_sig, m_sig, min(cierre_cfg.day, 28))
fecha_cierre = c3.date_input("Fecha de cierre ATC", value=cierre_def, format="DD/MM/YYYY",
                             key=f"cierre_{anio}_{mes}",
                             help="Adelanto Cierre = vencen después de esta fecha, más los que vencen del 1 hasta "
                                  "esta fecha pero se pagaron en un mes anterior.")
tabla1_hasta = c4.number_input("Tabla 1 hasta el día (del mes siguiente)", 1, 28,
                               int(CFG["tabla1_hasta_dia"]), step=1)

with st.expander("Ajustes avanzados"):
    a1, a2, a3, a4 = st.columns(4)
    anio_inicio = a1.number_input("Columnas por año desde (antes = Anteriores)", 2000, 2100,
                                  int(CFG["anio_inicio_columnas"]), step=1)
    per_ant = a2.number_input("PERIODO ≤ este valor = Anteriores", 2000, 2100,
                              int(CFG["periodo_anteriores"]), step=1)
    etiqueta = a3.text_input("Encabezado de las tablas", CFG["etiqueta_tabla"])
    filas_txt = a4.text_input("Filas de PERIODO (además de Anteriores)",
                              ", ".join(str(x) for x in CFG.get("periodos_filas", [2023, 2024, 2025, 2026])),
                              help="Siempre salen todas, aunque no haya pagos de ese año.")

# ------------------------------------------------------------------ 3. Asesoras
st.subheader("3 · Asesoras")
st.caption("El número de **Orden** define el orden de columnas en RES MES y de las pestañas. "
           "Puedes cambiar nombres, agregar filas (+) o borrar. El nombre debe ser idéntico al de la base.")
base_ases = pd.DataFrame({
    ORDEN_COL: range(1, len(CFG["asesoras"]) + 1),
    NOMBRE_COL: [a["nombre"] for a in CFG["asesoras"]],
    HOJA_COL: [a["hoja"] for a in CFG["asesoras"]],
})
editado = st.data_editor(base_ases, num_rows="dynamic", width="stretch", hide_index=True,
                         column_config={ORDEN_COL: st.column_config.NumberColumn(ORDEN_COL, step=1, width="small")})

# ------------------------------------------------------------------ Unir asesoras
with st.expander("Unir asesoras (sumar la cartera de una asesora a la de otra)"):
    st.caption("Ej.: si una asesora se fue y su cartera pasó a otras, elige quién se une y a la cartera de quién. "
               "La asesora destino debe estar en la tabla de arriba.")
    try:
        nombres_base = listar_coordinadoras(archivo, CFG["hojas_base"]) if archivo is not None else []
    except Exception:  # noqa: BLE001
        nombres_base = []
    destinos = sorted({n.strip() for n in editado[NOMBRE_COL].dropna()} | set(nombres_base))
    ini = [{"Se une": o, "A la cartera de": d} for o, d in (CFG.get("fusionar") or {}).items()]
    fus_ini = pd.DataFrame(ini, columns=["Se une", "A la cartera de"])
    fus_ed = st.data_editor(
        fus_ini, num_rows="dynamic", width="stretch", hide_index=True, key="fusion_editor",
        column_config={
            "Se une": (st.column_config.SelectboxColumn("Se une", options=nombres_base)
                       if nombres_base else st.column_config.TextColumn("Se une")),
            "A la cartera de": st.column_config.SelectboxColumn("A la cartera de", options=destinos),
        })

# ------------------------------------------------------------------ Generar
st.divider()
if st.button("▶ Generar reporte", type="primary", disabled=archivo is None):
    ases = (editado.dropna(subset=[NOMBRE_COL, HOJA_COL])
                   .assign(**{ORDEN_COL: lambda d: d[ORDEN_COL].fillna(999)})
                   .sort_values(ORDEN_COL, kind="stable"))
    ases = ases[(ases[NOMBRE_COL].str.strip() != "") & (ases[HOJA_COL].str.strip() != "")]
    lista = [{"nombre": n.strip(), "hoja": h.strip()} for n, h in zip(ases[NOMBRE_COL], ases[HOJA_COL])]

    hojas = [a["hoja"] for a in lista]
    if not lista:
        st.error("Agrega al menos una asesora.")
    elif len(set(hojas)) != len(hojas):
        st.error("Hay nombres de hoja repetidos. Cada asesora necesita un nombre de pestaña distinto.")
    elif any(len(h) > 31 or any(ch in h for ch in '[]:*?/\\') or h.upper() in ("RES MES", "DETALLE", "RESUMEN ADELANTOS", "DETALLE ADELANTOS") for h in hojas):
        st.error("Nombre de hoja no válido (máx. 31 caracteres, sin [ ] : * ? / \\, y distinto de 'RES MES', 'RESUMEN ADELANTOS', 'DETALLE ADELANTOS' y 'DETALLE').")
    else:
        try:
            filas_periodos = [int(x) for x in filas_txt.replace(";", ",").split(",") if x.strip()]
            if not filas_periodos:
                raise ValueError
        except ValueError:
            st.error("'Filas de PERIODO' debe ser una lista de años separados por coma (ej. 2023, 2024, 2025, 2026).")
            st.stop()
        cfg = {
            "hojas_base": CFG["hojas_base"], "mes_actual": f"{anio}-{mes:02d}",
            "fecha_cierre": fecha_cierre, "tabla1_hasta_dia": int(tabla1_hasta),
            "anio_inicio_columnas": int(anio_inicio), "periodo_anteriores": int(per_ant),
            "periodos_filas": filas_periodos,
            "etiqueta_tabla": etiqueta, "asesoras": lista,
            "fusionar": {o.strip(): d.strip() for o, d in zip(fus_ed["Se une"], fus_ed["A la cartera de"])
                         if isinstance(o, str) and isinstance(d, str) and o.strip() and d.strip()},
        }
        try:
            with st.spinner("Calculando..."):
                r = generar(archivo, cfg)
                buf = io.BytesIO()
                r.libro.save(buf)
            st.session_state["resultado"] = {
                "r": r, "excel": buf.getvalue(), "nombre": f"Reporte_cobranza_{anio}-{mes:02d}.xlsx",
                "yaml": yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True),
            }
        except ValueError as e:
            st.session_state.pop("resultado", None)
            st.error(str(e))
        except Exception as e:  # noqa: BLE001
            st.session_state.pop("resultado", None)
            st.error(f"No se pudo generar el reporte: {type(e).__name__}: {e}")
elif archivo is None:
    st.info("Sube el Excel para habilitar el botón.")

# ------------------------------------------------------------------ Resultado
res = st.session_state.get("resultado")
if res:
    r = res["r"]
    st.success(f"Listo · {r.n_filas:,} pagos procesados · Tabla 1 de {r.dias_tabla1[0]:%d/%m/%Y} "
               f"a {r.dias_tabla1[-1]:%d/%m/%Y}")
    for m in r.advertencias:
        st.warning(m)
    st.caption(buckets.descripcion_ciclo(r.params))

    d1, d2 = st.columns([1, 1])
    d1.download_button("⬇ Descargar Excel", res["excel"], file_name=res["nombre"], type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇ Descargar config.yaml con estos ajustes", res["yaml"], file_name="config.yaml",
                       mime="text/yaml", help="Úsalo para reemplazar config.yaml en GitHub si quieres que "
                       "estos valores queden como los iniciales de la app (el archivo sale sin comentarios).")

    t1, t3, t4, t2 = st.tabs(["RES MES", "Resumen adelantos", "Detalle adelantos", "Vista por asesora"])
    with t1:
        v = r.res_mes.copy()
        v.index = [i.strftime("%d/%m/%Y") if isinstance(i, pd.Timestamp) else i for i in v.index]
        st.dataframe(v.style.format("{:.0f}", na_rep=""), width="stretch")
    with t3:
        st.dataframe(r.resumen.style.format("{:.0f}"), width="stretch")
        st.caption("Montos en US$. Los 'Pagos vencidos' se clasifican solo por fecha de vencimiento; "
                   "las dos columnas de vencimiento suman el Adelanto mes natural.")
    with t4:
        st.dataframe(r.detalle_adelantos, width="stretch", hide_index=True,
                     column_config={"Fecha Pago": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                    "Fecha Vencimiento": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                    "Monto US$": st.column_config.NumberColumn(format="%.2f")})
        st.caption("Un pago por fila. SI/NO según su fecha de vencimiento y de pago; el Excel trae además los totales de SI.")
    with t2:
        hoja = st.selectbox("Asesora", list(r.vistas))
        for clave, titulo in (("monto", "Monto (suma de Real Pago Dolar) por PERIODO"),
                              ("cuotas", "Numero de cuotas por PERIODO")):
            st.markdown(f"**{titulo}**")
            st.dataframe(r.vistas[hoja][clave].style.format("{:.0f}"), width="stretch")
        st.caption("La Tabla 1 (por día) está en la hoja de cada asesora del Excel.")
