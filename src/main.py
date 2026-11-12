"""Genera el reporte de cobranza desde la línea de comandos.

Uso:
    python -m src.main --input data/base.xlsx
    python -m src.main --input data/base.xlsx --solo Katherin      # solo una(s) asesora(s) (pruebas)

(También puedes usar la app web:  streamlit run app.py)
"""
from __future__ import annotations

import argparse
from pathlib import Path

import yaml

from . import buckets
from .pipeline import generar


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Reporte de cobranza (RES MES + hojas por asesora)")
    ap.add_argument("--input", required=True, help="Excel con la hoja BASE FINAL")
    ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--output", default=None, help="Ruta del Excel de salida")
    ap.add_argument("--solo", nargs="*", default=None,
                    help="Nombres de hoja (ej. Katherin) para generar solo esas asesoras")
    args = ap.parse_args(argv)

    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    r = generar(args.input, cfg, args.solo)

    salida = Path(args.output or f"output/Reporte_cobranza_{cfg['mes_actual']}.xlsx")
    salida.parent.mkdir(parents=True, exist_ok=True)
    r.libro.save(salida)

    print(f"✔ Generado: {salida}")
    print(f"  Mes actual: {cfg['mes_actual']} | {buckets.descripcion_ciclo(r.params)}")
    print(f"  RES MES: pagos del {r.dias_base[0]:%d/%m/%Y} al {r.dias_base[-1]:%d/%m/%Y}")
    print(f"  Tabla 1 de cada asesora: {r.dias_tabla1[0]:%d/%m/%Y} a {r.dias_tabla1[-1]:%d/%m/%Y} "
          f"({len(r.dias_tabla1)} días)")
    for m in r.advertencias:
        print(f"  ⚠ {m}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
