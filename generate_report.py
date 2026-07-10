#!/usr/bin/env python3
"""
generate_report.py
-------------------
Genera el Reporte Semanal de NCRs (dashboard HTML interactivo) a partir del
Excel exportado de la plataforma de encuestas (SurveyMars).

USO:
    python3 generate_report.py "Reporte de NCRs.xlsx" [salida.html]

Si no se indica un archivo de salida, se genera "index.html" en la carpeta
actual (listo para subir a GitHub Pages).

Reglas de negocio aplicadas (ver /areas/ncr-report.md para contexto):
  - El campo real de planta viene de la columna "Detalles del colector":
        0706      -> Tarimas y Contenedores (Juárez)
        1122/8844 -> Custom Crates and Pallets (El Paso)
        081218Fa  -> Tarimas Regias (Monterrey)
  - EXCEPCIÓN: si el inspector es "Jahaziel Soto" o "Mario Flores", la planta
    siempre es Custom Crates and Pallets (El Paso), sin importar el colector.
  - Se corrige el proveedor "Misma" -> "Mimsa".
  - Se eliminan filas duplicadas (reenvíos accidentales del mismo NCR).
"""

import sys
import re
import json
import pandas as pd
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR / "template.html"

COLMAP = {
    'Índice': 'idx',
    'Detalles del colector': 'colector',
    '1. Fecha en que se levanta el NCR': 'fecha',
    '2. Nombre de la persona que encontro la No Conformidad': 'inspector',
    '3. Razón de la No Conformidad (NCR)': 'razon',
    '4. Numero de NCR': 'ncr_num',
    '5. Area en donde se encontro la No Conformidad': 'area',
    '6. Nombre de Proveedor': 'proveedor',
    '7. Nombre de Operador responsable (donde se encontro la No Conformidad), si no aplica poner N/A': 'operador',
    '8. Numero de Parte (medida) a Inspeccionar': 'parte',
    '9. Cantidad\xa0': 'cantidad',
    '10. Reportado a: (Indicar el Nombre de la Personal)': 'reportado_a',
    '12. Disposicion\xa0': 'disposicion',
    '13. Descripcion del Retrabajo': 'descripcion',
    '14. Verificado y Aceptado por': 'verificado_por',
    '15. Fotografia de evidencia de la No Conformidad': 'fotos',
}


def map_planta(colector, inspector):
    # Jahaziel Soto y Mario Flores son siempre Custom Crates and Pallets (El Paso),
    # sin importar qué código de colector traiga la fila (caso conocido de código cruzado).
    if inspector in ('Jahaziel Soto', 'Mario Flores'):
        return 'Custom Crates and Pallets (El Paso)'
    if colector == '0706':
        return 'Tarimas y Contenedores (Juárez)'
    elif colector in ('1122', '8844'):
        return 'Custom Crates and Pallets (El Paso)'
    elif colector == '081218Fa':
        return 'Tarimas Regias (Monterrey)'
    return 'Sin identificar'


def clean_razon(r):
    if pd.isna(r):
        return 'Sin especificar'
    r = str(r).strip()
    m = re.match(r'Otro\.?\s*Especificar:\[(.*)\]', r, re.IGNORECASE)
    if m:
        label = m.group(1).strip().rstrip('.')
        return label if label else 'Otro'
    return r


def split_fotos(f):
    if pd.isna(f) or not str(f).strip():
        return []
    return [u.strip() for u in str(f).split(';') if u.strip()]


def load_and_clean(xlsx_path):
    xls = pd.ExcelFile(xlsx_path)
    sheet_name = 'Registro de No Conformidad-NCR' if 'Registro de No Conformidad-NCR' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df = df.rename(columns=COLMAP)
    df = df[[c for c in COLMAP.values() if c in df.columns]]
    df = df.dropna(subset=['idx'])

    df['colector'] = df['colector'].astype(str).str.strip()
    df['planta'] = df.apply(lambda row: map_planta(row['colector'], row['inspector']), axis=1)
    df['fecha_dt'] = pd.to_datetime(df['fecha'], errors='coerce')
    df['razon_clean'] = df['razon'].apply(clean_razon)
    df['area'] = df['area'].fillna('Sin especificar').astype(str).str.strip()
    df['disposicion'] = df['disposicion'].fillna('Sin especificar').astype(str).str.strip()
    df['proveedor'] = df['proveedor'].replace('Misma', 'Mimsa')

    before = len(df)
    df = df.drop_duplicates(subset=['colector', 'ncr_num', 'razon', 'fecha', 'area', 'cantidad'], keep='first')
    after = len(df)
    print(f"[generate_report] Filas antes de deduplicar: {before}, despues: {after}")

    df['fotos_list'] = df['fotos'].apply(split_fotos)
    return df


def build_data_dict(df):
    records = []
    for _, r in df.iterrows():
        records.append({
            'idx': int(r['idx']) if pd.notna(r['idx']) else None,
            'ncr': str(r['ncr_num']),
            'fecha': r['fecha_dt'].strftime('%Y-%m-%d') if pd.notna(r['fecha_dt']) else str(r['fecha']),
            'planta': r['planta'],
            'inspector': r['inspector'] if pd.notna(r['inspector']) else '',
            'razon': r['razon_clean'],
            'area': r['area'],
            'proveedor': r['proveedor'] if pd.notna(r['proveedor']) else '',
            'disposicion': r['disposicion'],
            'descripcion': r['descripcion'] if pd.notna(r['descripcion']) else '',
            'cantidad': str(r['cantidad']) if pd.notna(r['cantidad']) else '',
            'operador': r['operador'] if pd.notna(r['operador']) else '',
            'reportado_a': r['reportado_a'] if pd.notna(r['reportado_a']) else '',
            'fotos': r['fotos_list'],
        })

    by_planta = df['planta'].value_counts().to_dict()

    return {
        'records': records,
        'by_planta': by_planta,
        'total': len(df),
        'fecha_min': df['fecha_dt'].min().strftime('%d %b %Y'),
        'fecha_max': df['fecha_dt'].max().strftime('%d %b %Y'),
    }


def render_html(data):
    tpl = TEMPLATE_PATH.read_text(encoding='utf-8')
    html = tpl.replace('__DATA_JSON__', json.dumps(data, ensure_ascii=False))
    html = html.replace('__FECHA_MIN__', data['fecha_min'])
    html = html.replace('__FECHA_MAX__', data['fecha_max'])
    html = html.replace('__TOTAL__', str(data['total']))
    return html


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    xlsx_path = sys.argv[1]
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path('index.html')

    df = load_and_clean(xlsx_path)
    data = build_data_dict(df)
    html = render_html(data)

    out_path.write_text(html, encoding='utf-8')
    print(f"[generate_report] Reporte generado: {out_path.resolve()}")
    print(f"[generate_report] Total NCRs: {data['total']}  |  Plantas: {data['by_planta']}")


if __name__ == '__main__':
    main()
