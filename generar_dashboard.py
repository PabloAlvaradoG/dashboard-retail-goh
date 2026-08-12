# -*- coding: utf-8 -*-
#
# Regenera Dashboard_Gerencial.html a partir del Historico_auditorias_retail.xlsx
# real del proyecto. Ejecutar con: python generar_dashboard.py
#
# Si el Excel esta abierto en Excel/OneDrive puede dar PermissionError al
# leerlo; en ese caso cierralo o trabaja sobre una copia y ajusta PATH abajo.
import os
import sys
import unicodedata
import pandas as pd
import json
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(os.path.dirname(BASE_DIR), 'Historico_auditorias_retail.xlsx')


def norm_name(s):
    # Nombres del Excel vienen con mezcla de mayúsculas/acentos entre hojas
    # (ej. 'JOSHUE REMACHE' en Registro_Auditorias vs 'Joshué Remache' en
    # Correos auditores) — normalizar para poder cruzarlos de forma confiable.
    s = str(s).strip().upper()
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'\s+', ' ', s)

# ---------- Registro_Auditorias ----------
reg = pd.read_excel(PATH, sheet_name='Registro_Auditorias', header=3, engine='openpyxl')
reg.columns = [str(c).replace('\n', ' ').strip() for c in reg.columns]
reg = reg[reg['No.'].notna()].copy()

# Guardia de calidad de datos — RIESGO, SCORE y RESULT. A COBRAR COSTO son
# formulas de Excel, no valores fijos. pandas/openpyxl leen el valor
# cacheado de la ultima vez que Excel calculo el archivo, no la formula en
# si. Si un proceso automatizado (ej. openpyxl) guardo el archivo sin pasar
# por Excel, ese cache puede perderse para TODAS las filas (no solo las
# nuevas) y el script generaria un dashboard con RIESGO="MEDIO" y montos en
# $0 en todos lados, sin ningun aviso. Mejor frenar aqui que publicar eso.
CRITICAL_FORMULA_COLS = ['RIESGO', 'SCORE (0/100)', 'RESULT. A COBRAR COSTO']
for _col in CRITICAL_FORMULA_COLS:
    _null_pct = reg[_col].isna().mean() * 100
    if _null_pct > 50:
        print('\n' + '=' * 72)
        print(f"ERROR: la columna '{_col}' viene vacia en {_null_pct:.0f}% de las {len(reg)} filas.")
        print("Esa columna es una formula de Excel (no un valor fijo) y su cache")
        print("de calculo parece haberse perdido -- probablemente porque un proceso")
        print("automatizado guardo el archivo sin pasar por Excel (openpyxl no")
        print("recalcula formulas al guardar).")
        print()
        print("SOLUCION: abre Historico_auditorias_retail.xlsx en Excel, presiona")
        print("Ctrl+Alt+F9 para forzar el recalculo completo, guarda el archivo,")
        print("y vuelve a correr este script.")
        print('=' * 72 + '\n')
        sys.exit(1)

reg['FECHA AUDITORÍA'] = pd.to_datetime(reg['FECHA AUDITORÍA'])
reg['ALMACÉN'] = reg['ALMACÉN'].astype(str).str.strip()
reg['LÍNEA'] = reg['LÍNEA'].astype(str).str.strip()
reg['ZONA'] = reg['ZONA'].astype(str).str.strip()
reg['AUDITOR'] = reg['AUDITOR'].astype(str).str.strip()
reg['ESTADO COBRO'] = reg['ESTADO COBRO'].fillna('No Aplica').astype(str).str.strip()
reg['RIESGO'] = reg['RIESGO'].astype(str).str.strip().str.upper()
RIESGO_FIX = {'CRITICO': 'CRÍTICO', 'CR\ufffdTICO': 'CRÍTICO'}
reg['RIESGO'] = reg['RIESGO'].replace(RIESGO_FIX)
RIESGO_ORDER = ['BAJO', 'MEDIO', 'ALTO', 'CRÍTICO']

for col in ['TOTAL INVENTARIO COSTO', 'RESULT. A COBRAR COSTO', 'VALOR COBRADO',
            'MAL ESTADO CADUCADOS PVP', 'SCORE (0/100)', 'FALTANTES PVP', 'SOBRANTES PVP',
            'DEFECTOS FÁBRICA PVP']:
    reg[col] = pd.to_numeric(reg[col], errors='coerce').fillna(0)

reg['Ajuste (Si/No)'] = reg['Ajuste (Si/No)'].astype(str).str.strip().str.upper()
reg['SOCIEDAD'] = reg['SOCIEDAD'].astype(str).str.strip()
reg['ESTADO REGISTRO'] = reg['ESTADO REGISTRO'].astype(str).str.strip()

# A "No procesado" audit was never reviewed, so it can't genuinely be
# "Pendiente" de cobro -- that reads as a collection backlog when it's
# actually an intake backlog. The Excel's own ESTADO COBRO formula was
# fixed (2026-08-10) to already return "No Aplica" for these, but enforce
# it here too as a defensive default so the dashboard stays correct even if
# that formula's cache is ever lost again (see the RIESGO/SCORE guard
# above -- same failure mode).
reg.loc[reg['ESTADO REGISTRO'] == 'No procesado', 'ESTADO COBRO'] = 'No Aplica'

# Simplified 4-stage collection-process funnel (board-legible), collapsing the
# 8 raw ESTADO REGISTRO values into stages that actually read as a pipeline.
FUNNEL_STAGES = ['Bloqueado / sin procesar', 'Registrado, sin revisar', 'En trámite', 'Cerrado']
FUNNEL_MAP = {
    'No procesado': 0, 'Informe físico no recibido': 0,
    'Registro automático': 1,
    'Datos correctos e informe firmado recibido': 2, 'En proceso de aprobación de ajustes': 2,
    'Procesado - Cobrado': 3, 'Procesado - No Cobrado': 3, 'Pasado a RRHH para cobro': 3,
}

stores = sorted(reg['ALMACÉN'].unique().tolist())
lines = sorted(reg['LÍNEA'].unique().tolist())
zones = sorted(reg['ZONA'].unique().tolist())
auditors = sorted(reg['AUDITOR'].unique().tolist())
estados = sorted(reg['ESTADO COBRO'].unique().tolist())
sociedades = sorted(reg['SOCIEDAD'].unique().tolist())

store_idx = {s: i for i, s in enumerate(stores)}
line_idx = {s: i for i, s in enumerate(lines)}
zone_idx = {s: i for i, s in enumerate(zones)}
auditor_idx = {s: i for i, s in enumerate(auditors)}
estado_idx = {s: i for i, s in enumerate(estados)}
sociedad_idx = {s: i for i, s in enumerate(sociedades)}

audits = []
# Montos con 2 decimales, no int redondeado por fila: redondear cada fila a
# entero antes de sumar 79+ auditorías acumula un desfase de un par de
# dolares frente a un calculo manual en Excel (visto en vivo: dashboard
# mostraba $21,114 vs $21,116.19 calculado a mano para faltante 2026 YTD).
# El redondeo a entero para mostrar ya lo hace fmtUSD() en el HTML, sobre la
# suma final -- no hace falta perder precision aqui.
for _, r in reg.iterrows():
    riesgo = r['RIESGO'] if r['RIESGO'] in RIESGO_ORDER else 'MEDIO'
    audits.append([
        r['FECHA AUDITORÍA'].strftime('%Y-%m-%d'),
        store_idx[r['ALMACÉN']],
        line_idx[r['LÍNEA']],
        zone_idx[r['ZONA']],
        auditor_idx[r['AUDITOR']],
        int(round(r['SCORE (0/100)'])),
        RIESGO_ORDER.index(riesgo),
        round(float(r['TOTAL INVENTARIO COSTO']), 2),
        round(float(r['RESULT. A COBRAR COSTO']), 2),
        round(float(r['VALOR COBRADO']), 2),
        estado_idx[r['ESTADO COBRO']],
        1 if r['Ajuste (Si/No)'] == 'SI' else 0,
        round(float(r['MAL ESTADO CADUCADOS PVP']), 2),
        sociedad_idx[r['SOCIEDAD']],
        FUNNEL_MAP.get(r['ESTADO REGISTRO'], 2),
        round(float(r['FALTANTES PVP']), 2),
        round(float(r['SOBRANTES PVP']), 2),
        round(float(r['DEFECTOS FÁBRICA PVP']), 2),
    ])

# ---------- Tiendas (coverage snapshot, fixed "as of" metric) ----------
tiendas = pd.read_excel(PATH, sheet_name='Tiendas', engine='openpyxl')
tiendas['Estado'] = tiendas['Estado'].astype(str).str.strip().str.upper()
activas = tiendas[tiendas['Estado'] == 'ACTIVA'].copy()
HOY = pd.Timestamp.today().normalize()

ultima_por_tienda = reg.groupby('ALMACÉN')['FECHA AUDITORÍA'].max()
ultima_norm = ultima_por_tienda.copy()
ultima_norm.index = ultima_norm.index.str.strip().str.upper()
activas_nombres = activas['Nombre tienda'].astype(str).str.strip().str.upper()
match = ultima_norm[ultima_norm.index.isin(set(activas_nombres))]
dias = (HOY - match).dt.days
n90 = int((dias <= 90).sum())
n180 = int((dias <= 180).sum())
nActivas = int(len(activas))

coverage = {
    'p90': round(n90 / nActivas * 100, 1),
    'p180': round(n180 / nActivas * 100, 1),
    'n90': n90, 'n180': n180, 'nActivas': nActivas,
}

# Roster completo de tiendas ACTIVAS (según hoja Tiendas), incluyendo las que
# nunca han sido auditadas — esas no aparecen en Registro_Auditorias, así que
# no existen en `stores`/`audits`. Necesario para "pendientes de auditar".
# 4to campo: fecha de apertura si la tienda es reciente (Fecha con Estado
# Activa) — contexto para no leer "nunca auditada" como una omisión.
active_roster = [
    [str(r['Nombre tienda']).strip(), str(r['Línea']).strip(), str(r['Zona']).strip(),
     r['Fecha'].strftime('%Y-%m-%d') if pd.notna(r['Fecha']) else None]
    for _, r in activas.iterrows()
]

# Active/closed flag + closure date per store, matched into the `stores` dict
# used by AUDITS, so the client can exclude closed stores from anything
# forward-looking (priority ranking, "pending audit" views).
#
# IMPORTANT: 'Fecha' in the Tiendas sheet is dual-purpose — for a store
# marked Cerrada it's the closure date, but for a store marked Activa it can
# hold an OPENING date instead (seen live: AIMA MAGDALENA, TCL MALL DEL
# ALTO, TCL MONAY, NAUTICA MALL DEL ALTO all carry a 2026 'Fecha' while
# still Activa). Only treat it as a closure date when the store is actually
# Cerrada — otherwise an active/newly-opened store gets wrongly listed as
# closed.
tiendas['NombreNorm'] = tiendas['Nombre tienda'].astype(str).str.strip().str.upper()
tiendas_map = tiendas.drop_duplicates('NombreNorm').set_index('NombreNorm')
stores_active = []
stores_close_date = []
for s in stores:
    key = s.strip().upper()
    if key in tiendas_map.index:
        row = tiendas_map.loc[key]
        is_active = str(row['Estado']).strip().upper() == 'ACTIVA'
        close_date = row['Fecha'].strftime('%Y-%m-%d') if (pd.notna(row['Fecha']) and not is_active) else None
    else:
        is_active = True  # not in Tiendas sheet -> assume active, don't silently exclude
        close_date = None
    stores_active.append(is_active)
    stores_close_date.append(close_date)

# ---------- Correos auditores (tipo de auditor, para KPI Auditores) ----------
try:
    correos = pd.read_excel(PATH, sheet_name='Correos auditores', engine='openpyxl')
    correos['Auditor responsable'] = correos['Auditor responsable'].astype(str).str.strip()
    correos['Tipo auditor'] = correos['Tipo auditor'].astype(str).str.strip()
    tipo_map_raw = {norm_name(k): v for k, v in zip(correos['Auditor responsable'], correos['Tipo auditor'])}
except Exception:
    tipo_map_raw = {}

tipos_auditor = sorted(set(tipo_map_raw.values())) + ['Sin clasificar']
tipo_idx_map = {t: i for i, t in enumerate(tipos_auditor)}
SIN_CLASIFICAR_IDX = tipo_idx_map['Sin clasificar']
auditor_tipo_idx = [tipo_idx_map.get(tipo_map_raw.get(norm_name(a)), SIN_CLASIFICAR_IDX) for a in auditors]

# ---------- Detalle_SKUs ----------
sku = pd.read_excel(PATH, sheet_name='Detalle_SKUs', header=2, engine='openpyxl')
sku.columns = [str(c).replace('\n', ' ').strip() for c in sku.columns]
sku = sku[sku['MATERIAL (SKU)'].notna()].copy()
sku['FECHA AUDITORÍA'] = pd.to_datetime(sku['FECHA AUDITORÍA'])
sku['COSTO TOTAL'] = pd.to_numeric(sku['COSTO TOTAL'], errors='coerce').fillna(0)

TIPOS = ['FALTANTE', 'SOBRANTE', 'FALTANTE CRUCE', 'SOBRANTE CRUCE',
         'MAL ESTADO', 'CADUCADOS', 'DEFECTO FABRICA', 'OTROS']
TIPO_MAP = {
    'FALTANTES': 'FALTANTE', 'FALTANTE': 'FALTANTE',
    'SOBRANTES': 'SOBRANTE', 'SOBRANTE': 'SOBRANTE',
    'FALTANTES CRUCES': 'FALTANTE CRUCE', 'FALTANTE CRUCE': 'FALTANTE CRUCE',
    'SOBRANTES CRUCES': 'SOBRANTE CRUCE', 'SOBRANTE CRUCE': 'SOBRANTE CRUCE',
    'MAL ESTADO': 'MAL ESTADO',
    'CADUCADOS': 'CADUCADOS',
    'DEFECTO DE F\ufffdBRICA': 'DEFECTO FABRICA', 'DEFECTO DE FABRICA': 'DEFECTO FABRICA',
    'DEFECTO DE F\u00c1BRICA': 'DEFECTO FABRICA',
    'OTROS': 'OTROS',
}


def norm_tipo(v):
    v = str(v).strip().upper()
    v = re.sub(r'\s+', ' ', v)
    return TIPO_MAP.get(v, 'OTROS')


sku['TIPO_N'] = sku['TIPO HALLAZGO'].apply(norm_tipo)

# join to Registro_Auditorias by REF. INFORME to recover a reliable store
reg_map = reg.drop_duplicates('REF. INFORME').set_index('REF. INFORME')['ALMACÉN']
sku['ALMACEN_JOIN'] = sku['REF. INFORME'].map(reg_map)


def resolve_store(row):
    if pd.notna(row['ALMACEN_JOIN']):
        return row['ALMACEN_JOIN']
    own = str(row['ALMACÉN']).strip()
    if own and own not in ('0001', 'nan', 'None') and own in store_idx:
        return own
    return None


sku['ALMACEN_R'] = sku.apply(resolve_store, axis=1)

sku['MATERIAL (SKU)'] = sku['MATERIAL (SKU)'].astype(str).str.strip()
sku['DESCRIPCIÓN'] = sku['DESCRIPCIÓN'].astype(str).str.strip().str.slice(0, 42)
sku['DESC_KEY'] = sku['MATERIAL (SKU)'] + ' \u2014 ' + sku['DESCRIPCIÓN']

sku_desc_list = sorted(sku['DESC_KEY'].unique().tolist())
sku_desc_idx = {s: i for i, s in enumerate(sku_desc_list)}
tipo_idx = {t: i for i, t in enumerate(TIPOS)}

sku_rows = []
for _, r in sku.iterrows():
    st_i = store_idx.get(r['ALMACEN_R'], -1) if r['ALMACEN_R'] else -1
    sku_rows.append([
        r['FECHA AUDITORÍA'].strftime('%Y-%m-%d'),
        tipo_idx[r['TIPO_N']],
        sku_desc_idx[r['DESC_KEY']],
        st_i,
        round(float(r['COSTO TOTAL']), 1),
    ])

payload = {
    'stores': stores, 'lines': lines, 'zones': zones, 'auditors': auditors,
    'estados': estados, 'riesgos': RIESGO_ORDER, 'tipos': TIPOS, 'skuDesc': sku_desc_list,
    'sociedades': sociedades, 'funnelStages': FUNNEL_STAGES,
    'storesActive': stores_active, 'storesCloseDate': stores_close_date,
    'tiposAuditor': tipos_auditor, 'auditorTipoIdx': auditor_tipo_idx,
    'activeRoster': active_roster,
    'audits': audits, 'skus': sku_rows, 'coverage': coverage, 'asOf': HOY.strftime('%Y-%m-%d'),
    'skuRange': [sku['FECHA AUDITORÍA'].min().strftime('%Y-%m-%d'), sku['FECHA AUDITORÍA'].max().strftime('%Y-%m-%d')],
}

data_json = json.dumps(payload, separators=(',', ':'), ensure_ascii=False)
print('audits:', len(audits), 'skus:', len(sku_rows), 'stores:', len(stores), 'skuDesc:', len(sku_desc_list))
print('json size bytes:', len(data_json.encode('utf-8')))
print('coverage:', coverage)
print('riesgo dist:', reg['RIESGO'].value_counts().to_dict())

TPL_PATH = os.path.join(BASE_DIR, 'plantilla_dashboard.html')
OUT_PATH = os.path.join(BASE_DIR, 'Dashboard_Gerencial.html')

with open(TPL_PATH, 'r', encoding='utf-8') as f:
    tpl = f.read()

out = tpl.replace('/*__DATA__*/', 'const DATA = ' + data_json + ';')
with open(OUT_PATH, 'w', encoding='utf-8') as f:
    f.write(out)
print('wrote', OUT_PATH, len(out), 'chars')
