"""
ETL CHIN CHIN — Streamlit App
==============================
Sube el reporte ODOO (y opcionalmente costos / SKU global actualizados)
→ descarga BASE_VENTAS_COMPLETA.xlsx lista para Power BI.

Archivos en repo (data/):
  - base_power_bi_siigo.xlsx  ← base estática SIIGO (se actualiza pocas veces al año)

Archivos que el usuario sube desde la UI:
  - Reporte_POS_CHIN_CHIN (XX).xlsx  ← SIEMPRE (cada actualización)
  - Producto (product.template).xlsx  ← cuando cambie la base de costos
  - SKU GLOBAL WIILOG.xlsx            ← cuando cambie el catálogo de tipos
"""

import streamlit as st
import pandas as pd
import io
import re
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"

# ── Helpers ───────────────────────────────────────────────────────────────────
def limpiar_vendedor(val):
    if not val or (isinstance(val, float) and pd.isna(val)): return None
    n = re.sub(r'^[Aa]tendió:\s*', '', str(val)).strip().upper()
    return n if n and n != 'NINGUNO' else None

def extraer_prefijo(orden):
    if not orden: return None
    s = str(orden); idx = s.rfind('/')
    return s[:idx] if idx > 0 else s

def extraer_consecutivo(orden):
    if not orden: return None
    s = str(orden); idx = s.rfind('/')
    if idx > 0:
        try: return float(s[idx+1:])
        except: return None
    return None

def normalizar_sku(sku):
    if not sku or (isinstance(sku, float) and pd.isna(sku)): return ''
    return re.sub(r'[^A-Za-z0-9\-]', '', str(sku).strip()).upper()

def mapear_impuesto(imp):
    if not imp or (isinstance(imp, float) and pd.isna(imp)): return None
    s = str(imp)
    if '19%' in s: return 'IVA 19%'
    if '0%'  in s: return 'IVA 0%'
    return s

# ── Loaders con cache ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando base SIIGO (archivo estático)…")
def cargar_siigo_repo():
    """Lee el SIIGO desde el repo. Se cachea hasta que cambie el archivo."""
    return pd.read_excel(DATA_DIR / "base_power_bi_siigo.xlsx", sheet_name=0, engine='openpyxl')

def cargar_siigo_upload(f):
    return pd.read_excel(f, sheet_name=0, engine='openpyxl')

def construir_costos(f):
    df = pd.read_excel(f, sheet_name=0, engine='openpyxl')
    df = df.rename(columns={'Referencia interna': 'SKU', 'Nombre': 'Nombre_costo', 'Costo': 'COSTO'})
    df = df[df['SKU'].notna()].drop_duplicates(subset=['SKU'], keep='first')
    costo_por_sku    = dict(zip(df['SKU'], df['COSTO']))
    costo_por_nombre = {str(n).upper().strip(): c for n, c in zip(df['Nombre_costo'], df['COSTO']) if pd.notna(n)}
    costo_por_sku['CHIN8002973-01'] = 26000.0
    costo_por_sku['W2000623-01']    = 29500.0
    for sku, nombre in {
        'S5':'Set aretes','CHIN9552714-01':'HEK AGUA SABOR A PERA 248G',
        'C26':'Consola de videojuegos','W9000726-00':'SET DE ARTE x 145 PIEZAS',
        'R7':'removedor de callos','W6002552-01':'MAQUINA AFEITADORA ELECTRICA RECARGABLE',
        'CHIN9552455-11':'CANTABILE BEBIDA SABOR A UVA 230ML','54':'TABLA PARA DESCONGELAR',
        'W5501139-03':'PANTY LINFÁTICO COLOR: NEGRO - TALLA: XL','B14':'Boxer faja',
        'W6500222-01':'HOT 69 VITAL HOMBRE BEBIDA x 500ml','AT66':'Estuche de Maquillaje',
        'C15':'Cepillo secador para mascotas','B13':'Bolso Onduly',
        'Q1':'Quita manchas quitte','CHIN9552455-10':'CANTABILE BEBIDA SABOR A MELOCOTÓN 230ML',
        'S1':'Sandalia Acupuntura','CHIN9552853-01':'JJLD CARAMELO SABOR ARANDANO: GUMMY OSO',
        'T2':'Tabla para descongelar','B11':'Bolso dualtone','R6':'Reloj vikec',
        'W9500537-01':'SET DE BROCAS x 6 UNIDADES','D5':'Dispositivo antirronquido',
        'S12':'SOPORTE CELULAR','A8':'ARETES','C24':'Collar luna',
    }.items():
        clave = nombre.upper().strip()
        for nom, costo in costo_por_nombre.items():
            if clave[:20] in nom or nom[:20] in clave:
                costo_por_sku[sku] = costo; break
    return costo_por_sku

def construir_tipos(f):
    df = pd.read_excel(f, sheet_name='TOTAL SKU', engine='openpyxl')
    df = df[df['SKU PRODUCTO WIILOG'].notna() & (df['SKU PRODUCTO WIILOG'] != '')]
    df = df.drop_duplicates(subset=['SKU PRODUCTO WIILOG'], keep='first')
    return {normalizar_sku(str(r['SKU PRODUCTO WIILOG'])): r['TIPO PRODUCTO'] for _, r in df.iterrows()}

# ── ETL ───────────────────────────────────────────────────────────────────────
def procesar_etl(df_siigo, odoo_file, costo_por_sku, tipo_por_sku, progress):
    vcols = list(df_siigo.columns) + ['COSTO', 'TIPO PRODUCTO']

    progress.progress(20, "Leyendo reporte ODOO…")
    df_o = pd.read_excel(odoo_file, sheet_name=0, engine='openpyxl')
    df_o = df_o.rename(columns={
        'FECHA_VENT':      'FECHA_VENTA',
        'CANTIDAD_VENDID': 'CANTIDAD_VENDIDA',
        'CANTIDAD_DEVUELT':'CANTIDAD_DEVUELTA',
    })
    df_o = df_o.drop_duplicates()
    df_o['fecha_dt'] = pd.to_datetime(df_o['FECHA_VENTA'], errors='coerce')
    df_o['mes'] = df_o['fecha_dt'].dt.month.astype(float)
    df_o['dia'] = df_o['fecha_dt'].dt.day.astype(float)
    fecha_min = df_o['fecha_dt'].dropna().min()
    fecha_max = df_o['fecha_dt'].dropna().max()

    progress.progress(45, "Procesando filas ODOO…")
    rows_odoo = []
    for _, row in df_o.iterrows():
        orden  = str(row.get('ORDEN') or '')
        es_r   = 'REEMBOLSO' in orden.upper()
        cant   = (row.get('CANTIDAD_VENDIDA') or 0) - (row.get('CANTIDAD_DEVUELTA') or 0)
        ref    = str(row.get('REFERENCIA_INTERNA') or '')
        prod   = str(row.get('PRODUCTO') or '')
        vtp    = float(row.get('VALOR_TOTAL_PRODUCTO') or 0)
        tneto  = float(row.get('TOTAL_NETO_LINEA') or 0)
        ref_n  = normalizar_sku(ref)
        fila = {col: None for col in vcols}
        fila['Tipo transacción']          = 'Nota crédito' if es_r else 'Factura de venta'
        fila['Número comprobante']        = extraer_prefijo(orden)
        fila['Consecutivo']               = extraer_consecutivo(orden)
        fila['Factura de venta completa'] = orden
        fila['Identificación']            = str(row.get('NUMERO_IDENTIFICACION') or '') or None
        fila['Nombre tercero']            = str(row.get('CLIENTE') or '') or None
        fila['Fecha creación']            = row['fecha_dt']
        fila['Fecha elaboración']         = row['fecha_dt']
        fila['mes']                       = row['mes']
        fila['dia']                       = row['dia']
        fila['Nombre contacto']           = str(row.get('CLIENTE') or '') or None
        fila['Correo electrónico']        = str(row.get('CORREO_CLIENTE') or '') or None
        fila['Tipo de registro']          = 'Secuencia'
        fila['Tipo clasificación']        = 'Producto'
        fila['Código']                    = ref or None
        fila['Nombre']                    = f'[{ref}] {prod}' if ref else prod
        fila['Nombre vendedor']           = limpiar_vendedor(row.get('EMPLEADO'))
        fila['Cantidad']                  = cant
        fila['Valor unitario']            = row.get('VALOR_UNITARIO')
        fila['Valor desc.']               = round(vtp - tneto, 2) if vtp > tneto else 0.0
        fila['Base AIU']                  = 0.0
        fila['Impuesto cargo']            = mapear_impuesto(row.get('IMPUESTOS'))
        fila['Valor Impuesto Cargo']      = row.get('VALOR IMPUESTO IVA APLICADO')
        fila['Valor Impuesto Cargo 2']    = 0.0
        fila['Valor Impuesto Retención']  = 0.0
        fila['Total']                     = tneto
        fila['Moneda']                    = 'COP'
        fila['Tasa de cambio']            = 0.0
        fila['Forma pago']                = str(row.get('METODO_PAGO') or '') or None
        fila['ORIGEN']                    = 'ODOO'
        fila['VENDEDOR']                  = limpiar_vendedor(row.get('VENDEDOR'))
        fila['COSTO']                     = costo_por_sku.get(ref_n) or costo_por_sku.get(ref)
        fila['TIPO PRODUCTO']             = tipo_por_sku.get(ref_n)
        rows_odoo.append(fila)

    progress.progress(70, "Combinando SIIGO + ODOO…")
    df_s = df_siigo.copy()
    df_s['COSTO']         = df_s['Código'].apply(lambda x: costo_por_sku.get(normalizar_sku(str(x))) if pd.notna(x) else None)
    df_s['TIPO PRODUCTO'] = df_s['Código'].apply(lambda x: tipo_por_sku.get(normalizar_sku(str(x))) if pd.notna(x) else None)
    df_o_pd  = pd.DataFrame(rows_odoo, columns=vcols)
    df_total = pd.concat([df_s, df_o_pd], ignore_index=True)
    for col in ['Fecha creación','Fecha elaboración','Fecha modificación','Fecha vencimiento']:
        if col in df_total.columns:
            df_total[col] = pd.to_datetime(df_total[col], errors='coerce').dt.tz_localize(None)

    progress.progress(85, "Generando Excel…")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy hh:mm') as writer:
        df_total.to_excel(writer, sheet_name='VENTAS_TOTALES', index=False)
        wb = writer.book; ws = writer.sheets['VENTAS_TOTALES']
        fh  = wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#2E75B6',
                              'border':1,'align':'center','text_wrap':True,'valign':'vcenter'})
        fco = wb.add_format({'num_format':'#,##0.00'})
        col_w = {'Tipo transacción':20,'Número comprobante':16,'Consecutivo':12,'Factura de venta completa':26,'Factura proveedor':16,'Identificación':18,'Sucursal':12,'Nombre tercero':28,'Centro costo':14,'Fecha creación':20,'Fecha modificación':20,'Fecha elaboración':20,'mes':7,'dia':7,'Nombre contacto':28,'Correo electrónico':26,'Tipo de registro':16,'Tipo clasificación':16,'Código':18,'Nombre':55,'Referencia fábrica':16,'Bodega':12,'Identificación Vendedor':18,'Nombre vendedor':22,'Cantidad':10,'Valor unitario':16,'Valor desc.':14,'Base AIU':12,'Impuesto cargo':14,'Valor Impuesto Cargo':18,'Impuesto Cargo 2':14,'Valor Impuesto Cargo 2':18,'Impuesto retención':16,'Valor Impuesto Retención':18,'Base retención (ICA/IVA)':20,'Cargo en totales':16,'Descuento en totales':18,'Total':16,'Moneda':8,'Tasa de cambio':12,'Forma pago':14,'Fecha vencimiento':18,'Observaciones':20,'ORIGEN':10,'VENDEDOR':22,'COSTO':14,'TIPO PRODUCTO':18}
        ws.set_row(0, 32)
        for i, col in enumerate(vcols):
            ws.set_column(i, i, col_w.get(col, 14)); ws.write(0, i, col, fh)
            if col == 'COSTO': ws.set_column(i, i, 14, fco)
        ws.autofilter(0, 0, len(df_total), len(vcols)-1); ws.freeze_panes(1, 0)

    progress.progress(100, "¡Listo!")
    output.seek(0)
    return output, {
        'total': len(df_total), 'siigo': len(df_s), 'odoo': len(df_o_pd),
        'costo_pct': df_total['COSTO'].notna().sum() / len(df_total) * 100,
        'fecha_min': fecha_min, 'fecha_max': fecha_max,
    }

# ── UI ────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="ETL CHIN CHIN", page_icon="🛒", layout="centered")
st.title("🛒 ETL CHIN CHIN")
st.caption("Combina SIIGO + ODOO → BASE_VENTAS_COMPLETA.xlsx")

# ── Sidebar: archivos que cambian ocasionalmente ──────────────────────────────
with st.sidebar:
    st.header("🔄 Actualizar archivos base")
    st.caption("Solo sube cuando haya una versión nueva. Si no subes nada, se usa la versión del repositorio.")

    f_siigo_up = st.file_uploader(
        "base_power_bi_siigo.xlsx",
        type=['xlsx'], key='siigo',
        help="Exportación completa desde SIIGO. Cambia pocas veces al año."
    )
    f_costos = st.file_uploader(
        "Producto (product.template).xlsx",
        type=['xlsx'], key='costos',
        help="Base de costos desde ODOO. Sube cuando haya productos nuevos o cambios de costo."
    )
    f_sku = st.file_uploader(
        "SKU GLOBAL WIILOG.xlsx",
        type=['xlsx'], key='sku',
        help="Catálogo de tipos de producto. Sube cuando haya SKUs nuevos."
    )

    # Indicadores de estado
    st.divider()
    st.markdown("**Estado archivos:**")
    siigo_ok  = f_siigo_up is not None or (DATA_DIR / "base_power_bi_siigo.xlsx").exists()
    costos_ok = f_costos is not None or  (DATA_DIR / "costos.xlsx").exists()
    sku_ok    = f_sku is not None or     (DATA_DIR / "sku_global.xlsx").exists()
    st.markdown(f"{'✅' if siigo_ok  else '❌'} SIIGO {'(nuevo)' if f_siigo_up else '(repo)'}")
    st.markdown(f"{'✅' if costos_ok else '❌'} Costos {'(nuevo)' if f_costos else '(repo)'}")
    st.markdown(f"{'✅' if sku_ok    else '❌'} SKU Global {'(nuevo)' if f_sku else '(repo)'}")

# ── Main: reporte ODOO ────────────────────────────────────────────────────────
st.markdown("### 📤 Reporte ODOO")
f_odoo = st.file_uploader(
    "Reporte_POS_CHIN_CHIN (XX).xlsx — histórico completo",
    type=['xlsx'],
    help="Descárgalo desde ODOO → Auditoría POS → Exportar. Usa siempre el histórico completo."
)

listo = f_odoo is not None and siigo_ok and costos_ok and sku_ok

if not listo:
    faltantes = []
    if not siigo_ok:  faltantes.append("SIIGO")
    if not costos_ok: faltantes.append("Costos")
    if not sku_ok:    faltantes.append("SKU Global")
    if f_odoo is None: faltantes.append("Reporte ODOO")
    st.warning(f"Faltan: **{', '.join(faltantes)}**")

if st.button("▶ Procesar ETL", type="primary", use_container_width=True, disabled=not listo):
    progress = st.progress(0, "Iniciando…")
    try:
        progress.progress(5, "Cargando archivos base…")

        # SIIGO
        df_siigo = cargar_siigo_upload(f_siigo_up) if f_siigo_up else cargar_siigo_repo()

        # Costos
        costos_src = f_costos if f_costos else (DATA_DIR / "costos.xlsx")
        costo_por_sku = construir_costos(costos_src)

        # Tipos
        sku_src = f_sku if f_sku else (DATA_DIR / "sku_global.xlsx")
        tipo_por_sku = construir_tipos(sku_src)

        output, stats = procesar_etl(df_siigo, f_odoo, costo_por_sku, tipo_por_sku, progress)

        st.success("✅ ETL completado")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total filas", f"{stats['total']:,}")
        c2.metric("SIIGO",       f"{stats['siigo']:,}")
        c3.metric("ODOO",        f"{stats['odoo']:,}")
        c4, c5 = st.columns(2)
        c4.metric("Cobertura COSTO", f"{stats['costo_pct']:.1f}%")
        c5.metric("Rango ODOO",
                  f"{stats['fecha_min'].strftime('%d/%m/%y')} → {stats['fecha_max'].strftime('%d/%m/%y')}"
                  if pd.notna(stats['fecha_min']) else "—")

        st.download_button(
            "⬇ Descargar BASE_VENTAS_COMPLETA.xlsx",
            data=output,
            file_name="BASE_VENTAS_COMPLETA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )
    except Exception as e:
        st.error(f"❌ Error: {e}")
        raise
