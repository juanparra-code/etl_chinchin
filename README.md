# ETL CHIN CHIN

App Streamlit que combina datos de ventas SIIGO + ODOO POS en `BASE_VENTAS_COMPLETA.xlsx`.

## Estructura

```
etl_chinchin/
├── app.py               ← App Streamlit
├── requirements.txt
├── data/
│   ├── base_power_bi_siigo.xlsx   ← Base histórica SIIGO
│   ├── costos.xlsx                ← Producto (product.template) *.xlsx
│   └── sku_global.xlsx            ← SKU GLOBAL WIILOG *.xlsx (hoja: TOTAL SKU)
└── ETL_CHIN_CHIN.py     ← Script local (alternativa sin Streamlit)
```

## Setup inicial

1. Copia los 3 archivos estáticos a la carpeta `data/` con los nombres exactos:
   - `data/base_power_bi_siigo.xlsx`
   - `data/costos.xlsx` (renombrar desde `Producto (product.template) XX.xlsx`)
   - `data/sku_global.xlsx` (renombrar desde `SKU GLOBAL WIILOG (1).xlsx`)

2. Haz commit y push al repositorio GitHub.

3. En [Streamlit Cloud](https://streamlit.io/cloud):
   - New app → selecciona el repo → `app.py` como archivo principal → Deploy

## Uso

1. Descarga el reporte histórico completo desde ODOO → Auditoría POS → Exportar
2. Sube el archivo `.xlsx` en la app
3. Clic en **Procesar ETL**
4. Descarga `BASE_VENTAS_COMPLETA.xlsx`

## Actualizar archivos estáticos

Cuando cambie la base SIIGO o los costos:
1. Reemplaza el archivo en `data/`
2. Haz commit y push → Streamlit se actualiza automáticamente
