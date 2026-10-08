from pathlib import Path

import pandas as pd
import streamlit as st


DATA_FILE = Path("Base_Limpia_Inventario_InfoMeta.xlsx")
SHEET_NAME = "Historico_Limpio"


st.set_page_config(
    page_title="Panel de Reposición - Info Meta",
    layout="wide",
)

st.title("📦 Panel de Reposición - Info Meta")

if not DATA_FILE.exists():
    st.error(
        "No se encontró el archivo local de datos. "
        "Coloca Base_Limpia_Inventario_InfoMeta.xlsx en la carpeta del proyecto."
    )
    st.stop()

df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)

st.write("Registros cargados:", len(df))

productos = sorted(df["Producto"].dropna().unique())

producto_elegido = st.selectbox(
    "Selecciona un producto",
    productos,
)

datos_producto = df[df["Producto"] == producto_elegido].copy()
datos_producto = datos_producto.sort_values("Mes_Num")

st.subheader(producto_elegido)

ventas_totales = datos_producto["Vendido_Mes"].fillna(0).sum()
promedio_mensual = datos_producto["Vendido_Mes"].fillna(0).mean()

fila_actual = datos_producto.iloc[-1]

stock_actual = fila_actual["Stock_Final"]
precio_usd = fila_actual["Precio_USD"]
tc_actual = fila_actual["TC"]

if promedio_mensual > 0:
    cobertura = stock_actual / promedio_mensual
else:
    cobertura = None

if promedio_mensual == 0:
    recomendacion = "SIN MOVIMIENTO"
elif stock_actual == 0:
    recomendacion = "🔴 REPONER URGENTE"
elif cobertura < 1.5:
    recomendacion = "🔴 REPONER"
elif cobertura < 3:
    recomendacion = "🟡 VIGILAR"
else:
    recomendacion = "🟢 STOCK SUFICIENTE"

col1, col2, col3, col4 = st.columns(4)

col1.metric("Stock actual", f"{stock_actual:.0f}")
col2.metric("Vendidos últimos meses", f"{ventas_totales:.0f}")
col3.metric("Promedio mensual", f"{promedio_mensual:.2f}")

if cobertura is not None:
    col4.metric("Cobertura estimada", f"{cobertura:.2f} meses")
else:
    col4.metric("Cobertura estimada", "Sin datos")

st.subheader("📊 Recomendación")
st.info(recomendacion)

col5, col6 = st.columns(2)

if pd.notna(precio_usd):
    col5.metric("Precio USD actual", "$" + f"{precio_usd:.2f}")
else:
    col5.metric("Precio USD actual", "Sin dato")

if pd.notna(tc_actual):
    col6.metric("Tipo de cambio", f"{tc_actual:.2f}")
else:
    col6.metric("Tipo de cambio", "Sin dato")

st.subheader("📈 Historial de ventas")

grafico = datos_producto[["Mes", "Vendido_Mes"]].set_index("Mes")
st.line_chart(grafico)

st.subheader("📋 Historial completo")

st.dataframe(
    datos_producto[
        [
            "Mes",
            "Producto",
            "Marca",
            "Stock_Inicial",
            "Vendido_Mes",
            "Stock_Final",
            "Precio_USD",
            "TC",
        ]
    ],
    use_container_width=True,
)
