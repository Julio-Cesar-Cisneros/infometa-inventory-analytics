from pathlib import Path
import re

import pandas as pd
import streamlit as st


# ==================================================
# CONFIGURACIÓN
# ==================================================

DATA_FILE = Path("Base_Limpia_Inventario_InfoMeta.xlsx")
SHEET_NAME = "Historico_Limpio"

st.set_page_config(
    page_title="Info Meta - Análisis de Inventario",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Info Meta - Análisis de Inventario")
st.caption(
    "Análisis de ventas, stock y apoyo a decisiones de reposición."
)


# ==================================================
# FUNCIONES
# ==================================================

def normalizar_producto(nombre):

    if pd.isna(nombre):
        return ""

    nombre = str(nombre).upper().strip()

    nombre = re.sub(
        r"[^A-Z0-9ÁÉÍÓÚÑ]+",
        "",
        nombre
    )

    return nombre


def calcular_estado(promedio, stock, cobertura):

    if promedio <= 0:
        return "⚪ SIN VENTAS REGISTRADAS"

    if stock <= 0:
        return "🔴 AGOTADO CON VENTAS"

    if cobertura < 1.5:
        return "🟠 STOCK BAJO"

    if cobertura < 3:
        return "🟡 VIGILAR"

    return "🟢 STOCK SUFICIENTE"


# ==================================================
# CARGA DEL EXCEL
# ==================================================

if not DATA_FILE.exists():

    st.error(
        "No se encontró Base_Limpia_Inventario_InfoMeta.xlsx "
        "en la carpeta del proyecto."
    )

    st.stop()


df = pd.read_excel(
    DATA_FILE,
    sheet_name=SHEET_NAME
)


# ==================================================
# LIMPIEZA
# ==================================================

columnas_numericas = [
    "Mes_Num",
    "Stock_Inicial",
    "Precio_USD",
    "TC",
    "Vendido_Mes",
    "Stock_Final",
]

for columna in columnas_numericas:

    df[columna] = pd.to_numeric(
        df[columna],
        errors="coerce"
    )


df["Vendido_Mes"] = (
    df["Vendido_Mes"]
    .fillna(0)
)

df["Stock_Final"] = (
    df["Stock_Final"]
    .fillna(0)
)


df["Mes_Fecha"] = pd.to_datetime(
    df["Mes"].astype(str) + "-01",
    errors="coerce"
)


df["Producto_Analitico"] = (
    df["Producto"]
    .apply(normalizar_producto)
)


# ==================================================
# SOLO INVENTARIO PROPIO
# ==================================================

df_propio = df[
    df["Tipo_Stock"] == "PROPIO"
].copy()


# ==================================================
# ÚLTIMO MES DISPONIBLE
# ==================================================

ultimo_mes = (
    df_propio["Mes_Fecha"]
    .max()
)


datos_actuales = df_propio[
    df_propio["Mes_Fecha"] == ultimo_mes
].copy()


# ==================================================
# CREAR RESUMEN
# ==================================================

def crear_resumen(
    datos_historicos,
    datos_mes_actual
):

    # INVENTARIO ACTUAL

    actual = (
        datos_mes_actual
        .groupby("Producto_Analitico")
        .agg(
            Producto=("Producto", "last"),
            Categoria=("Categoria", "last"),
            Marca=("Marca", "last"),
            Stock_Actual=("Stock_Final", "sum"),
            Precio_USD=("Precio_USD", "max"),
            TC=("TC", "max"),
        )
        .reset_index()
    )


    # HISTÓRICO DE VENTAS

    ventas = (
        datos_historicos
        .groupby("Producto_Analitico")
        .agg(
            Ventas_Totales=("Vendido_Mes", "sum"),
            Meses_Con_Datos=("Mes_Fecha", "nunique"),
        )
        .reset_index()
    )


    # UNIÓN

    resumen = actual.merge(
        ventas,
        on="Producto_Analitico",
        how="left"
    )


    resumen["Ventas_Totales"] = (
        resumen["Ventas_Totales"]
        .fillna(0)
    )


    resumen["Meses_Con_Datos"] = (
        resumen["Meses_Con_Datos"]
        .fillna(1)
    )


    # PROMEDIO MENSUAL

    resumen["Promedio_Mensual"] = (
        resumen["Ventas_Totales"]
        / resumen["Meses_Con_Datos"]
    )


    # COBERTURA

    resumen["Cobertura_Meses"] = (
        resumen["Stock_Actual"]
        / resumen["Promedio_Mensual"]
    )


    resumen.loc[
        resumen["Promedio_Mensual"] <= 0,
        "Cobertura_Meses"
    ] = None


    # ESTADO

    resumen["Estado"] = resumen.apply(
        lambda fila: calcular_estado(
            fila["Promedio_Mensual"],
            fila["Stock_Actual"],
            fila["Cobertura_Meses"],
        ),
        axis=1
    )


    # PRIORIDAD

    prioridad = {

        "🔴 AGOTADO CON VENTAS": 1,

        "🟠 STOCK BAJO": 2,

        "🟡 VIGILAR": 3,

        "🟢 STOCK SUFICIENTE": 4,

        "⚪ SIN VENTAS REGISTRADAS": 5,

    }


    resumen["Prioridad"] = (
        resumen["Estado"]
        .map(prioridad)
    )


    return resumen.sort_values(
        [
            "Prioridad",
            "Cobertura_Meses"
        ]
    )


resumen = crear_resumen(
    df_propio,
    datos_actuales
)


# ==================================================
# MENÚ
# ==================================================

vista = st.sidebar.radio(
    "Navegación",
    [
        "📊 Resumen general",
        "🔎 Detalle de producto",
    ],
)


# ==================================================
# RESUMEN GENERAL
# ==================================================

if vista == "📊 Resumen general":

    st.header(
        "📊 Situación general del inventario"
    )

    st.caption(
        "Se analizan únicamente productos propios "
        "que siguen registrados en el último inventario."
    )


    # ==================================================
    # EXPLICACIÓN
    # ==================================================

    with st.expander(
        "ℹ️ ¿Qué significa cada estado?"
    ):

        st.markdown(
            """
**🔴 Agotado con ventas**  
El producto tiene stock 0, pero registró ventas durante el periodo analizado.

**🟠 Stock bajo**  
Todavía queda producto, pero al ritmo promedio actual el stock alcanzaría para menos de 1,5 meses.

**🟡 Vigilar**  
El producto tiene aproximadamente entre 1,5 y 3 meses de cobertura.

**🟢 Stock suficiente**  
El inventario actual cubre más de 3 meses al ritmo promedio registrado.

**⚪ Sin ventas registradas**  
El producto está actualmente en inventario, pero no registró ventas durante los meses disponibles.

---

**Cobertura = Stock actual ÷ Venta promedio mensual**

Estos estados son indicadores de apoyo.  
La decisión final también puede depender del precio, proveedor, mercado, competencia y estrategia de la tienda.
"""
        )


    # ==================================================
    # CONTADORES
    # ==================================================

    agotados = (
        resumen["Estado"]
        == "🔴 AGOTADO CON VENTAS"
    ).sum()


    stock_bajo = (
        resumen["Estado"]
        == "🟠 STOCK BAJO"
    ).sum()


    vigilar = (
        resumen["Estado"]
        == "🟡 VIGILAR"
    ).sum()


    suficiente = (
        resumen["Estado"]
        == "🟢 STOCK SUFICIENTE"
    ).sum()


    sin_ventas = (
        resumen["Estado"]
        == "⚪ SIN VENTAS REGISTRADAS"
    ).sum()


    c1, c2, c3, c4, c5 = (
        st.columns(5)
    )


    c1.metric(
        "🔴 Agotados con ventas",
        agotados
    )


    c2.metric(
        "🟠 Stock bajo",
        stock_bajo
    )


    c3.metric(
        "🟡 Vigilar",
        vigilar
    )


    c4.metric(
        "🟢 Stock suficiente",
        suficiente
    )


    c5.metric(
        "⚪ Sin ventas",
        sin_ventas
    )


    st.divider()


    # ==================================================
    # PRODUCTOS QUE REQUIEREN ATENCIÓN
    # ==================================================

    st.subheader(
        "🚨 Productos que requieren atención"
    )


    st.caption(
        "Aquí aparecen agotados, productos con stock bajo "
        "y productos que conviene vigilar."
    )


    atencion = resumen[
        resumen["Prioridad"] <= 3
    ].copy()


    st.dataframe(

        atencion[
            [
                "Producto",
                "Categoria",
                "Marca",
                "Stock_Actual",
                "Promedio_Mensual",
                "Cobertura_Meses",
                "Estado",
            ]
        ].round(2),

        use_container_width=True,

        hide_index=True,

    )


    st.divider()


    # ==================================================
    # MAYOR ROTACIÓN
    # ==================================================

    st.subheader(
        "🔥 Productos con mayor rotación"
    )


    st.caption(
        "Productos con mayor promedio de unidades vendidas por mes."
    )


    top_rotacion = (

        resumen[
            resumen["Promedio_Mensual"] > 0
        ]

        .sort_values(
            "Promedio_Mensual",
            ascending=False
        )

        .head(15)

    )


    st.dataframe(

        top_rotacion[
            [
                "Producto",
                "Categoria",
                "Stock_Actual",
                "Ventas_Totales",
                "Promedio_Mensual",
                "Cobertura_Meses",
                "Estado",
            ]
        ].round(2),

        use_container_width=True,

        hide_index=True,

    )


    st.divider()


    # ==================================================
    # STOCK SUFICIENTE
    # ==================================================

    st.subheader(
        "🟢 Productos con stock suficiente"
    )


    st.caption(
        "Productos cuya cobertura estimada supera los 3 meses."
    )


    stock_ok = resumen[
        resumen["Estado"]
        == "🟢 STOCK SUFICIENTE"
    ].copy()


    stock_ok = (
        stock_ok
        .sort_values(
            "Cobertura_Meses",
            ascending=False
        )
    )


    st.dataframe(

        stock_ok[
            [
                "Producto",
                "Categoria",
                "Marca",
                "Stock_Actual",
                "Promedio_Mensual",
                "Cobertura_Meses",
            ]
        ].round(2),

        use_container_width=True,

        hide_index=True,

    )


    st.divider()


    # ==================================================
    # SIN VENTAS
    # ==================================================

    st.subheader(
        "⚪ Productos sin ventas registradas"
    )


    st.caption(
        "Estos productos están actualmente en inventario, "
        "pero no registraron ventas en el histórico disponible."
    )


    productos_sin_ventas = resumen[
        (
            resumen["Estado"]
            == "⚪ SIN VENTAS REGISTRADAS"
        )
        &
        (
            resumen["Stock_Actual"] > 0
        )
    ].copy()


    st.dataframe(

        productos_sin_ventas[
            [
                "Producto",
                "Categoria",
                "Marca",
                "Stock_Actual",
                "Precio_USD",
            ]
        ],

        use_container_width=True,

        hide_index=True,

    )


# ==================================================
# DETALLE DE PRODUCTO
# ==================================================

else:

    st.header(
        "🔎 Análisis por producto"
    )


    # ==================================================
    # BUSCADOR
    # ==================================================

    texto_busqueda = st.text_input(
        "🔍 Buscar producto",
        placeholder=(
            "Ejemplo: RTX 5060, B650, "
            "5600GT, SSD 1TB..."
        )
    )


    productos_actuales = (
        resumen["Producto"]
        .dropna()
        .sort_values()
        .unique()
        .tolist()
    )


    if texto_busqueda:

        productos_filtrados = [

            producto

            for producto
            in productos_actuales

            if texto_busqueda.lower()
            in producto.lower()

        ]

    else:

        productos_filtrados = (
            productos_actuales
        )


    if len(productos_filtrados) == 0:

        st.warning(
            "No se encontraron productos con ese texto."
        )

        st.stop()


    # ==================================================
    # SELECTOR
    # ==================================================

    producto_elegido = st.selectbox(
        "Selecciona un producto",
        productos_filtrados,
    )


    # ==================================================
    # RESUMEN DEL PRODUCTO
    # ==================================================

    fila_resumen = resumen[
        resumen["Producto"]
        == producto_elegido
    ].iloc[0]


    clave_producto = (
        fila_resumen[
            "Producto_Analitico"
        ]
    )


    # ==================================================
    # HISTORIAL
    # ==================================================

    datos_producto = df_propio[
        df_propio["Producto_Analitico"]
        == clave_producto
    ].copy()


    datos_producto = (
        datos_producto
        .sort_values("Mes_Fecha")
    )


    st.subheader(
        producto_elegido
    )


    categoria = (
        fila_resumen["Categoria"]
    )


    marca = (
        fila_resumen["Marca"]
    )


    st.caption(
        f"{categoria} | {marca}"
    )


    ventas_totales = (
        fila_resumen["Ventas_Totales"]
    )


    promedio_mensual = (
        fila_resumen["Promedio_Mensual"]
    )


    stock_actual = (
        fila_resumen["Stock_Actual"]
    )


    cobertura = (
        fila_resumen["Cobertura_Meses"]
    )


    precio_usd = (
        fila_resumen["Precio_USD"]
    )


    tc_actual = (
        fila_resumen["TC"]
    )


    recomendacion = (
        fila_resumen["Estado"]
    )


    # ==================================================
    # MÉTRICAS
    # ==================================================

    col1, col2, col3, col4 = (
        st.columns(4)
    )


    col1.metric(
        "Stock actual",
        f"{stock_actual:.0f}"
    )


    col2.metric(
        "Ventas registradas",
        f"{ventas_totales:.0f}"
    )


    col3.metric(
        "Venta promedio mensual",
        f"{promedio_mensual:.2f}"
    )


    if pd.notna(cobertura):

        col4.metric(
            "Cobertura estimada",
            f"{cobertura:.2f} meses"
        )

    else:

        col4.metric(
            "Cobertura estimada",
            "Sin datos"
        )


    # ==================================================
    # RECOMENDACIÓN
    # ==================================================

    st.subheader(
        "📊 Estado del inventario"
    )


    if recomendacion.startswith("🔴"):

        st.error(
            recomendacion
        )

    elif recomendacion.startswith("🟠"):

        st.warning(
            recomendacion
        )

    elif recomendacion.startswith("🟡"):

        st.warning(
            recomendacion
        )

    elif recomendacion.startswith("🟢"):

        st.success(
            recomendacion
        )

    else:

        st.info(
            recomendacion
        )


    # ==================================================
    # EXPLICACIÓN INDIVIDUAL
    # ==================================================

    if promedio_mensual > 0:

        st.write(
            f"Este producto vende aproximadamente "
            f"**{promedio_mensual:.2f} unidades por mes** "
            f"y actualmente quedan **{stock_actual:.0f} unidades**."
        )


        if pd.notna(cobertura):

            st.write(
                f"Al ritmo registrado, el stock podría cubrir "
                f"aproximadamente **{cobertura:.2f} meses**."
            )

    else:

        st.write(
            "No se registraron ventas de este producto "
            "durante el periodo disponible."
        )


    # ==================================================
    # PRECIO
    # ==================================================

    col5, col6 = st.columns(2)


    if pd.notna(precio_usd):

        col5.metric(
            "Precio USD actual",
            f"${precio_usd:.2f}"
        )

    else:

        col5.metric(
            "Precio USD actual",
            "Sin dato"
        )


    if pd.notna(tc_actual):

        col6.metric(
            "Tipo de cambio",
            f"{tc_actual:.2f}"
        )

    else:

        col6.metric(
            "Tipo de cambio",
            "Sin dato"
        )


    # ==================================================
    # HISTORIAL DE VENTAS
    # ==================================================

    st.subheader(
        "📈 Historial de ventas"
    )


    ventas_mes = (

        datos_producto

        .groupby("Mes")[
            "Vendido_Mes"
        ]

        .sum()

        .reset_index()

    )


    grafico = (
        ventas_mes
        .set_index("Mes")
    )


    st.line_chart(
        grafico
    )


    # ==================================================
    # HISTORIAL COMPLETO
    # ==================================================

    st.subheader(
        "📋 Historial"
    )


    historial = (

        datos_producto

        .groupby("Mes")

        .agg(

            Producto=("Producto", "last"),

            Marca=("Marca", "last"),

            Stock_Inicial=(
                "Stock_Inicial",
                "sum"
            ),

            Vendido_Mes=(
                "Vendido_Mes",
                "sum"
            ),

            Stock_Final=(
                "Stock_Final",
                "sum"
            ),

            Precio_USD=(
                "Precio_USD",
                "max"
            ),

            TC=("TC", "max"),

        )

        .reset_index()

    )


    st.dataframe(
        historial,
        use_container_width=True,
        hide_index=True,
    )