from pathlib import Path
from datetime import date
from calendar import monthrange
import math
import re

import pandas as pd
import streamlit as st


# ==================================================
# CONFIGURACIÓN GENERAL
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


def calcular_estado_general(
    promedio,
    stock,
    cobertura
):

    if promedio <= 0:
        return "⚪ SIN VENTAS REGISTRADAS"

    if stock <= 0:
        return "🔴 AGOTADO CON VENTAS"

    if cobertura < 1.5:
        return "🟠 STOCK BAJO"

    if cobertura < 3:
        return "🟡 VIGILAR"

    return "🟢 STOCK SUFICIENTE"


def clasificar_demanda(
    promedio,
    umbral_alta,
    umbral_media
):

    if promedio <= 0:
        return "⚪ SIN MOVIMIENTO"

    if promedio >= umbral_alta:
        return "🔥 ALTA DEMANDA"

    if promedio >= umbral_media:
        return "📦 DEMANDA MEDIA"

    return "🐢 BAJA DEMANDA"


# ==================================================
# CARGA DE DATOS
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
# SOLO STOCK PROPIO
# ==================================================

df_propio = df[
    df["Tipo_Stock"] == "PROPIO"
].copy()


ultimo_mes = (
    df_propio["Mes_Fecha"]
    .max()
)


datos_actuales = df_propio[
    df_propio["Mes_Fecha"] == ultimo_mes
].copy()


# ==================================================
# PARÁMETROS DE LA TIENDA
# ==================================================

st.sidebar.divider()

st.sidebar.subheader(
    "⚙️ Política de reposición"
)


umbral_alta = st.sidebar.number_input(
    "Alta demanda desde (unidades/mes)",
    min_value=1.0,
    value=6.0,
    step=1.0,
)


umbral_media = st.sidebar.number_input(
    "Demanda media desde (unidades/mes)",
    min_value=0.1,
    value=2.0,
    step=0.5,
)


stock_minimo_alta = st.sidebar.number_input(
    "Stock mínimo en alta demanda",
    min_value=0,
    value=2,
    step=1,
)


dias_alerta = st.sidebar.number_input(
    "Avisar si alcanza el mínimo en",
    min_value=1,
    value=7,
    step=1,
)


incluir_mes_siguiente = st.sidebar.checkbox(
    "Si hay presupuesto, cubrir también el próximo mes",
    value=False,
)


# ==================================================
# FECHA DE CORTE
# ==================================================

hoy = date.today()

anio_inventario = int(
    ultimo_mes.year
)

mes_inventario = int(
    ultimo_mes.month
)

dias_mes = monthrange(
    anio_inventario,
    mes_inventario
)[1]


if (
    hoy.year == anio_inventario
    and hoy.month == mes_inventario
):

    dia_default = hoy.day

else:

    dia_default = dias_mes


dia_corte = st.sidebar.number_input(
    "Día hasta el que está actualizado el inventario",
    min_value=1,
    max_value=dias_mes,
    value=dia_default,
    step=1,
)


dias_restantes = max(
    dias_mes - dia_corte,
    0
)


# ==================================================
# RESUMEN GENERAL
# ==================================================

def crear_resumen(
    historico,
    actual
):

    inventario = (
        actual
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


    ventas = (
        historico
        .groupby("Producto_Analitico")
        .agg(
            Ventas_Totales=("Vendido_Mes", "sum"),
            Meses_Con_Datos=("Mes_Fecha", "nunique"),
        )
        .reset_index()
    )


    resumen = inventario.merge(
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


    resumen["Promedio_Mensual"] = (
        resumen["Ventas_Totales"]
        / resumen["Meses_Con_Datos"]
    )


    resumen["Cobertura_Meses"] = (
        resumen["Stock_Actual"]
        / resumen["Promedio_Mensual"]
    )


    resumen.loc[
        resumen["Promedio_Mensual"] <= 0,
        "Cobertura_Meses"
    ] = None


    resumen["Estado"] = resumen.apply(
        lambda fila: calcular_estado_general(
            fila["Promedio_Mensual"],
            fila["Stock_Actual"],
            fila["Cobertura_Meses"],
        ),
        axis=1
    )


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


    with st.expander(
        "ℹ️ ¿Qué significa cada estado?"
    ):

        st.markdown(
            """
**🔴 Agotado con ventas**  
Stock actual igual a cero, pero existen ventas registradas.

**🟠 Stock bajo**  
La cobertura estimada es menor a 1,5 meses.

**🟡 Vigilar**  
La cobertura está aproximadamente entre 1,5 y 3 meses.

**🟢 Stock suficiente**  
La cobertura supera los 3 meses.

**⚪ Sin ventas registradas**  
El producto está en el inventario actual, pero no registró ventas en el periodo analizado.

La decisión final de compra también depende del precio, proveedor, presupuesto y estrategia comercial.
"""
        )


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


    c1, c2, c3, c4, c5 = st.columns(5)


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


    # ----------------------------------------------
    # REQUIEREN ATENCIÓN
    # ----------------------------------------------

    st.subheader(
        "🚨 Productos que requieren atención"
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


    # ----------------------------------------------
    # MAYOR ROTACIÓN
    # ----------------------------------------------

    st.subheader(
        "🔥 Productos con mayor rotación"
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


    # ----------------------------------------------
    # SIN VENTAS
    # ----------------------------------------------

    st.subheader(
        "⚪ Productos sin ventas registradas"
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
    ]


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


    # ----------------------------------------------
    # BUSCADOR
    # ----------------------------------------------

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

            for producto in productos_actuales

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


    producto_elegido = st.selectbox(
        "Selecciona un producto",
        productos_filtrados,
    )


    fila_resumen = resumen[
        resumen["Producto"]
        == producto_elegido
    ].iloc[0]


    clave_producto = (
        fila_resumen[
            "Producto_Analitico"
        ]
    )


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


    categoria = fila_resumen["Categoria"]
    marca = fila_resumen["Marca"]

    st.caption(
        f"{categoria} | {marca}"
    )


    # ==================================================
    # HISTÓRICO MENSUAL
    # ==================================================

    ventas_mensuales = (

        datos_producto

        .groupby("Mes_Fecha")

        .agg(
            Vendido_Mes=(
                "Vendido_Mes",
                "sum"
            )
        )

        .reset_index()

    )


    # Meses completos anteriores al inventario actual.
    meses_completos = ventas_mensuales[
        ventas_mensuales["Mes_Fecha"]
        < ultimo_mes
    ]


    ventas_mes_actual = (
        ventas_mensuales.loc[
            ventas_mensuales["Mes_Fecha"]
            == ultimo_mes,
            "Vendido_Mes"
        ]
        .sum()
    )


    # ==================================================
    # PROMEDIO HISTÓRICO
    # ==================================================

    if len(meses_completos) > 0:

        promedio_historico = (
            meses_completos["Vendido_Mes"]
            .mean()
        )

    else:

        promedio_historico = 0


    # ==================================================
    # PROYECCIÓN DEL MES ACTUAL
    # ==================================================

    if (
        dia_corte > 0
        and ventas_mes_actual > 0
    ):

        proyeccion_mes_actual = (
            ventas_mes_actual
            / dia_corte
            * dias_mes
        )

    else:

        proyeccion_mes_actual = 0


    # Si no existe histórico porque el producto es nuevo,
    # usamos la proyección actual como referencia.
    if promedio_historico > 0:

        demanda_base = promedio_historico

    else:

        demanda_base = proyeccion_mes_actual


    tipo_demanda = clasificar_demanda(
        demanda_base,
        umbral_alta,
        umbral_media
    )


    # ==================================================
    # DATOS ACTUALES
    # ==================================================

    stock_actual = fila_resumen["Stock_Actual"]

    precio_usd = fila_resumen["Precio_USD"]

    tc_actual = fila_resumen["TC"]


    # ==================================================
    # COBERTURA
    # ==================================================

    if demanda_base > 0:

        cobertura = (
            stock_actual
            / demanda_base
        )

    else:

        cobertura = None


    # ==================================================
    # DEMANDA RESTANTE DEL MES
    # ==================================================

    demanda_diaria_historica = (
        demanda_base / dias_mes
        if demanda_base > 0
        else 0
    )


    demanda_restante_base = (
        demanda_diaria_historica
        * dias_restantes
    )


    demanda_restante_ritmo_actual = 0

    if dia_corte > 0:

        ritmo_diario_actual = (
            ventas_mes_actual
            / dia_corte
        )

        demanda_restante_ritmo_actual = (
            ritmo_diario_actual
            * dias_restantes
        )


    # Opcional: cubrir también el próximo mes.
    demanda_objetivo = (
        demanda_restante_base
    )


    if incluir_mes_siguiente:

        demanda_objetivo += (
            demanda_base
        )


    pedido_base = max(
        0,
        math.ceil(
            demanda_objetivo
            - stock_actual
        )
    )


    pedido_ritmo_actual = max(
        0,
        math.ceil(
            demanda_restante_ritmo_actual
            - stock_actual
        )
    )


    # ==================================================
    # DÍAS HASTA STOCK MÍNIMO
    # ==================================================

    dias_hasta_minimo = None


    if (
        demanda_diaria_historica > 0
        and stock_actual > stock_minimo_alta
    ):

        dias_hasta_minimo = (
            stock_actual
            - stock_minimo_alta
        ) / demanda_diaria_historica


    # ==================================================
    # DECISIÓN SEGÚN POLÍTICA
    # ==================================================

    if tipo_demanda == "⚪ SIN MOVIMIENTO":

        accion = (
            "⚪ NO REPONER"
        )


    elif tipo_demanda == "🐢 BAJA DEMANDA":

        accion = (
            "🐢 NO REPONER AUTOMÁTICAMENTE"
        )


    elif tipo_demanda == "🔥 ALTA DEMANDA":

        if stock_actual <= stock_minimo_alta:

            accion = (
                "🔴 HACER PEDIDO"
            )

        elif (
            dias_hasta_minimo is not None
            and dias_hasta_minimo <= dias_alerta
        ):

            accion = (
                "🟠 PREPARAR PEDIDO"
            )

        else:

            accion = (
                "🟡 VIGILAR"
            )


    else:

        if pedido_base > 0:

            accion = (
                "🟠 REVISAR REPOSICIÓN"
            )

        else:

            accion = (
                "🟢 STOCK SUFICIENTE"
            )


    # ==================================================
    # MÉTRICAS PRINCIPALES
    # ==================================================

    c1, c2, c3, c4 = (
        st.columns(4)
    )


    c1.metric(
        "Stock actual",
        f"{stock_actual:.0f}"
    )


    c2.metric(
        "Ventas este mes",
        f"{ventas_mes_actual:.0f}"
    )


    c3.metric(
        "Promedio meses completos",
        f"{promedio_historico:.2f}"
    )


    if cobertura is not None:

        c4.metric(
            "Cobertura aproximada",
            f"{cobertura:.2f} meses"
        )

    else:

        c4.metric(
            "Cobertura aproximada",
            "Sin datos"
        )


    # ==================================================
    # CLASIFICACIÓN
    # ==================================================

    st.subheader(
        "📈 Nivel de demanda"
    )


    st.info(
        tipo_demanda
    )


    st.write(
        f"El análisis utiliza como referencia "
        f"**{demanda_base:.2f} unidades por mes**."
    )


    # ==================================================
    # POLÍTICA REAL DE REPOSICIÓN
    # ==================================================

    st.subheader(
        "🛒 Decisión de reposición"
    )


    if accion.startswith("🔴"):

        st.error(accion)

    elif accion.startswith("🟠"):

        st.warning(accion)

    elif accion.startswith("🟢"):

        st.success(accion)

    else:

        st.info(accion)


    if tipo_demanda == "🔥 ALTA DEMANDA":

        st.write(
            f"Para productos de alta demanda, "
            f"el stock mínimo configurado es de "
            f"**{stock_minimo_alta} unidades**."
        )


        if dias_hasta_minimo is not None:

            st.write(
                f"Al ritmo histórico podría llegar "
                f"a ese mínimo en aproximadamente "
                f"**{dias_hasta_minimo:.1f} días**."
            )


    # ==================================================
    # SUGERENCIA DE CANTIDAD
    # ==================================================

    st.subheader(
        "📦 Cantidad sugerida"
    )


    horizonte_texto = (
        "resto del mes y el próximo mes"
        if incluir_mes_siguiente
        else "resto del mes"
    )


    c5, c6 = st.columns(2)


    c5.metric(
        f"Pedido base para cubrir {horizonte_texto}",
        f"{pedido_base} unidades"
    )


    c6.metric(
        "Escenario si continúa el ritmo actual",
        f"{pedido_ritmo_actual} unidades"
    )


    st.caption(
        "El pedido base utiliza el comportamiento histórico. "
        "El segundo escenario utiliza el ritmo de ventas del mes actual "
        "y puede ser mucho más agresivo. "
        "Debe utilizarse como alerta, no como compra automática."
    )


    # ==================================================
    # PROYECCIÓN DEL MES ACTUAL
    # ==================================================

    st.subheader(
        "📊 Mes actual"
    )


    m1, m2, m3 = st.columns(3)


    m1.metric(
        "Ventas registradas",
        f"{ventas_mes_actual:.0f}"
    )


    m2.metric(
        "Días transcurridos",
        f"{dia_corte}"
    )


    m3.metric(
        "Proyección al cierre",
        f"{proyeccion_mes_actual:.1f}"
    )


    # ==================================================
    # PRECIO
    # ==================================================

    p1, p2 = st.columns(2)


    if pd.notna(precio_usd):

        p1.metric(
            "Precio USD actual",
            f"${precio_usd:.2f}"
        )


    if pd.notna(tc_actual):

        p2.metric(
            "Tipo de cambio",
            f"{tc_actual:.2f}"
        )


    # ==================================================
    # GRÁFICO
    # ==================================================

    st.subheader(
        "📈 Historial de ventas"
    )


    grafico = ventas_mensuales.copy()

    grafico["Mes"] = (
        grafico["Mes_Fecha"]
        .dt.strftime("%Y-%m")
    )


    grafico = (
        grafico[
            [
                "Mes",
                "Vendido_Mes"
            ]
        ]
        .set_index("Mes")
    )


    st.line_chart(
        grafico
    )


    # ==================================================
    # TABLA HISTÓRICA
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