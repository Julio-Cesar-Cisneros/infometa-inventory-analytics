from pathlib import Path
from datetime import datetime, date
from calendar import monthrange
import math
import re

import altair as alt
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
    "Análisis de inventario, ventas, calidad de datos "
    "y apoyo a decisiones de reposición."
)


# ==================================================
# FUNCIONES GENERALES
# ==================================================

def normalizar_producto(nombre):

    if pd.isna(nombre):
        return ""

    nombre = str(nombre).upper().strip()

    return re.sub(
        r"[^A-Z0-9ÁÉÍÓÚÑ]+",
        "",
        nombre
    )


def formato_mes(valor):

    return pd.Timestamp(valor).strftime("%Y-%m")


def formato_bs(valor):

    return (
        f"{valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# ==================================================
# NOMBRES AMIGABLES PARA EL CLIENTE
# ==================================================

NOMBRES_COLUMNAS_CLIENTE = {
    "Categoria": "Categoría",
    "Stock_Actual": "Stock actual",
    "Ventas_Totales": "Ventas registradas",
    "Promedio_Mensual": "Promedio mensual",
    "Cobertura_Meses": "Cobertura estimada (meses)",
    "Precio_USD": "Costo unitario USD",
    "Valor_USD": "Valor del inventario USD",
    "Valor_Stock_USD": "Valor acumulado USD",
    "Valor_Stock_Bs": "Valor acumulado Bs",
    "Stock_Inicial": "Stock inicial",
    "Vendido_Mes": "Unidades vendidas",
    "Vendido": "Unidades vendidas",
    "Stock_Final": "Stock final",
    "TC": "Tipo de cambio",
    "Meses_Con_Datos": "Meses analizados",
    "Unidades": "Unidades",
    "Estado": "Estado",
}


def mostrar_dataframe(datos, **kwargs):

    datos_cliente = datos.rename(
        columns=NOMBRES_COLUMNAS_CLIENTE
    )

    st.dataframe(
        datos_cliente,
        **kwargs
    )


def calcular_estado_general(
    promedio,
    stock,
    cobertura
):

    if promedio <= 0:
        return "⚪ SIN VENTAS REGISTRADAS"

    if stock <= 0:
        return "🔴 AGOTADO CON VENTAS"

    if pd.isna(cobertura):
        return "⚪ SIN DATOS"

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
# CREAR RESUMEN
# ==================================================

def crear_resumen(
    historico,
    inventario_actual
):

    actual = (
        inventario_actual
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

    resumen = actual.merge(
        ventas,
        on="Producto_Analitico",
        how="left"
    )

    resumen["Ventas_Totales"] = (
        resumen["Ventas_Totales"].fillna(0)
    )

    resumen["Meses_Con_Datos"] = (
        resumen["Meses_Con_Datos"].fillna(0)
    )

    resumen["Promedio_Mensual"] = 0.0

    tiene_historial = (
        resumen["Meses_Con_Datos"] > 0
    )

    resumen.loc[
        tiene_historial,
        "Promedio_Mensual"
    ] = (
        resumen.loc[
            tiene_historial,
            "Ventas_Totales"
        ]
        /
        resumen.loc[
            tiene_historial,
            "Meses_Con_Datos"
        ]
    )

    resumen["Cobertura_Meses"] = float("nan")

    con_demanda = (
        resumen["Promedio_Mensual"] > 0
    )

    resumen.loc[
        con_demanda,
        "Cobertura_Meses"
    ] = (
        resumen.loc[
            con_demanda,
            "Stock_Actual"
        ]
        /
        resumen.loc[
            con_demanda,
            "Promedio_Mensual"
        ]
    )

    resumen["Estado"] = resumen.apply(
        lambda fila:
        calcular_estado_general(
            fila["Promedio_Mensual"],
            fila["Stock_Actual"],
            fila["Cobertura_Meses"],
        ),
        axis=1,
    )

    prioridad = {
        "🔴 AGOTADO CON VENTAS": 1,
        "🟠 STOCK BAJO": 2,
        "🟡 VIGILAR": 3,
        "🟢 STOCK SUFICIENTE": 4,
        "⚪ SIN VENTAS REGISTRADAS": 5,
        "⚪ SIN DATOS": 6,
    }

    resumen["Prioridad"] = (
        resumen["Estado"].map(prioridad)
    )

    return resumen.sort_values(
        [
            "Prioridad",
            "Cobertura_Meses"
        ],
        na_position="last"
    )


# ==================================================
# ANALIZAR POSIBLE ARRASTRE
# ==================================================

def analizar_arrastre(
    datos,
    mes_anterior,
    mes_actual
):

    anterior = (
        datos[
            datos["Mes_Fecha"] == mes_anterior
        ]
        .groupby("Producto_Analitico")
        .agg(
            Producto_Anterior=("Producto", "last"),
            Stock_Inicial_Anterior=("Stock_Inicial", "sum"),
            Vendido_Anterior=("Vendido_Mes", "sum"),
            Stock_Final_Anterior=("Stock_Final", "sum"),
        )
        .reset_index()
    )

    actual = (
        datos[
            datos["Mes_Fecha"] == mes_actual
        ]
        .groupby("Producto_Analitico")
        .agg(
            Producto_Actual=("Producto", "last"),
            Stock_Inicial_Actual=("Stock_Inicial", "sum"),
            Vendido_Actual=("Vendido_Mes", "sum"),
            Stock_Final_Actual=("Stock_Final", "sum"),
        )
        .reset_index()
    )

    comparacion = anterior.merge(
        actual,
        on="Producto_Analitico",
        how="inner"
    )

    if comparacion.empty:
        return comparacion, 0.0, 0.0

    mismo_vendido = (
        comparacion[
            "Vendido_Anterior"
        ].fillna(-999999)
        ==
        comparacion[
            "Vendido_Actual"
        ].fillna(-999999)
    )

    mismo_registro = (
        (
            comparacion[
                "Stock_Inicial_Anterior"
            ].fillna(-999999)
            ==
            comparacion[
                "Stock_Inicial_Actual"
            ].fillna(-999999)
        )
        &
        mismo_vendido
        &
        (
            comparacion[
                "Stock_Final_Anterior"
            ].fillna(-999999)
            ==
            comparacion[
                "Stock_Final_Actual"
            ].fillna(-999999)
        )
    )

    comparacion["Mismo_Vendido"] = mismo_vendido
    comparacion["Mismo_Registro"] = mismo_registro

    porcentaje_vendido = (
        mismo_vendido.mean() * 100
    )

    porcentaje_registro = (
        mismo_registro.mean() * 100
    )

    return (
        comparacion,
        porcentaje_vendido,
        porcentaje_registro
    )


# ==================================================
# CARGAR EXCEL
# ==================================================

if not DATA_FILE.exists():

    st.error(
        "No se encontró "
        "Base_Limpia_Inventario_InfoMeta.xlsx "
        "en la carpeta del proyecto."
    )

    st.stop()


try:

    df = pd.read_excel(
        DATA_FILE,
        sheet_name=SHEET_NAME
    )

except Exception as exc:

    st.error(
        f"No se pudo leer la hoja {SHEET_NAME}. "
        f"Error: {exc}"
    )

    st.stop()


# ==================================================
# FECHA DEL ARCHIVO
# ==================================================

try:

    ultima_modificacion = datetime.fromtimestamp(
        DATA_FILE.stat().st_mtime
    )

    st.caption(
        "🕒 Base de inventario actualizada: "
        f"{ultima_modificacion.strftime('%d/%m/%Y %H:%M')}"
    )

except Exception:
    pass


# ==================================================
# VALIDAR COLUMNAS
# ==================================================

columnas_requeridas = [
    "Mes",
    "Tipo_Stock",
    "Producto",
    "Categoria",
    "Marca",
    "Stock_Inicial",
    "Vendido_Mes",
    "Stock_Final",
    "Precio_USD",
    "TC",
]


faltantes = [
    columna
    for columna in columnas_requeridas
    if columna not in df.columns
]


if faltantes:

    st.error(
        "Faltan columnas requeridas en la base: "
        + ", ".join(faltantes)
    )

    st.stop()


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

    if columna in df.columns:

        df[columna] = pd.to_numeric(
            df[columna],
            errors="coerce"
        )


df["Vendido_Mes"] = (
    df["Vendido_Mes"].fillna(0)
)

df["Stock_Final"] = (
    df["Stock_Final"].fillna(0)
)

df["Mes_Fecha"] = pd.to_datetime(
    df["Mes"].astype(str)
    + "-01",
    errors="coerce"
)

df["Producto_Analitico"] = (
    df["Producto"].apply(
        normalizar_producto
    )
)


# ==================================================
# SOLO INVENTARIO PROPIO
# ==================================================

# Los productos de terceros quedan completamente fuera.
# INFO META e INFO PC se conservan cuando Tipo_Stock = PROPIO.

df_propio = (
    df[
        df["Tipo_Stock"] == "PROPIO"
    ]
    .copy()
)


meses_disponibles_global = sorted(
    df_propio[
        "Mes_Fecha"
    ]
    .dropna()
    .unique()
)


if not meses_disponibles_global:

    st.error(
        "No se encontraron meses válidos."
    )

    st.stop()


ultimo_mes = pd.Timestamp(
    meses_disponibles_global[-1]
)


if len(meses_disponibles_global) >= 2:

    mes_anterior = pd.Timestamp(
        meses_disponibles_global[-2]
    )

else:

    mes_anterior = None


# ==================================================
# HISTÓRICO VS STOCK ACTUAL
# ==================================================

# El último mes:
# ✅ se usa para stock actual y análisis financiero.
# ⚠️ no entra automáticamente en ventas históricas.

df_historico = (
    df_propio[
        df_propio["Mes_Fecha"] < ultimo_mes
    ]
    .copy()
)


datos_actuales = (
    df_propio[
        df_propio["Mes_Fecha"] == ultimo_mes
    ]
    .copy()
)


resumen = crear_resumen(
    df_historico,
    datos_actuales
)


# ==================================================
# CALIDAD DE DATOS
# ==================================================

if mes_anterior is not None:

    (
        comparacion_calidad,
        pct_mismo_vendido,
        pct_mismo_registro

    ) = analizar_arrastre(
        df_propio,
        mes_anterior,
        ultimo_mes
    )

else:

    comparacion_calidad = pd.DataFrame()

    pct_mismo_vendido = 0.0
    pct_mismo_registro = 0.0


# ==================================================
# SIDEBAR
# ==================================================

st.sidebar.caption(
    f"📦 Stock actual: {formato_mes(ultimo_mes)}"
)

st.sidebar.caption(
    f"📈 Ventas históricas por defecto: hasta "
    f"{formato_mes(mes_anterior) if mes_anterior else '-'}"
)

st.sidebar.caption(
    f"⚠️ {formato_mes(ultimo_mes)}: "
    f"mes provisional para ventas"
)


vista = st.sidebar.radio(
    "Navegación",
    [
        "📊 Resumen general",
        "💰 Análisis financiero",
        "📈 Consulta histórica",
        "🔎 Detalle de producto",
        "🧪 Calidad de datos",
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
        f"El stock actual se toma de "
        f"{formato_mes(ultimo_mes)}. "
        f"Las métricas de ventas excluyen ese mes "
        f"por defecto."
    )

    st.warning(
        f"⚠️ {formato_mes(ultimo_mes)} "
        f"se considera provisional para ventas. "
        f"Sus valores de VENDIDO no entran "
        f"automáticamente en los promedios históricos."
    )


    with st.expander(
        "ℹ️ ¿Qué significa cada estado?"
    ):

        st.markdown(
            """
**🔴 Agotado con ventas**  
Actualmente no quedan unidades, pero el producto tiene movimiento histórico.

**🟠 Stock bajo**  
El stock disponible cubriría menos de aproximadamente 1,5 meses.

**🟡 Vigilar**  
El stock cubriría aproximadamente entre 1,5 y 3 meses.

**🟢 Stock suficiente**  
El inventario disponible supera aproximadamente los 3 meses de cobertura.

**⚪ Sin ventas registradas**  
El producto tiene stock actual, pero no registra movimiento en el histórico disponible.
"""
        )


    agotados = (
        resumen["Estado"]
        ==
        "🔴 AGOTADO CON VENTAS"
    ).sum()

    stock_bajo = (
        resumen["Estado"]
        ==
        "🟠 STOCK BAJO"
    ).sum()

    vigilar = (
        resumen["Estado"]
        ==
        "🟡 VIGILAR"
    ).sum()

    suficiente = (
        resumen["Estado"]
        ==
        "🟢 STOCK SUFICIENTE"
    ).sum()

    sin_ventas = (
        resumen["Estado"]
        ==
        "⚪ SIN VENTAS REGISTRADAS"
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


    st.subheader(
        "🚨 Productos que requieren atención"
    )

    atencion = (
        resumen[
            resumen["Prioridad"] <= 3
        ]
        .copy()
    )

    mostrar_dataframe(
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


    st.subheader(
        "🔥 Productos con mayor rotación histórica"
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

    mostrar_dataframe(
        top_rotacion[
            [
                "Producto",
                "Categoria",
                "Marca",
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


    st.subheader(
        "⚪ Productos sin ventas registradas"
    )

    productos_sin_ventas = (
        resumen[
            (
                resumen["Estado"]
                ==
                "⚪ SIN VENTAS REGISTRADAS"
            )
            &
            (
                resumen["Stock_Actual"] > 0
            )
        ]
    )

    mostrar_dataframe(
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
# ANÁLISIS FINANCIERO
# ==================================================

elif vista == "💰 Análisis financiero":

    st.header(
        "💰 Análisis financiero del stock actual"
    )

    st.caption(
        f"Este módulo utiliza el inventario más reciente "
        f"cargado: {formato_mes(ultimo_mes)}."
    )

    st.info(
        "Este módulo no depende de las ventas históricas. "
        "Calcula el costo del inventario usando "
        "Stock actual × Costo USD registrado."
    )

    st.success(
        "✅ Precio_USD corresponde al costo del producto. "
        "Los cálculos representan el costo del stock "
        "actual registrado en el inventario."
    )


    inventario_financiero = (
        datos_actuales
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


    inventario_financiero["Precio_USD"] = (
        inventario_financiero[
            "Precio_USD"
        ].fillna(0)
    )

    inventario_financiero["Stock_Actual"] = (
        inventario_financiero[
            "Stock_Actual"
        ].fillna(0)
    )

    inventario_financiero["Valor_Stock_USD"] = (
        inventario_financiero["Stock_Actual"]
        *
        inventario_financiero["Precio_USD"]
    )


    tc_disponibles = (
        datos_actuales[
            "TC"
        ]
        .dropna()
    )

    tc_disponibles = (
        tc_disponibles[
            tc_disponibles > 0
        ]
    )


    if not tc_disponibles.empty:

        tc_default = float(
            tc_disponibles.median()
        )

    else:

        tc_default = 12.0


    st.subheader(
        "⚙️ Parámetros de valorización"
    )

    tc_valor = st.number_input(
        "Tipo de cambio para convertir "
        "el costo del inventario a Bs",
        min_value=0.01,
        value=tc_default,
        step=0.05,
    )


    inventario_financiero["Valor_Stock_Bs"] = (
        inventario_financiero["Valor_Stock_USD"]
        *
        tc_valor
    )


    stock_positivo = (
        inventario_financiero[
            inventario_financiero["Stock_Actual"] > 0
        ]
        .copy()
    )


    unidades_stock = (
        stock_positivo[
            "Stock_Actual"
        ].sum()
    )

    productos_stock = (
        stock_positivo[
            "Producto_Analitico"
        ].nunique()
    )

    valor_total_usd = (
        stock_positivo[
            "Valor_Stock_USD"
        ].sum()
    )

    valor_total_bs = (
        stock_positivo[
            "Valor_Stock_Bs"
        ].sum()
    )


    # ==================================================
    # MÉTRICAS FINANCIERAS
    # ==================================================

    m1, m2 = st.columns(2)

    m1.metric(
        "📦 Unidades disponibles",
        f"{unidades_stock:,.0f}"
    )

    m2.metric(
        "🧾 Productos con stock",
        f"{productos_stock:,.0f}"
    )


    m3, m4 = st.columns(2)

    m3.metric(
        "💵 Costo del inventario USD",
        f"${valor_total_usd:,.2f}"
    )

    m4.metric(
        "🇧🇴 Costo equivalente en Bs",
        f"Bs {formato_bs(valor_total_bs)}"
    )


    st.divider()


    # ==================================================
    # VALOR POR CATEGORÍA
    # ==================================================

    st.subheader(
        "📦 Categorías donde hay más valor concentrado"
    )

    st.caption(
        "Muestra cuánto costo de inventario está "
        "concentrado en cada categoría."
    )


    valor_categoria = (
        stock_positivo
        .groupby("Categoria")
        .agg(
            Unidades=("Stock_Actual", "sum"),
            Valor_USD=("Valor_Stock_USD", "sum"),
        )
        .reset_index()
        .sort_values(
            "Valor_USD",
            ascending=False
        )
    )


    chart_categoria = (
        alt.Chart(
            valor_categoria
        )
        .mark_bar()
        .encode(
            y=alt.Y(
                "Categoria:N",
                sort="-x",
                title=None,
                axis=alt.Axis(
                    labelLimit=500
                )
            ),
            x=alt.X(
                "Valor_USD:Q",
                title="Costo del stock USD"
            ),
            tooltip=[
                alt.Tooltip(
                    "Categoria:N",
                    title="Categoría"
                ),
                alt.Tooltip(
                    "Unidades:Q",
                    title="Unidades",
                    format=".0f"
                ),
                alt.Tooltip(
                    "Valor_USD:Q",
                    title="Valor del inventario USD",
                    format=",.2f"
                ),
            ]
        )
        .properties(
            height=550
        )
    )

    st.altair_chart(
        chart_categoria,
        use_container_width=True
    )


    mostrar_dataframe(
        valor_categoria.round(2),
        use_container_width=True,
        hide_index=True,
    )


    st.divider()


    # ==================================================
    # VALOR POR MARCA
    # ==================================================

    st.subheader(
        "🏷️ Valor del inventario por marca"
    )


    valor_marca = (
        stock_positivo
        .groupby("Marca")
        .agg(
            Unidades=("Stock_Actual", "sum"),
            Valor_USD=("Valor_Stock_USD", "sum"),
        )
        .reset_index()
        .sort_values(
            "Valor_USD",
            ascending=False
        )
    )


    if not valor_marca.empty:

        max_marcas = min(
            30,
            len(valor_marca)
        )

        cantidad_marcas = st.slider(
            "Cantidad de marcas a mostrar",
            min_value=1,
            max_value=max_marcas,
            value=min(
                15,
                max_marcas
            ),
        )

        top_marcas = (
            valor_marca
            .head(cantidad_marcas)
        )

        chart_marca = (
            alt.Chart(
                top_marcas
            )
            .mark_bar()
            .encode(
                y=alt.Y(
                    "Marca:N",
                    sort="-x",
                    title=None,
                    axis=alt.Axis(
                        labelLimit=400
                    )
                ),
                x=alt.X(
                    "Valor_USD:Q",
                    title="Costo del stock USD"
                ),
                tooltip=[
                    alt.Tooltip(
                        "Marca:N",
                        title="Marca"
                    ),
                    alt.Tooltip(
                        "Unidades:Q",
                        title="Unidades",
                        format=".0f"
                    ),
                    alt.Tooltip(
                        "Valor_USD:Q",
                        title="Valor del inventario USD",
                        format=",.2f"
                    ),
                ]
            )
            .properties(
                height=max(
                    400,
                    cantidad_marcas * 35
                )
            )
        )

        st.altair_chart(
            chart_marca,
            use_container_width=True
        )


    st.divider()


    # ==================================================
    # PRODUCTOS CAROS
    # ==================================================

    st.subheader(
        "⚠️ Productos caros con muchas unidades"
    )

    st.caption(
        "Permite detectar productos de costo elevado "
        "donde hay varias unidades disponibles."
    )


    p1, p2 = st.columns(2)

    precio_minimo_caro = (
        p1.number_input(
            "Considerar producto caro desde USD",
            min_value=0.0,
            value=200.0,
            step=25.0,
        )
    )

    stock_minimo_alto = (
        p2.number_input(
            "Considerar muchas unidades desde",
            min_value=1,
            value=3,
            step=1,
        )
    )


    productos_caros = (
        inventario_financiero[
            (
                inventario_financiero[
                    "Precio_USD"
                ] >= precio_minimo_caro
            )
            &
            (
                inventario_financiero[
                    "Stock_Actual"
                ] >= stock_minimo_alto
            )
        ]
        .copy()
        .sort_values(
            "Valor_Stock_USD",
            ascending=False
        )
    )


    if productos_caros.empty:

        st.success(
            "No existen productos que cumplan "
            "los límites configurados."
        )

    else:

        mostrar_dataframe(
            productos_caros[
                [
                    "Producto",
                    "Categoria",
                    "Marca",
                    "Stock_Actual",
                    "Precio_USD",
                    "Valor_Stock_USD",
                ]
            ].round(2),
            use_container_width=True,
            hide_index=True,
        )


    st.divider()


    # ==================================================
    # PRODUCTOS CON 0-2 UNIDADES
    # ==================================================

    st.subheader(
        "🔴 Productos con 0–2 unidades"
    )

    st.caption(
        "Esta lista muestra únicamente el nivel físico "
        "de stock. No significa automáticamente que "
        "el producto deba reponerse."
    )


    stock_critico = (
        inventario_financiero[
            inventario_financiero[
                "Stock_Actual"
            ].between(
                0,
                2,
                inclusive="both"
            )
        ]
        .copy()
        .sort_values(
            [
                "Stock_Actual",
                "Precio_USD"
            ],
            ascending=[
                True,
                False
            ]
        )
    )


    sc1, sc2, sc3 = st.columns(3)

    sc1.metric(
        "Productos agotados",
        (
            stock_critico[
                "Stock_Actual"
            ] == 0
        ).sum()
    )

    sc2.metric(
        "Productos con 1 unidad",
        (
            stock_critico[
                "Stock_Actual"
            ] == 1
        ).sum()
    )

    sc3.metric(
        "Productos con 2 unidades",
        (
            stock_critico[
                "Stock_Actual"
            ] == 2
        ).sum()
    )


    mostrar_dataframe(
        stock_critico[
            [
                "Producto",
                "Categoria",
                "Marca",
                "Stock_Actual",
                "Precio_USD",
            ]
        ].round(2),
        use_container_width=True,
        hide_index=True,
    )


    st.divider()


    # ==================================================
    # MAYOR VALOR ACUMULADO
    # ==================================================

    st.subheader(
        "💎 Productos con mayor valor acumulado en stock"
    )

    st.caption(
        "Muestra en qué productos individuales está "
        "concentrada una mayor cantidad de dinero a costo."
    )


    cantidad_top_valor = st.slider(
        "Cantidad de productos a mostrar",
        min_value=5,
        max_value=30,
        value=15,
        key="top_valor_stock"
    )


    top_valor_producto = (
        stock_positivo
        .sort_values(
            "Valor_Stock_USD",
            ascending=False
        )
        .head(cantidad_top_valor)
        .copy()
    )


    top_valor_producto["Producto_Marca"] = (
        top_valor_producto[
            "Producto"
        ].astype(str)
        + " | "
        + top_valor_producto[
            "Marca"
        ].fillna("-").astype(str)
    )


    chart_producto_valor = (
        alt.Chart(
            top_valor_producto
        )
        .mark_bar()
        .encode(
            y=alt.Y(
                "Producto_Marca:N",
                sort="-x",
                title=None,
                axis=alt.Axis(
                    labelLimit=700,
                    labelPadding=10
                )
            ),
            x=alt.X(
                "Valor_Stock_USD:Q",
                title="Valor acumulado USD"
            ),
            tooltip=[
                alt.Tooltip(
                    "Producto:N",
                    title="Producto"
                ),
                alt.Tooltip(
                    "Marca:N",
                    title="Marca"
                ),
                alt.Tooltip(
                    "Stock_Actual:Q",
                    title="Stock actual",
                    format=".0f"
                ),
                alt.Tooltip(
                    "Precio_USD:Q",
                    title="Costo unitario USD",
                    format=",.2f"
                ),
                alt.Tooltip(
                    "Valor_Stock_USD:Q",
                    title="Valor acumulado USD",
                    format=",.2f"
                ),
            ]
        )
        .properties(
            height=max(
                420,
                cantidad_top_valor * 45
            )
        )
    )


    st.altair_chart(
        chart_producto_valor,
        use_container_width=True
    )


    st.divider()


    # ==================================================
    # DISTRIBUCIÓN
    # ==================================================

    st.subheader(
        "📊 Distribución del inventario"
    )


    d1, d2 = st.columns(2)


    agrupar_por = (
        d1.selectbox(
            "Agrupar por",
            [
                "Categoría",
                "Marca",
            ]
        )
    )


    medir_por = (
        d2.selectbox(
            "Medir distribución por",
            [
                "Unidades físicas",
                "Valor USD",
            ]
        )
    )


    if agrupar_por == "Categoría":
        columna_grupo = "Categoria"

    else:
        columna_grupo = "Marca"


    if medir_por == "Unidades físicas":

        distribucion = (
            stock_positivo
            .groupby(
                columna_grupo
            )[
                "Stock_Actual"
            ]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Stock_Actual":
                    "Valor"
                }
            )
        )

        titulo_valor = "Unidades"

    else:

        distribucion = (
            stock_positivo
            .groupby(
                columna_grupo
            )[
                "Valor_Stock_USD"
            ]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Valor_Stock_USD":
                    "Valor"
                }
            )
        )

        titulo_valor = "Valor USD"


    distribucion = (
        distribucion
        .sort_values(
            "Valor",
            ascending=False
        )
    )


    chart_distribucion = (
        alt.Chart(
            distribucion
        )
        .mark_bar()
        .encode(
            y=alt.Y(
                f"{columna_grupo}:N",
                sort="-x",
                title=None,
                axis=alt.Axis(
                    labelLimit=500
                )
            ),
            x=alt.X(
                "Valor:Q",
                title=titulo_valor
            ),
            tooltip=[
                alt.Tooltip(
                    f"{columna_grupo}:N",
                    title=agrupar_por
                ),
                alt.Tooltip(
                    "Valor:Q",
                    title=titulo_valor,
                    format=",.2f"
                ),
            ]
        )
        .properties(
            height=600
        )
    )


    st.altair_chart(
        chart_distribucion,
        use_container_width=True
    )


# ==================================================
# CONSULTA HISTÓRICA
# ==================================================

elif vista == "📈 Consulta histórica":

    st.header(
        "📈 Consulta histórica de ventas"
    )

    st.caption(
        "Por defecto se analizan los meses anteriores "
        "al último mes cargado. El último mes queda fuera "
        "porque puede contener valores arrastrados."
    )

    st.info(
        f"📌 Por defecto, la consulta llega hasta "
        f"{formato_mes(mes_anterior) if mes_anterior else '-'}."
    )


    incluir_provisional = st.checkbox(
        f"Incluir {formato_mes(ultimo_mes)} "
        f"(mes provisional / no validado)",
        value=False,
    )


    if incluir_provisional:

        base_consulta = df_propio.copy()

        st.warning(
            f"⚠️ Estás incluyendo "
            f"{formato_mes(ultimo_mes)}. "
            "Sus valores de ventas todavía "
            "no están validados."
        )

    else:

        base_consulta = df_historico.copy()


    meses_consulta = [
        pd.Timestamp(mes)
        for mes in sorted(
            base_consulta[
                "Mes_Fecha"
            ]
            .dropna()
            .unique()
        )
    ]


    if not meses_consulta:

        st.warning(
            "No hay meses históricos disponibles."
        )

        st.stop()


    f1, f2 = st.columns(2)


    desde = f1.selectbox(
        "Desde",
        meses_consulta,
        index=0,
        format_func=formato_mes,
    )


    hasta = f2.selectbox(
        "Hasta",
        meses_consulta,
        index=len(
            meses_consulta
        ) - 1,
        format_func=formato_mes,
    )


    if desde > hasta:

        st.error(
            "El mes inicial no puede ser posterior "
            "al mes final."
        )

        st.stop()


    periodo = (
        base_consulta[
            (
                base_consulta[
                    "Mes_Fecha"
                ] >= desde
            )
            &
            (
                base_consulta[
                    "Mes_Fecha"
                ] <= hasta
            )
        ]
        .copy()
    )


    st.subheader(
        "🔎 Filtros"
    )


    f3, f4 = st.columns(2)


    categorias = (
        ["TODAS"]
        +
        sorted(
            periodo[
                "Categoria"
            ]
            .dropna()
            .unique()
            .tolist()
        )
    )


    categoria_elegida = (
        f3.selectbox(
            "Categoría",
            categorias
        )
    )


    if categoria_elegida != "TODAS":

        periodo = (
            periodo[
                periodo[
                    "Categoria"
                ] == categoria_elegida
            ]
        )


    marcas = (
        ["TODAS"]
        +
        sorted(
            periodo[
                "Marca"
            ]
            .dropna()
            .unique()
            .tolist()
        )
    )


    marca_elegida = (
        f4.selectbox(
            "Marca",
            marcas
        )
    )


    if marca_elegida != "TODAS":

        periodo = (
            periodo[
                periodo[
                    "Marca"
                ] == marca_elegida
            ]
        )


    buscar = st.text_input(
        "🔍 Producto",
        placeholder=(
            "Ejemplo: 5600GT, RTX 5060, B650..."
        )
    )


    if buscar:

        periodo = (
            periodo[
                periodo[
                    "Producto"
                ]
                .str.contains(
                    buscar,
                    case=False,
                    na=False
                )
            ]
        )


    if periodo.empty:

        st.warning(
            "No existen datos que coincidan "
            "con los filtros."
        )

        st.stop()


    ventas_producto = (
        periodo
        .groupby(
            "Producto_Analitico"
        )
        .agg(
            Producto=("Producto", "last"),
            Marca=("Marca", "last"),
            Categoria=("Categoria", "last"),
            Unidades=("Vendido_Mes", "sum"),
        )
        .reset_index()
    )


    ventas_por_mes = (
        periodo
        .groupby(
            "Mes_Fecha"
        )[
            "Vendido_Mes"
        ]
        .sum()
        .reset_index()
    )


    ventas_periodo = (
        periodo[
            "Vendido_Mes"
        ].sum()
    )


    cantidad_meses = (
        periodo[
            "Mes_Fecha"
        ].nunique()
    )


    promedio_periodo = (
        ventas_periodo
        /
        cantidad_meses
        if cantidad_meses > 0
        else 0
    )


    productos_con_ventas = (
        ventas_producto[
            ventas_producto[
                "Unidades"
            ] > 0
        ][
            "Producto_Analitico"
        ]
        .nunique()
    )


    if not ventas_por_mes.empty:

        fila_pico = (
            ventas_por_mes.loc[
                ventas_por_mes[
                    "Vendido_Mes"
                ].idxmax()
            ]
        )

        mes_pico = formato_mes(
            fila_pico[
                "Mes_Fecha"
            ]
        )

        unidades_pico = (
            fila_pico[
                "Vendido_Mes"
            ]
        )

    else:

        mes_pico = "-"
        unidades_pico = 0


    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "Unidades vendidas",
        f"{ventas_periodo:.0f}"
    )

    m2.metric(
        "Promedio por mes",
        f"{promedio_periodo:.2f}"
    )

    m3.metric(
        "Productos con ventas",
        productos_con_ventas
    )

    m4.metric(
        "Mes con más ventas",
        f"{mes_pico} ({unidades_pico:.0f})"
    )


    productos_con_movimiento = (
        ventas_producto[
            ventas_producto[
                "Unidades"
            ] > 0
        ]
    )


    if not productos_con_movimiento.empty:

        producto_top = (
            productos_con_movimiento
            .sort_values(
                "Unidades",
                ascending=False
            )
            .iloc[0]
        )

        st.success(
            f"🏆 Producto más vendido del periodo: "
            f"**{producto_top['Producto']}** | "
            f"**Marca: {producto_top['Marca']}** — "
            f"**{producto_top['Unidades']:.0f} unidades**"
        )


    st.divider()


    st.subheader(
        "📊 Explorar gráficos"
    )


    tipo_analisis = st.selectbox(
        "¿Qué quieres visualizar?",
        [
            "Evolución mensual",
            "Comparar productos",
            "Top productos",
            "Ventas por categoría",
            "Ventas por marca",
        ],
    )


    if tipo_analisis == "Evolución mensual":

        grafico_mes = (
            ventas_por_mes.copy()
        )

        grafico_mes["Mes"] = (
            grafico_mes[
                "Mes_Fecha"
            ]
            .dt.strftime("%Y-%m")
        )

        chart = (
            alt.Chart(
                grafico_mes
            )
            .mark_line(
                point=True
            )
            .encode(
                x=alt.X(
                    "Mes:N",
                    title="Mes"
                ),
                y=alt.Y(
                    "Vendido_Mes:Q",
                    title="Unidades vendidas"
                ),
                tooltip=[
                    alt.Tooltip(
                        "Mes:N",
                        title="Mes"
                    ),
                    alt.Tooltip(
                        "Vendido_Mes:Q",
                        title="Unidades",
                        format=".0f"
                    ),
                ]
            )
            .properties(
                height=450
            )
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )


    elif tipo_analisis == "Top productos":

        cantidad_top = st.slider(
            "Cantidad de productos",
            min_value=5,
            max_value=30,
            value=10,
        )

        top = (
            ventas_producto[
                ventas_producto[
                    "Unidades"
                ] > 0
            ]
            .sort_values(
                "Unidades",
                ascending=False
            )
            .head(cantidad_top)
            .copy()
        )

        top["Producto_Marca"] = (
            top["Producto"].astype(str)
            + " | "
            + top["Marca"].fillna("-").astype(str)
        )

        chart = (
            alt.Chart(
                top
            )
            .mark_bar()
            .encode(
                y=alt.Y(
                    "Producto_Marca:N",
                    sort="-x",
                    title=None,
                    axis=alt.Axis(
                        labelLimit=700,
                        labelPadding=12,
                        labelFontSize=13,
                    ),
                ),
                x=alt.X(
                    "Unidades:Q",
                    title="Unidades vendidas"
                ),
                tooltip=[
                    alt.Tooltip(
                        "Producto:N",
                        title="Producto"
                    ),
                    alt.Tooltip(
                        "Marca:N",
                        title="Marca"
                    ),
                    alt.Tooltip(
                        "Unidades:Q",
                        title="Unidades",
                        format=".0f"
                    ),
                ],
            )
            .properties(
                height=max(
                    420,
                    cantidad_top * 48
                )
            )
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )


    elif tipo_analisis == "Ventas por categoría":

        categorias_grafico = (
            periodo
            .groupby(
                "Categoria"
            )[
                "Vendido_Mes"
            ]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Vendido_Mes":
                    "Unidades"
                }
            )
            .sort_values(
                "Unidades",
                ascending=False
            )
        )

        chart = (
            alt.Chart(
                categorias_grafico
            )
            .mark_bar()
            .encode(
                y=alt.Y(
                    "Categoria:N",
                    sort="-x",
                    title=None
                ),
                x=alt.X(
                    "Unidades:Q",
                    title="Unidades vendidas"
                ),
                tooltip=[
                    alt.Tooltip(
                        "Categoria:N",
                        title="Categoría"
                    ),
                    alt.Tooltip(
                        "Unidades:Q",
                        title="Unidades",
                        format=".0f"
                    ),
                ],
            )
            .properties(
                height=550
            )
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )


    elif tipo_analisis == "Ventas por marca":

        marcas_grafico = (
            periodo
            .groupby(
                "Marca"
            )[
                "Vendido_Mes"
            ]
            .sum()
            .reset_index()
            .rename(
                columns={
                    "Vendido_Mes":
                    "Unidades"
                }
            )
            .sort_values(
                "Unidades",
                ascending=False
            )
        )

        chart = (
            alt.Chart(
                marcas_grafico
            )
            .mark_bar()
            .encode(
                y=alt.Y(
                    "Marca:N",
                    sort="-x",
                    title=None
                ),
                x=alt.X(
                    "Unidades:Q",
                    title="Unidades vendidas"
                ),
                tooltip=[
                    alt.Tooltip(
                        "Marca:N",
                        title="Marca"
                    ),
                    alt.Tooltip(
                        "Unidades:Q",
                        title="Unidades",
                        format=".0f"
                    ),
                ],
            )
            .properties(
                height=550
            )
        )

        st.altair_chart(
            chart,
            use_container_width=True
        )


    elif tipo_analisis == "Comparar productos":

        productos_comparables = (
            periodo
            .groupby(
                "Producto_Analitico"
            )
            .agg(
                Producto=("Producto", "last"),
                Marca=("Marca", "last"),
            )
            .reset_index()
        )

        productos_comparables["Etiqueta"] = (
            productos_comparables[
                "Producto"
            ].astype(str)
            + " | "
            + productos_comparables[
                "Marca"
            ].fillna("-").astype(str)
        )

        etiquetas = sorted(
            productos_comparables[
                "Etiqueta"
            ].tolist()
        )

        comparar = st.multiselect(
            "Selecciona hasta 5 productos",
            etiquetas,
            max_selections=5,
        )

        if comparar:

            claves = (
                productos_comparables[
                    productos_comparables[
                        "Etiqueta"
                    ].isin(comparar)
                ][
                    "Producto_Analitico"
                ]
                .tolist()
            )

            comparacion = (
                periodo[
                    periodo[
                        "Producto_Analitico"
                    ].isin(claves)
                ]
                .copy()
            )

            comparacion = (
                comparacion
                .groupby(
                    [
                        "Mes_Fecha",
                        "Producto_Analitico"
                    ]
                )
                .agg(
                    Producto=("Producto", "last"),
                    Marca=("Marca", "last"),
                    Unidades=("Vendido_Mes", "sum"),
                )
                .reset_index()
            )

            comparacion["Mes"] = (
                comparacion[
                    "Mes_Fecha"
                ]
                .dt.strftime("%Y-%m")
            )

            comparacion["Producto_Marca"] = (
                comparacion["Producto"].astype(str)
                + " | "
                + comparacion["Marca"].fillna("-").astype(str)
            )

            chart = (
                alt.Chart(
                    comparacion
                )
                .mark_line(
                    point=True
                )
                .encode(
                    x=alt.X(
                        "Mes:N",
                        title="Mes"
                    ),
                    y=alt.Y(
                        "Unidades:Q",
                        title="Unidades vendidas"
                    ),
                    color=alt.Color(
                        "Producto_Marca:N",
                        title="Producto | Marca"
                    ),
                    tooltip=[
                        alt.Tooltip(
                            "Mes:N",
                            title="Mes"
                        ),
                        alt.Tooltip(
                            "Producto:N",
                            title="Producto"
                        ),
                        alt.Tooltip(
                            "Marca:N",
                            title="Marca"
                        ),
                        alt.Tooltip(
                            "Unidades:Q",
                            title="Unidades",
                            format=".0f"
                        ),
                    ],
                )
                .properties(
                    height=500
                )
            )

            st.altair_chart(
                chart,
                use_container_width=True
            )

        else:

            st.info(
                "Selecciona uno o más productos "
                "para compararlos."
            )


    st.divider()


    st.subheader(
        "📋 Detalle del periodo"
    )


    tabla_historica = (
        periodo
        .groupby(
            [
                "Mes",
                "Producto_Analitico"
            ]
        )
        .agg(
            Producto=("Producto", "last"),
            Categoria=("Categoria", "last"),
            Marca=("Marca", "last"),
            Stock_Inicial=("Stock_Inicial", "sum"),
            Vendido=("Vendido_Mes", "sum"),
            Stock_Final=("Stock_Final", "sum"),
            Precio_USD=("Precio_USD", "max"),
            TC=("TC", "max"),
        )
        .reset_index()
        .sort_values(
            [
                "Mes",
                "Vendido"
            ],
            ascending=[
                True,
                False
            ]
        )
    )


    tabla_visible = (
        tabla_historica
        .drop(
            columns=[
                "Producto_Analitico"
            ]
        )
    )


    mostrar_dataframe(
        tabla_visible,
        use_container_width=True,
        hide_index=True,
    )


    csv = (
        tabla_visible
        .rename(
            columns=NOMBRES_COLUMNAS_CLIENTE
        )
        .to_csv(
            index=False
        )
        .encode(
            "utf-8-sig"
        )
    )


    st.download_button(
        "⬇️ Descargar consulta en CSV",
        data=csv,
        file_name=(
            "consulta_historica_infometa.csv"
        ),
        mime="text/csv",
    )


# ==================================================
# DETALLE DE PRODUCTO
# ==================================================

elif vista == "🔎 Detalle de producto":

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
        "Si hay presupuesto, cubrir también "
        "el próximo mes",
        value=False,
    )

    confirmar_mes_actual = st.sidebar.checkbox(
        f"Confirmo que VENDIDO de "
        f"{formato_mes(ultimo_mes)} "
        f"fue reiniciado/actualizado",
        value=False,
    )


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
        and
        hoy.month == mes_inventario
    ):

        dia_default = hoy.day

    else:

        dia_default = dias_mes


    dia_corte = st.sidebar.number_input(
        "Día hasta el que está actualizado "
        "el inventario",
        min_value=1,
        max_value=dias_mes,
        value=dia_default,
        step=1,
    )


    dias_restantes = max(
        dias_mes - dia_corte,
        0
    )


    st.header(
        "🔎 Análisis por producto"
    )

    st.warning(
        "Las recomendaciones usan histórico heredado. "
        "Deben utilizarse como apoyo y no como una orden "
        "automática de compra."
    )


    texto_busqueda = st.text_input(
        "🔍 Buscar producto",
        placeholder=(
            "Ejemplo: RTX 5060, B650, "
            "5600GT, SSD 1TB..."
        )
    )


    productos_actuales = (
        resumen[
            "Producto"
        ]
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


    if not productos_filtrados:

        st.warning(
            "No se encontraron productos "
            "con ese texto."
        )

        st.stop()


    producto_elegido = st.selectbox(
        "Selecciona un producto",
        productos_filtrados,
    )


    fila_resumen = (
        resumen[
            resumen[
                "Producto"
            ] == producto_elegido
        ]
        .iloc[0]
    )


    clave_producto = (
        fila_resumen[
            "Producto_Analitico"
        ]
    )


    datos_producto = (
        df_propio[
            df_propio[
                "Producto_Analitico"
            ] == clave_producto
        ]
        .copy()
        .sort_values(
            "Mes_Fecha"
        )
    )


    st.subheader(
        producto_elegido
    )

    st.caption(
        f"{fila_resumen['Categoria']} | "
        f"{fila_resumen['Marca']}"
    )


    ventas_mensuales = (
        datos_producto
        .groupby(
            "Mes_Fecha"
        )
        .agg(
            Vendido_Mes=(
                "Vendido_Mes",
                "sum"
            )
        )
        .reset_index()
    )


    meses_historicos_producto = (
        ventas_mensuales[
            ventas_mensuales[
                "Mes_Fecha"
            ] < ultimo_mes
        ]
    )


    if not meses_historicos_producto.empty:

        promedio_historico = (
            meses_historicos_producto[
                "Vendido_Mes"
            ].mean()
        )

        ventas_historicas = (
            meses_historicos_producto[
                "Vendido_Mes"
            ].sum()
        )

    else:

        promedio_historico = 0
        ventas_historicas = 0


    vendido_actual_registrado = (
        ventas_mensuales.loc[
            ventas_mensuales[
                "Mes_Fecha"
            ] == ultimo_mes,
            "Vendido_Mes"
        ]
        .sum()
    )


    stock_actual = (
        fila_resumen[
            "Stock_Actual"
        ]
    )

    precio_usd = (
        fila_resumen[
            "Precio_USD"
        ]
    )

    tc_actual = (
        fila_resumen[
            "TC"
        ]
    )


    demanda_base = (
        promedio_historico
    )


    tipo_demanda = clasificar_demanda(
        demanda_base,
        umbral_alta,
        umbral_media
    )


    if demanda_base > 0:

        cobertura = (
            stock_actual
            /
            demanda_base
        )

    else:

        cobertura = None


    demanda_diaria_historica = (
        demanda_base
        /
        dias_mes
        if demanda_base > 0
        else 0
    )


    demanda_restante_base = (
        demanda_diaria_historica
        *
        dias_restantes
    )


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
            -
            stock_actual
        )
    )


    if tipo_demanda in [
        "🐢 BAJA DEMANDA",
        "⚪ SIN MOVIMIENTO",
    ]:

        pedido_base = 0


    proyeccion_mes_actual = None
    pedido_ritmo_actual = None


    if confirmar_mes_actual:

        ritmo_diario_actual = (
            vendido_actual_registrado
            /
            dia_corte
            if dia_corte > 0
            else 0
        )

        demanda_restante_actual = (
            ritmo_diario_actual
            *
            dias_restantes
        )

        proyeccion_mes_actual = (
            ritmo_diario_actual
            *
            dias_mes
        )

        pedido_ritmo_actual = max(
            0,
            math.ceil(
                demanda_restante_actual
                -
                stock_actual
            )
        )


    dias_hasta_minimo = None


    if (
        demanda_diaria_historica > 0
        and
        stock_actual > stock_minimo_alta
    ):

        dias_hasta_minimo = (
            stock_actual
            -
            stock_minimo_alta
        ) / demanda_diaria_historica


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
            dias_hasta_minimo
            is not None
            and
            dias_hasta_minimo <= dias_alerta
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


    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Stock actual",
        f"{stock_actual:.0f}"
    )

    c2.metric(
        "Ventas históricas",
        f"{ventas_historicas:.0f}"
    )

    c3.metric(
        "Promedio histórico mensual",
        f"{promedio_historico:.2f}"
    )

    c4.metric(
        "Cobertura histórica",
        (
            f"{cobertura:.2f} meses"
            if cobertura is not None
            else "Sin datos"
        )
    )


    st.subheader(
        "📈 Nivel de demanda"
    )

    st.info(
        tipo_demanda
    )

    st.write(
        f"Referencia utilizada: "
        f"**{demanda_base:.2f} unidades por mes**."
    )


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
            f"Stock mínimo configurado: "
            f"**{stock_minimo_alta} unidades**."
        )

        if dias_hasta_minimo is not None:

            st.write(
                f"Al ritmo histórico podría llegar "
                f"a ese mínimo en aproximadamente "
                f"**{dias_hasta_minimo:.1f} días**."
            )


    st.subheader(
        "📦 Cantidad sugerida"
    )


    if incluir_mes_siguiente:

        horizonte = (
            "resto del mes y próximo mes"
        )

    else:

        horizonte = (
            "resto del mes"
        )


    q1, q2 = st.columns(2)


    q1.metric(
        f"Pedido base para cubrir {horizonte}",
        f"{pedido_base} unidades"
    )


    if confirmar_mes_actual:

        q2.metric(
            "Escenario si continúa "
            "el ritmo actual",
            f"{pedido_ritmo_actual} unidades"
        )

    else:

        q2.metric(
            "Escenario con ritmo "
            "del mes actual",
            "No disponible"
        )


    st.caption(
        "El pedido base usa el histórico disponible. "
        "El escenario del mes actual solo se habilita "
        "cuando confirmas que las ventas del mes "
        "fueron correctamente actualizadas."
    )


    st.subheader(
        f"📊 {formato_mes(ultimo_mes)} "
        f"— mes provisional"
    )


    if confirmar_mes_actual:

        st.success(
            "Has confirmado que las ventas "
            "del mes actual fueron actualizadas."
        )

    else:

        st.warning(
            "Las ventas del mes actual no están validadas. "
            "Se muestran como información provisional "
            "y no se utilizan para proyectar demanda."
        )


    a1, a2, a3 = st.columns(3)

    a1.metric(
        "Ventas registradas",
        f"{vendido_actual_registrado:.0f}"
    )

    a2.metric(
        "Día de corte",
        f"{dia_corte}"
    )

    a3.metric(
        "Proyección al cierre",
        (
            f"{proyeccion_mes_actual:.1f}"
            if proyeccion_mes_actual is not None
            else "No calculada"
        )
    )


    p1, p2 = st.columns(2)

    p1.metric(
        "Costo USD actual",
        (
            f"${precio_usd:.2f}"
            if pd.notna(
                precio_usd
            )
            else "Sin dato"
        )
    )

    p2.metric(
        "Tipo de cambio registrado",
        (
            f"{tc_actual:.2f}"
            if pd.notna(
                tc_actual
            )
            else "Sin dato"
        )
    )


    st.subheader(
        "📈 Historial de ventas"
    )


    grafico_producto = (
        ventas_mensuales.copy()
    )


    grafico_producto["Mes"] = (
        grafico_producto[
            "Mes_Fecha"
        ]
        .dt.strftime("%Y-%m")
    )


    grafico_producto["Estado_Mes"] = (
        grafico_producto[
            "Mes_Fecha"
        ]
        .apply(
            lambda x:
            (
                "Provisional"
                if x == ultimo_mes
                else "Histórico"
            )
        )
    )


    chart = (
        alt.Chart(
            grafico_producto
        )
        .mark_line(
            point=True
        )
        .encode(
            x=alt.X(
                "Mes:N",
                title="Mes"
            ),
            y=alt.Y(
                "Vendido_Mes:Q",
                title="Unidades vendidas"
            ),
            tooltip=[
                alt.Tooltip(
                    "Mes:N",
                    title="Mes"
                ),
                alt.Tooltip(
                    "Estado_Mes:N",
                    title="Estado"
                ),
                alt.Tooltip(
                    "Vendido_Mes:Q",
                    title="Unidades vendidas",
                    format=".0f"
                ),
            ],
        )
        .properties(
            height=500
        )
    )


    st.altair_chart(
        chart,
        use_container_width=True
    )


    st.caption(
        f"⚠️ El punto de "
        f"{formato_mes(ultimo_mes)} "
        f"es provisional."
    )


    st.subheader(
        "📋 Historial del producto"
    )


    historial = (
        datos_producto
        .groupby(
            "Mes"
        )
        .agg(
            Producto=("Producto", "last"),
            Marca=("Marca", "last"),
            Stock_Inicial=("Stock_Inicial", "sum"),
            Vendido_Mes=("Vendido_Mes", "sum"),
            Stock_Final=("Stock_Final", "sum"),
            Precio_USD=("Precio_USD", "max"),
            TC=("TC", "max"),
        )
        .reset_index()
    )


    historial["Estado"] = (
        historial[
            "Mes"
        ]
        .apply(
            lambda x:
            (
                "⚠️ PROVISIONAL"
                if x == formato_mes(
                    ultimo_mes
                )
                else "HISTÓRICO"
            )
        )
    )


    mostrar_dataframe(
        historial,
        use_container_width=True,
        hide_index=True,
    )


# ==================================================
# CALIDAD DE DATOS
# ==================================================

else:

    st.header(
        "🧪 Calidad de datos"
    )

    st.caption(
        "Esta sección busca detectar señales "
        "de arrastre o falta de reinicio entre "
        "el último mes y el anterior."
    )


    if (
        mes_anterior is None
        or
        comparacion_calidad.empty
    ):

        st.info(
            "No existen suficientes meses "
            "para realizar la comparación."
        )

        st.stop()


    st.subheader(
        f"Comparación "
        f"{formato_mes(mes_anterior)} "
        f"→ "
        f"{formato_mes(ultimo_mes)}"
    )


    total_comparables = (
        len(
            comparacion_calidad
        )
    )


    mismo_vendido_n = int(
        comparacion_calidad[
            "Mismo_Vendido"
        ].sum()
    )


    mismo_registro_n = int(
        comparacion_calidad[
            "Mismo_Registro"
        ].sum()
    )


    q1, q2, q3 = st.columns(3)


    q1.metric(
        "Productos comparables",
        total_comparables
    )


    q2.metric(
        "Mismas ventas",
        f"{pct_mismo_vendido:.1f}%",
        help=(
            f"{mismo_vendido_n} productos "
            f"tienen exactamente el mismo número "
            f"de ventas en ambos meses."
        ),
    )


    q3.metric(
        "Registro completo idéntico",
        f"{pct_mismo_registro:.1f}%",
        help=(
            f"{mismo_registro_n} productos "
            f"mantienen iguales el stock inicial, "
            f"las ventas y el stock final."
        ),
    )


    if pct_mismo_registro >= 60:

        st.error(
            "🚨 Alta señal de posible arrastre de datos. "
            "Una gran proporción de productos mantiene "
            "exactamente los mismos valores que "
            "el mes anterior."
        )

    elif pct_mismo_registro >= 30:

        st.warning(
            "⚠️ Se detectó una proporción relevante "
            "de registros idénticos entre ambos meses. "
            "Conviene revisar cómo se cerró "
            "la hoja de inventario."
        )

    else:

        st.success(
            "No se detecta una señal fuerte "
            "de arrastre mediante esta comparación."
        )


    st.markdown(
        f"""
### Cómo interpreta la aplicación estos datos

- **{formato_mes(ultimo_mes)}** se considera provisional para ventas.
- Sus ventas no entran automáticamente en los promedios históricos.
- El **stock actual sí se toma del último mes**.
- Los meses anteriores se consideran histórico disponible.
- Una coincidencia entre dos meses no significa necesariamente que exista un error.
- Muchas coincidencias juntas sí son una señal para revisar la fuente.
"""
    )


    st.subheader(
        "🔍 Registros sospechosos"
    )


    sospechosos = (
        comparacion_calidad[
            comparacion_calidad[
                "Mismo_Registro"
            ]
        ]
        .copy()
    )


    tabla_sospechosos = (
        sospechosos[
            [
                "Producto_Actual",
                "Stock_Inicial_Anterior",
                "Vendido_Anterior",
                "Stock_Final_Anterior",
                "Stock_Inicial_Actual",
                "Vendido_Actual",
                "Stock_Final_Actual",
            ]
        ]
        .rename(
            columns={
                "Producto_Actual":
                "Producto",

                "Stock_Inicial_Anterior":
                f"Stock inicial {formato_mes(mes_anterior)}",

                "Vendido_Anterior":
                f"Unidades vendidas {formato_mes(mes_anterior)}",

                "Stock_Final_Anterior":
                f"Stock final {formato_mes(mes_anterior)}",

                "Stock_Inicial_Actual":
                f"Stock inicial {formato_mes(ultimo_mes)}",

                "Vendido_Actual":
                f"Unidades vendidas {formato_mes(ultimo_mes)}",

                "Stock_Final_Actual":
                f"Stock final {formato_mes(ultimo_mes)}",
            }
        )
    )


    mostrar_dataframe(
        tabla_sospechosos,
        use_container_width=True,
        hide_index=True,
    )