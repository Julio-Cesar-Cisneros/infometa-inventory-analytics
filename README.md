# InfoMeta Inventory Analytics

Sistema de análisis de inventario y apoyo a decisiones de reposición desarrollado con **Python**, **Pandas**, **Streamlit** y **OpenPyXL**.

El proyecto nació a partir de una necesidad real: transformar registros históricos de inventario de una tienda de tecnología en información más clara para apoyar decisiones sobre **stock, reposición, rotación, costos y calidad de los datos**.

## Objetivo

Convertir datos de inventario almacenados en Excel en un panel interactivo que permita responder preguntas como:

- ¿Qué productos requieren atención?
- ¿Qué artículos tienen stock bajo o están agotados?
- ¿Qué productos tienen mayor rotación histórica?
- ¿Cuántos meses de cobertura quedan según el ritmo de ventas?
- ¿Qué categorías y marcas concentran mayor valor de inventario?
- ¿Qué productos tienen mayor valor acumulado en stock?
- ¿Cuánto se vendió en un periodo determinado?
- ¿Existen señales de datos arrastrados o registros que necesitan revisión?

## Funcionalidades actuales

### 📊 Resumen general

- clasificación de productos por estado de inventario;
- detección de productos agotados con ventas históricas;
- identificación de stock bajo y productos a vigilar;
- cálculo de ventas registradas, promedio mensual y cobertura estimada;
- ranking de productos con mayor rotación;
- identificación de productos sin ventas registradas.

### 💰 Análisis financiero del inventario

El módulo financiero utiliza el **stock físico actual** y el **costo unitario en USD** registrado en la base.

Incluye:

- unidades disponibles;
- cantidad de productos con stock;
- costo aproximado total del inventario en USD;
- conversión del costo del inventario a bolivianos mediante un tipo de cambio configurable;
- categorías donde existe mayor valor concentrado;
- valor del inventario por marca;
- detección de productos de costo elevado con varias unidades;
- productos con 0, 1 o 2 unidades;
- productos con mayor valor acumulado en stock;
- distribución del inventario por categoría o marca.

### 📈 Consulta histórica

Permite analizar las ventas dentro de un periodo seleccionado mediante filtros por:

- mes inicial y final;
- categoría;
- marca;
- producto.

La aplicación muestra indicadores como:

- unidades vendidas;
- promedio mensual;
- productos con ventas;
- mes con mayor movimiento;
- producto más vendido del periodo;
- evolución mensual;
- ranking de productos.

### 🔎 Detalle de producto

Para cada producto se puede consultar:

- stock actual;
- ventas registradas;
- promedio mensual;
- cobertura estimada;
- costo unitario;
- historial de ventas;
- clasificación de demanda;
- estado del inventario;
- sugerencias orientativas de reposición.

La lógica de reposición busca apoyar una decisión humana y **no ejecutar compras automáticamente**.

### 🧪 Calidad de datos

Durante el desarrollo se detectó que algunos registros mensuales podían conservar valores del mes anterior.

Por ello se incorporó un módulo específico que:

- compara el último mes con el anterior;
- calcula el porcentaje de productos con valores de venta idénticos;
- detecta registros completos potencialmente repetidos;
- marca el último mes como provisional cuando corresponde;
- evita incluir automáticamente un mes no validado en los promedios históricos.

Este módulo permite distinguir entre el **stock físico actual**, que sí puede utilizarse, y los datos de ventas que todavía requieren validación.

## Reglas de negocio incorporadas

La aplicación trabaja únicamente con inventario identificado como **stock propio**.

Los productos pertenecientes a terceros quedan fuera de los cálculos de inventario y reposición.

Además:

- el último mes disponible se utiliza como referencia para el stock actual;
- las ventas del último mes pueden excluirse de los promedios cuando todavía son provisionales;
- la cobertura se calcula relacionando stock disponible y promedio histórico de ventas;
- los umbrales de demanda y reposición pueden adaptarse según el comportamiento del negocio.

## Tecnologías

- **Python**
- **Streamlit**
- **Pandas**
- **OpenPyXL**
- **Altair**
- **Git**
- **GitHub**
- **Excel**

## Estructura actual

```text
infometa-inventory-analytics/
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

El archivo real utilizado como fuente de datos se mantiene únicamente de forma local:

```text
Base_Limpia_Inventario_InfoMeta.xlsx
```

## Privacidad de los datos

Los datos reales de la empresa **no se publican en este repositorio**.

El archivo Excel que contiene inventario, costos y registros internos está excluido mediante `.gitignore`.

El repositorio contiene únicamente el código de la aplicación y la documentación necesaria para mostrar el desarrollo técnico del proyecto.

## Ejecutar localmente

### 1. Instalar las dependencias

```bash
pip install -r requirements.txt
```

### 2. Colocar la fuente de datos local

La aplicación espera encontrar en la raíz del proyecto:

```text
Base_Limpia_Inventario_InfoMeta.xlsx
```

con una hoja llamada:

```text
Historico_Limpio
```

### 3. Ejecutar Streamlit

```bash
streamlit run app.py
```

## Aprendizajes del proyecto

Este proyecto me ha permitido trabajar de forma práctica con:

- limpieza y transformación de datos reales;
- análisis exploratorio;
- construcción de indicadores de negocio;
- visualización de información;
- diseño de interfaces con Streamlit;
- tratamiento de datos inconsistentes;
- traducción de necesidades del usuario a reglas de negocio;
- control de versiones con Git y GitHub.

## Estado del proyecto

**En desarrollo activo.**

Actualmente el sistema ya permite analizar inventario, costos, rotación, ventas históricas y calidad de datos. Las siguientes mejoras estarán enfocadas en presentación para portafolio, datos de demostración y nuevas herramientas de apoyo a decisiones.

## Próximos pasos

- incorporar un dataset ficticio o anonimizado para demostración pública;
- añadir capturas de pantalla de la aplicación;
- mejorar la documentación técnica;
- incorporar nuevas métricas de reposición;
- evaluar modelos de pronóstico de demanda cuando exista un histórico suficientemente confiable.

---

**Proyecto de portafolio desarrollado por Julio Cesar Cisneros Calzadilla.**
