# InfoMeta Inventory Analytics

Sistema de análisis de inventario y apoyo a decisiones de reposición desarrollado con **Python**, **Pandas** y **Streamlit**.

## Objetivo

Transformar el inventario histórico de una tienda de computación en información útil para decidir:

- qué productos requieren reposición;
- cuáles deben vigilarse;
- qué productos tienen baja rotación;
- cuánto stock queda respecto al ritmo de ventas;
- cómo evolucionan las ventas, el stock y el precio en el tiempo.

## Estado actual — v0.1

El primer prototipo ya permite:

- cargar el histórico limpio desde Excel;
- seleccionar un producto;
- consultar stock actual;
- sumar ventas de los últimos meses;
- calcular promedio mensual;
- estimar meses de cobertura;
- generar una recomendación inicial de reposición;
- visualizar el historial mensual de ventas.

## Tecnologías

- Python
- Streamlit
- Pandas
- OpenPyXL
- Git / GitHub

## Datos

Los datos reales de la empresa **no se publican en este repositorio**.

La aplicación utiliza localmente el archivo:

`Base_Limpia_Inventario_InfoMeta.xlsx`

Ese archivo está excluido del control de versiones para proteger información interna. Más adelante se añadirá un dataset ficticio o anonimizado para demostración pública.

## Ejecutar localmente

1. Instalar dependencias:

```bash
pip install -r requirements.txt
```

2. Colocar el archivo de datos local en la raíz del proyecto.

3. Ejecutar:

```bash
streamlit run app.py
```

## Próximas etapas

- panel general de productos que requieren atención;
- ranking de prioridades de reposición;
- detección de sobrestock y capital inmovilizado;
- análisis por categoría y marca;
- sugerencia de cantidad a reponer;
- análisis de tendencia;
- incorporación futura de modelos de pronóstico de demanda / Machine Learning.

## Privacidad

Este proyecto forma parte de un portafolio profesional, pero el inventario real, precios internos y demás información sensible de la empresa se mantienen fuera del repositorio.
