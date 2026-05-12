

import streamlit as st      # Librería para crear aplicaciones web interactivas
import pandas as pd         # Librería para trabajar con datos (tablas)
from pathlib import Path    # Para manejar rutas de archivos de forma moderna
import io                   # Para crear archivos en memoria (necesario para descargar Excel)



# =============================================================================
# 1. CONFIGURACIÓN INICIAL DE LA APLICACIÓN
# =============================================================================

# Configuramos cómo se verá la página en el navegador
st.set_page_config(
    page_title="Análisis de Empresas",   # Título que aparece en la pestaña del navegador
    layout="centered",                   # Centra el contenido
    initial_sidebar_state="expanded"     # La barra lateral aparece abierta por defecto
)

# Título principal de la aplicación
st.title("Análisis de Empresas")
st.markdown("Visualización sencilla de facturación por cliente")  # Subtítulo descriptivo



# =============================================================================
# 2. CARGAR LOS DATOS (se ejecuta solo una vez gracias al cache)
# =============================================================================

@st.cache_data  # esto hecho con ayuda de ia.   # Guarda los datos en memoria para que no se lean cada vez
def cargar_datos():
    """Función que carga el archivo Excel"""
    # Definimos la ruta donde está el archivo (cámbiala si es necesario)
    ruta = Path(r"C:\Users\ingsoporte3\OneDrive - Syspotec\Documentos\empresas\emp_clean.xlsx")
    
    # Verificamos si el archivo existe
    if not ruta.exists():
        st.error("No se encontró el archivo 'emp_clean.xlsx'. Verifica la ruta.")
        st.stop()                     # Detiene la ejecución de la app
    
    # Leemos el archivo Excel
    df = pd.read_excel(
        ruta,
        engine="openpyxl",            # Motor recomendado para archivos .xlsx
        parse_dates=['Date']          # Convierte automáticamente la columna Date a fecha
    )
    
    # Creamos la columna Año_Mes (ej: 2026-03) si no existe
    if 'Año_Mes' not in df.columns:
        df['Año_Mes'] = df['Date'].dt.strftime('%Y-%m')
    
    return df

# Ejecutamos la función para cargar los datos
df = cargar_datos()


# =============================================================================
# 3. FILTROS DE FECHA (Sidebar)
# =============================================================================

# Todos los filtros van en la barra lateral para que quede más limpio
st.sidebar.header("Filtros de Fecha")

# Obtenemos la lista de años disponibles en los datos
años_disponibles = sorted(df['Date'].dt.year.unique())

# Filtro para seleccionar el año
año_seleccionado = st.sidebar.selectbox(
    "Seleccionar Año",
    options=años_disponibles,
    index=len(años_disponibles)-1          # Por defecto selecciona el año más reciente
)

# Filtramos los datos por el año seleccionado
df_filtrado = df[df['Date'].dt.year == año_seleccionado].copy()

# Opción para elegir cómo ver los datos
tipo_vista = st.sidebar.radio(
    "Ver por:",
    options=["Año completo", "Por trimestre", "Mes específico"],
    horizontal=True
)

# Aplicamos filtro adicional según la opción elegida
if tipo_vista == "Por trimestre":
    trimestre = st.sidebar.selectbox("Seleccionar Trimestre", [1, 2, 3, 4])
    df_filtrado = df_filtrado[df_filtrado['Date'].dt.quarter == trimestre]
    
elif tipo_vista == "Mes específico":
    meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
             "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    mes_nombre = st.sidebar.selectbox("Seleccionar Mes", meses)
    mes_numero = meses.index(mes_nombre) + 1
    df_filtrado = df_filtrado[df_filtrado['Date'].dt.month == mes_numero]


# =============================================================================
# 4. CALCULAR LOS KPIs (con los datos ya filtrados)
# =============================================================================


# Calculamos el cliente con mayor facturación
valor_por_cliente = df_filtrado.groupby('ClientName')['Total'].sum().sort_values(ascending=False)

cliente_top = valor_por_cliente.index[0]      # Nombre del cliente #1
valor_top = valor_por_cliente.iloc[0]         # Monto total de ese cliente

# Calculamos el mes con mayor valor de ese cliente
df_top_cliente = df_filtrado[df_filtrado['ClientName'] == cliente_top]
valor_por_mes = df_top_cliente.groupby('Año_Mes')['Total'].sum().sort_values(ascending=False)

mes_top = valor_por_mes.index[0]
valor_mes_top = valor_por_mes.iloc[0]


# =============================================================================
# 5. MOSTRAR LOS KPIs (tarjetas)
# =============================================================================


st.subheader("KPIs Principales")

col1, col2 = st.columns(2)   # Creamos dos columnas para mostrar los KPIs lado a lado

with col1:
    st.metric(
        label="Cliente con mayor valor total",
        value=cliente_top,
        delta=f"${valor_top:,.2f}"          # Muestra el monto destacado
    )

with col2:
    st.metric(
        label=f"Mes con mayor valor de {cliente_top}",
        value=mes_top,
        delta=f"${valor_mes_top:,.2f}"
    )

# =============================================================================
# 6. GRÁFICA: Valor mensual por cliente (Top 10)
# =============================================================================


st.subheader("Valor mensual por cliente (Top 10)")

# Seleccionamos los 10 clientes con más facturación
top_clientes = df_filtrado.groupby('ClientName')['Total'].sum().nlargest(10).index

# Filtramos solo esos clientes
df_grafico = df_filtrado[df_filtrado['ClientName'].isin(top_clientes)]

# Preparamos los datos para la gráfica (una columna por cliente)
grafico_data = df_grafico.groupby(['Año_Mes', 'ClientName'])['Total'].sum().unstack()

# Mostramos la gráfica de líneas
st.line_chart(grafico_data, use_container_width=True, height=450)


# =============================================================================
# 7. TABLA FILTRADA + BOTÓN DE DESCARGA EN EXCEL
# =============================================================================


st.subheader(f"Tabla de datos filtrados - {año_seleccionado}")

# Mostramos la tabla en pantalla
st.dataframe(
    df_filtrado.style.format({
        "Total": "${:,.2f}",
        "Subtotal": "${:,.2f}",
        "TaxSale": "${:,.2f}"
    }),
    use_container_width=True,
    hide_index=True
)

# ====================== DESCARGA EN EXCEL ======================

# Creamos el archivo Excel en memoria
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name=f"Datos_{año_seleccionado}")

output.seek(0)   # esto hecho con ayuda de ia.

# Botón de descarga
st.download_button(
    label="Descargar tabla filtrada (Excel)",
    data=output,
    file_name=f"datos_empresas_{año_seleccionado}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Datos cargados desde emp_clean.xlsx")