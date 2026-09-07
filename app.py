import os
import glob
import pandas as pd
import streamlit as st
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Configuración inicial de la página
st.set_page_config(
    page_title="Dashboard de Incidencias",
    page_icon="📊",
    layout="wide"
)

# Ruta del archivo Excel consolidado por defecto en la app
EXCEL_DESTINO = "resultado_unificado.xlsx"

# =========================================================================
# FUNCIÓN DE PROCESAMIENTO, FUSIÓN Y ACTUALIZACIÓN INTELIGENTE
# =========================================================================
def procesar_y_actualizar_incidencias(archivos_txt_subidos, ruta_excel_actual=EXCEL_DESTINO):
    indices_a_conservar = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26, 28, 29, 30, 31, 32, 34, 35, 36, 37, 38]
    nombres_finales = [
        "INCIDENCIA", "DIRECCIÓN", "MOTIVO", "RECEPCIÓN", "FINALIZACIÓN", 
        "TEC", "TCR", "TDV", "TEJ", "TAR", "ESTADO", "CLIENTE", 
        "TOTAL INCIDENCIAS", "SUMATORIAS TEC", "SUMATORIAS TCR", "SUMATORIAS TDV", 
        "SUMATORIAS TEJ", "SUMATORIAS TAR", "PROMEDIO TEC", "PROMEDIO TCR", 
        "PROMEDIO TDV", "PROMEDIO TEJ", "PROMEDIO TAR"
    ]

    mapeo_zonas = {
        "01": "NORTE", "02": "SUR", "03": "ESTE", "04": "OESTE", "05": "CENTRO",
        "06": "VÍA DUACA", "07": "VÍA RÍO CLARO", "08": "VÍA PAVIA", "09": "VÍA BUENA VISTA",
        "10": "VIA VIEJA CARORA", "11": "VÍA QUIBOR", "12": "VÍA AUTOPISTA CARORA"
    }

    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
        7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
    }

    lista_dataframes = []

    for archivo in archivos_txt_subidos:
        nombre_archivo = archivo.name if hasattr(archivo, "name") else os.path.basename(archivo)
        try:
            # Leer el archivo txt forzando todo como texto
            df = pd.read_csv(
                archivo, sep="|", header=None, dtype=str, 
                encoding="latin-1", engine="python", on_bad_lines="skip"
            )

            df = df.iloc[:, indices_a_conservar]
            df.columns = nombres_finales

            codigo = nombre_archivo[:2]
            nombre_zona = mapeo_zonas.get(codigo, nombre_archivo)
            df.insert(0, "ZONA", nombre_zona)

            # Limpieza general
            for col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
                df[col] = df[col].replace(["NAN", "NONE", "NAT", "NAN.0"], "", regex=False)
                if col == "INCIDENCIA":
                    df[col] = df[col].str.replace(r'\.0$', '', regex=True)

            df = df.replace(r'\*', '', regex=True)

            # Filtros estrictos
            df = df[df["INCIDENCIA"] != ""]
            df = df[df["INCIDENCIA"].str.match(r'^\d+$', na=False)]
            df = df[df["RECEPCIÓN"] != ""]

            dt_recepcion = pd.to_datetime(df["RECEPCIÓN"], format='%d-%m-%y %I:%M %p', errors='coerce')
            df["_DT_RECEPCION_TEMP"] = dt_recepcion
            df = df.dropna(subset=["_DT_RECEPCION_TEMP"])
            df = df.drop(columns=["_DT_RECEPCION_TEMP"])

            dt_recepcion = pd.to_datetime(df["RECEPCIÓN"], format='%d-%m-%y %I:%M %p', errors='coerce')
            mes_recibido = dt_recepcion.dt.month.map(meses_es).fillna("")
            anio_recibido = dt_recepcion.dt.year.fillna("").astype(str).str.replace(r'\.0$', '', regex=True)

            dt_finalizacion = pd.to_datetime(df["FINALIZACIÓN"], format='%d-%m-%y %I:%M %p', errors='coerce')
            mes_finalizado = dt_finalizacion.dt.month.map(meses_es).fillna("")
            anio_finalizado = dt_finalizacion.dt.year.fillna("").astype(str).str.replace(r'\.0$', '', regex=True)

            idx_rec = df.columns.get_loc("RECEPCIÓN")
            df.insert(idx_rec, "MES RECIBIDO", mes_recibido)
            df.insert(idx_rec, "AÑO RECIBIDO", anio_recibido)

            idx_fin = df.columns.get_loc("FINALIZACIÓN")
            df.insert(idx_fin, "MES FINALIZADO", mes_finalizado)
            df.insert(idx_fin, "AÑO FINALIZADO", anio_finalizado)

            df = df.drop_duplicates()
            lista_dataframes.append(df)

        except Exception as e:
            st.error(f"⚠️ Error procesando {nombre_archivo}: {e}")

    if not lista_dataframes:
        return False

    df_nuevos = pd.concat(lista_dataframes, ignore_index=True)

    # Fusión con la data ya existente sin perder el histórico
    if os.path.exists(ruta_excel_actual):
        try:
            df_existente = pd.read_excel(ruta_excel_actual, sheet_name="Consolidado Zonas", dtype=str)
            df_existente = df_existente.fillna("")

            for col in df_nuevos.columns:
                if col not in df_existente.columns:
                    df_existente[col] = ""
            for col in df_existente.columns:
                if col not in df_nuevos.columns:
                    df_nuevos[col] = ""

            df_existente["INCIDENCIA"] = df_existente["INCIDENCIA"].astype(str).str.replace(r'\.0$', '', regex=True)
            df_nuevos["INCIDENCIA"] = df_nuevos["INCIDENCIA"].astype(str).str.replace(r'\.0$', '', regex=True)

            df_existente.set_index("INCIDENCIA", inplace=True)
            df_nuevos.set_index("INCIDENCIA", inplace=True)

            df_existente.update(df_nuevos)
            df_final = pd.concat([df_nuevos[~df_nuevos.index.isin(df_existente.index)], df_existente])
            df_final.reset_index(inplace=True)
        except Exception:
            df_final = df_nuevos
    else:
        df_final = df_nuevos

    # Guardar manteniendo el formato profesional en Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Zonas"
    ws.views.sheetView[0].showGridLines = True

    headers = list(df_final.columns)
    ws.append(headers)

    for row in df_final.itertuples(index=False, name=None):
        ws.append(list(row))

    max_row = ws.max_row
    max_col = len(headers)
    end_cell = get_column_letter(max_col) + str(max_row)

    tab = Table(displayName="TablaIncidencias", ref=f"A1:{end_cell}")
    style = TableStyleInfo(
        name="TableStyleLight1", showFirstColumn=False, showLastColumn=False,
        showRowStripes=False, showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    data_font = Font(name="Calibri", size=10, color="333333")
    thin_border = Border(
        left=Side(style='thin', color='D5D8DC'), right=Side(style='thin', color='D5D8DC'),
        top=Side(style='thin', color='D5D8DC'), bottom=Side(style='thin', color='D5D8DC')
    )

    ws.row_dimensions[1].height = 25
    for row_num in range(2, max_row + 1):
        ws.row_dimensions[row_num].height = 20
        for col_num in range(1, max_col + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws.freeze_panes = 'A2'
    wb.save(ruta_excel_actual)
    return True

# =========================================================================
# CARGA DE DATOS PARA EL DASHBOARD
# =========================================================================
@st.cache_data(ttl=60)
def cargar_datos():
    if os.path.exists(EXCEL_DESTINO):
        df = pd.read_excel(EXCEL_DESTINO, sheet_name="Consolidado Zonas", dtype=str)
        return df.fillna("")
    return pd.DataFrame()

df_global = cargar_datos()

# =========================================================================
# INTERFAZ DE USUARIO (STREAMLIT)
# =========================================================================
st.title("📊 Dashboard de Gestión de Incidencias")

# Panel lateral para actualización y filtros
st.sidebar.header("⚙️ Actualización de Data")
archivos_cargados = st.sidebar.file_uploader(
    "Sube tus nuevos archivos TXT", 
    type=["txt"], 
    accept_multiple_files=True,
    help="Puedes seleccionar varios archivos de texto de las zonas para actualizar."
)

if st.sidebar.button("🔄 Procesar y Actualizar Data"):
    if archivos_cargados:
        with st.spinner("Procesando TXT, limpiando y fusionando con el histórico..."):
            exito = procesar_y_actualizar_incidencias(archivos_cargados)
            if exito:
                st.sidebar.success("¡Data actualizada correctamente!")
                st.cache_data.clear()
                df_global = cargar_datos()
                st.rerun()
            else:
                st.sidebar.error("No se pudo procesar la información de los archivos.")
    else:
        st.sidebar.warning("Por favor, selecciona al menos un archivo TXT.")

st.sidebar.markdown("---")
st.sidebar.header("🔍 Filtros de Visualización")

if not df_global.empty:
    # Filtros dinámicos en barra lateral
    zonas_disponibles = ["TODAS"] + sorted(df_global["ZONA"].unique().tolist())
    zona_seleccionada = st.sidebar.selectbox("Zona", zonas_disponibles)

    estados_disponibles = ["TODOS"] + sorted(df_global["ESTADO"].unique().tolist())
    estado_seleccionado = st.sidebar.selectbox("Estado", estados_disponibles)

    # Filtrado del DataFrame
    df_filtrado = df_global.copy()
    if zona_seleccionada != "TODAS":
        df_filtrado = df_filtrado[df_filtrado["ZONA"] == zona_seleccionada]
    if estado_seleccionado != "TODOS":
        df_filtrado = df_filtrado[df_filtrado["ESTADO"] == estado_seleccionado]

    # Métricas principales (KPIs)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Incidencias", len(df_filtrado))
    with col2:
        pendientes = len(df_filtrado[df_filtrado["ESTADO"].str.contains("PENDIENTE", case=False, na=False)])
        st.metric("Pendientes", pendientes)
    with col3:
        finalizadas = len(df_filtrado[df_filtrado["ESTADO"].str.contains("FINALIZAD|CERRAD", case=False, na=False)])
        st.metric("Finalizadas / Cerradas", finalizadas)
    with col4:
        st.metric("Zonas Activas", df_filtrado["ZONA"].nunique())

    st.markdown("---")

    # Gráficos interactivos
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("Incidencias por Zona")
        if "ZONA" in df_filtrado.columns and not df_filtrado.empty:
            df_zonas = df_filtrado["ZONA"].value_counts().reset_index()
            df_zonas.columns = ["ZONA", "CANTIDAD"]
            fig_zona = px.bar(df_zonas, x="ZONA", y="CANTIDAD", color="ZONA", text="CANTIDAD")
            st.plotly_chart(fig_zona, use_container_width=True)

    with c2:
        st.subheader("Distribución por Estado")
        if "ESTADO" in df_filtrado.columns and not df_filtrado.empty:
            df_estado = df_filtrado["ESTADO"].value_counts().reset_index()
            df_estado.columns = ["ESTADO", "CANTIDAD"]
            fig_estado = px.pie(df_estado, names="ESTADO", values="CANTIDAD", hole=0.4)
            st.plotly_chart(fig_estado, use_container_width=True)

    # Tabla de datos
    st.subheader("📋 Detalle de Incidencias Filtradas")
    st.dataframe(df_filtrado, use_container_width=True)

    # Botón para descargar el excel actualizado
    with open(EXCEL_DESTINO, "rb") as file:
        st.download_button(
            label="📥 Descargar Excel Consolidado Actualizado",
            data=file,
            file_name="resultado_unificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👋 No hay data cargada todavía. Sube tus archivos TXT en la barra lateral izquierda y haz clic en **'Procesar y Actualizar Data'** para comenzar.")
