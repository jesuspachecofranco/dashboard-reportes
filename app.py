import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Reportes 2026", page_icon="📊", layout="wide"
)

st.title("📊 Control y Seguimiento de Incidencias")


# ==========================================
# FUNCIÓN DE DEPURACIÓN Y UNIFICACIÓN DE TXT (CON PROTECCIÓN)
# ==========================================
def procesar_txts_seguro(archivos_subidos):
    """Procesa los TXT subidos, genera un archivo temporal de actualización

    y coteja con el histórico creando respaldos sin sobrescribir imprudentemente.
    """
    mapeo_zonas = {
        "01": "NORTE",
        "02": "SUR",
        "03": "ESTE",
        "04": "OESTE",
        "05": "CENTRO",
        "06": "VÍA DUACA",
        "07": "VÍA RÍO CLARO",
        "08": "VÍA PAVIA",
        "09": "VÍA BUENA VISTA",
        "10": "VIA VIEJA CARORA",
        "11": "VÍA QUIBOR",
        "12": "VÍA AUTOPISTA CARORA",
    }

    meses_es = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }

    indices_a_conservar = [
        13,
        14,
        15,
        16,
        17,
        18,
        19,
        20,
        21,
        22,
        23,
        25,
        26,
        28,
        29,
        30,
        31,
        32,
        34,
        35,
        36,
        37,
        38,
    ]
    nombres_finales = [
        "INCIDENCIA",
        "DIRECCIÓN",
        "MOTIVO",
        "RECEPCIÓN",
        "FINALIZACIÓN",
        "TEC",
        "TCR",
        "TDV",
        "TEJ",
        "TAR",
        "ESTADO",
        "CLIENTE",
        "TOTAL INCIDENCIAS",
        "SUMATORIAS TEC",
        "SUMATORIAS TCR",
        "SUMATORIAS TDV",
        "SUMATORIAS TEJ",
        "SUMATORIAS TAR",
        "PROMEDIO TEC",
        "PROMEDIO TCR",
        "PROMEDIO TDV",
        "PROMEDIO TEJ",
        "PROMEDIO TAR",
    ]

    lista_dataframes = []

    for archivo in archivos_subidos:
        nombre_archivo = archivo.name
        try:
            contenido = archivo.getvalue().decode("latin-1")
            df = pd.read_csv(
                io.StringIO(contenido),
                sep="|",
                header=None,
                dtype=str,
                engine="python",
                on_bad_lines="skip",
            )

            # Validar si el archivo tiene suficientes columnas para procesar (si le faltan, se asume sin reportes nuevos)
            total_columnas_archivo = df.shape[1]
            max_indice_requerido = max(indices_a_conservar)

            if total_columnas_archivo <= max_indice_requerido:
                st.info(
                    f"ℹ️ El archivo '{nombre_archivo}' no contiene reportes"
                    " nuevos o carece de la estructura completa (zona sin"
                    " incidencias). Se omitirá sin afectar el proceso."
                )
                continue

            df = df.iloc[:, indices_a_conservar]
            df.columns = nombres_finales

            codigo = nombre_archivo[:2]
            nombre_zona = mapeo_zonas.get(codigo, nombre_archivo)
            df.insert(0, "ZONA", nombre_zona)

            for col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
                df[col] = df[col].replace(
                    ["NAN", "NONE", "NAT", "NAN.0"], "", regex=False
                )
                if col == "INCIDENCIA":
                    df[col] = df[col].str.replace(r"\.0$", "", regex=True)

            df = df.replace(r"\*", "", regex=True)
            df = df[df["INCIDENCIA"] != ""]
            df = df[df["INCIDENCIA"].str.match(r"^\d+$", na=False)]
            df = df[df["RECEPCIÓN"] != ""]

            # Si después de limpiar no quedan filas válidas, saltamos al siguiente
            if df.empty:
                continue

            dt_recepcion = pd.to_datetime(
                df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
            )
            df["_DT_RECEPCION_TEMP"] = dt_recepcion
            df = df.dropna(subset=["_DT_RECEPCION_TEMP"])
            df = df.drop(columns=["_DT_RECEPCION_TEMP"])

            dt_recepcion = pd.to_datetime(
                df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
            )
            mes_recibido = dt_recepcion.dt.month.map(meses_es).fillna("")
            anio_recibido = (
                dt_recepcion.dt.year.fillna("")
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
            )

            dt_finalizacion = pd.to_datetime(
                df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
            )
            mes_finalizado = dt_finalizacion.dt.month.map(meses_es).fillna("")
            anio_finalizado = (
                dt_finalizacion.dt.year.fillna("")
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
            )

            idx_rec = df.columns.get_loc("RECEPCIÓN")
            df.insert(idx_rec, "MES RECIBIDO", mes_recibido)
            df.insert(idx_rec, "AÑO RECIBIDO", anio_recibido)

            idx_fin = df.columns.get_loc("FINALIZACIÓN")
            df.insert(idx_fin, "MES FINALIZADO", mes_finalizado)
            df.insert(idx_fin, "AÑO FINALIZADO", anio_finalizado)

            df = df.drop_duplicates()
            lista_dataframes.append(df)
        except Exception as e:
            st.error(
                f"⚠️ Error procesando el archivo '{nombre_archivo}': {e}"
            )
            return False
    if lista_dataframes:
        df_nuevo = pd.concat(lista_dataframes, ignore_index=True)

        archivo_actualizacion = "resultado_actualizacion.xlsx"
        df_nuevo.to_excel(archivo_actualizacion, index=False)

        if os.path.exists("resultado_unificado.xlsx"):
            try:
                timestamp_respaldo = datetime.now().strftime("%Y%m%d_%H%M%S")
                archivo_respaldo = (
                    f"resultado_unificado_respaldo_{timestamp_respaldo}.xlsx"
                )
                df_existente_seguridad = pd.read_excel(
                    "resultado_unificado.xlsx", dtype=str
                )
                df_existente_seguridad.to_excel(archivo_respaldo, index=False)

                df_final = pd.concat(
                    [df_existente_seguridad, df_nuevo], ignore_index=True
                )
                df_final = df_final.drop_duplicates(
                    subset=["INCIDENCIA"], keep="last"
                )
            except Exception:
                df_final = df_nuevo
        else:
            df_final = df_nuevo

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
            name="TableStyleLight1",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False,
        )
        tab.tableStyleInfo = style
        ws.add_table(tab)

        data_font = Font(name="Calibri", size=10, color="333333")
        thin_border = Border(
            left=Side(style="thin", color="D5D8DC"),
            right=Side(style="thin", color="D5D8DC"),
            top=Side(style="thin", color="D5D8DC"),
            bottom=Side(style="thin", color="D5D8DC"),
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

        ws.freeze_panes = "A2"
        wb.save("resultado_unificado.xlsx")
        return True
    return False

# ==========================================
# PANEL LATERAL: GESTIÓN Y ACTUALIZACIÓN
# ==========================================
st.sidebar.header("⚙️ Configuración y Datos")

# Inicializamos la clave del uploader en session_state para poder limpiarlo después de procesar
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# Limitamos estrictamente el uploader a archivos .txt
archivos_cargados = st.sidebar.file_uploader(
    "Sube los archivos .txt de las zonas",
    type=["txt", "TXT"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}",
)

# Si hay archivos cargados, mostramos el botón de ejecución de inmediato
if archivos_cargados:
    st.sidebar.info(
        f"📁 Se detectaron {len(archivos_cargados)} archivo(s) .txt listos."
    )

    if st.sidebar.button(
        "🚀 Ejecutar Depuración y Actualizar",
        key="btn_ejecutar_txt",
        use_container_width=True,
    ):
        with st.spinner(
            "Depurando, cotejando y generando respaldo seguro..."
        ):
            exito = procesar_txts_seguro(archivos_cargados)
            if exito:
                st.cache_data.clear()
                st.sidebar.success(
                    "¡Actualización exitosa! Se guardó"
                    " 'resultado_actualizacion.xlsx'."
                )
                # Incrementamos la clave para limpiar el uploader y recargar la vista
                st.session_state["uploader_key"] += 1
                st.rerun()

# Botón permanente justo abajo (después del bloque de ejecución) para descargar el archivo de actualización si existe
if os.path.exists("resultado_actualizacion.xlsx"):
    st.sidebar.markdown("---")
    with open("resultado_actualizacion.xlsx", "rb") as f:
        st.sidebar.download_button(
            "📥 Descargar Archivo de Actualización",
            f,
            file_name="resultado_actualizacion.xlsx",
            use_container_width=True,
        )
# ==========================================
# CARGA Y VISUALIZACIÓN DE DATOS (MANTIENE DISEÑO)
# ==========================================
@st.cache_data
def cargar_datos():
    archivo_entrada = "resultado_unificado.xlsx"
    if not os.path.exists(archivo_entrada):
        return None

    df = pd.read_excel(archivo_entrada, dtype=str)
    df["RECEPCION_DT"] = pd.to_datetime(
        df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )
    df["FINALIZACION_DT"] = pd.to_datetime(
        df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )
    return df


df = cargar_datos()

if df is None:
    st.error(
        "❌ No se encontró el archivo 'resultado_unificado.xlsx'. Sube tus"
        " archivos .txt en el panel lateral para comenzar."
    )
else:
    columnas_mostrar = [
        "INCIDENCIA",
        "RECEPCIÓN",
        "FINALIZACIÓN",
        "DIRECCIÓN",
        "CLIENTE",
        "MOTIVO",
        "ESTADO",
    ]

    tab1, tab2, tab3 = st.tabs([
        "📅 Seguimiento Diario",
        "📈 Seguimiento Anual",
        "🔍 Búsqueda por Fecha",
    ])

    # ==========================================
    # PESTAÑA 1: SEGUIMIENTO DIARIO
    # ==========================================
    with tab1:
        st.subheader("📅 Comportamiento Diario por Mes")
        col_s1, col_s2, _ = st.columns([1, 1, 2])
        with col_s1:
            anio_seleccionado = st.selectbox(
                "Seleccione el Año:", [2025, 2026, 2027], index=1
            )
        with col_s2:
            meses_dict = {
                1: "Enero",
                2: "Febrero",
                3: "Marzo",
                4: "Abril",
                5: "Mayo",
                6: "Junio",
                7: "Julio",
                8: "Agosto",
                9: "Septiembre",
                10: "Octubre",
                11: "Noviembre",
                12: "Diciembre",
            }
            mes_nombre_seleccionado = st.selectbox(
                "Seleccione el Mes:",
                list(meses_dict.values()),
                index=7,
            )
            mes_seleccionado = [
                k
                for k, v in meses_dict.items()
                if v == mes_nombre_seleccionado
            ][0]

        inicio_mes_dinamico = pd.Timestamp(
            year=anio_seleccionado, month=mes_seleccionado, day=1
        )
        ultimo_dia = calendar.monthrange(anio_seleccionado, mes_seleccionado)[1]
        fin_mes_dinamico = pd.Timestamp(
            year=anio_seleccionado,
            month=mes_seleccionado,
            day=ultimo_dia,
        ) + pd.Timedelta(days=1)

        acumulado_mes = int(
            (
                (df["RECEPCION_DT"] < inicio_mes_dinamico)
                & (
                    df["FINALIZACION_DT"].isna()
                    | (df["FINALIZACION_DT"] >= inicio_mes_dinamico)
                )
            ).sum()
        )
        recibidos_mes = int(
            (
                (df["RECEPCION_DT"] >= inicio_mes_dinamico)
                & (df["RECEPCION_DT"] < fin_mes_dinamico)
            ).sum()
        )
        total_mes = acumulado_mes + recibidos_mes
        finalizados_mes = int(
            (
                (df["FINALIZACION_DT"] >= inicio_mes_dinamico)
                & (df["FINALIZACION_DT"] < fin_mes_dinamico)
            ).sum()
        )
        pendientes_mes = total_mes - finalizados_mes

        dias_mes = pd.date_range(
            start=inicio_mes_dinamico,
            end=pd.Timestamp(
                year=anio_seleccionado,
                month=mes_seleccionado,
                day=ultimo_dia,
            ),
            freq="D",
        )
        datos_diarios = []

        for dia in dias_mes:
            inicio_dia = dia
            fin_dia = dia + pd.Timedelta(days=1)
            nombre_dia = dia.strftime("%d/%m/%Y")

            cant_recibidos = int(
                (
                    (df["RECEPCION_DT"] >= inicio_dia)
                    & (df["RECEPCION_DT"] < fin_dia)
                ).sum()
            )
            cant_acumulada = int(
                (
                    (df["RECEPCION_DT"] < inicio_dia)
                    & (
                        df["FINALIZACION_DT"].isna()
                        | (df["FINALIZACION_DT"] >= inicio_dia)
                    )
                ).sum()
            )
            total_rep = cant_recibidos + cant_acumulada
            cant_finalizados = int(
                (
                    (df["FINALIZACION_DT"] >= inicio_dia)
                    & (df["FINALIZACION_DT"] < fin_dia)
                ).sum()
            )
            efectividad_dia = (
                (cant_finalizados / total_rep * 100) if total_rep > 0 else 0.0
            )

            datos_diarios.append({
                "FECHA": nombre_dia,
                "REPORTES RECIBIDOS": cant_recibidos,
                "REPORTES ACUMULADOS AL INICIAR": cant_acumulada,
                "TOTAL REPORTES": total_rep,
                "REPORTES FINALIZADOS": cant_finalizados,
                "EFECTIVIDAD (%)": round(efectividad_dia, 2),
            })

        df_dia = pd.DataFrame(datos_diarios)
        efectividad_promedio_mes = (
            df_dia["EFECTIVIDAD (%)"].mean() if not df_dia.empty else 0.0
        )

        st.markdown(
            f"##### 📌 Resumen del Mes de {mes_nombre_seleccionado}"
            f" {anio_seleccionado}"
        )
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("📦 Acumulado Inicial", f"{acumulado_mes:,}")
        c2.metric("📥 Recibidos", f"{recibidos_mes:,}")
        c3.metric("📊 Total Incidencias", f"{total_mes:,}")
        c4.metric("✅ Finalizados", f"{finalizados_mes:,}")
        c5.metric("⏳ Pendientes", f"{pendientes_mes:,}")
        c6.metric("🎯 Efectividad Prom.", f"{efectividad_promedio_mes:.1f}%")
        st.divider()

        fig1 = go.Figure()
        fig1.add_trace(
            go.Bar(
                x=df_dia["FECHA"],
                y=df_dia["REPORTES ACUMULADOS AL INICIAR"],
                name="Acumulados al Iniciar",
                marker_color="#d8e4fc",
                text=df_dia["REPORTES ACUMULADOS AL INICIAR"],
                textposition="inside",
                textfont=dict(color="black", size=11),
            )
        )
        fig1.add_trace(
            go.Bar(
                x=df_dia["FECHA"],
                y=df_dia["REPORTES RECIBIDOS"],
                name="Reportes Recibidos",
                marker_color="#e9f056",
                text=df_dia["REPORTES RECIBIDOS"],
                textposition="inside",
                textfont=dict(color="black", size=11),
            )
        )
        fig1.add_trace(
            go.Scatter(
                x=df_dia["FECHA"],
                y=df_dia["TOTAL REPORTES"],
                name="Total Reportes",
                mode="text+markers",
                text=df_dia["TOTAL REPORTES"],
                textposition="top center",
                textfont=dict(color="#1F4E78", size=12),
                marker=dict(size=8, color="rgba(0,0,0,0)"),
                showlegend=False,
            )
        )
        fig1.add_trace(
            go.Scatter(
                x=df_dia["FECHA"],
                y=df_dia["REPORTES FINALIZADOS"],
                name="Reportes Finalizados",
                mode="lines+markers+text",
                text=df_dia["REPORTES FINALIZADOS"],
                textposition="bottom center",
                textfont=dict(color="#1d4ed8", size=11),
                marker=dict(size=6, color="#1d4ed8"),
                line=dict(color="#1d4ed8", width=2),
            )
        )

        fig1.update_layout(
            title=dict(
                text=(
                    "<b>Flujo Diario de Reportes -"
                    f" {mes_nombre_seleccionado} {anio_seleccionado}</b>"
                ),
                font=dict(size=18, color="#1f4e78"),
            ),
            barmode="stack",
            xaxis_title="<b>Día del Mes</b>",
            yaxis_title="<b>Cantidad de Reportes</b>",
            hovermode="x unified",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1
            ),
        )
        st.plotly_chart(fig1, use_container_width=True)

        with st.expander("Ver tabla de datos diarios"):
            st.dataframe(df_dia, use_container_width=True)

    # ==========================================
    # PESTAÑA 2: SEGUIMIENTO ANUAL
    # ==========================================
    with tab2:
        st.subheader("📈 Seguimiento Mensual")
        inicio_anio = pd.Timestamp("2026-01-01")
        fin_anio = pd.Timestamp("2026-09-01")

        acumulado_anio = int(
            (
                (df["RECEPCION_DT"] < inicio_anio)
                & (
                    df["FINALIZACION_DT"].isna()
                    | (df["FINALIZACION_DT"] >= inicio_anio)
                )
            ).sum()
        )
        recibidos_anio = int(
            (
                (df["RECEPCION_DT"] >= inicio_anio)
                & (df["RECEPCION_DT"] < fin_anio)
            ).sum()
        )
        total_anio = acumulado_anio + recibidos_anio
        finalizados_anio = int(
            (
                (df["FINALIZACION_DT"] >= inicio_anio)
                & (df["FINALIZACION_DT"] < fin_anio)
            ).sum()
        )
        pendientes_anio = total_anio - finalizados_anio
        efectividad_global_anio = (
            (finalizados_anio / total_anio * 100) if total_anio > 0 else 0.0
        )

        st.markdown("##### 📌 Resumen Acumulado Anual (Enero - Agosto)")
        ac1, ac2, ac3, ac4, ac5, ac6 = st.columns(6)
        ac1.metric("📦 Acumulado Inicial", f"{acumulado_anio:,}")
        ac2.metric("📥 Recibidos", f"{recibidos_anio:,}")
        ac3.metric("📊 Total Incidencias", f"{total_anio:,}")
        ac4.metric("✅ Finalizados", f"{finalizados_anio:,}")
        ac5.metric("⏳ Pendientes", f"{pendientes_anio:,}")
        ac6.metric("🎯 Efectividad Global", f"{efectividad_global_anio:.1f}%")
        st.divider()

        meses_2026_hasta_agosto = [
            (1, "ENERO"),
            (2, "FEBRERO"),
            (3, "MARZO"),
            (4, "ABRIL"),
            (5, "MAYO"),
            (6, "JUNIO"),
            (7, "JULIO"),
            (8, "AGOSTO"),
        ]
        datos_anual = []

        for num_mes, nombre_mes in meses_2026_hasta_agosto:
            inicio_mes = pd.Timestamp(year=2026, month=num_mes, day=1)
            fin_mes = (
                pd.Timestamp(year=2027, month=1, day=1)
                if num_mes == 12
                else pd.Timestamp(year=2026, month=num_mes + 1, day=1)
            )

            cant_recibidos = int(
                (
                    (df["RECEPCION_DT"] >= inicio_mes)
                    & (df["RECEPCION_DT"] < fin_mes)
                ).sum()
            )
            cant_acumulada = int(
                (
                    (df["RECEPCION_DT"] < inicio_mes)
                    & (
                        df["FINALIZACION_DT"].isna()
                        | (df["FINALIZACION_DT"] >= inicio_mes)
                    )
                ).sum()
            )
            total_rep = cant_recibidos + cant_acumulada
            cant_finalizados = int(
                (
                    (df["FINALIZACION_DT"] >= inicio_mes)
                    & (df["FINALIZACION_DT"] < fin_mes)
                ).sum()
            )

            datos_anual.append({
                "MES": nombre_mes,
                "AÑO": 2026,
                "REPORTES RECIBIDOS": cant_recibidos,
                "REPORTES ACUMULADOS AL INICIAR": cant_acumulada,
                "TOTAL REPORTES": total_rep,
                "REPORTES FINALIZADOS": cant_finalizados,
            })

        df_anual = pd.DataFrame(datos_anual)

        fig2 = go.Figure()
        fig2.add_trace(
            go.Bar(
                x=df_anual["MES"],
                y=df_anual["REPORTES ACUMULADOS AL INICIAR"],
                name="Acumulados al Iniciar",
                marker_color="#d8e4fc",
                text=df_anual["REPORTES ACUMULADOS AL INICIAR"],
                textposition="inside",
                textfont=dict(color="black", size=14),
            )
        )
        fig2.add_trace(
            go.Bar(
                x=df_anual["MES"],
                y=df_anual["REPORTES RECIBIDOS"],
                name="Reportes Recibidos",
                marker_color="#e9f056",
                text=df_anual["REPORTES RECIBIDOS"],
                textposition="inside",
                textfont=dict(color="black", size=14),
            )
        )
        fig2.add_trace(
            go.Scatter(
                x=df_anual["MES"],
                y=df_anual["TOTAL REPORTES"],
                name="Total Reportes",
                mode="text+markers",
                text=df_anual["TOTAL REPORTES"],
                textposition="top center",
                textfont=dict(color="#1F4E78", size=14),
                marker=dict(size=8, color="rgba(0,0,0,0)"),
                showlegend=False,
            )
        )
        fig2.add_trace(
            go.Scatter(
                x=df_anual["MES"],
                y=df_anual["REPORTES FINALIZADOS"],
                name="Reportes Finalizados",
                mode="lines+markers+text",
                text=df_anual["REPORTES FINALIZADOS"],
                textposition="bottom center",
                textfont=dict(color="#1d4ed8", size=14),
                marker=dict(size=6, color="#1d4ed8"),
                line=dict(color="#1d4ed8", width=2),
            )
        )

        fig2.update_layout(
            title=dict(
                text=(
                    "<b>Resumen Acumulado Mensual (Enero - Agosto 2026)</b>"
                ),
                font=dict(size=18, color="#1F4E78"),
            ),
            barmode="stack",
            xaxis_title="<b>Mes</b>",
            yaxis_title="<b>Cantidad de Reportes</b>",
            hovermode="x unified",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1
            ),
        )
        st.plotly_chart(fig2, use_container_width=True)

        with st.expander("Ver tabla de datos mensuales"):
            st.dataframe(df_anual, use_container_width=True)

    # ==========================================
    # PESTAÑA 3: BÚSQUEDA POR FECHA
    # ==========================================
    with tab3:
        st.subheader("🔍 Buscador de Incidencias por Fecha")
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            fecha_busqueda = st.date_input(
                "Seleccione la fecha a auditar:", value=datetime(2026, 8, 12)
            )

        inicio_sel = pd.Timestamp(fecha_busqueda)
        fin_sel = inicio_sel + pd.Timedelta(days=1)

        df_acumulados_dia = df[
            (df["RECEPCION_DT"] < inicio_sel)
            & (
                df["FINALIZACION_DT"].isna()
                | (df["FINALIZACION_DT"] >= inicio_sel)
            )
        ]
        df_recibidos_dia = df[
            (df["RECEPCION_DT"] >= inicio_sel) & (df["RECEPCION_DT"] < fin_sel)
        ]
        df_finalizados_dia = df[
            (df["FINALIZACION_DT"] >= inicio_sel)
            & (df["FINALIZACION_DT"] < fin_sel)
        ]

        st.markdown("---")
        st.markdown(
            f"##### 📊 Resumen para el día: {fecha_busqueda.strftime('%d/%m/%Y')}"
        )

        total_activos_dia = len(df_acumulados_dia) + len(df_recibidos_dia)
        total_finalizados_dia = len(df_finalizados_dia)
        efectividad_dia_busqueda = (
            (total_finalizados_dia / total_activos_dia * 100)
            if total_activos_dia > 0
            else 0.0
        )
        pendientes_dia_busqueda = total_activos_dia - total_finalizados_dia

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("📦 Acumulados Previos", f"{len(df_acumulados_dia):,}")
        m2.metric("📥 Recibidos", f"{len(df_recibidos_dia):,}")
        m3.metric("📊 Total Activos", f"{total_activos_dia:,}")
        m4.metric("✅ Finalizados", f"{total_finalizados_dia:,}")
        m5.metric("⏳ Pendientes", f"{pendientes_dia_busqueda:,}")
        m6.metric("🎯 Efectividad", f"{efectividad_dia_busqueda:.1f}%")
        st.markdown("---")

        with st.expander(
            f"📦 1. Incidencias Acumuladas Pendientes ({len(df_acumulados_dia)}"
            " registros)",
            expanded=True,
        ):
            if not df_acumulados_dia.empty:
                st.dataframe(
                    df_acumulados_dia[columnas_mostrar],
                    use_container_width=True,
                )
            else:
                st.info("No hay registros acumulados pendientes para esta fecha.")

        with st.expander(
            f"📥 2. Incidencias Recibidas ({len(df_recibidos_dia)} registros)"
        ):
            if not df_recibidos_dia.empty:
                st.dataframe(
                    df_recibidos_dia[columnas_mostrar],
                    use_container_width=True,
                )
            else:
                st.info(
                    "No se registraron incidencias recibidas en esta fecha."
                )

        with st.expander(
            f"✅ 3. Incidencias Finalizadas / Atendidas"
            f" ({len(df_finalizados_dia)} registros)"
        ):
            if not df_finalizados_dia.empty:
                st.dataframe(
                    df_finalizados_dia[columnas_mostrar],
                    use_container_width=True,
                )
            else:
                st.info(
                    "No hay incidencias finalizadas registradas en esta fecha."
                )
