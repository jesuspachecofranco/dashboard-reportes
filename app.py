import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import streamlit as st

# Configuración de la página con ancho completo
st.set_page_config(
    page_title="Dashboard de Reportes 2026", 
    page_icon="📊", 
    layout="wide"
)

st.title("📊 Control y Seguimiento de Incidencias")

# Inicializar variables de estado si no existen
if "lote_seleccionado" not in st.session_state:
    st.session_state["lote_seleccionado"] = pd.DataFrame()

if "lotes_guardados" not in st.session_state:
    st.session_state["lotes_guardados"] = {}


# ==========================================
# 1. FUNCIÓN DE PROCESAMIENTO DE TXT ROBUSTA
# ==========================================
def procesar_txts_seguro(archivos_subidos):
    mapeo_zonas = {
        "01": "NORTE", "02": "SUR", "03": "ESTE", "04": "OESTE",
        "05": "CENTRO", "06": "VÍA DUACA", "07": "VÍA RÍO CLARO",
        "08": "VÍA PAVIA", "09": "VÍA BUENA VISTA", "10": "VIA VIEJA CARORA",
        "11": "VÍA QUIBOR", "12": "VÍA AUTOPISTA CARORA",
    }

    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }

    indices_a_conservar = [
        13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26, 
        28, 29, 30, 31, 32, 34, 35, 36, 37, 38
    ]
    
    nombres_finales = [
        "INCIDENCIA", "DIRECCIÓN", "MOTIVO", "RECEPCIÓN", "FINALIZACIÓN",
        "TEC", "TCR", "TDV", "TEJ", "TAR", "ESTADO", "CLIENTE",
        "TOTAL INCIDENCIAS", "SUMATORIAS TEC", "SUMATORIAS TCR",
        "SUMATORIAS TDV", "SUMATORIAS TEJ", "SUMATORIAS TAR",
        "PROMEDIO TEC", "PROMEDIO TCR", "PROMEDIO TDV", "PROMEDIO TEJ", "PROMEDIO TAR"
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

            if df.shape[1] <= max(indices_a_conservar):
                continue

            df = df.iloc[:, indices_a_conservar]
            df.columns = nombres_finales

            codigo = nombre_archivo[:2]
            nombre_zona = mapeo_zonas.get(codigo, nombre_archivo)
            df.insert(0, "ZONA", nombre_zona)

            for col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
                df[col] = df[col].replace(["NAN", "NONE", "NAT", "NAN.0"], "", regex=False)
                if col == "INCIDENCIA":
                    df[col] = df[col].str.replace(r"\.0$", "", regex=True)

            df = df.replace(r"\*", "", regex=True)
            df = df[df["INCIDENCIA"] != ""]
            df = df[df["INCIDENCIA"].str.match(r"^\d+$", na=False)]
            df = df[df["RECEPCIÓN"] != ""]

            if df.empty:
                continue

            lista_dataframes.append(df)
        except Exception as e:
            st.error(f"⚠️ Error procesando el archivo '{nombre_archivo}': {e}")
            return False

    if lista_dataframes:
        df_nuevo = pd.concat(lista_dataframes, ignore_index=True)
        
        # Unificar con archivo existente si lo hay
        if os.path.exists("resultado_unificado.xlsx"):
            try:
                df_existente = pd.read_excel("resultado_unificado.xlsx", dtype=str)
                df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
                df_final = df_final.drop_duplicates(subset=["INCIDENCIA"], keep="last")
            except Exception:
                df_final = df_nuevo
        else:
            df_final = df_nuevo

        # Guardar archivo consolidado maestro
        df_final.to_excel("resultado_unificado.xlsx", index=False)
        return True
    return False


# ==========================================
# 2. CARGA Y CONVERSIÓN CRÍTICA DE FECHAS
# ==========================================
@st.cache_data
def cargar_datos():
    archivo_entrada = "resultado_unificado.xlsx"
    if not os.path.exists(archivo_entrada):
        return None

    df = pd.read_excel(archivo_entrada, dtype=str)

    # Limpieza exhaustiva de espacios en blanco en columnas de fecha
    for col in ["RECEPCIÓN", "FINALIZACIÓN"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(["NAN", "NONE", "NAT", ""], pd.NA)

    # Conversión estricta a Datetime con formato explícito de 2 dígitos en hora/minuto
    df["RECEPCION_DT"] = pd.to_datetime(
        df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )
    df["FINALIZACION_DT"] = pd.to_datetime(
        df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )

    # Columnas auxiliares de meses y años basados en la recepción
    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }
    
    df["MES RECIBIDO"] = df["RECEPCION_DT"].dt.month.map(meses_es).fillna("")
    df["AÑO RECIBIDO"] = df["RECEPCION_DT"].dt.year.fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

    return df


df = cargar_datos()

if df is None:
    st.warning("⚠️ No se encontró el archivo 'resultado_unificado.xlsx'. Ve a la pestaña **📦 Actualización** para subir tus archivos .txt.")

columnas_mostrar = ["INCIDENCIA", "RECEPCIÓN", "FINALIZACIÓN", "DIRECCIÓN", "CLIENTE", "MOTIVO", "ESTADO"]


# ==========================================
# 3. PESTAÑAS PRINCIPALES (TABS)
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📅 Seguimiento Diario",
    "📈 Seguimiento Anual",
    "🔍 Búsqueda Avanzada",
    "🎯 Análisis",
    "📄 Reportes",
    "📦 Actualización",
])


def renderizar_tabla_con_seleccion(df_sub, nombre_pestana, tipo_origen_etiqueta=None):
    if df is None:
        st.info("Primero debes cargar datos.")
        return
    if df_sub.empty:
        st.info("No hay registros para mostrar en esta sección.")
        return

    df_editable = df_sub[columnas_mostrar].copy()
    df_editable.insert(0, "Seleccionar", False)

    df_editado = st.data_editor(
        df_editable,
        use_container_width=True,
        key=f"editor_{nombre_pestana}",
        hide_index=True,
    )

    col_btn1, _ = st.columns([2, 4])
    with col_btn1:
        if st.button(f"➕ Agregar al Lote de Análisis", key=f"btn_agregar_{nombre_pestana}"):
            seleccionadas = df_editado[df_editado["Seleccionar"] == True]
            if not seleccionadas.empty:
                ids_seleccionados = seleccionadas["INCIDENCIA"].tolist()
                df_lote_nuevo = df[df["INCIDENCIA"].isin(ids_seleccionados)].copy()
                df_lote_nuevo["ORIGEN_TIPO"] = tipo_origen_etiqueta or "Recibidas en el Día"

                if "UNIDAD" not in df_lote_nuevo.columns:
                    df_lote_nuevo["UNIDAD"] = ""
                if "GERENCIA" not in df_lote_nuevo.columns:
                    df_lote_nuevo["GERENCIA"] = ""

                if not st.session_state["lote_seleccionado"].empty:
                    st.session_state["lote_seleccionado"] = pd.concat([
                        st.session_state["lote_seleccionado"],
                        df_lote_nuevo,
                    ]).drop_duplicates(subset=["INCIDENCIA"])
                else:
                    st.session_state["lote_seleccionado"] = df_lote_nuevo

                st.success(f"¡Se agregaron {len(ids_seleccionados)} incidencias al Lote de Análisis!")
            else:
                st.warning("⚠️ No has seleccionado ninguna incidencia.")


def renderizar_panel_graficos(df_lote_objetivo, sufijo=""):
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        if "ZONA" in df_lote_objetivo.columns and not df_lote_objetivo["ZONA"].isna().all():
            fig_zona = px.pie(df_lote_objetivo, names="ZONA", title="Distribución por Zonas", hole=0.4)
            st.plotly_chart(fig_zona, use_container_width=True, key=f"fig_zona_{sufijo}")
    with col_g2:
        if "MOTIVO" in df_lote_objetivo.columns:
            fig_motivo = px.bar(
                df_lote_objetivo["MOTIVO"].value_counts().reset_index(),
                x="MOTIVO", y="count",
                title="Incidencias por Motivo",
            )
            st.plotly_chart(fig_motivo, use_container_width=True, key=f"fig_motivo_{sufijo}")


# --- PESTAÑA 1: SEGUIMIENTO DIARIO ---
with tab1:
    st.subheader("📅 Comportamiento Diario por Mes")
    if df is not None:
        col_s1, col_s2, _ = st.columns([1, 1, 2])
        with col_s1:
            anio_seleccionado = st.selectbox("Seleccione el Año:", [2025, 2026, 2027], index=1, key="anio_s1")
        with col_s2:
            meses_dict = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
            }
            mes_nombre_seleccionado = st.selectbox("Seleccione el Mes:", list(meses_dict.values()), index=7, key="mes_s1")
            mes_seleccionado = [k for k, v in meses_dict.items() if v == mes_nombre_seleccionado][0]

        inicio_mes_dinamico = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=1)
        ultimo_dia = calendar.monthrange(anio_seleccionado, mes_seleccionado)[1]
        fin_mes_dinamico = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=ultimo_dia) + pd.Timedelta(days=1)

        # CÁLCULO ESTRICTO DE ACUMULADO INICIAL DEL MES:
        # Recepción antes de iniciar el mes Y (Finalización nula O finalización posterior/igual al inicio del mes)
        acumulado_mes = int((
            (df["RECEPCION_DT"] < inicio_mes_dinamico) & 
            (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= inicio_mes_dinamico))
        ).sum())

        recibidos_mes = int(((df["RECEPCION_DT"] >= inicio_mes_dinamico) & (df["RECEPCION_DT"] < fin_mes_dinamico)).sum())
        total_mes = acumulado_mes + recibidos_mes
        finalizados_mes = int(((df["FINALIZACION_DT"] >= inicio_mes_dinamico) & (df["FINALIZACION_DT"] < fin_mes_dinamico)).sum())
        pendientes_mes = total_mes - finalizados_mes

        dias_mes = pd.date_range(start=inicio_mes_dinamico, end=pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=ultimo_dia), freq="D")
        datos_diarios = []

        for dia in dias_mes:
            inicio_dia = dia
            fin_dia = dia + pd.Timedelta(days=1)
            
            cant_recibidos = int(((df["RECEPCION_DT"] >= inicio_dia) & (df["RECEPCION_DT"] < fin_dia)).sum())
            cant_acumulada = int(((df["RECEPCION_DT"] < inicio_dia) & (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= inicio_dia))).sum())
            total_rep = cant_recibidos + cant_acumulada
            cant_finalizados = int(((df["FINALIZACION_DT"] >= inicio_dia) & (df["FINALIZACION_DT"] < fin_dia)).sum())
            efectividad_dia = (cant_finalizados / total_rep * 100) if total_rep > 0 else 0.0

            datos_diarios.append({
                "FECHA": dia.strftime("%d/%m/%Y"),
                "REPORTES RECIBIDOS": cant_recibidos,
                "REPORTES ACUMULADOS AL INICIAR": cant_acumulada,
                "TOTAL REPORTES": total_rep,
                "REPORTES FINALIZADOS": cant_finalizados,
                "EFECTIVIDAD (%)": round(efectividad_dia, 2),
            })

        df_dia = pd.DataFrame(datos_diarios)
        efectividad_promedio_mes = df_dia["EFECTIVIDAD (%)"].mean() if not df_dia.empty else 0.0

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("📦 Acumulado Inicial", f"{acumulado_mes:,}")
        c2.metric("📥 Recibidos", f"{recibidos_mes:,}")
        c3.metric("📊 Total", f"{total_mes:,}")
        c4.metric("✅ Finalizados", f"{finalizados_mes:,}")
        c5.metric("⏳ Pendientes", f"{pendientes_mes:,}")
        c6.metric("🎯 Efectividad Prom.", f"{efectividad_promedio_mes:.1f}%")
        st.divider()

        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=df_dia["FECHA"], y=df_dia["REPORTES ACUMULADOS AL INICIAR"], name="Acumulados", marker_color="#d8e4fc"))
        fig1.add_trace(go.Bar(x=df_dia["FECHA"], y=df_dia["REPORTES RECIBIDOS"], name="Recibidos", marker_color="#e9f056"))
        fig1.update_layout(barmode="stack", title=f"Flujo Diario - {mes_nombre_seleccionado} {anio_seleccionado}")
        st.plotly_chart(fig1, use_container_width=True)

        with st.expander("Ver y seleccionar incidencias del mes"):
            df_mes_filtro = df[(df["RECEPCION_DT"] >= inicio_mes_dinamico) & (df["RECEPCION_DT"] < fin_mes_dinamico)].copy()
            renderizar_tabla_con_seleccion(df_mes_filtro, "diario_mes", "Recibidas en el Día")


# --- PESTAÑA 2: SEGUIMIENTO ANUAL ---
with tab2:
    st.subheader("📈 Seguimiento Mensual Acumulado (2026)")
    if df is not None:
        inicio_anio = pd.Timestamp("2026-01-01")
        fin_anio = pd.Timestamp("2026-09-01") # Hasta septiembre de 2026

        meses_2026 = [
            (1, "ENERO"), (2, "FEBRERO"), (3, "MARZO"), (4, "ABRIL"),
            (5, "MAYO"), (6, "JUNIO"), (7, "JULIO"), (8, "AGOSTO")
        ]
        datos_anual = []

        for num_mes, nombre_mes in meses_2026:
            inicio_mes = pd.Timestamp(year=2026, month=num_mes, day=1)
            fin_mes = pd.Timestamp(year=2026, month=num_mes + 1, day=1) if num_mes < 12 else pd.Timestamp(year=2027, month=1, day=1)

            cant_recibidos = int(((df["RECEPCION_DT"] >= inicio_mes) & (df["RECEPCION_DT"] < fin_mes)).sum())
            
            # FILTRO CORREGIDO Y VALIDADO PARA EL ACUMULADO MENSUAL (Evita el error de los 285 en agosto):
            cant_acumulada = int((
                (df["RECEPCION_DT"] < inicio_mes) & 
                (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= inicio_mes))
            ).sum())

            total_rep = cant_recibidos + cant_acumulada
            cant_finalizados = int(((df["FINALIZACION_DT"] >= inicio_mes) & (df["FINALIZACION_DT"] < fin_mes)).sum())

            datos_anual.append({
                "MES": nombre_mes,
                "REPORTES RECIBIDOS": cant_recibidos,
                "REPORTES ACUMULADOS AL INICIAR": cant_acumulada,
                "TOTAL REPORTES": total_rep,
                "REPORTES FINALIZADOS": cant_finalizados,
            })

        df_anual = pd.DataFrame(datos_anual)

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=df_anual["MES"], y=df_anual["REPORTES ACUMULADOS AL INICIAR"], name="Acumulados al Iniciar", marker_color="#d8e4fc"))
        fig2.add_trace(go.Bar(x=df_anual["MES"], y=df_anual["REPORTES RECIBIDOS"], name="Reportes Recibidos", marker_color="#e9f056"))
        fig2.update_layout(barmode="stack", title="<b>Resumen Acumulado Mensual 2026</b>")
        st.plotly_chart(fig2, use_container_width=True)

        with st.expander("Ver y seleccionar incidencias del acumulado anual"):
            df_anual_filtro = df[(df["RECEPCION_DT"] >= inicio_anio) & (df["RECEPCION_DT"] < fin_anio)].copy()
            renderizar_tabla_con_seleccion(df_anual_filtro, "anual", "Recibidas en el Día")


# --- PESTAÑA 3: BÚSQUEDA AVANZADA ---
with tab3:
    st.subheader("🔍 Buscador de Incidencias")
    if df is not None:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fecha_busqueda = st.date_input("Fecha a auditar:", value=datetime(2026, 8, 1))
        with col_f2:
            filtro_num = st.text_input("Nº de Incidencia:", placeholder="Ej. 12345")
        with col_f3:
            filtro_dir = st.text_input("Dirección:", placeholder="Ej. Av. Principal")

        inicio_sel = pd.Timestamp(fecha_busqueda)
        fin_sel = inicio_sel + pd.Timedelta(days=1)

        df_base_filtrada = df.copy()
        if filtro_num:
            df_base_filtrada = df_base_filtrada[df_base_filtrada["INCIDENCIA"].str.contains(filtro_num, case=False, na=False)]
        if filtro_dir:
            df_base_filtrada = df_base_filtrada[df_base_filtrada["DIRECCIÓN"].str.contains(filtro_dir, case=False, na=False)]

        df_acumulados_dia = df_base_filtrada[
            (df_base_filtrada["RECEPCION_DT"] < inicio_sel) & 
            (df_base_filtrada["FINALIZACION_DT"].isna() | (df_base_filtrada["FINALIZACION_DT"] >= inicio_sel))
        ].copy()

        df_recibidos_dia = df_base_filtrada[
            (df_base_filtrada["RECEPCION_DT"] >= inicio_sel) & 
            (df_base_filtrada["RECEPCION_DT"] < fin_sel)
        ].copy()

        st.markdown(f"##### 📊 Resultados para el: {fecha_busqueda.strftime('%d/%m/%Y')}")
        m1, m2 = st.columns(2)
        m1.metric("📦 Acumulados Previos Activos", f"{len(df_acumulados_dia):,}")
        m2.metric("📥 Recibidos en el Día", f"{len(df_recibidos_dia):,}")
        st.divider()

        st.markdown("##### 📦 Detalle Acumulados Previos")
        renderizar_tabla_con_seleccion(df_acumulados_dia, "buscador_acumulados", "Acumuladas")

        st.markdown("##### 📥 Detalle Recibidos en el Día")
        renderizar_tabla_con_seleccion(df_recibidos_dia, "buscador_recibidos", "Recibidas en el Día")


# --- PESTAÑA 4: ANÁLISIS ---
with tab4:
    st.subheader("🎯 Análisis y Gestión del Lote Seleccionado")
    if not st.session_state["lote_seleccionado"].empty:
        df_lote = st.session_state["lote_seleccionado"]
        st.write(f"Total en el lote: **{len(df_lote)}**")

        if "UNIDAD" not in df_lote.columns:
            df_lote["UNIDAD"] = ""
        if "GERENCIA" not in df_lote.columns:
            df_lote["GERENCIA"] = ""

        columnas_analisis = ["INCIDENCIA", "RECEPCIÓN", "FINALIZACIÓN", "DIRECCIÓN", "CLIENTE", "MOTIVO", "ESTADO", "UNIDAD", "GERENCIA"]
        df_lote_editado = st.data_editor(df_lote[columnas_analisis], use_container_width=True, key="editor_lote_analisis", hide_index=True)

        if st.button("💾 Guardar Cambios"):
            st.session_state["lote_seleccionado"].update(df_lote_editado)
            st.success("¡Guardado exitosamente!")

        st.markdown("---")
        renderizar_panel_graficos(st.session_state["lote_seleccionado"], sufijo="lote_actual")

        if st.button("🗑️ Vaciar Lote"):
            st.session_state["lote_seleccionado"] = pd.DataFrame()
            st.rerun()
    else:
        st.info("El lote de análisis está vacío. Selecciona registros desde las pestañas de seguimiento o búsqueda.")


# --- PESTAÑA 5: REPORTES ---
with tab5:
    st.subheader("📄 Generación de Reportes en Excel")
    if not st.session_state["lote_seleccionado"].empty:
        if st.button("📥 Descargar Reporte Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                st.session_state["lote_seleccionado"].to_excel(writer, index=False, sheet_name="Reporte")
            st.download_button("⬇️ Descargar", output.getvalue(), file_name="reporte.xlsx")
    else:
        st.info("No hay datos en el lote para exportar.")


# --- PESTAÑA 6: ACTUALIZACIÓN ---
with tab6:
    st.subheader("📦 Actualización de Datos (.txt)")
    archivos_txt = st.file_uploader("Sube tus archivos .txt:", type=["txt"], accept_multiple_files=True)
    if archivos_txt and st.button("🚀 Procesar"):
        if procesar_txts_seguro(archivos_txt):
            st.success("¡Actualizado correctamente!")
            st.cache_data.clear()
            st.rerun()
