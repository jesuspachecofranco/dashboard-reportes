import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from supabase import create_client, Client
import streamlit as st

# Configuración de la página con ancho completo
st.set_page_config(
    page_title="Dashboard de Reportes 2026", page_icon="📊", layout="wide"
)

# ==========================================
# CONEXIÓN A SUPABASE
# ==========================================
@st.cache_resource
def init_supabase():
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
        return create_client(url, key)
    except Exception as e:
        st.error(f"⚠️ Error conectando a Supabase. Verifica tus st.secrets: {e}")
        return None

supabase: Client = init_supabase()


# ==========================================
# SISTEMA DE AUTENTICACIÓN POR CONTRASEÑA
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False


def verificar_credenciales():
    usuario_ingresado = st.session_state.get("input_user", "")
    password_ingresada = st.session_state.get("input_pass", "")

    try:
        usuario_valido = st.secrets["auth"]["username"]
        password_valida = st.secrets["auth"]["password"]
    except Exception:
        usuario_valido = "admin"
        password_valida = "12345"

    if (
        usuario_ingresado == usuario_valido
        and password_ingresada == password_valida
    ):
        st.session_state["autenticado"] = True
    else:
        st.error("❌ Usuario o contraseña incorrectos.")


if not st.session_state["autenticado"]:
    st.title("🔐 Acceso Restringido")
    st.markdown("Por favor, introduce tus credenciales para acceder al Dashboard.")

    with st.form("form_login"):
        st.text_input("Usuario", key="input_user")
        st.text_input("Contraseña", type="password", key="input_pass")
        st.form_submit_button("Ingresar", on_click=verificar_credenciales)

    st.stop()


# ==========================================
# A PARTIR DE AQUÍ VA LA APLICACIÓN NORMAL
# ==========================================
st.title("📊 Control y Seguimiento de Incidencias")

with st.sidebar:
    st.write(f"Conectado como: **Administrador**")
    if st.button("🚪 Cerrar Sesión"):
        st.session_state["autenticado"] = False
        st.rerun()

if "lote_seleccionado" not in st.session_state:
    st.session_state["lote_seleccionado"] = pd.DataFrame()

if "lotes_guardados" not in st.session_state:
    st.session_state["lotes_guardados"] = {}


# ==========================================
# FUNCIÓN DE PROCESAMIENTO Y GUARDADO EN SUPABASE
# ==========================================
def procesar_txts_en_supabase(archivos_subidos):
    if not supabase:
        st.error("No hay conexión activa con Supabase.")
        return False

    mapeo_zonas = {
        "01": "NORTE", "02": "SUR", "03": "ESTE", "04": "OESTE",
        "05": "CENTRO", "06": "VÍA DUACA", "07": "VÍA RÍO CLARO",
        "08": "VÍA PAVIA", "09": "VÍA BUENA VISTA", "10": "VIA VIEJA CARORA",
        "11": "VÍA QUIBOR", "12": "VÍA AUTOPISTA CARORA",
    }
    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
        7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }
    indices_a_conservar = [
        13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26, 28, 29, 30, 31, 32, 34, 35, 36, 37, 38
    ]
    nombres_finales = [
        "INCIDENCIA", "DIRECCIÓN", "MOTIVO", "RECEPCIÓN", "FINALIZACIÓN", "TEC", "TCR", "TDV", "TEJ", "TAR", "ESTADO", "CLIENTE",
        "TOTAL INCIDENCIAS", "SUMATORIAS TEC", "SUMATORIAS TCR", "SUMATORIAS TDV", "SUMATORIAS TEJ", "SUMATORIAS TAR",
        "PROMEDIO TEC", "PROMEDIO TCR", "PROMEDIO TDV", "PROMEDIO TEJ", "PROMEDIO TAR"
    ]

    lista_dataframes = []

    for archivo in archivos_subidos:
        nombre_archivo = archivo.name
        try:
            contenido = archivo.getvalue().decode("latin-1")
            df = pd.read_csv(
                io.StringIO(contenido), sep="|", header=None, dtype=str, engine="python", on_bad_lines="skip"
            )
            if df.shape[1] <= max(indices_a_conservar):
                continue

            df = df.iloc[:, indices_a_conservar]
            df.columns = nombres_finales
            df.insert(0, "ZONA", mapeo_zonas.get(nombre_archivo[:2], nombre_archivo))

            for col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()
                df[col] = df[col].replace(["NAN", "NONE", "NAT", "NAN.0"], "", regex=False)
                if col == "INCIDENCIA":
                    df[col] = df[col].str.replace(r"\.0$", "", regex=True)

            df = df.replace(r"\*", "", regex=True)
            df = df[df["INCIDENCIA"] != ""]
            df = df[df["INCIDENCIA"].str.match(r"^\d+$", na=False)]
            df = df[df["RECEPCIÓN"] != ""]

            dt_recepcion = pd.to_datetime(df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce")
            df["RECEPCION_DT"] = dt_recepcion
            df = df.dropna(subset=["RECEPCION_DT"])

            mes_recibido = df["RECEPCION_DT"].dt.month.map(meses_es).fillna("")
            anio_recibido = df["RECEPCION_DT"].dt.year.fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

            dt_finalizacion = pd.to_datetime(df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce")
            df["FINALIZACION_DT"] = dt_finalizacion
            mes_finalizado = df["FINALIZACION_DT"].dt.month.map(meses_es).fillna("")
            anio_finalizado = df["FINALIZACION_DT"].dt.year.fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

            idx_rec = df.columns.get_loc("RECEPCIÓN")
            df.insert(idx_rec, "MES RECIBIDO", mes_recibido)
            df.insert(idx_rec, "AÑO RECIBIDO", anio_recibido)

            idx_fin = df.columns.get_loc("FINALIZACIÓN")
            df.insert(idx_fin, "MES FINALIZADO", mes_finalizado)
            df.insert(idx_fin, "AÑO FINALIZADO", anio_finalizado)

            df = df.drop_duplicates(subset=["INCIDENCIA"])
            lista_dataframes.append(df)
        except Exception as e:
            st.error(f"⚠️ Error procesando '{nombre_archivo}': {e}")
            return False

    if lista_dataframes:
        df_nuevo = pd.concat(lista_dataframes, ignore_index=True)
        
        # Preparamos los registros para enviarlos a Supabase
        # Supabase maneja upsert (si la incidencia ya existe, la actualiza; si no, la crea)
        registros = []
        for _, row in df_nuevo.iterrows():
            reg = {
                "incidencia": str(row["INCIDENCIA"]),
                "zona": str(row.get("ZONA", "")),
                "mes_recibido": str(row.get("MES RECIBIDO", "")),
                "anio_recibido": str(row.get("AÑO RECIBIDO", "")),
                "recepcion": str(row.get("RECEPCIÓN", "")),
                "mes_finalizado": str(row.get("MES FINALIZADO", "")),
                "anio_finalizado": str(row.get("AÑO FINALIZADO", "")),
                "finalizacion": str(row.get("FINALIZACIÓN", "")),
                "direccion": str(row.get("DIRECCIÓN", "")),
                "motivo": str(row.get("MOTIVO", "")),
                "tec": str(row.get("TEC", "")),
                "tcr": str(row.get("TCR", "")),
                "tdv": str(row.get("TDV", "")),
                "tej": str(row.get("TEJ", "")),
                "tar": str(row.get("TAR", "")),
                "estado": str(row.get("ESTADO", "")),
                "cliente": str(row.get("CLIENTE", "")),
                "total_incidencias": str(row.get("TOTAL INCIDENCIAS", "")),
                "sumatorias_tec": str(row.get("SUMATORIAS TEC", "")),
                "sumatorias_tcr": str(row.get("SUMATORIAS TCR", "")),
                "sumatorias_tdv": str(row.get("SUMATORIAS TDV", "")),
                "sumatorias_tej": str(row.get("SUMATORIAS TEJ", "")),
                "sumatorias_tar": str(row.get("SUMATORIAS TAR", "")),
                "promedio_tec": str(row.get("PROMEDIO TEC", "")),
                "promedio_tcr": str(row.get("PROMEDIO TCR", "")),
                "promedio_tdv": str(row.get("PROMEDIO TDV", "")),
                "promedio_tej": str(row.get("PROMEDIO TEJ", "")),
                "promedio_tar": str(row.get("PROMEDIO TAR", "")),
                "unidad": "",
                "gerencia": "",
                "origen_tipo": "Recibidas en el Día",
                "recepcion_dt": row["RECEPCION_DT"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(row["RECEPCION_DT"]) else None,
                "finalizacion_dt": row["FINALIZACION_DT"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(row["FINALIZACION_DT"]) else None,
            }
            registros.append(reg)

        # Inserción por lotes en Supabase para evitar sobrecargar la red
        chunk_size = 500
        for i in range(0, len(registros), chunk_size):
            chunk = registros[i:i + chunk_size]
            supabase.table("incidencias").upsert(chunk).execute()

        return True
    return False


# ==========================================
# CARGA DE DATOS DESDE SUPABASE
# ==========================================
@st.cache_data(ttl=60) # Se actualiza automáticamente cada minuto o al limpiar caché
def cargar_datos_supabase():
    if not supabase:
        return None
    try:
        response = supabase.table("incidencias").select("*").execute()
        data = response.data
        if not data:
            return None
        
        df = pd.DataFrame(data)
        # Normalizamos nombres de columnas de vuelta a mayúsculas para que el resto del dashboard funcione igual
        df.columns = [c.upper() for c in df.columns]
        
        # Mapeo inverso de nombres específicos que espera el dashboard
        if "RECEPCION" not in df.columns and "RECEPCION" in df.columns:
            pass # Ya están en mayúsculas
            
        df["RECEPCION_DT"] = pd.to_datetime(df["RECEPCION_DT"], errors="coerce")
        df["FINALIZACION_DT"] = pd.to_datetime(df["FINALIZACION_DT"], errors="coerce")
        return df
    except Exception as e:
        st.error(f"Error al leer datos de Supabase: {e}")
        return None


df = cargar_datos_supabase()

columnas_mostrar = [
    "INCIDENCIA", "RECEPCIÓN", "FINALIZACIÓN", "DIRECCIÓN", "CLIENTE", "MOTIVO", "ESTADO"
]

# ==========================================
# VISTAS PRINCIPALES EN TABS (6 TABS)
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📅 Seguimiento Diario",
    "📈 Seguimiento Anual",
    "🔍 Búsqueda Avanzada",
    "🎯 Análisis",
    "📁 Lotes Guardados",
    "📁 Carga y Actualización",
])


def renderizar_tabla_con_seleccion(df_sub, nombre_pestana, tipo_origen_etiqueta=None):
    if df is None or df.empty:
        st.info("No hay registros para mostrar en esta sección.")
        return

    df_editable = df_sub[[c for c in columnas_mostrar if c in df_sub.columns]].copy()
    df_editable.insert(0, "Seleccionar", False)

    df_editado = st.data_editor(
        df_editable, use_container_width=True, key=f"editor_{nombre_pestana}", hide_index=True
    )

    col_btn1, _ = st.columns([2, 4])
    with col_btn1:
        if st.button(f"➕ Agregar al Lote de Análisis", key=f"btn_agregar_{nombre_pestana}"):
            seleccionadas = df_editado[df_editado["Seleccionar"] == True]
            if not seleccionadas.empty:
                ids_seleccionados = seleccionadas["INCIDENCIA"].tolist()
                df_lote_nuevo = df[df["INCIDENCIA"].isin(ids_seleccionados)].copy()
                df_lote_nuevo["ORIGEN_TIPO"] = tipo_origen_etiqueta or "Recibidas en el Día"

                if "UNIDAD" not in df_lote_nuevo.columns: df_lote_nuevo["UNIDAD"] = ""
                if "GERENCIA" not in df_lote_nuevo.columns: df_lote_nuevo["GERENCIA"] = ""

                if not st.session_state["lote_seleccionado"].empty:
                    df_acumulado = pd.concat([
                        st.session_state["lote_seleccionado"], df_lote_nuevo
                    ]).drop_duplicates(subset=["INCIDENCIA"])
                    st.session_state["lote_seleccionado"] = df_acumulado
                else:
                    st.session_state["lote_seleccionado"] = df_lote_nuevo

                st.success(f"¡Se agregaron {len(ids_seleccionados)} incidencias al Análisis!")
            else:
                st.warning("⚠️ No has seleccionado ninguna incidencia.")


def renderizar_panel_graficos(df_lote_objetivo, sufijo=""):
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown("##### 📍 1. Gráfico de Zonas")
        if "ZONA" in df_lote_objetivo.columns and not df_lote_objetivo["ZONA"].isna().all():
            fig_zona = px.pie(df_lote_objetivo, names="ZONA", title="Distribución por Zonas", hole=0.4)
            st.plotly_chart(fig_zona, use_container_width=True, key=f"fig_zona_{sufijo}")
        else:
            st.info("Sin datos de zona.")

    with col_g2:
        st.markdown("##### 📅 2. Gráfico por Mes de Recepción")
        if "RECEPCION_DT" in df_lote_objetivo.columns and not df_lote_objetivo["RECEPCION_DT"].isna().all():
            df_meses_lote = df_lote_objetivo.groupby(df_lote_objetivo["RECEPCION_DT"].dt.to_period("M")).size().reset_index(name="CANTIDAD")
            df_meses_lote["MES_STR"] = df_meses_lote["RECEPCION_DT"].dt.strftime("%B %Y").str.upper()
            fig_mes = px.bar(df_meses_lote, x="MES_STR", y="CANTIDAD", title="Incidencias por Mes de Recepción", color="CANTIDAD", color_continuous_scale="Blues")
            st.plotly_chart(fig_mes, use_container_width=True, key=f"fig_mes_{sufijo}")
        else:
            st.info("Sin fechas válidas.")

    col_g3, col_g4 = st.columns(2)
    with col_g3:
        st.markdown("##### 🏷️ 3. Gráfico de Motivos")
        if "MOTIVO" in df_lote_objetivo.columns:
            fig_motivo = px.bar(df_lote_objetivo["MOTIVO"].value_counts().reset_index(), x="MOTIVO", y="count", title="Incidencias por Motivo")
            st.plotly_chart(fig_motivo, use_container_width=True, key=f"fig_motivo_{sufijo}")

    with col_g4:
        st.markdown("##### 🔄 4. Gráfico por Origen")
        if "ORIGEN_TIPO" in df_lote_objetivo.columns:
            fig_origen = px.pie(df_lote_objetivo, names="ORIGEN_TIPO", title="Clasificación por Origen", hole=0.4)
            st.plotly_chart(fig_origen, use_container_width=True, key=f"fig_origen_{sufijo}")


with tab1:
    st.subheader("📅 Comportamiento Diario por Mes")
    if df is not None:
        col_s1, col_s2, _ = st.columns([1, 1, 2])
        with col_s1:
            anio_seleccionado = st.selectbox("Año:", [2025, 2026, 2027], index=1, key="anio_s1")
        with col_s2:
            meses_dict = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
            mes_nombre_seleccionado = st.selectbox("Mes:", list(meses_dict.values()), index=7, key="mes_s1")
            mes_seleccionado = [k for k, v in meses_dict.items() if v == mes_nombre_seleccionado][0]

        inicio_mes_dinamico = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=1)
        ultimo_dia = calendar.monthrange(anio_seleccionado, mes_seleccionado)[1]
        fin_mes_dinamico = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=ultimo_dia) + pd.Timedelta(days=1)

        acumulado_mes = int(((df["RECEPCION_DT"] < inicio_mes_dinamico) & (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= inicio_mes_dinamico))).sum())
        recibidos_mes = int(((df["RECEPCION_DT"] >= inicio_mes_dinamico) & (df["RECEPCION_DT"] < fin_mes_dinamico)).sum())
        total_mes = acumulado_mes + recibidos_mes
        finalizados_mes = int(((df["FINALIZACION_DT"] >= inicio_mes_dinamico) & (df["FINALIZACION_DT"] < fin_mes_dinamico)).sum())
        pendientes_mes = total_mes - finalizados_mes

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("📦 Acumulado", f"{acumulado_mes:,}")
        c2.metric("📥 Recibidos", f"{recibidos_mes:,}")
        c3.metric("📊 Total", f"{total_mes:,}")
        c4.metric("✅ Finalizados", f"{finalizados_mes:,}")
        c5.metric("⏳ Pendientes", f"{pendientes_mes:,}")
        st.divider()

        with st.expander("Ver y seleccionar incidencias del mes"):
            df_mes_filtro = df[(df["RECEPCION_DT"] >= inicio_mes_dinamico) & (df["RECEPCION_DT"] < fin_mes_dinamico)].copy()
            renderizar_tabla_con_seleccion(df_mes_filtro, "diario_mes", "Recibidas en el Día")
    else:
        st.info("No hay datos cargados en Supabase.")


with tab2:
    st.subheader("📈 Seguimiento Mensual")
    if df is not None:
        inicio_anio = pd.Timestamp("2026-01-01")
        fin_anio = pd.Timestamp("2026-09-01")
        
        total_anio = int(((df["RECEPCION_DT"] >= inicio_anio) & (df["RECEPCION_DT"] < fin_anio)).sum())
        st.metric("Total Recibidas (Ene - Ago 2026)", f"{total_anio:,}")
        
        with st.expander("Ver y seleccionar incidencias anuales"):
            df_anual_filtro = df[(df["RECEPCION_DT"] >= inicio_anio) & (df["RECEPCION_DT"] < fin_anio)].copy()
            renderizar_tabla_con_seleccion(df_anual_filtro, "anual", "Recibidas en el Día")
    else:
        st.info("No hay datos cargados.")


with tab3:
    st.subheader("🔍 Buscador Avanzado")
    if df is not None:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_num = st.text_input("Filtrar por Nº de Incidencia:")
        with col_f2:
            filtro_dir = st.text_input("Filtrar por Dirección:")

        df_base = df.copy()
        if filtro_num:
            df_base = df_base[df_base["INCIDENCIA"].str.contains(filtro_num, case=False, na=False)]
        if filtro_dir:
            df_base = df_base[df_base["DIRECCIÓN"].str.contains(filtro_dir, case=False, na=False)]

        renderizar_tabla_con_seleccion(df_base, "busq_avanzada", "Recibidas en el Día")
    else:
        st.info("No hay datos.")


with tab4:
    st.header("🎯 Panel de Análisis y Asignación de Unidades")
    df_lote = st.session_state["lote_seleccionado"]

    if df_lote.empty:
        st.info("Agrega incidencias desde las otras pestañas para analizarlas aquí.")
    else:
        st.markdown(f"**Total en lote actual:** {len(df_lote)} incidencias.")
        
        if "UNIDAD" not in df_lote.columns: df_lote["UNIDAD"] = ""
        if "GERENCIA" not in df_lote.columns: df_lote["GERENCIA"] = ""

        cols_edicion = ["INCIDENCIA", "MOTIVO", "ESTADO", "UNIDAD", "GERENCIA"]
        df_editado_manual = st.data_editor(
            df_lote[[c for c in cols_edicion if c in df_lote.columns]],
            use_container_width=True, key="editor_manual", hide_index=True, disabled=["INCIDENCIA", "MOTIVO", "ESTADO"]
        )

        if st.button("💾 Guardar Cambios en Unidades y Sincronizar"):
            for _, row in df_editado_manual.iterrows():
                inc_id = str(row["INCIDENCIA"])
                mask = df_lote["INCIDENCIA"] == inc_id
                df_lote.loc[mask, "UNIDAD"] = str(row["UNIDAD"]).strip().upper()
                df_lote.loc[mask, "GERENCIA"] = str(row["GERENCIA"]).strip().upper()
                
                # Actualizamos de una vez en Supabase para que los demás lo vean
                if supabase:
                    supabase.table("incidencias").update({
                        "unidad": str(row["UNIDAD"]).strip().upper(),
                        "gerencia": str(row["GERENCIA"]).strip().upper()
                    }).eq("incidencia", inc_id).execute()

            st.session_state["lote_seleccionado"] = df_lote
            st.success("¡Guardado y sincronizado globalmente con éxito!")
            st.rerun()

        st.divider()
        renderizar_panel_graficos(df_lote, sufijo="activo")


with tab5:
    st.header("📁 Lotes y Análisis Guardados")
    st.info("Los lotes activos se gestionan globalmente en las vistas anteriores.")


with tab6:
    st.subheader("📁 Carga de Zonas y Actualización Global en Supabase")
    st.markdown("Sube tus archivos `.txt`. Al procesarlos, se actualizarán en la base de datos en la nube para **todos los usuarios** de la aplicación.")

    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    archivos_cargados = st.file_uploader(
        "Selecciona archivos .txt", type=["txt", "TXT"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
    )

    if st.button("🚀 Procesar y Subir a Supabase", type="primary"):
        if archivos_cargados:
            with st.spinner("Subiendo y sincronizando datos en la nube..."):
                exito = procesar_txts_en_supabase(archivos_cargados)
                if exito:
                    st.cache_data.clear() # Limpia la caché para refrescar los datos nuevos
                    st.success("¡Actualización exitosa y disponible para todos los usuarios!")
                    st.session_state["uploader_key"] += 1
                    st.rerun()
        else:
            st.warning("⚠️ Sube al menos un archivo .txt.")
