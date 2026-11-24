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
    layout="wide",
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

    indices_a_conservar = [
        13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26,
        28, 29, 30, 31, 32, 34, 35, 36, 37, 38,
    ]

    nombres_finales = [
        "INCIDENCIA", "DIRECCIÓN", "MOTIVO", "RECEPCIÓN", "FINALIZACIÓN",
        "TEC", "TCR", "TDV", "TEJ", "TAR", "ESTADO", "CLIENTE",
        "TOTAL INCIDENCIAS", "SUMATORIAS TEC", "SUMATORIAS TCR",
        "SUMATORIAS TDV", "SUMATORIAS TEJ", "SUMATORIAS TAR",
        "PROMEDIO TEC", "PROMEDIO TCR", "PROMEDIO TDV", "PROMEDIO TEJ", "PROMEDIO TAR",
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

        if os.path.exists("resultado_unificado.xlsx"):
            try:
                df_existente = pd.read_excel("resultado_unificado.xlsx", dtype=str)
                df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
                df_final = df_final.drop_duplicates(subset=["INCIDENCIA"], keep="last")
            except Exception:
                df_final = df_nuevo
        else:
            df_final = df_nuevo

        df_final.to_excel("resultado_unificado.xlsx", index=False)
        return True
    return False


# ==========================================
# 2. CARGA, TIEMPOS (MINUTOS) Y FECHAS POR ESTATUS
# ==========================================
def parsear_minutos(serie):
    limpio = (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpio, errors="coerce").fillna(0.0)


@st.cache_data
def cargar_datos():
    archivo_entrada = "resultado_unificado.xlsx"
    if not os.path.exists(archivo_entrada):
        return None

    df = pd.read_excel(archivo_entrada, dtype=str)

    for col in ["RECEPCIÓN", "FINALIZACIÓN"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(["NAN", "NONE", "NAT", ""], pd.NA)

    if "ESTADO" in df.columns:
        df["ESTADO"] = df["ESTADO"].fillna("").astype(str).str.strip().str.upper()
    else:
        df["ESTADO"] = ""

    df["RECEPCION_DT"] = pd.to_datetime(
        df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )
    df["FINALIZACION_DT"] = pd.to_datetime(
        df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
    )

    # Tiempos en minutos
    for col_t in ["TEC", "TCR", "TDV", "TEJ", "TAR"]:
        if col_t in df.columns:
            df[f"{col_t}_MIN"] = parsear_minutos(df[col_t])
        else:
            df[f"{col_t}_MIN"] = 0.0

    # Fecha y hora estimada de entrada a cada etapa sumando minutos desde RECEPCION_DT
    df["DT_ZONIFICACION"] = df["RECEPCION_DT"] + pd.to_timedelta(df["TEC_MIN"], unit="m")
    df["DT_PA"] = df["DT_ZONIFICACION"] + pd.to_timedelta(df["TCR_MIN"], unit="m")
    df["DT_AS"] = df["DT_PA"] + pd.to_timedelta(df["TDV_MIN"], unit="m")
    df["DT_EE"] = df["DT_AS"] + pd.to_timedelta(df["TEJ_MIN"], unit="m")

    # Fecha/hora de atención (FO): usa FINALIZACION_DT si existe; si no, DT_EE + TAR_MIN cuando TAR_MIN > 0
    dt_fo_estimado = df["DT_EE"] + pd.to_timedelta(df["TAR_MIN"], unit="m")
    tiene_fo = (df["ESTADO"] == "FO") | df["FINALIZACION_DT"].notna() | (df["TAR_MIN"] > 0)
    df["DT_FO"] = df["FINALIZACION_DT"].where(
        df["FINALIZACION_DT"].notna(),
        dt_fo_estimado.where(tiene_fo, pd.NaT),
    )

    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }

    df["MES RECIBIDO"] = df["RECEPCION_DT"].dt.month.map(meses_es).fillna("")
    df["AÑO RECIBIDO"] = (
        df["RECEPCION_DT"].dt.year.fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    )

    return df


def calcular_metricas_turno(df_data, fecha_dia):
    """
    Calcula las métricas del turno de 07:00 AM del día `fecha_dia`
    hasta las 07:00 AM del día siguiente (`fecha_dia + 1 día`).
    """
    base_dia = pd.Timestamp(year=fecha_dia.year, month=fecha_dia.month, day=fecha_dia.day)
    turno_inicio = base_dia + pd.Timedelta(hours=7)
    turno_fin = turno_inicio + pd.Timedelta(days=1)

    estado_limpio = df_data["ESTADO"].fillna("").astype(str).str.strip().str.upper()

    # 1. Pendiente: acumulado del día anterior en estatus PA, corte 7 AM a 7 AM del día siguiente
    mask_pendiente = (
        (estado_limpio == "PA")
        & (df_data["RECEPCION_DT"] < turno_fin)
        & (df_data["FINALIZACION_DT"].isna() | (df_data["FINALIZACION_DT"] >= turno_inicio))
    )

    # 2. Recibido: lo recibido en el corte de 7 AM a 7 AM del día siguiente
    mask_recibido = (
        (df_data["RECEPCION_DT"] >= turno_inicio)
        & (df_data["RECEPCION_DT"] < turno_fin)
    )

    # 3. En Ejecución: lo que quedó en el turno con estatus AS y EE, corte 7 AM a 7 AM del día siguiente
    mask_ejecucion = (
        estado_limpio.isin(["AS", "EE"])
        & (df_data["RECEPCION_DT"] < turno_fin)
        & (df_data["FINALIZACION_DT"].isna() | (df_data["FINALIZACION_DT"] >= turno_inicio))
    )

    # 4. Atendido: lo atendido en el turno, estatus FO (o finalizado en el turno cuando es FO)
    en_turno_fo = df_data["DT_FO"].notna() & (df_data["DT_FO"] >= turno_inicio) & (df_data["DT_FO"] < turno_fin)
    en_turno_fin = df_data["FINALIZACION_DT"].notna() & (df_data["FINALIZACION_DT"] >= turno_inicio) & (df_data["FINALIZACION_DT"] < turno_fin)
    mask_atendido = (estado_limpio == "FO") & (en_turno_fo | en_turno_fin | (df_data["DT_FO"].isna() & df_data["FINALIZACION_DT"].isna()))

    pendiente = int(mask_pendiente.sum())
    recibido = int(mask_recibido.sum())
    ejecucion = int(mask_ejecucion.sum())
    atendido = int(mask_atendido.sum())
    # 5. Total: sumatoria de lo recibido más lo pendiente, sin considerar ejecución
    total = pendiente + recibido

    return {
        "turno_inicio": turno_inicio,
        "turno_fin": turno_fin,
        "pendiente": pendiente,
        "recibido": recibido,
        "ejecucion": ejecucion,
        "atendido": atendido,
        "total": total,
        "df_pendiente": df_data[mask_pendiente].copy(),
        "df_recibido": df_data[mask_recibido].copy(),
        "df_ejecucion": df_data[mask_ejecucion].copy(),
        "df_atendido": df_data[mask_atendido].copy(),
    }


df = cargar_datos()

if df is None:
    st.warning(
        "⚠️ No se encontró el archivo 'resultado_unificado.xlsx'. Ve a la pestaña **📦 Actualización** para subir tus archivos .txt."
    )

columnas_mostrar = [
    "INCIDENCIA", "RECEPCIÓN", "FINALIZACIÓN", "DIRECCIÓN", "CLIENTE",
    "MOTIVO", "ESTADO", "TEC", "TCR", "TDV", "TEJ", "TAR",
]


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

    cols_visibles = [c for c in columnas_mostrar if c in df_sub.columns]
    df_editable = df_sub[cols_visibles].copy()
    df_editable.insert(0, "Seleccionar", False)

    df_editado = st.data_editor(
        df_editable,
        use_container_width=True,
        key=f"editor_{nombre_pestana}",
        hide_index=True,
    )

    col_btn1, _ = st.columns([2, 4])
    with col_btn1:
        if st.button("➕ Agregar al Lote de Análisis", key=f"btn_agregar_{nombre_pestana}"):
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
    st.subheader("📅 Comportamiento Diario por Turno (07:00 AM a 07:00 AM)")
    if df is not None:
        col_s1, col_s2, _ = st.columns([1, 1, 2])
        with col_s1:
            anio_seleccionado = st.selectbox("Seleccione el Año:", [2025, 2026, 2027], index=1, key="anio_s1")
        with col_s2:
            meses_dict = {
                1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
                5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
                9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
            }
            mes_nombre_seleccionado = st.selectbox("Seleccione el Mes:", list(meses_dict.values()), index=10, key="mes_s1")
            mes_seleccionado = [k for k, v in meses_dict.items() if v == mes_nombre_seleccionado][0]

        ultimo_dia = calendar.monthrange(anio_seleccionado, mes_seleccionado)[1]
        dias_mes = pd.date_range(
            start=pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=1),
            end=pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=ultimo_dia),
            freq="D",
        )

        fecha_maxima_datos = df["RECEPCION_DT"].max()
        datos_diarios = []

        for dia in dias_mes:
            turno_ini_dia = dia + pd.Timedelta(hours=7)
            if pd.notna(fecha_maxima_datos) and turno_ini_dia > fecha_maxima_datos + pd.Timedelta(days=1):
                datos_diarios.append({
                    "FECHA": dia.strftime("%d/%m/%Y"),
                    "PENDIENTE (PA)": None,
                    "RECIBIDO": None,
                    "EN EJECUCION (AS+EE)": None,
                    "ATENDIDO (FO)": None,
                    "TOTAL": None,
                })
            else:
                res_turno = calcular_metricas_turno(df, dia)
                datos_diarios.append({
                    "FECHA": dia.strftime("%d/%m/%Y"),
                    "PENDIENTE (PA)": res_turno["pendiente"],
                    "RECIBIDO": res_turno["recibido"],
                    "EN EJECUCION (AS+EE)": res_turno["ejecucion"],
                    "ATENDIDO (FO)": res_turno["atendido"],
                    "TOTAL": res_turno["total"],
                })

        df_dia = pd.DataFrame(datos_diarios)

        recibidos_validos = pd.to_numeric(df_dia["RECIBIDO"], errors="coerce").dropna()
        atendidos_validos = pd.to_numeric(df_dia["ATENDIDO (FO)"], errors="coerce").dropna()

        promedio_recibido_diario = float(recibidos_validos.mean()) if not recibidos_validos.empty else 0.0
        promedio_atendido_diario = float(atendidos_validos.mean()) if not atendidos_validos.empty else 0.0

        total_recibido_mes = int(recibidos_validos.sum()) if not recibidos_validos.empty else 0
        total_atendido_mes = int(atendidos_validos.sum()) if not atendidos_validos.empty else 0
        pendiente_validos = pd.to_numeric(df_dia["PENDIENTE (PA)"], errors="coerce").dropna()
        ejecucion_validos = pd.to_numeric(df_dia["EN EJECUCION (AS+EE)"], errors="coerce").dropna()

        ultimo_pendiente_mes = int(pendiente_validos.iloc[-1]) if not pendiente_validos.empty else 0
        ultimo_ejecucion_mes = int(ejecucion_validos.iloc[-1]) if not ejecucion_validos.empty else 0
        total_acum_mes = ultimo_pendiente_mes + total_recibido_mes

        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        c1.metric("⏳ Pendiente (PA)", f"{ultimo_pendiente_mes:,}")
        c2.metric("📥 Recibido (Mes)", f"{total_recibido_mes:,}")
        c3.metric("⚙️ En Ejecución (AS+EE)", f"{ultimo_ejecucion_mes:,}")
        c4.metric("✅ Atendido (FO Mes)", f"{total_atendido_mes:,}")
        c5.metric("📊 Total (Pend+Rec)", f"{total_acum_mes:,}")
        c6.metric("📈 Prom. Recibido/Día", f"{promedio_recibido_diario:.1f}")
        c7.metric("📉 Prom. Atendido/Día", f"{promedio_atendido_diario:.1f}")
        st.divider()

        curva_prom_recibido = [
            round(promedio_recibido_diario, 1) if pd.notna(v) else None for v in df_dia["RECIBIDO"]
        ]
        curva_prom_atendido = [
            round(promedio_atendido_diario, 1) if pd.notna(v) else None for v in df_dia["ATENDIDO (FO)"]
        ]

        tope_barra = (
            pd.to_numeric(df_dia["PENDIENTE (PA)"], errors="coerce").fillna(0)
            + pd.to_numeric(df_dia["RECIBIDO"], errors="coerce").fillna(0)
            + pd.to_numeric(df_dia["EN EJECUCION (AS+EE)"], errors="coerce").fillna(0)
        )
        tope_barra = tope_barra.where(df_dia["TOTAL"].notna(), None)

        fig1 = go.Figure()
        fig1.add_trace(go.Bar(
            x=df_dia["FECHA"], y=df_dia["PENDIENTE (PA)"],
            name="Pendiente (PA)", marker_color="#fca5a5",
            text=df_dia["PENDIENTE (PA)"], textposition="inside",
            textfont=dict(color="black", size=11),
        ))
        fig1.add_trace(go.Bar(
            x=df_dia["FECHA"], y=df_dia["RECIBIDO"],
            name="Recibido", marker_color="#e9f056",
            text=df_dia["RECIBIDO"], textposition="inside",
            textfont=dict(color="black", size=11),
        ))
        fig1.add_trace(go.Bar(
            x=df_dia["FECHA"], y=df_dia["EN EJECUCION (AS+EE)"],
            name="En Ejecución (AS+EE)", marker_color="#93c5fd",
            text=df_dia["EN EJECUCION (AS+EE)"], textposition="inside",
            textfont=dict(color="black", size=11),
        ))
        fig1.add_trace(go.Scatter(
            x=df_dia["FECHA"], y=tope_barra,
            name="Total (Pendiente + Recibido)", mode="text+markers",
            text=df_dia["TOTAL"], textposition="top center",
            textfont=dict(color="#1F4E78", size=12),
            marker=dict(size=8, color="rgba(0,0,0,0)"), showlegend=False,
            connectgaps=False, cliponaxis=False,
        ))
        fig1.add_trace(go.Scatter(
            x=df_dia["FECHA"], y=df_dia["ATENDIDO (FO)"],
            name="Atendido (FO)", mode="lines+markers+text",
            text=df_dia["ATENDIDO (FO)"], textposition="bottom center",
            textfont=dict(color="#1d4ed8", size=11),
            marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
            connectgaps=False,
        ))
        fig1.add_trace(go.Scatter(
            x=df_dia["FECHA"], y=curva_prom_recibido,
            name=f"Prom. Recibido Diario ({promedio_recibido_diario:.1f})",
            mode="lines", line=dict(color="#d97706", width=2.5, dash="dash"),
            connectgaps=False,
        ))
        fig1.add_trace(go.Scatter(
            x=df_dia["FECHA"], y=curva_prom_atendido,
            name=f"Prom. Atendido Diario ({promedio_atendido_diario:.1f})",
            mode="lines", line=dict(color="#16a34a", width=2.5, dash="dot"),
            connectgaps=False,
        ))

        fig1.update_layout(
            title=dict(
                text=f"<b>Flujo Diario por Turno (07:00 AM - 07:00 AM) - {mes_nombre_seleccionado} {anio_seleccionado}</b>",
                font=dict(size=18, color="#1f4e78"),
            ),
            barmode="stack",
            xaxis_title="<b>Día del Turno (Inicia 07:00 AM)</b>",
            yaxis_title="<b>Cantidad de Reportes</b>",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
        )
        st.plotly_chart(fig1, use_container_width=True)

        st.divider()
        st.markdown(f"### 📌 Resumen Mensual por Motivo — {mes_nombre_seleccionado} {anio_seleccionado}")
        inicio_mes = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=1, hour=7)
        fin_mes = pd.Timestamp(year=anio_seleccionado, month=mes_seleccionado, day=ultimo_dia, hour=7) + pd.Timedelta(days=1)

        df_mes_motivos = df[
            (df["RECEPCION_DT"] >= inicio_mes) & (df["RECEPCION_DT"] < fin_mes)
        ].copy()

        if not df_mes_motivos.empty and "MOTIVO" in df_mes_motivos.columns:
            motivos_limpios = df_mes_motivos["MOTIVO"].fillna("").astype(str).str.strip().replace("", "SIN MOTIVO")
            resumen_motivos = motivos_limpios.value_counts().reset_index()
            resumen_motivos.columns = ["MOTIVO", "TOTAL TURNOS MES"]
            total_mot = int(resumen_motivos["TOTAL TURNOS MES"].sum())
            resumen_motivos["% DEL MES"] = (resumen_motivos["TOTAL TURNOS MES"] / total_mot * 100).round(1)

            tabla_res = resumen_motivos.copy()
            tabla_res["% DEL MES"] = tabla_res["% DEL MES"].map(lambda x: f"{x:.1f}%")
            tabla_res = pd.concat([
                tabla_res,
                pd.DataFrame([{"MOTIVO": "TOTAL GENERAL", "TOTAL TURNOS MES": total_mot, "% DEL MES": "100.0%"}]),
            ], ignore_index=True)

            col_t, col_g = st.columns([1.1, 1.4])
            with col_t:
                st.dataframe(tabla_res, use_container_width=True, hide_index=True)
            with col_g:
                resumen_ord = resumen_motivos.sort_values("TOTAL TURNOS MES", ascending=True)
                fig_mot = px.bar(
                    resumen_ord, x="TOTAL TURNOS MES", y="MOTIVO",
                    orientation="h", text="TOTAL TURNOS MES",
                    title=f"Totales Mensuales por Motivo - {mes_nombre_seleccionado} {anio_seleccionado}",
                )
                fig_mot.update_traces(marker_color="#1f4e78", textposition="outside", cliponaxis=False)
                fig_mot.update_layout(
                    xaxis_title="<b>Cantidad</b>", yaxis_title="<b>Motivo</b>",
                    height=max(320, len(resumen_ord) * 34), margin=dict(r=40),
                )
                st.plotly_chart(fig_mot, use_container_width=True, key="fig_motivos_turno_mes")
        else:
            st.info("No hay incidencias en el rango del mes seleccionado.")

        with st.expander("Ver y seleccionar incidencias del mes"):
            renderizar_tabla_con_seleccion(df_mes_motivos, "diario_mes", "Turno del Mes")


# --- PESTAÑA 2: SEGUIMIENTO ANUAL ---
with tab2:
    st.subheader("📊 Seguimiento Mensual Acumulado (Turno 07:00 AM)")
    if df is not None:
        col_y1, _ = st.columns([1, 3])
        with col_y1:
            anio_anual = st.selectbox("Seleccione el Año:", [2025, 2026, 2027], index=1, key="anio_s2")

        df_anio_actual = df[df["RECEPCION_DT"].dt.year == anio_anual]
        max_mes_datos = df_anio_actual["RECEPCION_DT"].max().month if not df_anio_actual.empty else 11

        meses_dict_anual = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
            5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
            9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
        }

        datos_anuales = []
        for m in range(1, max_mes_datos + 1):
            ult_d = calendar.monthrange(anio_anual, m)[1]
            t_ini = pd.Timestamp(year=anio_anual, month=m, day=1, hour=7)
            t_fin = pd.Timestamp(year=anio_anual, month=m, day=ult_d, hour=7) + pd.Timedelta(days=1)

            estado_limpio = df["ESTADO"].fillna("").astype(str).str.strip().str.upper()
            pend_m = int((
                (estado_limpio == "PA")
                & (df["RECEPCION_DT"] < t_fin)
                & (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= t_ini))
            ).sum())
            rec_m = int(((df["RECEPCION_DT"] >= t_ini) & (df["RECEPCION_DT"] < t_fin)).sum())
            ej_m = int((
                estado_limpio.isin(["AS", "EE"])
                & (df["RECEPCION_DT"] < t_fin)
                & (df["FINALIZACION_DT"].isna() | (df["FINALIZACION_DT"] >= t_ini))
            ).sum())
            at_m = int((
                (estado_limpio == "FO")
                & (df["RECEPCION_DT"] >= t_ini)
                & (df["RECEPCION_DT"] < t_fin)
            ).sum())
            tot_m = pend_m + rec_m

            datos_anuales.append({
                "MES": meses_dict_anual[m],
                "PENDIENTE (PA)": pend_m,
                "RECIBIDO": rec_m,
                "EN EJECUCION (AS+EE)": ej_m,
                "ATENDIDO (FO)": at_m,
                "TOTAL (PEND+REC)": tot_m,
            })

        df_anual = pd.DataFrame(datos_anuales)
        st.dataframe(df_anual, use_container_width=True, hide_index=True)

        fig_anual = go.Figure()
        fig_anual.add_trace(go.Bar(
            x=df_anual["MES"], y=df_anual["PENDIENTE (PA)"],
            name="Pendiente (PA)", marker_color="#fca5a5",
            text=df_anual["PENDIENTE (PA)"], textposition="inside",
        ))
        fig_anual.add_trace(go.Bar(
            x=df_anual["MES"], y=df_anual["RECIBIDO"],
            name="Recibido", marker_color="#e9f056",
            text=df_anual["RECIBIDO"], textposition="inside",
        ))
        fig_anual.add_trace(go.Bar(
            x=df_anual["MES"], y=df_anual["EN EJECUCION (AS+EE)"],
            name="En Ejecución (AS+EE)", marker_color="#93c5fd",
            text=df_anual["EN EJECUCION (AS+EE)"], textposition="inside",
        ))
        fig_anual.add_trace(go.Scatter(
            x=df_anual["MES"], y=df_anual["ATENDIDO (FO)"],
            name="Atendido (FO)", mode="lines+markers+text",
            text=df_anual["ATENDIDO (FO)"], textposition="bottom center",
            marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
        ))
        fig_anual.update_layout(
            title=dict(text=f"<b>Evolución Mensual - {anio_anual}</b>", font=dict(size=18, color="#1f4e78")),
            barmode="stack", hovermode="x unified",
        )
        st.plotly_chart(fig_anual, use_container_width=True)


# --- PESTAÑA 3: BÚSQUEDA AVANZADA ---
with tab3:
    st.subheader("🔍 Buscador de Incidencias por Turno (07:00 AM a 07:00 AM)")
    if df is not None:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fecha_busqueda = st.date_input("Fecha del turno (inicia 07:00 AM):", value=datetime(2026, 11, 24))
        with col_f2:
            filtro_num = st.text_input("Nº de Incidencia:", placeholder="Ej. 12345")
        with col_f3:
            filtro_dir = st.text_input("Dirección:", placeholder="Ej. Av. Principal")

        df_base_filtrada = df.copy()
        if filtro_num:
            df_base_filtrada = df_base_filtrada[
                df_base_filtrada["INCIDENCIA"].str.contains(filtro_num, case=False, na=False)
            ]
        if filtro_dir:
            df_base_filtrada = df_base_filtrada[
                df_base_filtrada["DIRECCIÓN"].str.contains(filtro_dir, case=False, na=False)
            ]

        res = calcular_metricas_turno(df_base_filtrada, fecha_busqueda)

        st.markdown(
            f"##### 📊 Turno: {res['turno_inicio'].strftime('%d/%m/%Y %I:%M %p')} → {res['turno_fin'].strftime('%d/%m/%Y %I:%M %p')}"
        )
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("⏳ Pendiente (PA)", f"{res['pendiente']:,}")
        m2.metric("📥 Recibido", f"{res['recibido']:,}")
        m3.metric("⚙️ En Ejecución (AS+EE)", f"{res['ejecucion']:,}")
        m4.metric("✅ Atendido (FO)", f"{res['atendido']:,}")
        m5.metric("📊 Total (Pend+Rec)", f"{res['total']:,}")
        st.divider()

        st.markdown("##### ⏳ Detalle Pendiente (PA)")
        renderizar_tabla_con_seleccion(res["df_pendiente"], "buscador_pendiente", "Pendiente PA")

        st.markdown("##### 📥 Detalle Recibido en el Turno")
        renderizar_tabla_con_seleccion(res["df_recibido"], "buscador_recibido", "Recibido Turno")

        st.markdown("##### ⚙️ Detalle En Ejecución (AS + EE)")
        renderizar_tabla_con_seleccion(res["df_ejecucion"], "buscador_ejecucion", "En Ejecución AS+EE")

        st.markdown("##### ✅ Detalle Atendido (FO)")
        renderizar_tabla_con_seleccion(res["df_atendido"], "buscador_atendido", "Atendido FO")


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

        columnas_analisis = [
            "INCIDENCIA", "RECEPCIÓN", "FINALIZACIÓN", "DIRECCIÓN", "CLIENTE",
            "MOTIVO", "ESTADO", "UNIDAD", "GERENCIA",
        ]
        df_lote_editado = st.data_editor(
            df_lote[columnas_analisis], use_container_width=True,
            key="editor_lote_analisis", hide_index=True,
        )

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
