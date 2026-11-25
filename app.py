import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Reportes 2026",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Control y Seguimiento de Incidencias (Turno 07:00 AM - 07:00 AM)")

if "lote_seleccionado" not in st.session_state:
    st.session_state["lote_seleccionado"] = pd.DataFrame()


def _parsear_dt(df, col):
    if col not in df.columns:
        return pd.Series(pd.NaT, index=df.index)
    s = df[col].fillna("").astype(str).str.strip()
    s = s.replace(["", "NAN", "NONE", "NAT"], pd.NA)
    return pd.to_datetime(s, format="%d-%m-%y %I:%M %p", errors="coerce")


@st.cache_data
def cargar_datos():
    archivo_entrada = "resultado_unificado.xlsx"
    if not os.path.exists(archivo_entrada):
        return None

    df = pd.read_excel(archivo_entrada, dtype=str)

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
        df[col] = df[col].replace(["NAN", "NONE", "NAT"], "", regex=False)

    if "ESTADO" in df.columns:
        df["ESTADO"] = df["ESTADO"].str.upper()
    else:
        df["ESTADO"] = ""

    df["RECEPCION_DT"] = _parsear_dt(df, "RECEPCIÓN")
    df["FINALIZACION_DT"] = _parsear_dt(df, "FINALIZACIÓN")

    # Horas calculadas desde el unificador (con respaldo si se lee un Excel anterior)
    df["H_PA_DT"] = _parsear_dt(df, "H_PA")
    if df["H_PA_DT"].isna().all():
        df["H_PA_DT"] = df["RECEPCION_DT"]

    df["H_EE_DT"] = _parsear_dt(df, "H_EE")
    df["H_FO_DT"] = _parsear_dt(df, "H_FO")
    if df["H_FO_DT"].isna().all():
        df["H_FO_DT"] = df["FINALIZACION_DT"]

    return df


def evaluar_turno_cadena(df_data, fecha_dia, pendiente_arrastre_previo=None):
    """
    Evalúa el turno de 07:00 AM de `fecha_dia` hasta 07:00 AM del día siguiente.
    Permite encadenar el arrastre del turno anterior.
    """
    base = pd.Timestamp(year=fecha_dia.year, month=fecha_dia.month, day=fecha_dia.day)
    inicio_turno = base + pd.Timedelta(hours=7)
    fin_turno = inicio_turno + pd.Timedelta(days=1)

    estado = df_data["ESTADO"].fillna("").str.upper()
    rec_dt = df_data["RECEPCION_DT"]
    h_pa_dt = df_data["H_PA_DT"]
    h_ee_dt = df_data["H_EE_DT"]
    h_fo_dt = df_data["H_FO_DT"]

    # Lo que ya pasó a ejecución (AS/EE) o se atendió (FO) antes de iniciar este turno
    paso_ejecucion_antes = h_ee_dt.notna() & (h_ee_dt <= inicio_turno)
    atendido_antes = h_fo_dt.notna() & (h_fo_dt <= inicio_turno)

    # Candidatos en estatus PA acumulados hasta el corte del turno
    candidatos_pa = (
        (estado == "PA")
        & h_pa_dt.notna()
        & (h_pa_dt < fin_turno)
        & (~paso_ejecucion_antes)
        & (~atendido_antes)
    )

    if pendiente_arrastre_previo is not None:
        # En cadena: conserva lo que seguía pendiente del día anterior y suma nuevos PA previos a fin_turno
        mask_pendiente = (
            (df_data.index.isin(pendiente_arrastre_previo) & (~paso_ejecucion_antes) & (~atendido_antes))
            | candidatos_pa
        )
    else:
        mask_pendiente = candidatos_pa

    mask_recibido = rec_dt.notna() & (rec_dt >= inicio_turno) & (rec_dt < fin_turno)

    # En ejecución: estatus AS o EE en el turno (no atendido antes de fin_turno)
    atendido_antes_fin = h_fo_dt.notna() & (h_fo_dt <= fin_turno)
    mask_ejecucion = (
        estado.isin(["AS", "EE"])
        & rec_dt.notna()
        & (rec_dt < fin_turno)
        & (~atendido_antes_fin)
    )

    # Atendido: estatus FO con hora de finalización dentro del turno [inicio_turno, fin_turno)
    en_ventana_fo = h_fo_dt.notna() & (h_fo_dt >= inicio_turno) & (h_fo_dt < fin_turno)
    mask_atendido = (estado == "FO") & en_ventana_fo

    pendiente = int(mask_pendiente.sum())
    recibido = int(mask_recibido.sum())
    ejecucion = int(mask_ejecucion.sum())
    atendido = int(mask_atendido.sum())
    total = pendiente + recibido

    # Lo que queda pendiente al cierre del turno (7 AM del día siguiente) para alimentar el siguiente eslabón
    paso_ejecucion_cierre = h_ee_dt.notna() & (h_ee_dt <= fin_turno)
    atendido_cierre = h_fo_dt.notna() & (h_fo_dt <= fin_turno)
    siguiente_arrastre = df_data.index[
        mask_pendiente & (~paso_ejecucion_cierre) & (~atendido_cierre)
    ]

    return {
        "inicio_turno": inicio_turno,
        "fin_turno": fin_turno,
        "pendiente": pendiente,
        "recibido": recibido,
        "ejecucion": ejecucion,
        "atendido": atendido,
        "total": total,
        "siguiente_arrastre": siguiente_arrastre,
        "df_pendiente": df_data[mask_pendiente].copy(),
        "df_recibido": df_data[mask_recibido].copy(),
        "df_ejecucion": df_data[mask_ejecucion].copy(),
        "df_atendido": df_data[mask_atendido].copy(),
    }


df = cargar_datos()

if df is None:
    st.warning("⚠️ No se encontró 'resultado_unificado.xlsx'. Genéralo primero con tu script unificador.")
    st.stop()

columnas_visibles = [
    "INCIDENCIA", "ZONA", "RECEPCIÓN", "H_PA", "H_EE", "H_FO",
    "FINALIZACIÓN", "ESTADO", "MOTIVO", "DIRECCIÓN",
]


def mostrar_tabla(df_sub, key_suffix):
    if df_sub.empty:
        st.info("No hay registros en este bloque.")
        return
    cols = [c for c in columnas_visibles if c in df_sub.columns]
    st.dataframe(df_sub[cols], use_container_width=True, hide_index=True, key=f"tbl_{key_suffix}")


tab1, tab2, tab3 = st.tabs([
    "📅 Seguimiento Diario (Cadena)",
    "🔍 Auditoría por Turno",
    "📄 Exportar",
])

with tab1:
    c1, c2, _ = st.columns([1, 1, 2])
    with c1:
        anio_sel = st.selectbox("Año:", [2025, 2026, 2027], index=1)
    with c2:
        meses_dict = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
            5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
            9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
        }
        mes_nombre = st.selectbox("Mes:", list(meses_dict.values()), index=10)
        mes_sel = [k for k, v in meses_dict.items() if v == mes_nombre][0]

    ultimo_dia = calendar.monthrange(anio_sel, mes_sel)[1]
    dias_mes = pd.date_range(
        start=pd.Timestamp(year=anio_sel, month=mes_sel, day=1),
        end=pd.Timestamp(year=anio_sel, month=mes_sel, day=ultimo_dia),
        freq="D",
    )

    fecha_max_dato = df["RECEPCION_DT"].max()
    arrastre_cadena = None
    filas_cadena = []

    for dia in dias_mes:
        turno_ini = dia + pd.Timedelta(hours=7)
        if pd.notna(fecha_max_dato) and turno_ini > (fecha_max_dato + pd.Timedelta(days=1)):
            filas_cadena.append({
                "TURNO (07:00 AM)": dia.strftime("%d/%m/%Y"),
                "PENDIENTE (PA)": None,
                "RECIBIDO": None,
                "EN EJECUCION (AS+EE)": None,
                "ATENDIDO (FO)": None,
                "TOTAL (PEND+REC)": None,
            })
        else:
            res = evaluar_turno_cadena(df, dia, pendiente_arrastre_previo=arrastre_cadena)
            arrastre_cadena = res["siguiente_arrastre"]
            filas_cadena.append({
                "TURNO (07:00 AM)": dia.strftime("%d/%m/%Y"),
                "PENDIENTE (PA)": res["pendiente"],
                "RECIBIDO": res["recibido"],
                "EN EJECUCION (AS+EE)": res["ejecucion"],
                "ATENDIDO (FO)": res["atendido"],
                "TOTAL (PEND+REC)": res["total"],
            })

    df_cadena = pd.DataFrame(filas_cadena)

    rec_vals = pd.to_numeric(df_cadena["RECIBIDO"], errors="coerce").dropna()
    at_vals = pd.to_numeric(df_cadena["ATENDIDO (FO)"], errors="coerce").dropna()
    pend_vals = pd.to_numeric(df_cadena["PENDIENTE (PA)"], errors="coerce").dropna()
    ej_vals = pd.to_numeric(df_cadena["EN EJECUCION (AS+EE)"], errors="coerce").dropna()

    prom_rec = float(rec_vals.mean()) if not rec_vals.empty else 0.0
    prom_at = float(at_vals.mean()) if not at_vals.empty else 0.0
    tot_rec = int(rec_vals.sum()) if not rec_vals.empty else 0
    tot_at = int(at_vals.sum()) if not at_vals.empty else 0
    ult_pend = int(pend_vals.iloc[-1]) if not pend_vals.empty else 0
    ult_ej = int(ej_vals.iloc[-1]) if not ej_vals.empty else 0

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("⏳ Pendiente (PA)", f"{ult_pend:,}")
    m2.metric("📥 Recibido", f"{tot_rec:,}")
    m3.metric("⚙️ En Ejecución (AS+EE)", f"{ult_ej:,}")
    m4.metric("✅ Atendido (FO)", f"{tot_at:,}")
    m5.metric("📊 Total (Pend+Rec)", f"{ult_pend + tot_rec:,}")
    m6.metric("📈 Prom. Recibido/Día", f"{prom_rec:.1f}")
    m7.metric("📉 Prom. Atendido/Día", f"{prom_at:.1f}")

    st.dataframe(df_cadena, use_container_width=True, hide_index=True)

    tope_barra = (
        pd.to_numeric(df_cadena["PENDIENTE (PA)"], errors="coerce").fillna(0)
        + pd.to_numeric(df_cadena["RECIBIDO"], errors="coerce").fillna(0)
        + pd.to_numeric(df_cadena["EN EJECUCION (AS+EE)"], errors="coerce").fillna(0)
    )
    tope_barra = tope_barra.where(df_cadena["TOTAL (PEND+REC)"].notna(), None)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_cadena["TURNO (07:00 AM)"], y=df_cadena["PENDIENTE (PA)"],
        name="Pendiente (PA)", marker_color="#fca5a5",
        text=df_cadena["PENDIENTE (PA)"], textposition="inside",
    ))
    fig.add_trace(go.Bar(
        x=df_cadena["TURNO (07:00 AM)"], y=df_cadena["RECIBIDO"],
        name="Recibido", marker_color="#e9f056",
        text=df_cadena["RECIBIDO"], textposition="inside",
    ))
    fig.add_trace(go.Bar(
        x=df_cadena["TURNO (07:00 AM)"], y=df_cadena["EN EJECUCION (AS+EE)"],
        name="En Ejecución (AS+EE)", marker_color="#93c5fd",
        text=df_cadena["EN EJECUCION (AS+EE)"], textposition="inside",
    ))
    fig.add_trace(go.Scatter(
        x=df_cadena["TURNO (07:00 AM)"], y=tope_barra,
        name="Total (Pendiente + Recibido)", mode="text",
        text=df_cadena["TOTAL (PEND+REC)"], textposition="top center",
        textfont=dict(color="#1F4E78", size=12), showlegend=False, cliponaxis=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_cadena["TURNO (07:00 AM)"], y=df_cadena["ATENDIDO (FO)"],
        name="Atendido (FO)", mode="lines+markers+text",
        text=df_cadena["ATENDIDO (FO)"], textposition="bottom center",
        marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
        connectgaps=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_cadena["TURNO (07:00 AM)"],
        y=[round(prom_rec, 1) if pd.notna(v) else None for v in df_cadena["RECIBIDO"]],
        name=f"Prom. Recibido ({prom_rec:.1f}/día)",
        mode="lines", line=dict(color="#d97706", width=2.5, dash="dash"), connectgaps=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_cadena["TURNO (07:00 AM)"],
        y=[round(prom_at, 1) if pd.notna(v) else None for v in df_cadena["ATENDIDO (FO)"]],
        name=f"Prom. Atendido ({prom_at:.1f}/día)",
        mode="lines", line=dict(color="#16a34a", width=2.5, dash="dot"), connectgaps=False,
    ))
    fig.update_layout(
        title=f"<b>Revisión en Cadena por Turno (07:00 AM - 07:00 AM) — {mes_nombre} {anio_sel}</b>",
        barmode="stack", hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fecha_aud = st.date_input("Día de inicio del turno (07:00 AM):", value=datetime(2026, 11, 24))
    res_aud = evaluar_turno_cadena(df, fecha_aud)
    st.markdown(
        f"**Ventana evaluada:** `{res_aud['inicio_turno'].strftime('%d/%m/%Y %I:%M %p')}` → `{res_aud['fin_turno'].strftime('%d/%m/%Y %I:%M %p')}`"
    )
    a1, a2, a3, a4, a5 = st.columns(5)
    a1.metric("Pendiente (PA)", f"{res_aud['pendiente']:,}")
    a2.metric("Recibido", f"{res_aud['recibido']:,}")
    a3.metric("En Ejecución (AS+EE)", f"{res_aud['ejecucion']:,}")
    a4.metric("Atendido (FO)", f"{res_aud['atendido']:,}")
    a5.metric("Total (Pend+Rec)", f"{res_aud['total']:,}")

    st.markdown("#### ⏳ Pendiente (PA)")
    mostrar_tabla(res_aud["df_pendiente"], "pend")
    st.markdown("#### 📥 Recibido en Turno")
    mostrar_tabla(res_aud["df_recibido"], "rec")
    st.markdown("#### ⚙️ En Ejecución (AS + EE)")
    mostrar_tabla(res_aud["df_ejecucion"], "ej")
    st.markdown("#### ✅ Atendido (FO)")
    mostrar_tabla(res_aud["df_atendido"], "at")

with tab3:
    if not df_cadena.empty:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_cadena.to_excel(writer, index=False, sheet_name="Cadena_7AM")
        st.download_button(
            "⬇️ Descargar Resumen de Cadena (.xlsx)",
            buf.getvalue(),
            file_name=f"resumen_cadena_{mes_nombre.lower()}_{anio_sel}.xlsx",
        )
