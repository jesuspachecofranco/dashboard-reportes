import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Trazabilidad de Incidencias (07:00 AM - 07:00 AM)",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Trazabilidad Día a Día por Turno (07:00 AM - 07:00 AM)")


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

    df["H_PA_DT"] = _parsear_dt(df, "H_PA")
    if df["H_PA_DT"].isna().all():
        df["H_PA_DT"] = df["RECEPCION_DT"]

    df["H_EE_DT"] = _parsear_dt(df, "H_EE")

    df["H_FO_DT"] = _parsear_dt(df, "H_FO")
    if df["H_FO_DT"].isna().all():
        df["H_FO_DT"] = df["FINALIZACION_DT"]

    return df


def trazar_turno(df_data, fecha_dia):
    """
    Reconstruye el comportamiento del turno [07:00 AM día D -> 07:00 AM día D+1]
    con evaluación independiente al inicio (07:00 AM) y al cierre (07:00 AM siguiente),
    permitiendo que una incidencia abierta en PA y cerrada en EE cuente en ambos.
    """
    base = pd.Timestamp(year=fecha_dia.year, month=fecha_dia.month, day=fecha_dia.day)
    inicio_turno = base + pd.Timedelta(hours=7)
    fin_turno = inicio_turno + pd.Timedelta(days=1)

    rec_dt = df_data["RECEPCION_DT"]
    h_pa_dt = df_data["H_PA_DT"]
    h_ee_dt = df_data["H_EE_DT"]
    h_fo_dt = df_data["H_FO_DT"]

    # 1. Pendiente al iniciar el corte (07:00 AM):
    # Estaba en PA al momento de abrir el turno (aún no había pasado a EE ni a FO a las 07:00 AM)
    entro_pa_inicio = h_pa_dt.notna() & (h_pa_dt <= inicio_turno)
    no_en_ee_inicio = h_ee_dt.isna() | (h_ee_dt > inicio_turno)
    no_atendida_inicio = h_fo_dt.isna() | (h_fo_dt > inicio_turno)
    mask_pendiente_inicio = entro_pa_inicio & no_en_ee_inicio & no_atendida_inicio

    # 2. Recibido durante el turno [inicio_turno, fin_turno)
    mask_recibido_turno = rec_dt.notna() & (rec_dt >= inicio_turno) & (rec_dt < fin_turno)

    # 3. Atendido al cerrar el turno (07:00 AM del día siguiente):
    # Hora de atención/finalización dentro de [inicio_turno, fin_turno)
    mask_atendido_cierre = h_fo_dt.notna() & (h_fo_dt >= inicio_turno) & (h_fo_dt < fin_turno)

    # 4. En Ejecución al cerrar el turno (07:00 AM del día siguiente) - CÁLCULO INDEPENDIENTE:
    # Había entrado a ejecución antes o al cierre del turno (H_EE_DT <= fin_turno)
    # y aún no estaba atendida al cierre del turno (H_FO_DT.isna() | (H_FO_DT > fin_turno))
    entro_ee_cierre = h_ee_dt.notna() & (h_ee_dt <= fin_turno)
    no_atendida_cierre = h_fo_dt.isna() | (h_fo_dt > fin_turno)
    mask_ee_cierre = entro_ee_cierre & no_atendida_cierre

    # 5. Total: Pendiente al iniciar + Recibido en el turno (sin sumar En Ejecución)
    pendiente = int(mask_pendiente_inicio.sum())
    recibido = int(mask_recibido_turno.sum())
    atendido = int(mask_atendido_cierre.sum())
    en_ejecucion = int(mask_ee_cierre.sum())
    total = pendiente + recibido

    return {
        "inicio_turno": inicio_turno,
        "fin_turno": fin_turno,
        "pendiente_inicio": pendiente,
        "recibido": recibido,
        "atendido_cierre": atendido,
        "ee_cierre": en_ejecucion,
        "total": total,
        "df_pendiente_inicio": df_data[mask_pendiente_inicio].copy(),
        "df_recibido": df_data[mask_recibido_turno].copy(),
        "df_atendido_cierre": df_data[mask_atendido_cierre].copy(),
        "df_ee_cierre": df_data[mask_ee_cierre].copy(),
    }


df = cargar_datos()

if df is None:
    st.warning("⚠️ No se encontró 'resultado_unificado.xlsx'. Genéralo primero con tu script unificador.")
    st.stop()

columnas_traza = [
    "INCIDENCIA", "ZONA", "RECEPCIÓN", "H_PA", "H_EE", "H_FO",
    "FINALIZACIÓN", "ESTADO", "MOTIVO", "DIRECCIÓN",
]


def mostrar_tabla(df_sub, key_suffix):
    if df_sub.empty:
        st.info("No hay incidencias en este corte.")
        return
    cols = [c for c in columnas_traza if c in df_sub.columns]
    st.dataframe(df_sub[cols], use_container_width=True, hide_index=True, key=f"tbl_{key_suffix}")


tab1, tab2, tab3 = st.tabs([
    "📅 Trazabilidad Diaria",
    "🔍 Detalle de un Turno",
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
    filas = []

    for dia in dias_mes:
        turno_ini = dia + pd.Timedelta(hours=7)
        if pd.notna(fecha_max_dato) and turno_ini > (fecha_max_dato + pd.Timedelta(days=1)):
            filas.append({
                "TURNO (07:00 AM)": dia.strftime("%d/%m/%Y"),
                "PENDIENTE INICIO (PA)": None,
                "RECIBIDO": None,
                "ATENDIDO CIERRE (FO)": None,
                "EN EJECUCION CIERRE (EE)": None,
                "TOTAL (PEND+REC)": None,
            })
        else:
            res = trazar_turno(df, dia)
            filas.append({
                "TURNO (07:00 AM)": dia.strftime("%d/%m/%Y"),
                "PENDIENTE INICIO (PA)": res["pendiente_inicio"],
                "RECIBIDO": res["recibido"],
                "ATENDIDO CIERRE (FO)": res["atendido_cierre"],
                "EN EJECUCION CIERRE (EE)": res["ee_cierre"],
                "TOTAL (PEND+REC)": res["total"],
            })

    df_traza = pd.DataFrame(filas)

    rec_vals = pd.to_numeric(df_traza["RECIBIDO"], errors="coerce").dropna()
    at_vals = pd.to_numeric(df_traza["ATENDIDO CIERRE (FO)"], errors="coerce").dropna()
    pend_vals = pd.to_numeric(df_traza["PENDIENTE INICIO (PA)"], errors="coerce").dropna()
    ee_vals = pd.to_numeric(df_traza["EN EJECUCION CIERRE (EE)"], errors="coerce").dropna()

    prom_rec = float(rec_vals.mean()) if not rec_vals.empty else 0.0
    prom_at = float(at_vals.mean()) if not at_vals.empty else 0.0
    tot_rec = int(rec_vals.sum()) if not rec_vals.empty else 0
    tot_at = int(at_vals.sum()) if not at_vals.empty else 0
    ult_pend = int(pend_vals.iloc[-1]) if not pend_vals.empty else 0
    ult_ee = int(ee_vals.iloc[-1]) if not ee_vals.empty else 0

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("⏳ Pendiente Inicio (PA)", f"{ult_pend:,}")
    m2.metric("📥 Recibido (Mes)", f"{tot_rec:,}")
    m3.metric("✅ Atendido Cierre (FO)", f"{tot_at:,}")
    m4.metric("⚙️ En Ejecución Cierre (EE)", f"{ult_ee:,}")
    m5.metric("📊 Total (Pend+Rec)", f"{ult_pend + tot_rec:,}")
    m6.metric("📈 Prom. Recibido/Día", f"{prom_rec:.1f}")
    m7.metric("📉 Prom. Atendido/Día", f"{prom_at:.1f}")

    st.dataframe(df_traza, use_container_width=True, hide_index=True)

    tope_barra = (
        pd.to_numeric(df_traza["PENDIENTE INICIO (PA)"], errors="coerce").fillna(0)
        + pd.to_numeric(df_traza["RECIBIDO"], errors="coerce").fillna(0)
        + pd.to_numeric(df_traza["EN EJECUCION CIERRE (EE)"], errors="coerce").fillna(0)
    )
    tope_barra = tope_barra.where(df_traza["TOTAL (PEND+REC)"].notna(), None)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_traza["TURNO (07:00 AM)"], y=df_traza["PENDIENTE INICIO (PA)"],
        name="Pendiente Inicio (PA)", marker_color="#fca5a5",
        text=df_traza["PENDIENTE INICIO (PA)"], textposition="inside",
    ))
    fig.add_trace(go.Bar(
        x=df_traza["TURNO (07:00 AM)"], y=df_traza["RECIBIDO"],
        name="Recibido", marker_color="#e9f056",
        text=df_traza["RECIBIDO"], textposition="inside",
    ))
    fig.add_trace(go.Bar(
        x=df_traza["TURNO (07:00 AM)"], y=df_traza["EN EJECUCION CIERRE (EE)"],
        name="En Ejecución Cierre (EE)", marker_color="#93c5fd",
        text=df_traza["EN EJECUCION CIERRE (EE)"], textposition="inside",
    ))
    fig.add_trace(go.Scatter(
        x=df_traza["TURNO (07:00 AM)"], y=tope_barra,
        name="Total (Pendiente + Recibido)", mode="text",
        text=df_traza["TOTAL (PEND+REC)"], textposition="top center",
        textfont=dict(color="#1F4E78", size=12), showlegend=False, cliponaxis=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_traza["TURNO (07:00 AM)"], y=df_traza["ATENDIDO CIERRE (FO)"],
        name="Atendido Cierre (FO)", mode="lines+markers+text",
        text=df_traza["ATENDIDO CIERRE (FO)"], textposition="bottom center",
        marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
        connectgaps=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_traza["TURNO (07:00 AM)"],
        y=[round(prom_rec, 1) if pd.notna(v) else None for v in df_traza["RECIBIDO"]],
        name=f"Prom. Recibido ({prom_rec:.1f}/día)",
        mode="lines", line=dict(color="#d97706", width=2.5, dash="dash"), connectgaps=False,
    ))
    fig.add_trace(go.Scatter(
        x=df_traza["TURNO (07:00 AM)"],
        y=[round(prom_at, 1) if pd.notna(v) else None for v in df_traza["ATENDIDO CIERRE (FO)"]],
        name=f"Prom. Atendido ({prom_at:.1f}/día)",
        mode="lines", line=dict(color="#16a34a", width=2.5, dash="dot"), connectgaps=False,
    ))
    fig.update_layout(
        title=f"<b>Trazabilidad por Turno (07:00 AM - 07:00 AM) — {mes_nombre} {anio_sel}</b>",
        barmode="stack", hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fecha_aud = st.date_input("Día del turno (inicia 07:00 AM):", value=datetime(2026, 11, 24))
    res = trazar_turno(df, fecha_aud)
    st.markdown(
        f"**Turno:** `{res['inicio_turno'].strftime('%d/%m/%Y %I:%M %p')}` → `{res['fin_turno'].strftime('%d/%m/%Y %I:%M %p')}`"
    )
    a1, a2, a3, a4, a5 = st.columns(5)
    a1.metric("Pendiente Inicio (PA)", f"{res['pendiente_inicio']:,}")
    a2.metric("Recibido", f"{res['recibido']:,}")
    a3.metric("Atendido Cierre (FO)", f"{res['atendido_cierre']:,}")
    a4.metric("En Ejecución Cierre (EE)", f"{res['ee_cierre']:,}")
    a5.metric("Total (Pend+Rec)", f"{res['total']:,}")

    st.markdown("#### ⏳ Pendiente al Iniciar el Corte (07:00 AM)")
    mostrar_tabla(res["df_pendiente_inicio"], "pend_ini")

    st.markdown("#### 📥 Recibido en el Turno")
    mostrar_tabla(res["df_recibido"], "rec_turn")

    st.markdown("#### ✅ Atendido al Cerrar el Turno (07:00 AM siguiente)")
    mostrar_tabla(res["df_atendido_cierre"], "at_cie")

    st.markdown("#### ⚙️ En Ejecución al Cerrar el Turno (07:00 AM siguiente)")
    mostrar_tabla(res["df_ee_cierre"], "ee_cie")

with tab3:
    if not df_traza.empty:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_traza.to_excel(writer, index=False, sheet_name="Trazabilidad_7AM")
        st.download_button(
            "⬇️ Descargar Trazabilidad (.xlsx)",
            buf.getvalue(),
            file_name=f"trazabilidad_{mes_nombre.lower()}_{anio_sel}.xlsx",
        )
