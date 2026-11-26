import calendar
from datetime import datetime
import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Incidencias 2026 (Corte 6:00 AM)",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Control y Seguimiento de Incidencias — Corte 6:00 AM")

ARCHIVO_EXCEL = "resultado_unificado.xlsx"


# ==========================================
# 1. CARGA DE LAS 3 HOJAS DEL EXCEL
# ==========================================
@st.cache_data
def cargar_excel_completo():
    if not os.path.exists(ARCHIVO_EXCEL):
        return None, None, None

    xl = pd.ExcelFile(ARCHIVO_EXCEL)
    hojas = {h.strip().lower(): h for h in xl.sheet_names}

    # 1. Hoja Consolidado Zonas
    df_consolidado = pd.DataFrame()
    nombre_cons = hojas.get("consolidado zonas")
    if nombre_cons:
        df_consolidado = pd.read_excel(xl, sheet_name=nombre_cons, dtype=str)
        for col in df_consolidado.columns:
            df_consolidado[col] = df_consolidado[col].fillna("").astype(str).str.strip()
            df_consolidado[col] = df_consolidado[col].replace(["NAN", "NONE", "NAT"], "", regex=False)
        if "ESTADO" in df_consolidado.columns:
            df_consolidado["ESTADO"] = df_consolidado["ESTADO"].str.upper()
        if "ZONA" in df_consolidado.columns:
            df_consolidado["ZONA"] = df_consolidado["ZONA"].str.upper()
        if "RECEPCIÓN" in df_consolidado.columns:
            df_consolidado["RECEPCION_DT"] = pd.to_datetime(
                df_consolidado["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
            )
        if "FINALIZACIÓN" in df_consolidado.columns:
            df_consolidado["FINALIZACION_DT"] = pd.to_datetime(
                df_consolidado["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
            )

    # 2. Hoja Resumen Diario 6AM
    df_resumen = pd.DataFrame()
    nombre_res = hojas.get("resumen diario 6am")
    if nombre_res:
        df_resumen = pd.read_excel(xl, sheet_name=nombre_res)
        df_resumen.columns = [str(c).strip().upper() for c in df_resumen.columns]
        if "FECHA OPERATIVA" in df_resumen.columns:
            df_resumen["FECHA_DT"] = pd.to_datetime(
                df_resumen["FECHA OPERATIVA"].astype(str).str.strip(),
                format="%d-%m-%Y",
                errors="coerce",
            )
            # Respaldo si viene en formato YYYY-MM-DD
            mask_nat = df_resumen["FECHA_DT"].isna()
            if mask_nat.any():
                df_resumen.loc[mask_nat, "FECHA_DT"] = pd.to_datetime(
                    df_resumen.loc[mask_nat, "FECHA OPERATIVA"], errors="coerce"
                )
            df_resumen = df_resumen.dropna(subset=["FECHA_DT"]).sort_values("FECHA_DT").reset_index(drop=True)
            df_resumen["ANIO"] = df_resumen["FECHA_DT"].dt.year
            df_resumen["MES"] = df_resumen["FECHA_DT"].dt.month

        cols_num = [
            "PENDIENTES INICIO (6 AM)",
            "RECIBIDOS",
            "FINALIZADOS",
            "EN EJECUCIÓN CIERRE (6 AM)",
            "PENDIENTES CIERRE (6 AM)",
        ]
        for c in cols_num:
            if c in df_resumen.columns:
                df_resumen[c] = pd.to_numeric(df_resumen[c], errors="coerce").fillna(0).astype(int)

        if "PENDIENTES INICIO (6 AM)" in df_resumen.columns and "RECIBIDOS" in df_resumen.columns:
            df_resumen["TOTAL"] = df_resumen["PENDIENTES INICIO (6 AM)"] + df_resumen["RECIBIDOS"]

    # 3. Hoja Matriz Zonas 6AM
    df_matriz = pd.DataFrame()
    nombre_mat = hojas.get("matriz zonas 6am")
    if nombre_mat:
        df_matriz = pd.read_excel(xl, sheet_name=nombre_mat)
        df_matriz.columns = [str(c).strip() for c in df_matriz.columns]
        if "ZONA" in df_matriz.columns:
            df_matriz["ZONA"] = df_matriz["ZONA"].fillna("").astype(str).str.strip().str.upper()

    return df_consolidado, df_resumen, df_matriz


df_consolidado, df_resumen, df_matriz = cargar_excel_completo()

if df_consolidado is None:
    st.warning(f"⚠️ No se encontró el archivo `{ARCHIVO_EXCEL}`. Genera el archivo desde tu procesador de TXT y recarga.")
    st.stop()

meses_es = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
meses_upper = {k: v.upper() for k, v in meses_es.items()}


# ==========================================
# 2. PESTAÑAS PRINCIPALES
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📅 Seguimiento Diario (6 AM)",
    "📈 Seguimiento Anual",
    "⏳ PA por Motivo",
    "🗺️ Matriz Zonas (6 AM)",
    "🔍 Buscador Consolidado",
])


# ------------------------------------------
# TAB 1: SEGUIMIENTO DIARIO (Resumen Diario 6AM)
# ------------------------------------------
with tab1:
    st.subheader("📅 Comportamiento Diario — Corte 6:00 AM")
    if df_resumen.empty:
        st.info("La hoja 'Resumen Diario 6AM' no tiene registros.")
    else:
        anios_disp = sorted(df_resumen["ANIO"].dropna().unique().astype(int).tolist())
        idx_anio = anios_disp.index(2026) if 2026 in anios_disp else len(anios_disp) - 1

        col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
        with col_f1:
            anio_sel = st.selectbox("Año:", anios_disp, index=idx_anio, key="t1_anio")

        meses_disp = sorted(
            df_resumen[df_resumen["ANIO"] == anio_sel]["MES"].dropna().unique().astype(int).tolist()
        )
        nombres_meses_disp = [meses_es[m] for m in meses_disp]
        idx_mes = nombres_meses_disp.index("Noviembre") if "Noviembre" in nombres_meses_disp else len(nombres_meses_disp) - 1

        with col_f2:
            mes_nombre_sel = st.selectbox("Mes:", nombres_meses_disp, index=idx_mes, key="t1_mes")
            mes_sel = [k for k, v in meses_es.items() if v == mes_nombre_sel][0]

        df_mes = df_resumen[
            (df_resumen["ANIO"] == anio_sel) & (df_resumen["MES"] == mes_sel)
        ].copy()

        # Filtro opcional por rango de días dentro del mes
        if not df_mes.empty:
            min_f = df_mes["FECHA_DT"].min().date()
            max_f = df_mes["FECHA_DT"].max().date()
            with col_f3:
                rango_fechas = st.date_input(
                    "Rango de fechas:",
                    value=(min_f, max_f),
                    min_value=min_f,
                    max_value=max_f,
                    key="t1_rango",
                )
            if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
                df_mes = df_mes[
                    (df_mes["FECHA_DT"].dt.date >= rango_fechas[0])
                    & (df_mes["FECHA_DT"].dt.date <= rango_fechas[1])
                ].copy()

        if not df_mes.empty:
            pend_inicio_mes = int(df_mes["PENDIENTES INICIO (6 AM)"].iloc[0])
            recibidos_mes = int(df_mes["RECIBIDOS"].sum())
            finalizados_mes = int(df_mes["FINALIZADOS"].sum())
            ee_cierre_ultimo = int(df_mes["EN EJECUCIÓN CIERRE (6 AM)"].iloc[-1])
            pend_cierre_ultimo = int(df_mes["PENDIENTES CIERRE (6 AM)"].iloc[-1])
            total_mes = pend_inicio_mes + recibidos_mes

            prom_rec = float(df_mes["RECIBIDOS"].mean())
            prom_fin = float(df_mes["FINALIZADOS"].mean())

            c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
            c1.metric("📦 Pend. Inicio (6 AM)", f"{pend_inicio_mes:,}")
            c2.metric("📥 Recibidos", f"{recibidos_mes:,}")
            c3.metric("📊 Total (Ini+Rec)", f"{total_mes:,}")
            c4.metric("✅ Finalizados", f"{finalizados_mes:,}")
            c5.metric("⏳ Pend. Cierre (6 AM)", f"{pend_cierre_ultimo:,}")
            c6.metric("⚙️ En Ejec. Cierre (6 AM)", f"{ee_cierre_ultimo:,}")
            c7.metric("📈 Prom. Rec / Atend", f"{prom_rec:.1f} / {prom_fin:.1f}")
            st.divider()

            df_mes["FECHA_STR"] = df_mes["FECHA_DT"].dt.strftime("%d/%m/%Y")
            tope_barra = (
                df_mes["PENDIENTES INICIO (6 AM)"]
                + df_mes["RECIBIDOS"]
                + df_mes["EN EJECUCIÓN CIERRE (6 AM)"]
            )

            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["PENDIENTES INICIO (6 AM)"],
                name="Pendientes Inicio (6 AM)", marker_color="#fca5a5",
                text=df_mes["PENDIENTES INICIO (6 AM)"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["RECIBIDOS"],
                name="Recibidos", marker_color="#e9f056",
                text=df_mes["RECIBIDOS"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["EN EJECUCIÓN CIERRE (6 AM)"],
                name="En Ejecución Cierre (6 AM)", marker_color="#93c5fd",
                text=df_mes["EN EJECUCIÓN CIERRE (6 AM)"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig1.add_trace(go.Scatter(
                x=df_mes["FECHA_STR"], y=tope_barra,
                name="Total (Pend. Inicio + Recibidos)", mode="text",
                text=df_mes["TOTAL"], textposition="top center",
                textfont=dict(color="#1F4E78", size=12),
                showlegend=False, cliponaxis=False,
            ))
            fig1.add_trace(go.Scatter(
                x=df_mes["FECHA_STR"], y=df_mes["FINALIZADOS"],
                name="Finalizados", mode="lines+markers+text",
                text=df_mes["FINALIZADOS"], textposition="bottom center",
                textfont=dict(color="#1d4ed8", size=11),
                marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
            ))
            fig1.add_trace(go.Scatter(
                x=df_mes["FECHA_STR"], y=[round(prom_rec, 1)] * len(df_mes),
                name=f"Prom. Recibidos ({prom_rec:.1f}/día)",
                mode="lines", line=dict(color="#d97706", width=2.5, dash="dash"),
            ))
            fig1.add_trace(go.Scatter(
                x=df_mes["FECHA_STR"], y=[round(prom_fin, 1)] * len(df_mes),
                name=f"Prom. Finalizados ({prom_fin:.1f}/día)",
                mode="lines", line=dict(color="#16a34a", width=2.5, dash="dot"),
            ))
            fig1.update_layout(
                title=f"<b>Flujo Diario de Incidencias (Corte 6:00 AM) — {mes_nombre_sel} {anio_sel}</b>",
                barmode="stack",
                xaxis_title="<b>Fecha Operativa</b>",
                yaxis_title="<b>Cantidad</b>",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
            )
            st.plotly_chart(fig1, use_container_width=True)

            cols_tabla_res = [
                "FECHA OPERATIVA",
                "VENTANA TURNO",
                "PENDIENTES INICIO (6 AM)",
                "RECIBIDOS",
                "TOTAL",
                "FINALIZADOS",
                "EN EJECUCIÓN CIERRE (6 AM)",
                "PENDIENTES CIERRE (6 AM)",
            ]
            cols_tabla_res = [c for c in cols_tabla_res if c in df_mes.columns]
            st.dataframe(df_mes[cols_tabla_res], use_container_width=True, hide_index=True)

            buf_d = io.BytesIO()
            with pd.ExcelWriter(buf_d, engine="openpyxl") as writer:
                df_mes[cols_tabla_res].to_excel(writer, index=False, sheet_name="Seguimiento_Diario")
            st.download_button(
                "⬇️ Descargar Seguimiento Diario (.xlsx)",
                buf_d.getvalue(),
                file_name=f"seguimiento_diario_{mes_nombre_sel.lower()}_{anio_sel}.xlsx",
            )


# ------------------------------------------
# TAB 2: SEGUIMIENTO ANUAL (Resumen Diario 6AM agrupado por mes)
# ------------------------------------------
with tab2:
    st.subheader("📈 Seguimiento Mensual Acumulado")
    if df_resumen.empty:
        st.info("La hoja 'Resumen Diario 6AM' no tiene registros.")
    else:
        anios_disp_y = sorted(df_resumen["ANIO"].dropna().unique().astype(int).tolist())
        idx_anio_y = anios_disp_y.index(2026) if 2026 in anios_disp_y else len(anios_disp_y) - 1

        col_y1, _ = st.columns([1, 3])
        with col_y1:
            anio_anual = st.selectbox("Seleccione el Año:", anios_disp_y, index=idx_anio_y, key="t2_anio")

        df_y = df_resumen[df_resumen["ANIO"] == anio_anual].sort_values("FECHA_DT")
        filas_anuales = []

        for m, grp in df_y.groupby("MES", sort=True):
            pend_ini_m = int(grp["PENDIENTES INICIO (6 AM)"].iloc[0])
            rec_m = int(grp["RECIBIDOS"].sum())
            fin_m = int(grp["FINALIZADOS"].sum())
            ee_cierre_m = int(grp["EN EJECUCIÓN CIERRE (6 AM)"].iloc[-1])
            pend_cierre_m = int(grp["PENDIENTES CIERRE (6 AM)"].iloc[-1])
            tot_m = pend_ini_m + rec_m

            filas_anuales.append({
                "MES_NUM": int(m),
                "MES": meses_upper[int(m)],
                "PENDIENTES INICIO (6 AM)": pend_ini_m,
                "RECIBIDOS": rec_m,
                "TOTAL (INI+REC)": tot_m,
                "FINALIZADOS": fin_m,
                "EN EJECUCIÓN CIERRE (6 AM)": ee_cierre_m,
                "PENDIENTES CIERRE (6 AM)": pend_cierre_m,
            })

        df_anual = pd.DataFrame(filas_anuales).sort_values("MES_NUM")

        if not df_anual.empty:
            st.dataframe(df_anual.drop(columns=["MES_NUM"]), use_container_width=True, hide_index=True)

            fig_anual = go.Figure()
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["PENDIENTES INICIO (6 AM)"],
                name="Pendientes Inicio (6 AM)", marker_color="#fca5a5",
                text=df_anual["PENDIENTES INICIO (6 AM)"], textposition="inside",
            ))
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["RECIBIDOS"],
                name="Recibidos", marker_color="#e9f056",
                text=df_anual["RECIBIDOS"], textposition="inside",
            ))
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["EN EJECUCIÓN CIERRE (6 AM)"],
                name="En Ejecución Cierre (6 AM)", marker_color="#93c5fd",
                text=df_anual["EN EJECUCIÓN CIERRE (6 AM)"], textposition="inside",
            ))
            tope_anual = (
                df_anual["PENDIENTES INICIO (6 AM)"]
                + df_anual["RECIBIDOS"]
                + df_anual["EN EJECUCIÓN CIERRE (6 AM)"]
            )
            fig_anual.add_trace(go.Scatter(
                x=df_anual["MES"], y=tope_anual,
                name="Total (Ini+Rec)", mode="text",
                text=df_anual["TOTAL (INI+REC)"], textposition="top center",
                textfont=dict(color="#1F4E78", size=12), showlegend=False, cliponaxis=False,
            ))
            fig_anual.add_trace(go.Scatter(
                x=df_anual["MES"], y=df_anual["FINALIZADOS"],
                name="Finalizados", mode="lines+markers+text",
                text=df_anual["FINALIZADOS"], textposition="bottom center",
                marker=dict(size=6, color="#1d4ed8"), line=dict(color="#1d4ed8", width=2),
            ))
            fig_anual.update_layout(
                title=f"<b>Evolución Mensual Acumulada — {anio_anual}</b>",
                barmode="stack",
                xaxis_title="<b>Mes</b>",
                yaxis_title="<b>Cantidad</b>",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
            )
            st.plotly_chart(fig_anual, use_container_width=True)


# ------------------------------------------
# TAB 3: PENDIENTES (PA) POR MOTIVO (Consolidado Zonas)
# ------------------------------------------
with tab3:
    st.subheader("⏳ Incidencias Pendientes (PA) por Motivo")
    if df_consolidado.empty:
        st.info("La hoja 'Consolidado Zonas' está vacía.")
    else:
        df_pa = df_consolidado[df_consolidado["ESTADO"] == "PA"].copy()

        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            zonas_pa = ["TODAS"] + sorted([z for z in df_pa["ZONA"].unique() if z])
            zona_pa_sel = st.selectbox("Filtrar por Zona:", zonas_pa, key="t3_zona")
        with col_p2:
            anios_pa = ["TODOS"] + sorted([a for a in df_pa.get("AÑO RECIBIDO", pd.Series()).unique() if a])
            anio_pa_sel = st.selectbox("Año Recibido:", anios_pa, key="t3_anio")
        with col_p3:
            meses_pa = ["TODOS"] + [m for m in meses_upper.values() if m in df_pa.get("MES RECIBIDO", pd.Series()).unique()]
            mes_pa_sel = st.selectbox("Mes Recibido:", meses_pa, key="t3_mes")

        if zona_pa_sel != "TODAS":
            df_pa = df_pa[df_pa["ZONA"] == zona_pa_sel]
        if anio_pa_sel != "TODOS" and "AÑO RECIBIDO" in df_pa.columns:
            df_pa = df_pa[df_pa["AÑO RECIBIDO"] == anio_pa_sel]
        if mes_pa_sel != "TODOS" and "MES RECIBIDO" in df_pa.columns:
            df_pa = df_pa[df_pa["MES RECIBIDO"] == mes_pa_sel]

        st.metric("⏳ Total Incidencias PA Filtradas", f"{len(df_pa):,}")

        if not df_pa.empty and "MOTIVO" in df_pa.columns:
            motivos = df_pa["MOTIVO"].replace("", "SIN MOTIVO")
            res_motivos = motivos.value_counts().reset_index()
            res_motivos.columns = ["MOTIVO", "PENDIENTES (PA)"]
            tot_pa = int(res_motivos["PENDIENTES (PA)"].sum())
            res_motivos["% PARTICIPACIÓN"] = (res_motivos["PENDIENTES (PA)"] / tot_pa * 100).round(1)

            tabla_mot = res_motivos.copy()
            tabla_mot["% PARTICIPACIÓN"] = tabla_mot["% PARTICIPACIÓN"].map(lambda x: f"{x:.1f}%")
            tabla_mot = pd.concat([
                tabla_mot,
                pd.DataFrame([{"MOTIVO": "TOTAL GENERAL", "PENDIENTES (PA)": tot_pa, "% PARTICIPACIÓN": "100.0%"}]),
            ], ignore_index=True)

            c_tbl, c_fig = st.columns([1.1, 1.4])
            with c_tbl:
                st.markdown("**Cuadro Resumen de PA por Motivo**")
                st.dataframe(tabla_mot, use_container_width=True, hide_index=True)

            with c_fig:
                res_ord = res_motivos.sort_values("PENDIENTES (PA)", ascending=True)
                fig_pa_mot = px.bar(
                    res_ord,
                    x="PENDIENTES (PA)",
                    y="MOTIVO",
                    orientation="h",
                    text="PENDIENTES (PA)",
                    title="Distribución de Pendientes (PA) por Motivo",
                )
                fig_pa_mot.update_traces(marker_color="#b91c1c", textposition="outside", cliponaxis=False)
                fig_pa_mot.update_layout(
                    xaxis_title="<b>Cantidad PA</b>",
                    yaxis_title="<b>Motivo</b>",
                    height=max(320, len(res_ord) * 34),
                    margin=dict(r=40),
                )
                st.plotly_chart(fig_pa_mot, use_container_width=True)

            with st.expander("Ver listado detallado de incidencias PA"):
                cols_pa_det = [c for c in ["INCIDENCIA", "ZONA", "RECEPCIÓN", "MOTIVO", "DIRECCIÓN", "CLIENTE", "TEC", "TCR"] if c in df_pa.columns]
                st.dataframe(df_pa[cols_pa_det], use_container_width=True, hide_index=True)
        else:
            st.info("No hay incidencias en estado PA para los filtros seleccionados.")


# ------------------------------------------
# TAB 4: MATRIZ ZONAS (6 AM)
# ------------------------------------------
with tab4:
    st.subheader("🗺️ Matriz de Pendientes por Zona (Corte 6:00 AM)")
    if df_matriz.empty:
        st.info("La hoja 'Matriz Zonas 6AM' no tiene datos.")
    else:
        cols_fecha_mat = [c for c in df_matriz.columns if c != "ZONA"]
        # Pasar a formato largo para filtrar por mes/año y graficar
        df_mat_zonas = df_matriz[df_matriz["ZONA"] != "TOTALES"].copy()
        df_long = df_mat_zonas.melt(id_vars=["ZONA"], value_vars=cols_fecha_mat, var_name="FECHA_STR", value_name="CANTIDAD")
        df_long["FECHA_DT"] = pd.to_datetime(df_long["FECHA_STR"], format="%d/%m/%Y", errors="coerce")
        df_long["CANTIDAD"] = pd.to_numeric(df_long["CANTIDAD"], errors="coerce").fillna(0).astype(int)
        df_long = df_long.dropna(subset=["FECHA_DT"])

        if not df_long.empty:
            df_long["ANIO"] = df_long["FECHA_DT"].dt.year
            df_long["MES"] = df_long["FECHA_DT"].dt.month

            cm1, cm2, cm3 = st.columns(3)
            anios_m = sorted(df_long["ANIO"].unique().tolist())
            idx_am = anios_m.index(2026) if 2026 in anios_m else len(anios_m) - 1
            with cm1:
                anio_mat = st.selectbox("Año:", anios_m, index=idx_am, key="t4_anio")
            meses_m = sorted(df_long[df_long["ANIO"] == anio_mat]["MES"].unique().tolist())
            nombres_mm = [meses_es[m] for m in meses_m]
            idx_mm = nombres_mm.index("Noviembre") if "Noviembre" in nombres_mm else len(nombres_mm) - 1
            with cm2:
                mes_mat_nombre = st.selectbox("Mes:", nombres_mm, index=idx_mm, key="t4_mes")
                mes_mat = [k for k, v in meses_es.items() if v == mes_mat_nombre][0]
            with cm3:
                zonas_m = ["TODAS"] + sorted(df_long["ZONA"].unique().tolist())
                zona_mat = st.selectbox("Zona:", zonas_m, key="t4_zona")

            df_long_f = df_long[(df_long["ANIO"] == anio_mat) & (df_long["MES"] == mes_mat)].copy()
            if zona_mat != "TODAS":
                df_long_f = df_long_f[df_long_f["ZONA"] == zona_mat]

            if not df_long_f.empty:
                fig_mat = px.line(
                    df_long_f.sort_values("FECHA_DT"),
                    x="FECHA_DT",
                    y="CANTIDAD",
                    color="ZONA",
                    markers=True,
                    title=f"<b>Comportamiento Diario por Zona (6:00 AM) — {mes_mat_nombre} {anio_mat}</b>",
                )
                fig_mat.update_layout(xaxis_title="<b>Fecha</b>", yaxis_title="<b>Cantidad</b>", hovermode="x unified")
                st.plotly_chart(fig_mat, use_container_width=True)

            # Mostrar tabla pivoteada filtrada por el mes seleccionado
            fechas_cols_mes = (
                df_long[(df_long["ANIO"] == anio_mat) & (df_long["MES"] == mes_mat)]
                .sort_values("FECHA_DT")["FECHA_STR"]
                .unique()
                .tolist()
            )
            cols_ver = ["ZONA"] + [c for c in fechas_cols_mes if c in df_matriz.columns]
            df_matriz_vista = df_matriz[cols_ver].copy()
            if zona_mat != "TODAS":
                df_matriz_vista = df_matriz_vista[df_matriz_vista["ZONA"].isin([zona_mat, "TOTALES"])]
            st.dataframe(df_matriz_vista, use_container_width=True, hide_index=True)


# ------------------------------------------
# TAB 5: BUSCADOR EN CONSOLIDADO ZONAS
# ------------------------------------------
with tab5:
    st.subheader("🔍 Buscador en Consolidado Zonas")
    if df_consolidado.empty:
        st.info("No hay datos en 'Consolidado Zonas'.")
    else:
        cb1, cb2, cb3, cb4 = st.columns(4)
        with cb1:
            f_inc = st.text_input("Nº Incidencia:", placeholder="Ej. 12345")
        with cb2:
            f_zona = st.selectbox("Zona:", ["TODAS"] + sorted([z for z in df_consolidado["ZONA"].unique() if z]))
        with cb3:
            f_est = st.selectbox("Estado:", ["TODOS"] + sorted([e for e in df_consolidado["ESTADO"].unique() if e]))
        with cb4:
            f_mot = st.text_input("Motivo:", placeholder="Ej. SIN SERVICIO")

        df_busq = df_consolidado.copy()
        if f_inc:
            df_busq = df_busq[df_busq["INCIDENCIA"].str.contains(f_inc, case=False, na=False)]
        if f_zona != "TODAS":
            df_busq = df_busq[df_busq["ZONA"] == f_zona]
        if f_est != "TODOS":
            df_busq = df_busq[df_busq["ESTADO"] == f_est]
        if f_mot and "MOTIVO" in df_busq.columns:
            df_busq = df_busq[df_busq["MOTIVO"].str.contains(f_mot, case=False, na=False)]

        st.write(f"Registros encontrados: **{len(df_busq):,}**")
        cols_busq = [
            c for c in [
                "INCIDENCIA", "ZONA", "RECEPCIÓN", "FINALIZACIÓN", "ESTADO",
                "MOTIVO", "DIRECCIÓN", "CLIENTE", "TEC", "TCR", "TDV", "TEJ", "TAR",
            ] if c in df_busq.columns
        ]
        st.dataframe(df_busq[cols_busq].head(1000), use_container_width=True, hide_index=True)
