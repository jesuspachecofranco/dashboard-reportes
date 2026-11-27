import io
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Incidencias (Corte 6:00 AM)",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Control y Seguimiento de Incidencias — Corte 6:00 AM")

ARCHIVO_EXCEL = "resultado_unificado.xlsx"
COLOR_PA_INICIO = "#D9E2F3"
COLOR_RECIBIDOS = "#e9f056"
COLOR_EE_CIERRE = "#93c5fd"
COLOR_FINALIZADOS = "#1d4ed8"

meses_es = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
meses_upper = {k: v.upper() for k, v in meses_es.items()}


# ==========================================
# 1. FUNCIONES DE LIMPIEZA Y CARGA DE HOJAS
# ==========================================
def _limpiar_consolidado(df, col_entidad):
    if df.empty:
        return df
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
        df[col] = df[col].replace(["NAN", "NONE", "NAT"], "", regex=False)
    if "ESTADO" in df.columns:
        df["ESTADO"] = df["ESTADO"].str.upper()
    if col_entidad in df.columns:
        df[col_entidad] = df[col_entidad].str.upper()
    if "RECEPCIÓN" in df.columns:
        df["RECEPCION_DT"] = pd.to_datetime(
            df["RECEPCIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
        )
    if "FINALIZACIÓN" in df.columns:
        df["FINALIZACION_DT"] = pd.to_datetime(
            df["FINALIZACIÓN"], format="%d-%m-%y %I:%M %p", errors="coerce"
        )
    return df


def _limpiar_resumen(df, col_entidad=None):
    if df.empty:
        return df
    df.columns = [str(c).strip().upper() for c in df.columns]
    if col_entidad and col_entidad in df.columns:
        df[col_entidad] = df[col_entidad].fillna("").astype(str).str.strip().str.upper()
    if "FECHA OPERATIVA" in df.columns:
        df["FECHA_DT"] = pd.to_datetime(
            df["FECHA OPERATIVA"].astype(str).str.strip(),
            format="%d-%m-%Y",
            errors="coerce",
        )
        mask_nat = df["FECHA_DT"].isna()
        if mask_nat.any():
            df.loc[mask_nat, "FECHA_DT"] = pd.to_datetime(
                df.loc[mask_nat, "FECHA OPERATIVA"], errors="coerce"
            )
        df = df.dropna(subset=["FECHA_DT"]).sort_values("FECHA_DT").reset_index(drop=True)
        df["ANIO"] = df["FECHA_DT"].dt.year
        df["MES"] = df["FECHA_DT"].dt.month

    cols_num = [
        "PENDIENTES INICIO (6 AM)",
        "RECIBIDOS",
        "FINALIZADOS",
        "EN EJECUCIÓN CIERRE (6 AM)",
        "PENDIENTES CIERRE (6 AM)",
    ]
    for c in cols_num:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    if "PENDIENTES INICIO (6 AM)" in df.columns and "RECIBIDOS" in df.columns:
        df["TOTAL"] = df["PENDIENTES INICIO (6 AM)"] + df["RECIBIDOS"]

    return df


def _limpiar_matriz(df, col_entidad):
    if df.empty:
        return df
    df.columns = [str(c).strip() for c in df.columns]
    if col_entidad in df.columns:
        df[col_entidad] = df[col_entidad].fillna("").astype(str).str.strip().str.upper()
    return df


@st.cache_data
def cargar_excel_completo():
    if not os.path.exists(ARCHIVO_EXCEL):
        return None

    xl = pd.ExcelFile(ARCHIVO_EXCEL)
    hojas = {h.strip().lower(): h for h in xl.sheet_names}

    datos = {
        "iribarren": {
            "consolidado": pd.DataFrame(),
            "resumen": pd.DataFrame(),
            "matriz": pd.DataFrame(),
        },
        "foraneos": {
            "consolidado": pd.DataFrame(),
            "resumen": pd.DataFrame(),
            "matriz": pd.DataFrame(),
        },
    }

    # --- BLOQUE IRIBARREN (Hojas de Zonas) ---
    if "consolidado zonas" in hojas:
        datos["iribarren"]["consolidado"] = _limpiar_consolidado(
            pd.read_excel(xl, sheet_name=hojas["consolidado zonas"], dtype=str), "ZONA"
        )
    if "resumen diario 6am" in hojas:
        datos["iribarren"]["resumen"] = _limpiar_resumen(
            pd.read_excel(xl, sheet_name=hojas["resumen diario 6am"])
        )
    if "matriz zonas 6am" in hojas:
        datos["iribarren"]["matriz"] = _limpiar_matriz(
            pd.read_excel(xl, sheet_name=hojas["matriz zonas 6am"]), "ZONA"
        )

    # --- BLOQUE FORÁNEOS (Hojas de Municipios) ---
    for posible_cons in ["consolidado municipios", "consolidado municipio"]:
        if posible_cons in hojas:
            datos["foraneos"]["consolidado"] = _limpiar_consolidado(
                pd.read_excel(xl, sheet_name=hojas[posible_cons], dtype=str), "MUNICIPIO"
            )
            break

    for posible_res in ["resumen municipios 6am", "resumen municipio 6am"]:
        if posible_res in hojas:
            datos["foraneos"]["resumen"] = _limpiar_resumen(
                pd.read_excel(xl, sheet_name=hojas[posible_res]), col_entidad="MUNICIPIO"
            )
            break

    for posible_mat in ["matriz municipios 6am", "matriz municipio 6am"]:
        if posible_mat in hojas:
            datos["foraneos"]["matriz"] = _limpiar_matriz(
                pd.read_excel(xl, sheet_name=hojas[posible_mat]), "MUNICIPIO"
            )
            break

    return datos


datos = cargar_excel_completo()

if datos is None:
    st.warning(f"⚠️ No se encontró el archivo `{ARCHIVO_EXCEL}`. Genera el archivo desde tu procesador de TXT y recarga.")
    st.stop()


# ==========================================
# 2. SELECTOR PRINCIPAL DE ÁMBITO (SEPARADO)
# ==========================================
with st.sidebar:
    st.header("🧭 Ámbito de Análisis")
    ambito = st.radio(
        "Seleccione el conjunto de datos:",
        options=["Iribarren", "Foráneos"],
        index=0,
        key="selector_ambito",
    )
    st.caption("Los datos se analizan por separado.")

if ambito == "Iribarren":
    df_consolidado = datos["iribarren"]["consolidado"]
    df_resumen = datos["iribarren"]["resumen"]
    df_matriz = datos["iribarren"]["matriz"]
    col_entidad = "ZONA"
    etiqueta_entidad = "Zona"
    etiqueta_listado = "Zonas"
    sufijo_key = "iribarren"
else:
    df_consolidado = datos["foraneos"]["consolidado"]
    df_resumen = datos["foraneos"]["resumen"]
    df_matriz = datos["foraneos"]["matriz"]
    col_entidad = "MUNICIPIO"
    etiqueta_entidad = "Municipio"
    etiqueta_listado = "Municipios"
    sufijo_key = "foraneos"


# ==========================================
# 3. PESTAÑAS PRINCIPALES
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📅 Seguimiento Diario",
    "📈 Seguimiento Anual",
    "⏳ PA por Motivo",
    f"🗺️ Matriz {etiqueta_listado} (6 AM)",
    "🔍 Buscador",
])


# ------------------------------------------
# TAB 1: SEGUIMIENTO DIARIO
# ------------------------------------------
with tab1:
    st.subheader(f"📅 Comportamiento Diario ({ambito}) — Corte 6:00 AM")
    if df_resumen.empty:
        st.info(f"No hay registros cargados en el resumen diario de {ambito}.")
    else:
        tiene_col_entidad = col_entidad in df_resumen.columns

        if tiene_col_entidad:
            col_f0, col_f1, col_f2, col_f3 = st.columns([1.2, 0.9, 1.0, 1.6])
            entidades_res = ["TODOS"] + sorted([e for e in df_resumen[col_entidad].unique() if e])
            with col_f0:
                entidad_sel = st.selectbox(f"{etiqueta_entidad}:", entidades_res, key=f"t1_ent_{sufijo_key}")
            if entidad_sel != "TODOS":
                df_res_base = df_resumen[df_resumen[col_entidad] == entidad_sel].copy()
            else:
                df_res_base = (
                    df_resumen.groupby(["FECHA_DT", "ANIO", "MES"], as_index=False)
                    .agg({
                        "FECHA OPERATIVA": "first",
                        "VENTANA TURNO": "first",
                        "PENDIENTES INICIO (6 AM)": "sum",
                        "RECIBIDOS": "sum",
                        "FINALIZADOS": "sum",
                        "EN EJECUCIÓN CIERRE (6 AM)": "sum",
                        "PENDIENTES CIERRE (6 AM)": "sum",
                        "TOTAL": "sum",
                    })
                    .sort_values("FECHA_DT")
                )
        else:
            col_f1, col_f2, col_f3 = st.columns([1, 1, 2])
            entidad_sel = "TODOS"
            df_res_base = df_resumen.copy()

        anios_disp = sorted(df_res_base["ANIO"].dropna().unique().astype(int).tolist())
        idx_anio = anios_disp.index(2026) if 2026 in anios_disp else len(anios_disp) - 1
        with col_f1:
            anio_sel = st.selectbox("Año:", anios_disp, index=idx_anio, key=f"t1_anio_{sufijo_key}")

        meses_disp = sorted(
            df_res_base[df_res_base["ANIO"] == anio_sel]["MES"].dropna().unique().astype(int).tolist()
        )
        nombres_meses_disp = [meses_es[m] for m in meses_disp]
        idx_mes = nombres_meses_disp.index("Noviembre") if "Noviembre" in nombres_meses_disp else len(nombres_meses_disp) - 1
        with col_f2:
            mes_nombre_sel = st.selectbox("Mes:", nombres_meses_disp, index=idx_mes, key=f"t1_mes_{sufijo_key}")
            mes_sel = [k for k, v in meses_es.items() if v == mes_nombre_sel][0]

        df_mes = df_res_base[
            (df_res_base["ANIO"] == anio_sel) & (df_res_base["MES"] == mes_sel)
        ].copy()

        rango_activo = None
        if not df_mes.empty:
            min_f = df_mes["FECHA_DT"].min().date()
            max_f = df_mes["FECHA_DT"].max().date()
            with col_f3:
                rango_fechas = st.date_input(
                    "Rango de fechas:",
                    value=(min_f, max_f),
                    min_value=min_f,
                    max_value=max_f,
                    key=f"t1_rango_{sufijo_key}",
                )
            if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
                rango_activo = rango_fechas
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

            subtitulo_ent = f" — {entidad_sel}" if tiene_col_entidad and entidad_sel != "TODOS" else ""
            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["PENDIENTES INICIO (6 AM)"],
                name="Pendientes Inicio (6 AM)", marker_color=COLOR_PA_INICIO,
                text=df_mes["PENDIENTES INICIO (6 AM)"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["RECIBIDOS"],
                name="Recibidos", marker_color=COLOR_RECIBIDOS,
                text=df_mes["RECIBIDOS"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig1.add_trace(go.Bar(
                x=df_mes["FECHA_STR"], y=df_mes["EN EJECUCIÓN CIERRE (6 AM)"],
                name="En Ejecución Cierre (6 AM)", marker_color=COLOR_EE_CIERRE,
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
                textfont=dict(color=COLOR_FINALIZADOS, size=11),
                marker=dict(size=6, color=COLOR_FINALIZADOS), line=dict(color=COLOR_FINALIZADOS, width=2),
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
                title=f"<b>Flujo Diario ({ambito}{subtitulo_ent}) — {mes_nombre_sel} {anio_sel}</b>",
                barmode="stack",
                xaxis_title="<b>Fecha Operativa</b>",
                yaxis_title="<b>Cantidad</b>",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
            )
            st.plotly_chart(fig1, use_container_width=True, key=f"fig_diario_{sufijo_key}")

            # --- RANKING COMPARATIVO POR MUNICIPIO (EN ÁMBITO FORÁNEOS Y "TODOS") ---
            if ambito == "Foráneos" and tiene_col_entidad and entidad_sel == "TODOS":
                st.divider()
                st.markdown(f"### 🏆 Ranking Comparativo por Municipio — {mes_nombre_sel} {anio_sel}")
                df_mun_periodo = df_resumen[
                    (df_resumen["ANIO"] == anio_sel) & (df_resumen["MES"] == mes_sel)
                ].copy()
                if rango_activo is not None:
                    df_mun_periodo = df_mun_periodo[
                        (df_mun_periodo["FECHA_DT"].dt.date >= rango_activo[0])
                        & (df_mun_periodo["FECHA_DT"].dt.date <= rango_activo[1])
                    ].copy()

                if not df_mun_periodo.empty:
                    df_mun_periodo = df_mun_periodo.sort_values("FECHA_DT")
                    ranking_mun = (
                        df_mun_periodo.groupby(col_entidad, as_index=False)
                        .agg(
                            PENDIENTES_CIERRE=("PENDIENTES CIERRE (6 AM)", "last"),
                            RECIBIDOS_TOTAL=("RECIBIDOS", "sum"),
                            FINALIZADOS_TOTAL=("FINALIZADOS", "sum"),
                        )
                    )

                    col_rk1, col_rk2 = st.columns(2)
                    with col_rk1:
                        rk_pend = ranking_mun.sort_values("PENDIENTES_CIERRE", ascending=True)
                        fig_rk_pend = px.bar(
                            rk_pend,
                            x="PENDIENTES_CIERRE",
                            y=col_entidad,
                            orientation="h",
                            text="PENDIENTES_CIERRE",
                            title="Pendientes al Cierre (6 AM) por Municipio",
                        )
                        fig_rk_pend.update_traces(
                            marker_color="#b91c1c",
                            textposition="outside",
                            cliponaxis=False,
                        )
                        fig_rk_pend.update_layout(
                            xaxis_title="<b>Pendientes Cierre (6 AM)</b>",
                            yaxis_title="<b>Municipio</b>",
                            height=max(320, len(rk_pend) * 34),
                            margin=dict(r=40),
                        )
                        st.plotly_chart(fig_rk_pend, use_container_width=True, key="fig_rk_mun_pend")

                    with col_rk2:
                        rk_rec = ranking_mun.sort_values("RECIBIDOS_TOTAL", ascending=True)
                        fig_rk_rec = px.bar(
                            rk_rec,
                            x="RECIBIDOS_TOTAL",
                            y=col_entidad,
                            orientation="h",
                            text="RECIBIDOS_TOTAL",
                            title="Recibidos Acumulados por Municipio",
                        )
                        fig_rk_rec.update_traces(
                            marker_color="#1f4e78",
                            textposition="outside",
                            cliponaxis=False,
                        )
                        fig_rk_rec.update_layout(
                            xaxis_title="<b>Total Recibidos</b>",
                            yaxis_title="<b>Municipio</b>",
                            height=max(320, len(rk_rec) * 34),
                            margin=dict(r=40),
                        )
                        st.plotly_chart(fig_rk_rec, use_container_width=True, key="fig_rk_mun_rec")

            cols_tabla_res = [
                col_entidad if (tiene_col_entidad and entidad_sel != "TODOS") else None,
                "FECHA OPERATIVA",
                "VENTANA TURNO",
                "PENDIENTES INICIO (6 AM)",
                "RECIBIDOS",
                "TOTAL",
                "FINALIZADOS",
                "EN EJECUCIÓN CIERRE (6 AM)",
                "PENDIENTES CIERRE (6 AM)",
            ]
            cols_tabla_res = [c for c in cols_tabla_res if c and c in df_mes.columns]
            st.dataframe(df_mes[cols_tabla_res], use_container_width=True, hide_index=True)

            buf_d = io.BytesIO()
            with pd.ExcelWriter(buf_d, engine="openpyxl") as writer:
                df_mes[cols_tabla_res].to_excel(writer, index=False, sheet_name="Seguimiento_Diario")
            st.download_button(
                f"⬇️ Descargar Seguimiento Diario {ambito} (.xlsx)",
                buf_d.getvalue(),
                file_name=f"seguimiento_diario_{sufijo_key}_{mes_nombre_sel.lower()}_{anio_sel}.xlsx",
                key=f"dl_diario_{sufijo_key}",
            )


# ------------------------------------------
# TAB 2: SEGUIMIENTO ANUAL
# ------------------------------------------
with tab2:
    st.subheader(f"📈 Seguimiento Mensual Acumulado ({ambito})")
    if df_resumen.empty:
        st.info(f"No hay registros cargados en el resumen diario de {ambito}.")
    else:
        tiene_col_entidad = col_entidad in df_resumen.columns
        if tiene_col_entidad:
            col_y0, col_y1, _ = st.columns([1.2, 1, 2])
            entidades_y = ["TODOS"] + sorted([e for e in df_resumen[col_entidad].unique() if e])
            with col_y0:
                entidad_y_sel = st.selectbox(f"{etiqueta_entidad}:", entidades_y, key=f"t2_ent_{sufijo_key}")
            if entidad_y_sel != "TODOS":
                df_y_base = df_resumen[df_resumen[col_entidad] == entidad_y_sel].copy()
            else:
                df_y_base = (
                    df_resumen.groupby(["FECHA_DT", "ANIO", "MES"], as_index=False)
                    .agg({
                        "PENDIENTES INICIO (6 AM)": "sum",
                        "RECIBIDOS": "sum",
                        "FINALIZADOS": "sum",
                        "EN EJECUCIÓN CIERRE (6 AM)": "sum",
                        "PENDIENTES CIERRE (6 AM)": "sum",
                    })
                    .sort_values("FECHA_DT")
                )
        else:
            col_y1, _ = st.columns([1, 3])
            entidad_y_sel = "TODOS"
            df_y_base = df_resumen.copy()

        anios_disp_y = sorted(df_y_base["ANIO"].dropna().unique().astype(int).tolist())
        idx_anio_y = anios_disp_y.index(2026) if 2026 in anios_disp_y else len(anios_disp_y) - 1
        with col_y1:
            anio_anual = st.selectbox("Seleccione el Año:", anios_disp_y, index=idx_anio_y, key=f"t2_anio_{sufijo_key}")

        df_y = df_y_base[df_y_base["ANIO"] == anio_anual].sort_values("FECHA_DT")
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

            subtitulo_y = f" — {entidad_y_sel}" if tiene_col_entidad and entidad_y_sel != "TODOS" else ""
            fig_anual = go.Figure()
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["PENDIENTES INICIO (6 AM)"],
                name="Pendientes Inicio (6 AM)", marker_color=COLOR_PA_INICIO,
                text=df_anual["PENDIENTES INICIO (6 AM)"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["RECIBIDOS"],
                name="Recibidos", marker_color=COLOR_RECIBIDOS,
                text=df_anual["RECIBIDOS"], textposition="inside",
                textfont=dict(color="black", size=11),
            ))
            fig_anual.add_trace(go.Bar(
                x=df_anual["MES"], y=df_anual["EN EJECUCIÓN CIERRE (6 AM)"],
                name="En Ejecución Cierre (6 AM)", marker_color=COLOR_EE_CIERRE,
                text=df_anual["EN EJECUCIÓN CIERRE (6 AM)"], textposition="inside",
                textfont=dict(color="black", size=11),
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
                marker=dict(size=6, color=COLOR_FINALIZADOS), line=dict(color=COLOR_FINALIZADOS, width=2),
            ))
            fig_anual.update_layout(
                title=f"<b>Evolución Mensual Acumulada ({ambito}{subtitulo_y}) — {anio_anual}</b>",
                barmode="stack",
                xaxis_title="<b>Mes</b>",
                yaxis_title="<b>Cantidad</b>",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
            )
            st.plotly_chart(fig_anual, use_container_width=True, key=f"fig_anual_{sufijo_key}")


# ------------------------------------------
# TAB 3: PENDIENTES (PA) POR MOTIVO
# ------------------------------------------
with tab3:
    st.subheader(f"⏳ Incidencias Pendientes (PA) por Motivo — {ambito}")
    if df_consolidado.empty:
        st.info(f"La hoja de consolidado de {ambito} está vacía.")
    else:
        df_pa = df_consolidado[df_consolidado["ESTADO"] == "PA"].copy() if "ESTADO" in df_consolidado.columns else df_consolidado.copy()

        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            opciones_ent = ["TODAS" if ambito == "Iribarren" else "TODOS"] + (
                sorted([e for e in df_pa[col_entidad].unique() if e]) if col_entidad in df_pa.columns else []
            )
            ent_pa_sel = st.selectbox(f"Filtrar por {etiqueta_entidad}:", opciones_ent, key=f"t3_ent_{sufijo_key}")
        with col_p2:
            anios_pa = ["TODOS"] + sorted([a for a in df_pa.get("AÑO RECIBIDO", pd.Series()).unique() if a])
            anio_pa_sel = st.selectbox("Año Recibido:", anios_pa, key=f"t3_anio_{sufijo_key}")
        with col_p3:
            meses_pa = ["TODOS"] + [m for m in meses_upper.values() if m in df_pa.get("MES RECIBIDO", pd.Series()).unique()]
            mes_pa_sel = st.selectbox("Mes Recibido:", meses_pa, key=f"t3_mes_{sufijo_key}")

        if ent_pa_sel not in ["TODAS", "TODOS"] and col_entidad in df_pa.columns:
            df_pa = df_pa[df_pa[col_entidad] == ent_pa_sel]
        if anio_pa_sel != "TODOS" and "AÑO RECIBIDO" in df_pa.columns:
            df_pa = df_pa[df_pa["AÑO RECIBIDO"] == anio_pa_sel]
        if mes_pa_sel != "TODOS" and "MES RECIBIDO" in df_pa.columns:
            df_pa = df_pa[df_pa["MES RECIBIDO"] == mes_pa_sel]

        st.metric(f"⏳ Total Incidencias PA ({ambito})", f"{len(df_pa):,}")

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
                    title=f"Distribución de Pendientes (PA) por Motivo ({ambito})",
                )
                fig_pa_mot.update_traces(marker_color="#b91c1c", textposition="outside", cliponaxis=False)
                fig_pa_mot.update_layout(
                    xaxis_title="<b>Cantidad PA</b>",
                    yaxis_title="<b>Motivo</b>",
                    height=max(320, len(res_ord) * 34),
                    margin=dict(r=40),
                )
                st.plotly_chart(fig_pa_mot, use_container_width=True, key=f"fig_pa_mot_{sufijo_key}")

            with st.expander(f"Ver listado detallado de incidencias PA ({ambito})"):
                cols_pa_det = [
                    c for c in ["INCIDENCIA", col_entidad, "RECEPCIÓN", "MOTIVO", "DIRECCIÓN", "CLIENTE", "TEC", "TCR"]
                    if c in df_pa.columns
                ]
                st.dataframe(df_pa[cols_pa_det], use_container_width=True, hide_index=True)
        else:
            st.info("No hay incidencias en estado PA para los filtros seleccionados.")


# ------------------------------------------
# TAB 4: MATRIZ ZONAS / MUNICIPIOS (6 AM)
# ------------------------------------------
with tab4:
    st.subheader(f"🗺️ Matriz de Pendientes por {etiqueta_entidad} — {ambito} (Corte 6:00 AM)")
    if df_matriz.empty:
        st.info(f"La matriz de {etiqueta_listado} ({ambito}) no tiene datos.")
    else:
        cols_fecha_mat = [c for c in df_matriz.columns if c != col_entidad]
        df_mat_ent = df_matriz[df_matriz[col_entidad] != "TOTALES"].copy()
        df_long = df_mat_ent.melt(
            id_vars=[col_entidad],
            value_vars=cols_fecha_mat,
            var_name="FECHA_STR",
            value_name="CANTIDAD",
        )
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
                anio_mat = st.selectbox("Año:", anios_m, index=idx_am, key=f"t4_anio_{sufijo_key}")
            meses_m = sorted(df_long[df_long["ANIO"] == anio_mat]["MES"].unique().tolist())
            nombres_mm = [meses_es[m] for m in meses_m]
            idx_mm = nombres_mm.index("Noviembre") if "Noviembre" in nombres_mm else len(nombres_mm) - 1
            with cm2:
                mes_mat_nombre = st.selectbox("Mes:", nombres_mm, index=idx_mm, key=f"t4_mes_{sufijo_key}")
                mes_mat = [k for k, v in meses_es.items() if v == mes_mat_nombre][0]
            with cm3:
                todas_lbl = "TODAS" if ambito == "Iribarren" else "TODOS"
                entidades_m = [todas_lbl] + sorted(df_long[col_entidad].unique().tolist())
                ent_mat = st.selectbox(f"{etiqueta_entidad}:", entidades_m, key=f"t4_ent_{sufijo_key}")

            df_long_f = df_long[(df_long["ANIO"] == anio_mat) & (df_long["MES"] == mes_mat)].copy()
            if ent_mat != todas_lbl:
                df_long_f = df_long_f[df_long_f[col_entidad] == ent_mat]

            if not df_long_f.empty:
                fig_mat = px.line(
                    df_long_f.sort_values("FECHA_DT"),
                    x="FECHA_DT",
                    y="CANTIDAD",
                    color=col_entidad,
                    markers=True,
                    title=f"<b>Comportamiento Diario por {etiqueta_entidad} (6:00 AM) — {mes_mat_nombre} {anio_mat}</b>",
                )
                fig_mat.update_layout(xaxis_title="<b>Fecha</b>", yaxis_title="<b>Cantidad</b>", hovermode="x unified")
                st.plotly_chart(fig_mat, use_container_width=True, key=f"fig_mat_{sufijo_key}")

            fechas_cols_mes = (
                df_long[(df_long["ANIO"] == anio_mat) & (df_long["MES"] == mes_mat)]
                .sort_values("FECHA_DT")["FECHA_STR"]
                .unique()
                .tolist()
            )
            cols_ver = [col_entidad] + [c for c in fechas_cols_mes if c in df_matriz.columns]
            df_matriz_vista = df_matriz[cols_ver].copy()
            if ent_mat != todas_lbl:
                df_matriz_vista = df_matriz_vista[df_matriz_vista[col_entidad].isin([ent_mat, "TOTALES"])]
            st.dataframe(df_matriz_vista, use_container_width=True, hide_index=True)


# ------------------------------------------
# TAB 5: BUSCADOR DE INCIDENCIAS
# ------------------------------------------
with tab5:
    st.subheader("🔍 Buscador de Incidencias")
    if df_consolidado.empty:
        st.info(f"No hay datos en el consolidado de {ambito}.")
    else:
        todas_b = "TODAS" if ambito == "Iribarren" else "TODOS"
        cb1, cb2, cb3, cb4 = st.columns(4)
        with cb1:
            f_inc = st.text_input("Nº Incidencia:", placeholder="Ej. 926117", key=f"t5_inc_{sufijo_key}")
        with cb2:
            lista_ent_b = [todas_b] + (
                sorted([e for e in df_consolidado[col_entidad].unique() if e]) if col_entidad in df_consolidado.columns else []
            )
            f_ent = st.selectbox(f"{etiqueta_entidad}:", lista_ent_b, key=f"t5_ent_{sufijo_key}")
        with cb3:
            lista_est_b = ["TODOS"] + (
                sorted([e for e in df_consolidado["ESTADO"].unique() if e]) if "ESTADO" in df_consolidado.columns else []
            )
            f_est = st.selectbox("Estado:", lista_est_b, key=f"t5_est_{sufijo_key}")
        with cb4:
            f_mot = st.text_input("Motivo:", placeholder="Ej. SECTOR SIN SERVICIO", key=f"t5_mot_{sufijo_key}")

        df_busq = df_consolidado.copy()
        if f_inc and "INCIDENCIA" in df_busq.columns:
            df_busq = df_busq[df_busq["INCIDENCIA"].str.contains(f_inc, case=False, na=False)]
        if f_ent != todas_b and col_entidad in df_busq.columns:
            df_busq = df_busq[df_busq[col_entidad] == f_ent]
        if f_est != "TODOS" and "ESTADO" in df_busq.columns:
            df_busq = df_busq[df_busq["ESTADO"] == f_est]
        if f_mot and "MOTIVO" in df_busq.columns:
            df_busq = df_busq[df_busq["MOTIVO"].str.contains(f_mot, case=False, na=False)]

        st.write(f"Registros encontrados: **{len(df_busq):,}**")
        cols_busq = [
            c for c in [
                "INCIDENCIA", col_entidad, "RECEPCIÓN", "FINALIZACIÓN", "ESTADO",
                "MOTIVO", "DIRECCIÓN", "CLIENTE", "TEC", "TCR", "TDV", "TEJ", "TAR",
            ] if c in df_busq.columns
        ]
        st.dataframe(df_busq[cols_busq].head(1000), use_container_width=True, hide_index=True)
