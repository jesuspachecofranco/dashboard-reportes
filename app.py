import calendar
from datetime import datetime
import os
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Reportes 2026", page_icon="📊", layout="wide"
)

st.title("📊 Control y Seguimiento de Incidencias")


# Función para cargar y procesar los datos
@st.cache_data
def cargar_datos():
    archivo_entrada = "resultado_unificado.xlsx"
    if not os.path.exists(archivo_entrada):
        return None

    df = pd.read_excel(archivo_entrada)
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
        "❌ No se encontró el archivo 'resultado_unificado.xlsx' en el"
        " directorio."
    )
else:
    # Columnas requeridas para mostrar en las tablas detalladas
    columnas_mostrar = [
        "INCIDENCIA",
        "RECEPCIÓN",
        "FINALIZACIÓN",
        "DIRECCIÓN",
        "CLIENTE",
        "MOTIVO",
        "ESTADO",
    ]

    # 3 Pestañas principales unificadas arriba
    tab1, tab2, tab3 = st.tabs([
        "📅 Seguimiento Diario",
        "📈 Seguimiento Anual",
        "🔍 Búsqueda por Fecha",
    ])

    # ==========================================
    # PESTAÑA 1: SEGUIMIENTO DIARIO DINÁMICO
    # ==========================================
    with tab1:
        st.subheader("📅 Comportamiento Diario por Mes")

        # Selectores para elegir el Año y el Mes que desees visualizar
        col_s1, col_s2, col_s3,col_s4,col_s5,col_s6, = st.columns([1, 1, 1, 1, 1, 1])
        with col_s1:
            # Puedes ajustar el rango de años según tus necesidades
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
            )  # Por defecto Agosto (índice 7)
            # Recuperar el número del mes seleccionado
            mes_seleccionado = [
                k
                for k, v in meses_dict.items()
                if v == mes_nombre_seleccionado
            ][0]

        # Definir dinámicamente el inicio y fin del mes seleccionado
        inicio_mes_dinamico = pd.Timestamp(
            year=anio_seleccionado, month=mes_seleccionado, day=1
        )

        # Calcular el último día del mes de forma exacta
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

        st.markdown(
            f"##### 📌 Resumen del Mes de {mes_nombre_seleccionado}"
            f" {anio_seleccionado}"
        )
        c1, c2, c3, c4,c5, c6 = st.columns(6)
        c3.metric("📦 Acumulado Inicial", f"{acumulado_mes:,}")
        c4.metric("📥 Recibidos", f"{recibidos_mes:,}")
        c5.metric("📊 Total Incidencias", f"{total_mes:,}")
        c6.metric("✅ Atendidos / Finalizados", f"{finalizados_mes:,}")
        st.divider()

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

            datos_diarios.append({
                "FECHA": nombre_dia,
                "REPORTES RECIBIDOS": cant_recibidos,
                "REPORTES ACUMULADOS AL INICIAR": cant_acumulada,
                "TOTAL REPORTES": total_rep,
                "REPORTES FINALIZADOS": cant_finalizados,
            })

        df_dia = pd.DataFrame(datos_diarios)

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
    # PESTAÑA 2: ANUAL HASTA AGOSTO
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

        st.markdown("##### 📌 Resumen Acumulado Anual (Enero - Agosto)")
        ac1, ac2, ac3, ac4 = st.columns(4)
        ac1.metric("📦 Acumulado Inicial", f"{acumulado_anio:,}")
        ac2.metric("📥 Recibidos", f"{recibidos_anio:,}")
        ac3.metric("📊 Total Incidencias", f"{total_anio:,}")
        ac4.metric("✅ Atendidos / Finalizados", f"{finalizados_anio:,}")
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
        st.markdown(
            "Selecciona **cualquier fecha** para consultar a detalle las"
            " incidencias acumuladas pendientes, recibidas y finalizadas."
        )

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

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📦 Acumulados Previos", f"{len(df_acumulados_dia):,}")
        m2.metric("📥 Recibidos en el día", f"{len(df_recibidos_dia):,}")
        m3.metric(
            "📊 Total Activos",
            f"{len(df_acumulados_dia) + len(df_recibidos_dia):,}",
        )
        m4.metric("✅ Finalizados en el día", f"{len(df_finalizados_dia):,}")
        st.markdown("---")

        with st.expander(
            f"📦 1. Incidencias Acumuladas Pendientes ({len(df_acumulados_dia)}"
            " registros)",
            expanded=True,
        ):
            st.markdown(
                "*Casos que venían de fechas anteriores y seguían activos al"
                " iniciar este día.*"
            )
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
            st.markdown(
                "*Casos que ingresaron exactamente durante el transcurso de este"
                " día.*"
            )
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
            st.markdown(
                "*Casos cuya atención o finalización se registró durante este"
                " día.*"
            )
            if not df_finalizados_dia.empty:
                st.dataframe(
                    df_finalizados_dia[columnas_mostrar],
                    use_container_width=True,
                )
            else:
                st.info(
                    "No hay incidencias finalizadas registradas en esta fecha."
                )
