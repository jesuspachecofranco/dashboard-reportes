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
      "❌ No se encontró el archivo 'resultado_unificado.xlsx' en el directorio."
  )
else:
  # Pestañas principales de navegación
  tab1, tab2 = st.tabs(["📅 Agosto Diario", "📈 Acumulados Anual (Hasta Agosto)"])

  # ==========================================
  # PESTAÑA 1: AGOSTO DIARIO
  # ==========================================
  with tab1:
    st.subheader("Comportamiento Diario - Agosto 2026")

    dias_agosto = pd.date_range(
        start="2026-08-01", end="2026-08-31", freq="D"
    )
    datos_diarios = []

    for dia in dias_agosto:
      inicio_dia = dia
      fin_dia = dia + pd.Timedelta(days=1)
      nombre_dia = dia.strftime("%d/%m/%Y")

      cant_recibidos = int(
          ((df["RECEPCION_DT"] >= inicio_dia) & (df["RECEPCION_DT"] < fin_dia))
          .sum()
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

    # Gráfica Plotly Interactiva para Diario
    fig1 = go.Figure()

    # 1. Base de la barra: Acumulados al Iniciar (Fondo claro -> Texto NEGRO para contraste)
    fig1.add_trace(
        go.Bar(
            x=df_dia["FECHA"],
            y=df_dia["REPORTES ACUMULADOS AL INICIAR"],
            name="Acumulados al Iniciar",
            marker_color="#d8e4fc",
            text=df_dia["REPORTES ACUMULADOS AL INICIAR"],
            textposition="inside",
            insidetextanchor="top center",
            textfont=dict(
                color="black", size=11, family="Arial", weight="bold"
            ),
        )
    )

    # 2. Cima de la barra: Reportes Recibidos (Fondo oscuro -> Texto BLANCO para contraste)
    fig1.add_trace(
        go.Bar(
            x=df_dia["FECHA"],
            y=df_dia["REPORTES RECIBIDOS"],
            name="Reportes Recibidos",
            marker_color="#6391f4",
            text=df_dia["REPORTES RECIBIDOS"],
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(
                color="white", size=11, family="Arial", weight="bold"
            ),
        )
    )

    # 3. Línea invisible de soporte para mostrar el Total en la cúspide
    fig1.add_trace(
        go.Scatter(
            x=df_dia["FECHA"],
            y=df_dia["TOTAL REPORTES"],
            name="Total Reportes",
            mode="text+markers",
            text=df_dia["TOTAL REPORTES"],
            textposition="top center",
            textfont=dict(
                color="#1F4E78", size=12, family="Arial", weight="bold"
            ),
            marker=dict(size=8, color="rgba(0,0,0,0)"),
            showlegend=False,
        )
    )

    # 4. Línea de Reportes Finalizados
    fig1.add_trace(
        go.Scatter(
            x=df_dia["FECHA"],
            y=df_dia["REPORTES FINALIZADOS"],
            name="Reportes Finalizados",
            mode="lines+markers+text",
            text=df_dia["REPORTES FINALIZADOS"],
            textposition="top center",
            textfont=dict(
                color="#087333", size=11, family="Arial", weight="bold"
            ),
            marker=dict(size=6, color="#27AE60"),
            line=dict(color="#087333", width=2),
        )
    )

    # Personalización de Títulos y Diseño General
    fig1.update_layout(
        title=dict(
            text="<b>Flujo Diario de Reportes - Agosto 2026</b>",
            font=dict(size=18, color="#1F4E78"),
        ),
        barmode="stack",
        xaxis_title="<b>Día del Mes</b>",
        yaxis_title="<b>Cantidad de Reportes</b>",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
    )
    st.plotly_chart(fig1, use_container_width=True)

    with st.expander("Ver tabla de datos diarios"):
      st.dataframe(df_dia, use_container_width=True)

  # ==========================================
  # PESTAÑA 2: ANUAL HASTA AGOSTO
  # ==========================================
  with tab2:
    st.subheader("Evolución Mensual (Hasta Agosto 2026)")

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
          ((df["RECEPCION_DT"] >= inicio_mes) & (df["RECEPCION_DT"] < fin_mes))
          .sum()
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

    # Gráfica Plotly Interactiva para Anual
    fig2 = go.Figure()

    # 1. Base de la barra: Acumulados al Iniciar (Fondo claro -> Texto NEGRO)
    fig2.add_trace(
        go.Bar(
            x=df_anual["MES"],
            y=df_anual["REPORTES ACUMULADOS AL INICIAR"],
            name="Acumulados al Iniciar",
            marker_color="#5DADE2",
            text=df_anual["REPORTES ACUMULADOS AL INICIAR"],
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(
                color="black", size=12, family="Arial", weight="bold"
            ),
        )
    )

    # 2. Cima de la barra: Reportes Recibidos (Fondo oscuro -> Texto BLANCO)
    fig2.add_trace(
        go.Bar(
            x=df_anual["MES"],
            y=df_anual["REPORTES RECIBIDOS"],
            name="Reportes Recibidos",
            marker_color="#1F4E78",
            text=df_anual["REPORTES RECIBIDOS"],
            textposition="inside",
            insidetextanchor="middle",
            textfont=dict(
                color="white", size=12, family="Arial", weight="bold"
            ),
        )
    )

    # 3. Línea de soporte para el Total en la cúspide
    fig2.add_trace(
        go.Scatter(
            x=df_anual["MES"],
            y=df_anual["TOTAL REPORTES"],
            name="Total Reportes",
            mode="text+markers",
            text=df_anual["TOTAL REPORTES"],
            textposition="top center",
            textfont=dict(
                color="#1F4E78", size=12, family="Arial", weight="bold"
            ),
            marker=dict(size=8, color="rgba(0,0,0,0)"),
            showlegend=False,
        )
    )

    # 4. Línea de Reportes Finalizados
    fig2.add_trace(
        go.Scatter(
            x=df_anual["MES"],
            y=df_anual["REPORTES FINALIZADOS"],
            name="Reportes Finalizados",
            mode="lines+markers+text",
            text=df_anual["REPORTES FINALIZADOS"],
            textposition="top center",
            textfont=dict(
                color="#27AE60", size=12, family="Arial", weight="bold"
            ),
            marker=dict(size=6, color="#27AE60"),
            line=dict(color="#27AE60", width=2),
        )
    )

    # Personalización de Títulos y Diseño General
    fig2.update_layout(
        title=dict(
            text="<b>Resumen Acumulado Mensual (Enero - Agosto 2026)</b>",
            font=dict(size=18, color="#1F4E78"),
        ),
        barmode="stack",
        xaxis_title="<b>Mes</b>",
        yaxis_title="<b>Cantidad de Reportes</b>",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
    )
    st.plotly_chart(fig2, use_container_width=True)

    with st.expander("Ver tabla de datos mensuales"):
      st.dataframe(df_anual, use_container_width=True)
