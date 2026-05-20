import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuración de la página del dashboard
st.set_page_config(
    page_title="Dashboard Notas de Crédito", layout="wide", page_icon="📊"
)
st.title("📊 Dashboard Único de Notas de Crédito")
st.markdown(
    "Este panel lee los datos directamente de la ruta local actualizada por Power Query."
)

# ==============================================================================
# CONFIGURACIÓN DE LA RUTA LOCAL
# ==============================================================================
RUTA_ARCHIVO = RUTA_ARCHIVO = "NOTAS CREDITO ACTUALIZABLE BD SIESA.xlsx"

if os.path.exists(RUTA_ARCHIVO):
    try:
        # Carga inteligente de pestañas (Acepta mayúsculas y minúsculas)
        excel_file = pd.ExcelFile(RUTA_ARCHIVO)
        todos_los_nombres = excel_file.sheet_names

        pestana_c1 = next(
            (s for s in todos_los_nombres if s.strip().lower() == "consulta1"), None
        )
        pestana_c2 = next(
            (s for s in todos_los_nombres if s.strip().lower() == "consulta2"), None
        )

        if not pestana_c1 or not pestana_c2:
            st.error("❌ Error de lectura: No se encontraron las pestañas requeridas.")
            st.stop()

        df1_raw = pd.read_excel(RUTA_ARCHIVO, sheet_name=pestana_c1)
        df2_raw = pd.read_excel(RUTA_ARCHIVO, sheet_name=pestana_c2)

        # Detección automática de la columna de motivos
        col_motivo_c1 = next(
            (c for c in df1_raw.columns if "motivo" in c.lower() or "20_0000186" in c),
            None,
        )
        col_motivo_c2 = next(
            (c for c in df2_raw.columns if "motivo" in c.lower() or "20_0000186" in c),
            None,
        )

        if not col_motivo_c1:
            col_motivo_c1 = df1_raw.columns[-1]
        if not col_motivo_c2:
            col_motivo_c2 = df2_raw.columns[-1]

        # Mapeos específicos de columnas
        columnas_c1 = {
            "f_nrodocto": "Numero_NC",
            "f_fecha": "Fecha",
            "f_cliente": "NIT_Cliente",
            "f_cliente_razon_soc": "Cliente",
            "f_notas": "Observaciones",
            "f_valor_neto_docto": "Valor_Neto",
            "f_estado": "Estado",
            col_motivo_c1: "Motivo_Anulacion",
        }

        columnas_c2 = {
            "f_nrodocto": "Numero_NC",
            "f_fecha": "Fecha",
            "f_cliente_fact": "NIT_Cliente",
            "f_cliente_fact_razon_soc": "Cliente",
            "f_notes": "Observaciones",
            "F_valor_neto_alt": "Valor_Neto",
            "f_estado": "Estado",
            col_motivo_c2: "Motivo_Anulacion",
        }

        df1 = df1_raw.rename(columns=columnas_c1)
        df2 = df2_raw.rename(columns=columnas_c2)

        df1["Origen_Data"] = "Consulta 1"
        df2["Origen_Data"] = "Consulta 2"

        columnas_finales = [
            "Numero_NC",
            "Fecha",
            "NIT_Cliente",
            "Cliente",
            "Observaciones",
            "Valor_Neto",
            "Estado",
            "Motivo_Anulacion",
            "Origen_Data",
        ]
        df1 = df1[[c for c in columnas_finales if c in df1.columns]]
        df2 = df2[[c for c in columnas_finales if c in df2.columns]]

        df_total = pd.concat([df1, df2], ignore_index=True)

        # Limpieza de datos básica
        df_total["Estado"] = df_total["Estado"].astype(str).str.strip().str.upper()
        df_total = df_total[df_total["Estado"] == "APROBADAS"]

        df_total["Fecha"] = pd.to_datetime(df_total["Fecha"], errors="coerce")
        df_total = df_total.dropna(subset=["Fecha"])
        df_total["Año_Mes"] = df_total["Fecha"].dt.to_period("M").astype(str)

        # Corrección definitiva de formato de dinero
        def limpiar_monto(val):
            if pd.isna(val):
                return 0.0
            val_str = str(val).strip().replace(" ", "")
            if "," in val_str and "." in val_str:
                val_str = val_str.replace(".", "").replace(",", ".")
            elif "," in val_str:
                val_str = val_str.replace(",", ".")
            try:
                return abs(float(val_str))
            except ValueError:
                return 0.0

        df_total["Valor_Neto"] = df_total["Valor_Neto"].apply(limpiar_monto)
        df_total = df_total[df_total["Valor_Neto"] > 0]

        # Botón de actualización en barra lateral
        if st.sidebar.button("🔄 Actualizar Datos de la Ruta Local"):
            st.rerun()

        # ==============================================================================
        # SECCIÓN DE FILTROS AVANZADOS EN LA BARRA LATERAL
        # ==============================================================================
        st.sidebar.header("🎯 Filtros Globales")

        # 1. FILTRO DE FECHAS (Calendario)
        min_fecha = df_total["Fecha"].min().date()
        max_fecha = df_total["Fecha"].max().date()

        rango_fechas = st.sidebar.date_input(
            "Selecciona Rango de Fechas:",
            value=(min_fecha, max_fecha),
            min_value=min_fecha,
            max_value=max_fecha,
        )

        # Filtrado temporal inicial por fechas para actualizar las métricas de los siguientes filtros
        df_fechas = df_total.copy()
        if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
            f_inicio, f_fin = rango_fechas
            df_fechas = df_fechas[
                (df_fechas["Fecha"].dt.date >= f_inicio)
                & (df_fechas["Fecha"].dt.date <= f_fin)
            ]

        # 2. FILTRO DE CLIENTES CON MÉTRICAS (Monto y Cantidad de NC)
        resumen_clientes = (
            df_fechas.groupby("Cliente")
            .agg(Total=("Valor_Neto", "sum"), Cantidad=("Numero_NC", "count"))
            .reset_index()
        )

        opciones_clientes = []
        mapeo_clientes = {}
        for _, row in resumen_clientes.sort_values(
            by="Total", ascending=False
        ).iterrows():
            label = f"{row['Cliente']} ($ {row['Total']:,.0f} | {row['Cantidad']} NC)"
            opciones_clientes.append(label)
            mapeo_clientes[label] = row["Cliente"]

        filtro_clientes_labels = st.sidebar.multiselect(
            "Filtrar por Cliente:", options=opciones_clientes
        )
        clientes_seleccionados = [mapeo_clientes[l] for l in filtro_clientes_labels]

        # Aplicar filtro de cliente para delimitar los motivos disponibles
        df_motivos_prev = df_fechas.copy()
        if clientes_seleccionados:
            df_motivos_prev = df_motivos_prev[
                df_motivos_prev["Cliente"].isin(clientes_seleccionados)
            ]

        # 3. FILTRO DE MOTIVOS DINÁMICOS POR CLIENTE (Con métricas integradas)
        resumen_motivos = (
            df_motivos_prev.groupby("Motivo_Anulacion")
            .agg(Total=("Valor_Neto", "sum"), Cantidad=("Numero_NC", "count"))
            .reset_index()
        )

        opciones_motivos = []
        mapeo_motivos = {}
        for _, row in resumen_motivos.sort_values(
            by="Total", ascending=False
        ).iterrows():
            label = f"{row['Motivo_Anulacion']} ($ {row['Total']:,.0f} | {row['Cantidad']} NC)"
            opciones_motivos.append(label)
            mapeo_motivos[label] = row["Motivo_Anulacion"]

        filtro_motivos_labels = st.sidebar.multiselect(
            "Filtrar por Motivo de Nota:", options=opciones_motivos
        )
        motivos_seleccionados = [mapeo_motivos[l] for l in filtro_motivos_labels]

        # APLICACIÓN DE TODOS LOS FILTROS COMBINADOS AL DATAFRAME FINAL
        df_filtrado = df_fechas.copy()
        if clientes_seleccionados:
            df_filtrado = df_filtrado[
                df_filtrado["Cliente"].isin(clientes_seleccionados)
            ]
        if motivos_seleccionados:
            df_filtrado = df_filtrado[
                df_filtrado["Motivo_Anulacion"].isin(motivos_seleccionados)
            ]

        # ==============================================================================
        # DESPLIEGUE DEL DASHBOARD (KPIs, Gráficos y Tabla)
        # ==============================================================================
        monto_total = df_filtrado["Valor_Neto"].sum()
        conteo_notas = df_filtrado.shape[0]
        promedio_nota = monto_total / conteo_notas if conteo_notas > 0 else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("💰 Total Crédito Consolidado", f"$ {monto_total:,.2f}")
        col2.metric("📄 Volumen Total de Notas", f"{conteo_notas:,}")
        col3.metric("🧮 Valor Promedio por NC", f"$ {promedio_nota:,.2f}")

        st.markdown("---")

    
        # ==============================================================================
        # GRÁFICOS INTERACTIVOS (CON CORRECCIÓN DE UNIDADES FINANCIERAS)
        # ==============================================================================
        g1, g2 = st.columns(2)

        with g1:
            st.subheader("📈 Tendencia Temporal de Créditos")
            
            # Agrupar por la fecha exacta formateada como día para ver el movimiento cronológico real
            df_tiempo = df_filtrado.groupby(df_filtrado['Fecha'].dt.date)['Valor_Neto'].sum().reset_index()
            df_tiempo = df_tiempo.sort_values(by='Fecha') # Forzar orden cronológico estricto
            
            fig_linea = px.line(
                df_tiempo, x='Fecha', y='Valor_Neto', markers=True,
                labels={'Fecha': 'Fecha de Emisión', 'Valor_Neto': 'Monto ($)'},
                template="plotly_white", color_discrete_sequence=['#2E7D32']
            )
            
            # Forzar formato con puntos para miles y comas para decimales en el cuadro flotante (Hover)
            fig_linea.update_traces(
                hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Monto:</b> $ %{y:,.0f}<extra></extra>"
            )
            
            # Formatear el eje X para que muestre las fechas de forma estética y el eje Y con dinero regional
            fig_linea.update_layout(
                xaxis_tickformat="%d/%m/%Y",
                yaxis_tickformat="$ ,.0f",
                separators=",.",
                xaxis_title="Línea de Tiempo (Día a Día)"
            )
            st.plotly_chart(fig_linea, use_container_width=True)



        with g2:
            st.subheader("📋 Distribución por Cantidad de Notas Crédito")
            # Agrupar por motivo y contar cuántos números de NC existen por cada uno
            df_motivos_graf = df_filtrado.groupby('Motivo_Anulacion')['Numero_NC'].count().reset_index()
            df_motivos_graf = df_motivos_graf.rename(columns={'Numero_NC': 'Cantidad_NC'})
            df_motivos_graf = df_motivos_graf.sort_values(by='Cantidad_NC', ascending=True)
            
            fig_barra = px.bar(
                df_motivos_graf, x='Cantidad_NC', y='Motivo_Anulacion', orientation='h',
                labels={'Cantidad_NC': 'Cantidad de NC', 'Motivo_Anulacion': 'Motivo'},
                template="plotly_white", color_discrete_sequence=['#4CAF50']
            )
            # Como son enteros de cantidades, le quitamos el signo de pesos ($) al formato del eje X
            fig_barra.update_layout(
                xaxis_tickformat=",.0f",
                separators=",."
            )
            st.plotly_chart(fig_barra, use_container_width=True)

        # Tabla Formateada con Puntos en Miles y Comas en Decimales
        st.subheader("🔍 Explorador de Datos Integrado")
        columnas_tabla = [
            "Numero_NC",
            "Fecha",
            "NIT_Cliente",
            "Cliente",
            "Motivo_Anulacion",
            "Valor_Neto",
            "Origen_Data",
        ]
        df_vista = df_filtrado[columnas_tabla].sort_values(by="Fecha", ascending=False)

        st.dataframe(
            df_vista.style.format(
                {
                    "Valor_Neto": lambda x: f"$ {x:,.2f}".replace(",", "X")
                    .replace(".", ",")
                    .replace("X", "."),
                    "Fecha": lambda t: (
                        t.strftime("%d/%m/%Y %H:%M") if pd.notnull(t) else ""
                    ),
                }
            ),
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Error procesando los datos. Detalle: {e}")
