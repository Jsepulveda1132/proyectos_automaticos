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
        # GRÁFICOS INTERACTIVOS (CUADRÍCULA DE 2X2)
        # ==============================================================================
        # FILA 1: TENDENCIAS TEMPORALES (CRONOLÓGICAS)
        st.markdown("### 📈 Análisis de Tendencia Temporal")
        fila1_col1, fila1_col2 = st.columns(2)

        # Agrupar datos cronológicos base
        df_tiempo = (
            df_filtrado.groupby(df_filtrado["Fecha"].dt.date)
            .agg(
                Monto_Total=("Valor_Neto", "sum"), Cantidad_Total=("Numero_NC", "count")
            )
            .reset_index()
        )
        df_tiempo = df_tiempo.sort_values(by="Fecha")

        with fila1_col1:
            st.subheader("Tendencia Temporal por Valor")
            fig_linea_valor = px.line(
                df_tiempo,
                x="Fecha",
                y="Monto_Total",
                markers=True,
                labels={"Fecha": "Fecha de Emisión", "Monto_Total": "Monto ($)"},
                template="plotly_white",
                color_discrete_sequence=["#2E7D32"],
            )
            fig_linea_valor.update_traces(
                hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Monto:</b> $ %{y:,.0f}<extra></extra>"
            )
            fig_linea_valor.update_layout(
                xaxis_tickformat="%d/%m/%Y",
                yaxis_tickformat="$ ,.0f",
                separators=",.",
                xaxis_title="Línea de Tiempo (Día a Día)",
            )
            st.plotly_chart(fig_linea_valor, use_container_width=True)

        with fila1_col2:
            st.subheader("Tendencia Temporal por Cantidad de NC")
            fig_linea_cant = px.line(
                df_tiempo,
                x="Fecha",
                y="Cantidad_Total",
                markers=True,
                labels={
                    "Fecha": "Fecha de Emisión",
                    "Cantidad_Total": "Cantidad de NC",
                },
                template="plotly_white",
                color_discrete_sequence=["#1565C0"],  # Color azul para diferenciar
            )
            fig_linea_cant.update_traces(
                hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Notas:</b> %{y:,.0f} NC<extra></extra>"
            )
            fig_linea_cant.update_layout(
                xaxis_tickformat="%d/%m/%Y",
                yaxis_tickformat=",.0f",
                separators=",.",
                xaxis_title="Línea de Tiempo (Día a Día)",
            )
            st.plotly_chart(fig_linea_cant, use_container_width=True)

        st.markdown("---")

        # FILA 2: DISTRIBUCIÓN POR MOTIVOS DE ANULACIÓN
        st.markdown("### 📋 Análisis por Motivos de Anulación")
        fila2_col1, fila2_col2 = st.columns(2)

        # Agrupar datos de motivos base
        df_motivos_graf = (
            df_filtrado.groupby("Motivo_Anulacion")
            .agg(
                Monto_Total=("Valor_Neto", "sum"), Cantidad_Total=("Numero_NC", "count")
            )
            .reset_index()
        )

        with fila2_col1:
            st.subheader("Distribución por Valor Total")
            df_motivos_valor = df_motivos_graf.sort_values(
                by="Monto_Total", ascending=True
            )
            fig_barra_valor = px.bar(
                df_motivos_valor,
                x="Monto_Total",
                y="Motivo_Anulacion",
                orientation="h",
                labels={"Monto_Total": "Total ($)", "Motivo_Anulacion": "Motivo"},
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"],
            )
            # Forzar formato regional colombiano en el cuadro flotante (Hover)
            fig_barra_valor.update_traces(
                hovertemplate="<b>Motivo:</b> %{y}<br><b>Monto:</b> $ %{x:,.0f}<extra></extra>"
            )
            fig_barra_valor.update_layout(xaxis_tickformat="$ ,.0f", separators=",.")
            st.plotly_chart(fig_barra_valor, use_container_width=True)

        with fila2_col2:
            st.subheader("Distribución por Cantidad de NC")
            df_motivos_cant = df_motivos_graf.sort_values(
                by="Cantidad_Total", ascending=True
            )
            fig_barra_cant = px.bar(
                df_motivos_cant,
                x="Cantidad_Total",
                y="Motivo_Anulacion",
                orientation="h",
                labels={
                    "Cantidad_Total": "Cantidad de NC",
                    "Motivo_Anulacion": "Motivo",
                },
                template="plotly_white",
                color_discrete_sequence=["#1E88E5"],
            )
            # Forzar formato regional de cantidad pura en el cuadro flotante (Hover)
            fig_barra_cant.update_traces(
                hovertemplate="<b>Motivo:</b> %{y}<br><b>Notas:</b> %{x:,.0f} NC<extra></extra>"
            )
            fig_barra_cant.update_layout(xaxis_tickformat=",.0f", separators=",.")
            st.plotly_chart(fig_barra_cant, use_container_width=True)

            st.markdown("---")
            st.markdown("### 🎯 Análisis de Clientes Críticos)")

        # 1. Calcular la concentración del dinero por cliente
        df_pareto = df_filtrado.groupby('Cliente')['Valor_Neto'].sum().reset_index()
        df_pareto = df_pareto.sort_values(by='Valor_Neto', ascending=False)

        monto_total_global = df_pareto['Valor_Neto'].sum()

        if monto_total_global > 0:
            # Calcular porcentajes individuales y acumulados
            df_pareto['Porcentaje'] = (df_pareto['Valor_Neto'] / monto_total_global) * 100
            df_pareto['Acumulado'] = df_pareto['Porcentaje'].cumsum()

            # Identificar quiénes representan el primer 80% del problema financiero
            clientes_criticos = df_pareto[df_pareto['Acumulado'] <= 85] # Margen seguro del 80-85%
            num_criticos = len(clientes_criticos) if len(clientes_criticos) > 0 else 1

            # Mostrar insights ejecutivos en tarjetas de diseño
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                st.info(f"💡 **Insight Financiero:** Solo **{num_criticos} cliente(s)** concentran la gran mayoría del dinero devuelto por Notas de Crédito en este periodo filtrado.")

            # 2. Dibujar la gráfica de barras combinada de Pareto
            fig_pareto = px.bar(
                df_pareto.head(10), x='Cliente', y='Valor_Neto',
                text=df_pareto.head(10)['Porcentaje'].apply(lambda x: f"{x:.1f}%"),
                title="Top 10 Clientes con Mayor Impacto en Cartera",
                template="plotly_white", color_discrete_sequence=['#B71C1C'] # Color rojo corporativo de alerta
            )
            fig_pareto.update_layout(yaxis_tickformat="$ ,.0f", separators=",.")
            fig_pareto.update_traces(textposition='outside', hovertemplate="<b>Cliente:</b> %{x}<br><b>Monto:</b> $ %{y:,.0f}<extra></extra>")
            st.plotly_chart(fig_pareto, use_container_width=True)

            st.markdown("---")

        # Tabla Formateada con Puntos en Miles y Comas en Decimales (Sin Origen_Data)
        st.subheader("🔍 Explorador de Datos Integrado")

        # Removimos 'Origen_Data' de esta lista para ocultarla del usuario
        columnas_tabla = ['Numero_NC', 'Fecha', 'NIT_Cliente', 'Cliente', 'Motivo_Anulacion', 'Valor_Neto']

        df_vista = df_filtrado[columnas_tabla].sort_values(by='Fecha', ascending=False)

        st.dataframe(
            df_vista.style.format({
                'Valor_Neto': lambda x: f"$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
                'Fecha': lambda t: t.strftime('%d/%m/%Y %H:%M') if pd.notnull(t) else ""
            }), 
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Error procesando los datos. Detalle: {e}")
