import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime
import io
from O365 import Account

# Configuración de la página del dashboard
st.set_page_config(page_title="Análisis Facturación DIAN", layout="wide", page_icon="🧾")
st.title("🧾 Análisis de Facturación DIAN - 2026")
st.markdown("Estacionalidad y segmentación semanal de proveedores.")

# ==============================================================================
# CONFIGURACIÓN DE PARÁMETROS Y ÁREAS CORPORATIVAS CONTROLADAS
# ==============================================================================
EXCEL_PROV = "INFORME FACTURACIÓN DIAN 2026.xlsx"
HOJA_PROV = "CONTROL GENERAL"

# Lista oficial de áreas para excluir Cajas Menores, Tarjetas de Crédito y Gastos Varios
AREAS_SUPERPACK = [
    "BIENESTAR", "CALIDAD", "COMPRAS", "CONTABILIDAD", "FINANCIERA", 
    "GESTION HUMANA", "MANTENIMIENTO", "NOMINA", "OPERACIONES", 
    "SSTA", "TI", "COMERCIAL", "GERENCIA", "TESORERIA"
]

# ==============================================================================
# MOTOR DE CONEXIÓN HÍBRIDO (OneDrive API / GitHub Backup)
# ==============================================================================
def descargar_excel_onedrive():
    try:
        tenant_id = st.secrets["microsoft"]["tenant_id"]
        client_id = st.secrets["microsoft"]["client_id"]
        client_secret = st.secrets["microsoft"]["client_secret"]
        credentials = (client_id, client_secret)
        
        if "token_data" in st.secrets["microsoft"]:
            with open("o365_token.txt", "w") as token_file:
                token_file.write(st.secrets["microsoft"]["token_data"])
        
        account = Account(credentials, tenant_id=tenant_id)
        
        if account.is_authenticated:
            storage = account.storage()  
            folder = storage.get_root_folder().get_folder(by_path='PROYECTOS AUTOMATICOS/Informe Notas Credito')
            file = folder.get_items(search=EXCEL_PROV, limit=1)
            
            if file:
                target_file = file if isinstance(file, list) else file
                out_stream = io.BytesIO()
                target_file.download(out_stream)
                out_stream.seek(0)
                st.sidebar.success("🔄 Conectado en vivo a OneDrive")
                return out_stream
        raise Exception("Token inactivo")
        
    except Exception:
        ruta_local = r"C:\Users\jsepulveda\Empaques y Servicios Superiores S.A.S\CONTABILIDAD SPK - DOCUMENTOS COMPARTIDOS - DOCUMENTOS COMPARTIDOS\CONTROL DOCUMENTOS ELECTRONICOS DIAN\INFORME FACTURACIÓN DIAN 2026.xlsx"
        if os.path.exists(ruta_local):
            st.sidebar.info("💻 Leyendo ruta física local (PC)")
            return ruta_local
        elif os.path.exists(EXCEL_PROV):
            st.sidebar.warning("📦 Leyendo base de datos de respaldo (GitHub)")
            return EXCEL_PROV
        else:
            st.error("❌ Error crítico: No se encontró ninguna fuente de datos disponible.")
            return None

# Ejecutar la carga inteligente
origen_datos = descargar_excel_onedrive()

if origen_datos is not None:
    try:
        # Cargar los datos brutos
        df = pd.read_excel(origen_datos, sheet_name=HOJA_PROV, engine="openpyxl")
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Limpieza básica de la base de datos
        df["TOTAL"] = pd.to_numeric(df["TOTAL"], errors="coerce").fillna(0)
        df["RESPONSABLE"] = df["RESPONSABLE"].astype(str).str.strip().str.upper()
        df["FECHA EMISIÓN"] = pd.to_datetime(df["FECHA EMISIÓN"], errors="coerce")
        
        # Eliminar registros sin fecha de emisión válida para el análisis temporal
        df = df.dropna(subset=["FECHA EMISIÓN"])

        # ==============================================================================
        # 🚨 FILTRO FILTRADO DE EXCLUSIÓN CRÍTICO (Solo áreas de la agenda)
        # ==============================================================================
        df_filtrado = df[df["RESPONSABLE"].isin(AREAS_SUPERPACK)].copy()

        # Extraer variables de tiempo básicas
        df_filtrado["MES_EMISION"] = df_filtrado["FECHA EMISIÓN"].dt.to_period("M").astype(str)
        df_filtrado["DIA_DEL_MES"] = df_filtrado["FECHA EMISIÓN"].dt.day

        # ==============================================================================
        # 🧠 LÓGICA DE SEGMENTACIÓN SEMANAL DENTRO DEL MES (Requerimiento Adicional)
        # ==============================================================================
        def clasificar_semana_mes(dia):
            if dia <= 7:
                return "1. Días 01 al 07"
            elif dia <= 15:
                return "2. Días 08 al 15"
            elif dia <= 21:
                return "3. Días 16 al 21"
            else:
                return "4. Días 22 al Cierre"

        df_filtrado["RANGO_SEMANAL"] = df_filtrado["DIA_DEL_MES"].apply(clasificar_semana_mes)

        # Métricas Macro Globales
        total_facturas_procesadas = len(df_filtrado)
        total_monto_gasto = df_filtrado["TOTAL"].sum()
        promedio_por_factura = total_monto_gasto / total_facturas_procesadas if total_facturas_procesadas > 0 else 0

        # Botón de actualización en barra lateral
        if st.sidebar.button("🔄 Actualizar Reporte"):
            st.rerun()

         # ==============================================================================
        # SECCIÓN DE FILTROS AVANZADOS EN LA BARRA LATERAL (NUEVO)
        # ==============================================================================
        st.sidebar.header("🎯 Filtros Avanzados")

        # 1. FILTRO DE MESES CON MÉTRICAS (Monto y Cantidad)
        resumen_meses = (
            df_filtrado.groupby("MES_EMISION")
            .agg(Total=("TOTAL", "sum"), Cantidad=("FOLIO", "count"))
            .reset_index()
        )

        opciones_meses = []
        mapeo_meses = {}
        for _, row in resumen_meses.sort_values(
            by="MES_EMISION", ascending=True
        ).iterrows():
            label = (
                f"{row['MES_EMISION']} ($ {row['Total']:,.0f} | {row['Cantidad']} Fac)"
            )
            opciones_meses.append(label)
            mapeo_meses[label] = row["MES_EMISION"]

        filtro_meses_labels = st.sidebar.multiselect(
            "Filtrar por Mes de Emisión:", options=opciones_meses
        )
        meses_seleccionados = [mapeo_meses[l] for l in filtro_meses_labels]

        # Filtrado intermedio por meses para delimitar proveedores y áreas
        df_mes_prev = df_filtrado.copy()
        if meses_seleccionados:
            df_mes_prev = df_mes_prev[
                df_mes_prev["MES_EMISION"].isin(meses_seleccionados)
            ]

        # 2. FILTRO DE RESPONSABLES / ÁREAS CON MÉTRICAS
        resumen_areas = (
            df_mes_prev.groupby("RESPONSABLE")
            .agg(Total=("TOTAL", "sum"), Cantidad=("FOLIO", "count"))
            .reset_index()
        )

        opciones_areas = []
        mapeo_areas = {}
        for _, row in resumen_areas.sort_values(by="Total", ascending=False).iterrows():
            label = (
                f"{row['RESPONSABLE']} ($ {row['Total']:,.0f} | {row['Cantidad']} Fac)"
            )
            opciones_areas.append(label)
            mapeo_areas[label] = row["RESPONSABLE"]

        filtro_areas_labels = st.sidebar.multiselect(
            "Filtrar por Área Responsable:", options=opciones_areas
        )
        areas_seleccionadas = [mapeo_areas[l] for l in filtro_areas_labels]

        # Filtrado intermedio por áreas
        df_area_prev = df_mes_prev.copy()
        if areas_seleccionadas:
            df_area_prev = df_area_prev[
                df_area_prev["RESPONSABLE"].isin(areas_seleccionadas)
            ]

        # 3. FILTRO DE PROVEEDORES CON MÉTRICAS
        resumen_provs = (
            df_area_prev.groupby("NOMBRE EMISOR")
            .agg(Total=("TOTAL", "sum"), Cantidad=("FOLIO", "count"))
            .reset_index()
        )

        opciones_provs = []
        mapeo_provs = {}
        for _, row in resumen_provs.sort_values(by="Total", ascending=False).iterrows():
            label = f"{row['NOMBRE EMISOR']} ($ {row['Total']:,.0f} | {row['Cantidad']} Fac)"
            opciones_provs.append(label)
            mapeo_provs[label] = row["NOMBRE EMISOR"]

        filtro_provs_labels = st.sidebar.multiselect(
            "Filtrar por Proveedor (Emisor):", options=opciones_provs
        )
        proveedores_seleccionados = [mapeo_provs[l] for l in filtro_provs_labels]

        # APLICACIÓN DE TODOS LOS FILTROS SELECCIONADOS AL DATAFRAME FINAL
        df_final = df_filtrado.copy()
        if meses_seleccionados:
            df_final = df_final[df_final["MES_EMISION"].isin(meses_seleccionados)]
        if areas_seleccionadas:
            df_final = df_final[df_final["RESPONSABLE"].isin(areas_seleccionadas)]
        if proveedores_seleccionados:
            df_final = df_final[
                df_final["NOMBRE EMISOR"].isin(proveedores_seleccionados)
            ]

        # RECALCULAR MÉTRICAS BASADAS EN LOS FILTROS SELECCIONADOS
        total_facturas_procesadas = len(df_final)
        total_monto_gasto = df_final["TOTAL"].sum()
        promedio_por_factura = (
            total_monto_gasto / total_facturas_procesadas
            if total_facturas_procesadas > 0
            else 0
        )

        # ==============================================================================
        # DESPLIEGUE VISUAL (MÉTRICAS PRINCIPALES DINÁMICAS)
        # ==============================================================================
        st.markdown("### 📊 Métricas de facturación de Proveedores (Año 2026)")
        col1, col2, col3 = st.columns(3)
        col1.metric("💰 Facturación Consolidada DIAN", f"$ {total_monto_compras:,.0f}")
        col2.metric(
            "📄 Volumen de Documentos Emitidos",
            f"{total_facturas_procesadas:,} Facturas",
        )
        col3.metric("🧮 Ticket Promedio por Factura", f"$ {promedio_por_factura:,.0f}")

        st.markdown("---")

        # ==============================================================================
        # SECCIÓN 1: ANALÍTICA DE COMPORTAMIENTO SEMANAL (RANGOS DE DÍAS)
        # ==============================================================================
        st.markdown("### ⏱️ Análisis del Comportamiento Semanal dentro del Mes")
        sem_col1, sem_col2 = st.columns(2)

        df_semanal_kpis = (
            df_final.groupby("RANGO_SEMANAL")
            .agg(Monto_Total=("TOTAL", "sum"), Cantidad_Total=("FOLIO", "count"))
            .reset_index()
            .sort_values(by="RANGO_SEMANAL")
        )

        with sem_col1:
            st.subheader("Distribución del Pasivo por Rangos de Días")
            fig_sem_val = px.bar(
                df_semanal_kpis,
                x="RANGO_SEMANAL",
                y="Monto_Total",
                template="plotly_white",
                color_discrete_sequence=["#0D47A1"],
                labels={
                    "RANGO_SEMANAL": "Semana del Mes",
                    "Monto_Total": "Valor Neto Total ($)",
                },
            )
            fig_sem_val.update_layout(yaxis_tickformat="$ ,.0f", separators=",.")
            fig_sem_val.update_traces(
                hovertemplate="<b>Rango:</b> %{x}<br><b>Pasivo Total:</b> $ %{y:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_sem_val, use_container_width=True)

        with sem_col2:
            st.subheader("Carga Operativa: Volumen de Facturas por Rangos de Días")
            fig_sem_cant = px.bar(
                df_semanal_kpis,
                x="RANGO_SEMANAL",
                y="Cantidad_Total",
                template="plotly_white",
                color_discrete_sequence=["#1565C0"],
                labels={
                    "RANGO_SEMANAL": "Semana del Mes",
                    "Cantidad_Total": "Cantidad de Facturas",
                },
            )
            fig_sem_cant.update_layout(yaxis_tickformat=",.0f", separators=",.")
            fig_sem_cant.update_traces(
                hovertemplate="<b>Rango:</b> %{x}<br><b>Cantidad:</b> %{y:,} Facturas<extra></extra>"
            )
            st.plotly_chart(fig_sem_cant, use_container_width=True)

        st.markdown("---")

        # ==============================================================================
        # SECCIÓN 2: ESTACIONALIDAD OPERATIVA Y TENDENCIA DINÁMICA
        # ==============================================================================
        st.markdown("---")
        st.markdown("### 📅 Estacionalidad Cronológica del Gasto")
        t_col1, t_col2 = st.columns(2)

        with t_col1:
            st.subheader(
                "Días del Mes con Mayor Emisión (Carga por Cantidad de Facturas)"
            )
            # Cambiamos la métrica a conteo ('count') de folios por cada número de día
            df_dia_mes = (
                df_final.groupby("DIA_DEL_MES")["FOLIO"]
                .count()
                .reset_index()
                .sort_values(by="DIA_DEL_MES")
            )
            df_dia_mes = df_dia_mes.rename(columns={"FOLIO": "Cantidad_Facturas"})

            # Cambiamos el color a azul corporativo (#1E88E5) para mantener la identidad de cantidad
            fig_dia = px.bar(
                df_dia_mes,
                x="DIA_DEL_MES",
                y="Cantidad_Facturas",
                template="plotly_white",
                color_discrete_sequence=["#1E88E5"],
                labels={
                    "DIA_DEL_MES": "Día del Calendario (1 al 31)",
                    "Cantidad_Facturas": "Facturas Emitidas",
                },
            )
            fig_dia.update_layout(yaxis_tickformat=",.0f", separators=",.")
            fig_dia.update_traces(
                hovertemplate="<b>Día del Mes:</b> %{x}<br><b>Volumen Recibido:</b> %{y:,} Facturas<extra></extra>"
            )
            st.plotly_chart(fig_dia, use_container_width=True)

        with t_col2:
            # LÓGICA DE ZOOM DINÁMICO: Mantiene el dinero por valor neto para contrastar
            if len(meses_seleccionados) == 1:
                st.subheader(
                    f"Evolución Diaria Detallada por Valor - {meses_seleccionados}"
                )
                df_mes_dinamico = (
                    df_final.groupby("FECHA EMISIÓN")["TOTAL"]
                    .sum()
                    .reset_index()
                    .sort_values(by="FECHA EMISIÓN")
                )
                fig_mes = px.line(
                    df_mes_dinamico,
                    x="FECHA EMISIÓN",
                    y="TOTAL",
                    markers=True,
                    template="plotly_white",
                    color_discrete_sequence=["#4CAF50"],
                    labels={"FECHA EMISIÓN": "Fecha", "TOTAL": "Monto ($)"},
                )
                fig_mes.update_layout(
                    xaxis_tickformat="%d/%m/%Y",
                    yaxis_tickformat="$ ,.0f",
                    separators=",.",
                    xaxis_title="Línea de Tiempo Diaria",
                )
                fig_mes.update_traces(
                    hovertemplate="<b>Fecha:</b> %{x|%d/%m/%Y}<br><b>Total Compras:</b> $ %{y:,.0f}<extra></extra>"
                )
            else:
                st.subheader("Evolución de la Facturación Mensual por Valor 2026")
                df_mes_macro = (
                    df_final.groupby("MES_EMISION")["TOTAL"]
                    .sum()
                    .reset_index()
                    .sort_values(by="MES_EMISION")
                )
                fig_mes = px.line(
                    df_mes_macro,
                    x="MES_EMISION",
                    y="TOTAL",
                    markers=True,
                    template="plotly_white",
                    color_discrete_sequence=["#4CAF50"],
                    labels={"MES_EMISION": "Mes", "TOTAL": "Monto Facturado ($)"},
                )
                fig_mes.update_layout(
                    yaxis_tickformat="$ ,.0f",
                    separators=",.",
                    xaxis_title="Mes de Emisión",
                )
                fig_mes.update_traces(
                    hovertemplate="<b>Mes:</b> %{x}<br><b>Total Gasto:</b> $ %{y:,.0f}<extra></extra>"
                )

            st.plotly_chart(fig_mes, use_container_width=True)

        st.markdown("---")

        # ==============================================================================
        # SECCIÓN 3: CONCENTRACIÓN POR ÁREAS (VALOR VS CANTIDAD) Y PROVEEDORES
        # ==============================================================================
        st.markdown("### 🏢 Concentración Estructural de las Compras por Área")
        g1, g2 = st.columns(2)

        # Agrupar datos macro por área responsable una sola vez
        df_resumen_area = (
            df_final.groupby("RESPONSABLE")
            .agg(Monto_Total=("TOTAL", "sum"), Cantidad_Total=("FOLIO", "count"))
            .reset_index()
        )

        with g1:
            st.subheader("Distribución del Gasto Total por Área ($)")
            df_resumen_area_val = df_resumen_area.sort_values(
                by="Monto_Total", ascending=False
            )
            fig_bar_val = px.bar(
                df_resumen_area_val,
                x="RESPONSABLE",
                y="Monto_Total",
                template="plotly_white",
                color_discrete_sequence=["#4CAF50"],
                labels={"RESPONSABLE": "Área", "Monto_Total": "Monto ($)"},
            )
            fig_bar_val.update_layout(yaxis_tickformat="$ ,.0f", separators=",.")
            fig_bar_val.update_traces(
                hovertemplate="<b>Área:</b> %{x}<br><b>Total Compras:</b> $ %{y:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_bar_val, use_container_width=True)

        with g2:
            st.subheader("Carga Operativa: Cantidad de Facturas por Área (Unidades)")
            df_resumen_area_cant = df_resumen_area.sort_values(
                by="Cantidad_Total", ascending=False
            )
            fig_bar_cant = px.bar(
                df_resumen_area_cant,
                x="RESPONSABLE",
                y="Cantidad_Total",
                template="plotly_white",
                color_discrete_sequence=["#1E88E5"],
                labels={
                    "RESPONSABLE": "Área",
                    "Cantidad_Total": "Cantidad de Facturas",
                },
            )
            fig_bar_cant.update_layout(yaxis_tickformat=",.0f", separators=",.")
            fig_bar_cant.update_traces(
                hovertemplate="<b>Área:</b> %{x}<br><b>Facturas Recibidas:</b> %{y:,} Fac<extra></extra>"
            )
            st.plotly_chart(fig_bar_cant, use_container_width=True)

        # ==============================================================================
        # SECCIÓN 3: CONCENTRACIÓN POR ÁREAS Y ANÁLISIS DE PROVEEDORES (VALOR VS CANTIDAD)
        # ==============================================================================
        st.markdown("---")
        st.markdown("### 🏬 Análisis de Principales Proveedores (Top 10)")
        p_g1, p_g2 = st.columns(2)

        # Agrupar datos macro por proveedor (Emisor) una sola vez para optimizar rendimiento
        df_prov_consolidado = (
            df_final.groupby("NOMBRE EMISOR")
            .agg(Monto_Total=("TOTAL", "sum"), Cantidad_Total=("FOLIO", "count"))
            .reset_index()
        )

        with p_g1:
            st.subheader("Top 10 Proveedores con Mayor Volumen de Facturación ($)")
            df_prov_top_val = df_prov_consolidado.sort_values(
                by="Monto_Total", ascending=False
            ).head(10)
            fig_prov_top_val = px.bar(
                df_prov_top_val,
                x="Monto_Total",
                y="NOMBRE EMISOR",
                orientation="h",
                template="plotly_white",
                color_discrete_sequence=["#2E7D32"],
                labels={
                    "Monto_Total": "Total Facturado ($)",
                    "NOMBRE EMISOR": "Proveedor",
                },
            )
            fig_prov_top_val.update_layout(xaxis_tickformat="$ ,.0f", separators=",.")
            fig_prov_top_val.update_traces(
                hovertemplate="<b>Proveedor:</b> %{y}<br><b>Total Facturado:</b> $ %{x:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_prov_top_val, use_container_width=True)

        with p_g2:
            st.subheader("Top 10 Proveedores con Mayor Cantidad de Facturas Emitidas")
            df_prov_top_cant = df_prov_consolidado.sort_values(
                by="Cantidad_Total", ascending=False
            ).head(10)
            fig_prov_top_cant = px.bar(
                df_prov_top_cant,
                x="Cantidad_Total",
                y="NOMBRE EMISOR",
                orientation="h",
                template="plotly_white",
                color_discrete_sequence=["#1E88E5"],
                labels={
                    "Cantidad_Total": "Cantidad de Facturas",
                    "NOMBRE EMISOR": "Proveedor",
                },
            )
            fig_prov_top_cant.update_layout(xaxis_tickformat=",.0f", separators=",.")
            fig_prov_top_cant.update_traces(
                hovertemplate="<b>Proveedor:</b> %{y}<br><b>Facturas Recibidas:</b> %{x:,} Fac<extra></extra>"
            )
            st.plotly_chart(fig_prov_top_cant, use_container_width=True)

        # ==============================================================================
        # TABLA GENERAL DE AUDITORÍA DIAN
        # ==============================================================================
        st.markdown("---")
        st.subheader(
            "🔍 Explorador de Documentos DIAN (Áreas Agendadas - Corte al 21 de Mayo)"
        )

        columnas_vista = [
            "FOLIO",
            "NOMBRE EMISOR",
            "TOTAL",
            "RESPONSABLE",
            "FECHA EMISIÓN",
        ]
        df_tabla = df_final[columnas_vista].sort_values(
            by="FECHA EMISIÓN", ascending=False
        )

        st.dataframe(
            df_tabla.style.format(
                {
                    "TOTAL": lambda x: f"$ {x:,.0f}".replace(",", "X")
                    .replace(".", ",")
                    .replace("X", "."),
                    "FECHA EMISIÓN": lambda t: (
                        t.strftime("%d/%m/%Y") if pd.notnull(t) else ""
                    ),
                }
            ),
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Error procesando Facturación DIAN. Detalle: {e}")
else:
    st.error(
        f"❌ No se encontró el archivo de Excel en la ruta especificada: {EXCEL_PROV}"
    )
    
