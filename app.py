import streamlit as st
import pandas as pd
from datetime import datetime, date
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

# ==============================
# CONFIGURACIÓN
# ==============================
st.set_page_config(
    page_title="Portal Kaizen",
    layout="wide"
)

# ==============================
# CONEXIÓN A GOOGLE SHEETS
# ==============================
conn = st.connection("gsheets", type=GSheetsConnection)
URL_SHEET = st.secrets["spreadsheet"]

COLS_PLAN = ['id', 'fecha_hallazgo', 'sede', 'objetivo', 'objetivo_especifico',
             'origen', 'hallazgo', 'ocurrencia', 'consecuencia', 'valor',
             'prioridad', 'accion', 'tipo_accion', 'responsable', 'sucursal', 'deadline']
COLS_SEG  = ['id', 'seguimiento', 'fecha_seguimiento', 'porc_avance', 'estado', 'evidencia']

@st.cache_data(ttl=300)
def load_data():
    try:
        df_plan = conn.read(spreadsheet=URL_SHEET, worksheet="plan_kaizen", usecols=list(range(len(COLS_PLAN))))
        df_plan = df_plan.dropna(how='all')
    except Exception:
        df_plan = pd.DataFrame(columns=COLS_PLAN)

    try:
        df_seguimiento = conn.read(spreadsheet=URL_SHEET, worksheet="seguimiento", usecols=list(range(len(COLS_SEG))))
        df_seguimiento = df_seguimiento.dropna(how='all')
    except Exception:
        df_seguimiento = pd.DataFrame(columns=COLS_SEG)

    # Normalizar nombres de columnas
    df_plan.columns = df_plan.columns.str.strip().str.lower()
    df_seguimiento.columns = df_seguimiento.columns.str.strip().str.lower()

    # Fechas plan
    for col in ['fecha_hallazgo', 'deadline']:
        if col in df_plan.columns:
            df_plan[col] = pd.to_datetime(df_plan[col], errors="coerce").dt.normalize()

    # Fechas seguimiento
    if 'fecha_seguimiento' in df_seguimiento.columns:
        df_seguimiento['fecha_seguimiento'] = pd.to_datetime(df_seguimiento['fecha_seguimiento'], errors="coerce").dt.normalize()

    # Limpiar % avance
    if 'porc_avance' in df_seguimiento.columns:
        df_seguimiento['porc_avance'] = (
            df_seguimiento['porc_avance']
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
            .replace({"": "0", "nan": "0"})
            .astype(float)
        )

    # Tipos numéricos plan
    for col in ['id', 'ocurrencia', 'consecuencia', 'valor']:
        if col in df_plan.columns:
            df_plan[col] = pd.to_numeric(df_plan[col], errors='coerce').fillna(0).astype('int64')

    return df_plan, df_seguimiento

def save_data(df_plan, df_seguimiento):
    """Guardar ambos DataFrames en Google Sheets"""
    df_plan_save = df_plan.copy()
    df_seg_save  = df_seguimiento.copy()

    # Convertir fechas a string ISO para que Sheets las almacene bien
    for col in ['fecha_hallazgo', 'deadline']:
        if col in df_plan_save.columns:
            df_plan_save[col] = pd.to_datetime(df_plan_save[col], errors='coerce').dt.strftime('%Y-%m-%d')

    if 'fecha_seguimiento' in df_seg_save.columns:
        df_seg_save['fecha_seguimiento'] = pd.to_datetime(df_seg_save['fecha_seguimiento'], errors='coerce').dt.strftime('%Y-%m-%d')

    conn.update(spreadsheet=URL_SHEET, worksheet="plan_kaizen", data=df_plan_save)
    conn.update(spreadsheet=URL_SHEET, worksheet="seguimiento", data=df_seg_save)

    st.cache_data.clear()

# Cargar datos
if 'df_plan' not in st.session_state:
    df_plan, df_seguimiento = load_data()
    st.session_state.df_plan = df_plan
    st.session_state.df_seguimiento = df_seguimiento
else:
    df_plan = st.session_state.df_plan
    df_seguimiento = st.session_state.df_seguimiento

# ==============================
# FUNCIONES AUXILIARES
# ==============================
def get_next_id(df_plan):
    """Obtener el siguiente ID disponible"""
    if df_plan.empty or 'id' not in df_plan.columns:
        return 1
    return int(df_plan['id'].max()) + 1

def actualizar_estado_accion(id_accion, df_seguimiento):
    """Actualizar estado general de una acción basado en sus seguimientos"""
    seguimientos_accion = df_seguimiento[df_seguimiento['id'] == id_accion]
    
    if seguimientos_accion.empty:
        return "No iniciado"
    
    ultimo_estado = seguimientos_accion.iloc[-1]['estado'] if 'estado' in seguimientos_accion.columns else "Proceso"
    
    if ultimo_estado == "Finalizado":
        ultimo_avance = seguimientos_accion.iloc[-1]['porc_avance'] if 'porc_avance' in seguimientos_accion.columns else 0
        if ultimo_avance >= 100:
            return "Cerrado"
        else:
            return "En proceso"
    else:
        return "En proceso"

def calcular_avance_total(id_accion, df_seguimiento):
    """Calcular el avance total de una acción"""
    seguimientos_accion = df_seguimiento[df_seguimiento['id'] == id_accion]
    if seguimientos_accion.empty or 'porc_avance' not in seguimientos_accion.columns:
        return 0
    return float(seguimientos_accion.iloc[-1]['porc_avance'])

def calcular_semaforo(row, df_seguimiento):
    if row['estado_general'] == "Cerrado":
        return "🟢 Cerrado"
    
    if pd.isna(row.get('deadline', None)):
        return "⚪ Sin fecha"
    
    hoy = pd.Timestamp.now().normalize()
    deadline = pd.Timestamp(row['deadline']).normalize()
    
    if deadline < hoy:
        return "🔴 Atrasado"
    elif (deadline - hoy).days <= 7:
        return "🟡 En riesgo"
    else:
        return "🟢 OK"

def preparar_datos_completos(df_plan, df_seguimiento):
    """Preparar datos con estados y avances"""
    df_completo = df_plan.copy()
    df_completo['estado_general'] = df_completo['id'].apply(lambda x: actualizar_estado_accion(x, df_seguimiento))
    df_completo['porc_avance'] = df_completo['id'].apply(lambda x: calcular_avance_total(x, df_seguimiento))
    df_completo['SEMAFORO'] = df_completo.apply(lambda row: calcular_semaforo(row, df_seguimiento), axis=1)
    df_completo['dias_restantes'] = df_completo['deadline'].apply(
        lambda x: (pd.Timestamp(x).normalize() - pd.Timestamp.now().normalize()).days if pd.notna(x) else None
    )
    return df_completo

# ==============================
# PESTAÑA 1: MIS ACCIONES (USUARIO)
# ==============================
def pestania_mis_acciones():
    st.header("👤 Mis Acciones")
    st.caption("Busca y da seguimiento a tus acciones asignadas")
    
    if st.session_state.df_plan.empty:
        st.warning("⚠️ No hay acciones registradas aún")
        return
    
    # Preparar datos completos
    df_completo = preparar_datos_completos(st.session_state.df_plan, st.session_state.df_seguimiento)
    
    # Selección de usuario
    responsables = sorted(df_completo['responsable'].dropna().unique())
    usuario_seleccionado = st.selectbox("Selecciona tu nombre", responsables, key="usuario_select")
    
    # Filtrar por usuario
    df_usuario = df_completo[df_completo['responsable'] == usuario_seleccionado].copy()
    
    if df_usuario.empty:
        st.info(f"No hay acciones asignadas a {usuario_seleccionado}")
        return
    
    # ===== FILTROS =====
    st.subheader("🔍 Filtros de búsqueda")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        rango_fechas = st.date_input(
            "Rango de fechas (hallazgo)",
            value=(),
            key="fechas_usuario"
        )
    
    with col2:
        estados_disponibles = sorted(df_usuario['estado_general'].unique())
        estados_seleccionados = st.multiselect("Estado", estados_disponibles, default=estados_disponibles, key="estados_usuario")
    
    with col3:
        prioridades_disponibles = sorted(df_usuario['prioridad'].dropna().unique())
        if prioridades_disponibles:
            prioridades_seleccionadas = st.multiselect("Prioridad", prioridades_disponibles, default=prioridades_disponibles, key="prioridades_usuario")
        else:
            prioridades_seleccionadas = []
            st.multiselect("Prioridad", [], default=[], key="prioridades_usuario_empty")
    
    with col4:
        # Manejar el caso cuando todos los valores son iguales
        valor_min = int(df_usuario['valor'].min())
        valor_max = int(df_usuario['valor'].max())
        
        if valor_min == valor_max:
            # Si son iguales, mostrar un número único en lugar de un slider
            st.write(f"**Valor:** {valor_min}")
            rango_valor = (valor_min, valor_max)
        else:
            rango_valor = st.slider(
                "Valor (ocurrencia × consecuencia)",
                min_value=valor_min,
                max_value=valor_max,
                value=(valor_min, valor_max),
                key="valor_slider_usuario"
            )
    
    # Aplicar filtros
    df_filtrado = df_usuario.copy()
    
    if len(rango_fechas) == 2:
        fecha_ini, fecha_fin = rango_fechas
        df_filtrado = df_filtrado[
            (df_filtrado['fecha_hallazgo'].dt.date >= fecha_ini) & 
            (df_filtrado['fecha_hallazgo'].dt.date <= fecha_fin)
        ]
    
    if estados_seleccionados:
        df_filtrado = df_filtrado[df_filtrado['estado_general'].isin(estados_seleccionados)]
    
    if prioridades_seleccionadas:
        df_filtrado = df_filtrado[df_filtrado['prioridad'].isin(prioridades_seleccionadas)]
    
    if isinstance(rango_valor, tuple) and len(rango_valor) == 2:
        df_filtrado = df_filtrado[
            (df_filtrado['valor'] >= rango_valor[0]) & 
            (df_filtrado['valor'] <= rango_valor[1])
        ]
    
    # ===== ESTADÍSTICAS =====
    st.subheader("📊 Mis Estadísticas")
    
    if not df_filtrado.empty:
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_acciones = len(df_filtrado)
        cerradas = len(df_filtrado[df_filtrado['estado_general'] == "Cerrado"])
        en_proceso = len(df_filtrado[df_filtrado['estado_general'] == "En proceso"])
        no_iniciado = len(df_filtrado[df_filtrado['estado_general'] == "No iniciado"])
        atrasadas = len(df_filtrado[df_filtrado['SEMAFORO'].str.contains("🔴", na=False)])
        
        col1.metric("📋 Total acciones", total_acciones)
        col2.metric("✅ Cerradas", cerradas, delta=f"{cerradas/total_acciones*100:.0f}%" if total_acciones > 0 else "0%")
        col3.metric("🟡 En proceso", en_proceso)
        col4.metric("⚪ No iniciado", no_iniciado)
        col5.metric("🔴 Atrasadas", atrasadas, delta="⚠️" if atrasadas > 0 else None)
        
        # Gráficos
        col_graf1, col_graf2 = st.columns(2)
        
        with col_graf1:
            # Gráfico de estado
            estado_counts = df_filtrado['estado_general'].value_counts()
            if not estado_counts.empty:
                fig_pie = px.pie(values=estado_counts.values, names=estado_counts.index, 
                               title="Distribución por Estado", color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_pie, use_container_width=True, key="estado_pie")
            else:
                st.info("No hay datos para mostrar")
        
        with col_graf2:
            # Gráfico de prioridades
            prioridad_counts = df_filtrado['prioridad'].value_counts()
            if not prioridad_counts.empty:
                fig_bar = px.bar(x=prioridad_counts.index, y=prioridad_counts.values, 
                               title="Acciones por Prioridad", color=prioridad_counts.index,
                               color_discrete_sequence=px.colors.qualitative.Set3)
                st.plotly_chart(fig_bar, use_container_width=True, key="prioridad_bar")
            else:
                st.info("No hay datos para mostrar")
        
        # Tabla de acciones
        st.subheader("📋 Detalle de mis acciones")
        
        columnas_mostrar = ['id', 'fecha_hallazgo', 'hallazgo', 'prioridad', 'estado_general', 
                           'porc_avance', 'SEMAFORO', 'deadline', 'dias_restantes', 'valor']
        
        columnas_disponibles = [col for col in columnas_mostrar if col in df_filtrado.columns]
        
        df_display = df_filtrado[columnas_disponibles].copy()
        df_display['porc_avance'] = df_display['porc_avance'].round(0).astype('int64')
        
        # Formatear fechas
        if 'fecha_hallazgo' in df_display.columns:
            df_display['fecha_hallazgo'] = df_display['fecha_hallazgo'].dt.date
        if 'deadline' in df_display.columns:
            df_display['deadline'] = df_display['deadline'].dt.date
        
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Detalle expandido de acción seleccionada
        st.subheader("🔍 Ver detalle completo")
        if not df_filtrado.empty:
            acciones_opciones = {f"#{row['id']} - {row.get('hallazgo', 'Sin descripción')[:80]}": row['id'] 
                                for _, row in df_filtrado.iterrows()}
            
            accion_seleccionada = st.selectbox("Selecciona una acción para ver detalles", list(acciones_opciones.keys()), key="detalle_accion_usuario")
            id_seleccionado = acciones_opciones[accion_seleccionada]
            
            detalle = df_filtrado[df_filtrado['id'] == id_seleccionado].iloc[0]
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**📝 Hallazgo**")
                st.write(detalle['hallazgo'])
                st.markdown("**🎯 Acción propuesta**")
                st.write(detalle.get('accion', 'N/A'))
                st.markdown("**👤 Responsable**")
                st.write(detalle['responsable'])
            
            with col2:
                st.markdown("**⚡ Prioridad**")
                st.write(detalle['prioridad'])
                st.markdown("**📊 Avance actual**")
                st.progress(detalle['porc_avance']/100)
                st.write(f"{detalle['porc_avance']:.0f}%")
                st.markdown("**🚦 Estado**")
                st.write(detalle['SEMAFORO'])
    else:
        st.warning("No hay acciones que coincidan con los filtros seleccionados")

# ==============================
# PESTAÑA 2: PLAN KAIZEN (ADMINISTRADOR)
# ==============================
def pestania_plan_kaizen():
    st.header("📝 Plan Kaizen - Administración")
    st.caption("Gestión completa del plan de acción (solo administradores)")
    
    # ===== SECCIÓN DE ALTA =====
    with st.expander("➕ Registrar nueva acción", expanded=False):
        with st.form("nueva_accion", clear_on_submit=True):
            st.subheader("Nueva Acción")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                fecha_hallazgo = st.date_input("Fecha del hallazgo", datetime.now())
                sede = st.text_input("Sede *", placeholder="Ej: Nacional, Cochabamba...")
                objetivo = st.text_input("Objetivo", placeholder="Ej: Objetivos nacionales, Objetivos Kaizen...")
                objetivo_especifico = st.text_input("Objetivo específico", placeholder="Ej: Retention, Q KAIZEN...")
                origen = st.text_input("Origen", placeholder="Ej: 5Ss, Estructuración del área...")
            
            with col2:
                hallazgo = st.text_area("Hallazgo *", height=100, placeholder="Describe el hallazgo encontrado...")
                ocurrencia = st.number_input("Ocurrencia (1-5)", min_value=1, max_value=5, value=3)
                consecuencia = st.number_input("Consecuencia (1-5)", min_value=1, max_value=5, value=3)
                valor = ocurrencia * consecuencia
                st.metric("Valor (ocurrencia × consecuencia)", valor)
            
            with col3:
                prioridad = st.selectbox("Prioridad", ["Baja", "Media", "Alta", "Crítica"])
                accion = st.text_area("Acción propuesta", height=100, placeholder="Describe la acción a realizar...")
                tipo_accion = st.selectbox("Tipo de acción", ["Correctiva", "Mejora", "Preventiva"])
                responsable = st.text_input("Responsable *")
                sucursal = st.text_input("Sucursal", placeholder="Ej: Centro de Servicio, Nacional...")
                deadline = st.date_input("Deadline (fecha límite)", datetime.now())
            
            submitted = st.form_submit_button("✅ Registrar Acción")
            
            if submitted:
                if not sede or not hallazgo or not responsable:
                    st.error("❌ Por favor completa los campos obligatorios (*)")
                else:
                    nuevo_id = get_next_id(st.session_state.df_plan)
                    
                    nueva_accion = pd.DataFrame([{
                        'id': nuevo_id,
                        'fecha_hallazgo': pd.Timestamp(fecha_hallazgo).normalize(),
                        'sede': sede,
                        'objetivo': objetivo,
                        'objetivo_especifico': objetivo_especifico,
                        'origen': origen,
                        'hallazgo': hallazgo,
                        'ocurrencia': ocurrencia,
                        'consecuencia': consecuencia,
                        'valor': valor,
                        'prioridad': prioridad,
                        'accion': accion,
                        'tipo_accion': tipo_accion,
                        'responsable': responsable,
                        'sucursal': sucursal,
                        'deadline': pd.Timestamp(deadline).normalize()
                    }])
                    
                    st.session_state.df_plan = pd.concat([st.session_state.df_plan, nueva_accion], ignore_index=True)
                    save_data(st.session_state.df_plan, st.session_state.df_seguimiento)
                    st.success(f"✅ Acción #{nuevo_id} registrada exitosamente!")
                    st.rerun()
    
    # ===== SECCIÓN DE BÚSQUEDA Y FILTROS =====

    col_titulo, col_actualizar = st.columns([4, 1])

    with col_titulo:
        st.subheader("🔍 Búsqueda y filtros del Plan")

    with col_actualizar:
        if st.button(
            "🔄 Actualizar datos",
            key="actualizar_datos_plan",
            use_container_width=True
        ):
            # Limpiar la caché
            st.cache_data.clear()

            # Recargar los datos directamente desde Google Sheets
            df_plan_actualizado, df_seguimiento_actualizado = load_data()

            # Actualizar los DataFrames de la sesión
            st.session_state.df_plan = df_plan_actualizado
            st.session_state.df_seguimiento = df_seguimiento_actualizado

            # Reiniciar los selectores para incorporar las opciones nuevas
            filtros_a_reiniciar = [
                "responsable_admin",
                "estado_admin",
                "prioridad_admin",
                "sede_admin",
                "origen_admin",
                "objetivo_admin",
                "objetivo_especifico_admin"
            ]

            for filtro in filtros_a_reiniciar:
                st.session_state.pop(filtro, None)

            # Actualizar la interfaz
            st.rerun()

    # Preparar datos completos
    df_completo = preparar_datos_completos(
        st.session_state.df_plan,
        st.session_state.df_seguimiento
    )

    # Función auxiliar para obtener valores únicos de forma segura
    def obtener_opciones(df, columna):
        if columna not in df.columns:
            return []
        return sorted(
            df[columna].dropna().unique().tolist(),
            key=lambda x: str(x).lower()
        )

    # ===== FILTROS =====
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        # Búsqueda por texto
        texto_busqueda = st.text_input(
            "🔎 Buscar por ID, hallazgo o acción",
            placeholder="Ej: 123, mantenimiento, limpieza...",
            key="texto_busqueda_admin"
        )

        # Filtro por responsable
        responsables_all = obtener_opciones(
            df_completo, "responsable"
        )

        responsables_filter = st.multiselect(
            "Responsable",
            responsables_all,
            key="responsable_admin"
        )

    with col2:
        # Filtro por estado
        estados_all = obtener_opciones(
            df_completo, "estado_general"
        )

        estados_filter = st.multiselect(
            "Estado",
            estados_all,
            default=estados_all,
            key="estado_admin"
        )

        # Filtro por prioridad
        prioridades_all = obtener_opciones(
            df_completo, "prioridad"
        )

        prioridades_filter = st.multiselect(
            "Prioridad",
            prioridades_all,
            default=prioridades_all,
            key="prioridad_admin"
        )

    with col3:
        # Filtro por sede
        sedes_all = obtener_opciones(
            df_completo, "sede"
        )

        sedes_filter = st.multiselect(
            "Sede",
            sedes_all,
            default=sedes_all,
            key="sede_admin"
        )

        # Rango de fechas
        rango_fechas_admin = st.date_input(
            "Rango de fechas (hallazgo)",
            value=(),
            key="fechas_admin"
        )

    with col4:
        # Filtro por origen
        origenes_all = obtener_opciones(
            df_completo, "origen"
        )

        origen_filter = st.multiselect(
            "Origen",
            origenes_all,
            default=origenes_all,
            key="origen_admin"
        )

        # Filtro por objetivo
        objetivos_all = obtener_opciones(
            df_completo, "objetivo"
        )

        objetivo_filter = st.multiselect(
            "Objetivo",
            objetivos_all,
            default=objetivos_all,
            key="objetivo_admin"
        )

        # Filtro por objetivo específico
        objetivos_especificos_all = obtener_opciones(
            df_completo, "objetivo_especifico"
        )

        objetivo_especifico_filter = st.multiselect(
            "Objetivo específico",
            objetivos_especificos_all,
            default=objetivos_especificos_all,
            key="objetivo_especifico_admin"
        )


    # ===== APLICAR FILTROS =====
    df_filtrado_admin = df_completo.copy()

    # 1. Búsqueda por texto
    if texto_busqueda:
        texto = texto_busqueda.strip()

        if texto:
            # Columnas disponibles para la búsqueda
            columnas_busqueda = [
                col for col in ["id", "hallazgo", "accion"]
                if col in df_filtrado_admin.columns
            ]

            if columnas_busqueda:
                mascara_busqueda = pd.Series(
                    False,
                    index=df_filtrado_admin.index
                )

                for columna in columnas_busqueda:
                    mascara_busqueda |= (
                        df_filtrado_admin[columna]
                        .astype(str)
                        .str.contains(
                            texto,
                            case=False,
                            na=False,
                            regex=False
                        )
                    )

                df_filtrado_admin = df_filtrado_admin[
                    mascara_busqueda
                ]

    # 2. Filtro por responsable
    if responsables_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["responsable"].isin(
                responsables_filter
            )
        ]

    # 3. Filtro por estado
    if estados_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["estado_general"].isin(
                estados_filter
            )
        ]

    # 4. Filtro por prioridad
    if prioridades_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["prioridad"].isin(
                prioridades_filter
            )
        ]

    # 5. Filtro por sede
    if sedes_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["sede"].isin(
                sedes_filter
            )
        ]

    # 6. Filtro por rango de fechas
    if len(rango_fechas_admin) == 2:
        fecha_ini, fecha_fin = rango_fechas_admin

        if "fecha_hallazgo" in df_filtrado_admin.columns:
            fechas = pd.to_datetime(
                df_filtrado_admin["fecha_hallazgo"],
                errors="coerce"
            )

            df_filtrado_admin = df_filtrado_admin[
                fechas.notna()
                & (fechas.dt.date >= fecha_ini)
                & (fechas.dt.date <= fecha_fin)
            ]

    # 7. Filtro por origen
    if origen_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["origen"].isin(
                origen_filter
            )
        ]

    # 8. Filtro por objetivo
    if objetivo_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["objetivo"].isin(
                objetivo_filter
            )
        ]

    # 9. Filtro por objetivo específico
    if objetivo_especifico_filter:
        df_filtrado_admin = df_filtrado_admin[
            df_filtrado_admin["objetivo_especifico"].isin(
                objetivo_especifico_filter
            )
        ]

    # ===== RESULTADO DE LOS FILTROS =====
    st.caption(
        f"📊 Registros encontrados: {len(df_filtrado_admin):,} "
        f"de {len(df_completo):,}"
    )

    # Estadísticas generales
    if not df_filtrado_admin.empty:
        st.subheader("📊 Estadísticas del Plan")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_plan = len(df_filtrado_admin)
        cerradas_plan = len(df_filtrado_admin[df_filtrado_admin['estado_general'] == "Cerrado"])
        atrasadas_plan = len(df_filtrado_admin[df_filtrado_admin['SEMAFORO'].str.contains("🔴", na=False)])
        riesgo_plan = len(df_filtrado_admin[df_filtrado_admin['SEMAFORO'].str.contains("🟡", na=False)])
        avance_promedio = df_filtrado_admin['porc_avance'].mean()
        
        col1.metric("📋 Total acciones", total_plan)
        col2.metric("✅ Cerradas", cerradas_plan, delta=f"{cerradas_plan/total_plan*100:.0f}%" if total_plan > 0 else "0%")
        col3.metric("🔴 Atrasadas", atrasadas_plan, delta="⚠️ Crítico" if atrasadas_plan > 0 else None)
        col4.metric("🟡 En riesgo", riesgo_plan)
        col5.metric("📈 Avance promedio", f"{avance_promedio:.0f}%")
        
        # Tabla completa
        st.subheader("📋 Todas las acciones")
        
        columnas_admin = ['id', 'fecha_hallazgo', 'sede', 'accion', 'responsable', 'prioridad', 
                         'estado_general', 'porc_avance', 'SEMAFORO', 'deadline', 'valor']
        
        columnas_admin_disponibles = [col for col in columnas_admin if col in df_filtrado_admin.columns]
        
        df_admin_display = df_filtrado_admin[columnas_admin_disponibles].copy()
        df_admin_display['porc_avance'] = df_admin_display['porc_avance'].round(0).astype('int64')
        
        if 'fecha_hallazgo' in df_admin_display.columns:
            df_admin_display['fecha_hallazgo'] = df_admin_display['fecha_hallazgo'].dt.date
        if 'deadline' in df_admin_display.columns:
            df_admin_display['deadline'] = df_admin_display['deadline'].dt.date
        
        st.dataframe(df_admin_display, use_container_width=True, height=500)
    else:
        st.warning("No hay acciones que coincidan con los filtros seleccionados")

# ==============================
# PESTAÑA 3: SEGUIMIENTO
# ==============================
def pestania_seguimiento():
    st.header("📊 Seguimiento de Acciones")
    st.caption("Registra avances línea por línea hasta el cierre")
    
    if st.session_state.df_plan.empty:
        st.warning("⚠️ No hay acciones registradas. Ve a la pestaña 'Plan Kaizen' para crear una acción.")
        return
    
    # Preparar datos para búsqueda
    df_completo = preparar_datos_completos(st.session_state.df_plan, st.session_state.df_seguimiento)
    
    # ===== BÚSQUEDA DE ACCIÓN =====
    st.subheader("🔍 Buscar acción")
    
    col_busq1, col_busq2 = st.columns(2)
    
    with col_busq1:
        texto_busqueda_seg = st.text_input("Buscar por ID, hallazgo o acción", 
                                           placeholder="Ej: 123, mantenimiento, limpieza...",
                                           key="busqueda_seguimiento")
    
    # Filtrar acciones
    df_busqueda = df_completo.copy()
    if texto_busqueda_seg:
        df_busqueda = df_busqueda[
            df_busqueda['id'].astype(str).str.contains(texto_busqueda_seg, case=False, na=False) |
            df_busqueda['hallazgo'].str.contains(texto_busqueda_seg, case=False, na=False) |
            df_busqueda['accion'].str.contains(texto_busqueda_seg, case=False, na=False)
        ]
    
    if df_busqueda.empty:
        st.warning("No se encontraron acciones con ese criterio de búsqueda")
        return
    
    # Seleccionar acción
    with col_busq2:
        opciones_acciones = {f"#{row['id']} - {row.get('accion', 'Sin descripción')[:80]}": row['id']
                            for _, row in df_busqueda.iterrows()}
        accion_seleccionada = st.selectbox("Selecciona una acción", list(opciones_acciones.keys()), key="accion_seguimiento_select")
        id_accion = opciones_acciones[accion_seleccionada]
    
    # Mostrar información de la acción seleccionada
    accion_info = st.session_state.df_plan[st.session_state.df_plan['id'] == id_accion].iloc[0]
    
    with st.expander("📋 Ver detalles de la acción", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Hallazgo:** {accion_info.get('hallazgo', 'N/A')}")
            st.markdown(f"**Acción propuesta:** {accion_info.get('accion', 'N/A')}")
            st.markdown(f"**Responsable:** {accion_info.get('responsable', 'N/A')}")
        with col2:
            st.markdown(f"**Prioridad:** {accion_info.get('prioridad', 'N/A')}")
            deadline = accion_info.get('deadline', 'N/A')
            if pd.notna(deadline):
                deadline = pd.Timestamp(deadline).strftime('%d/%m/%Y')
            st.markdown(f"**Deadline:** {deadline}")
            st.markdown(f"**Sede:** {accion_info.get('sede', 'N/A')}")
    
    # Registro de nuevo seguimiento
    st.subheader("📝 Registrar nuevo seguimiento")
    
    with st.form("nuevo_seguimiento", clear_on_submit=True):
        seguimiento_texto = st.text_area("Seguimiento / Avance *", height=100, 
                                         placeholder="Describe el avance realizado...")
        fecha_seguimiento = st.date_input("Fecha del seguimiento", datetime.now())
        
        # Permitir ingresar % de avance
        seguimientos_actuales = st.session_state.df_seguimiento[st.session_state.df_seguimiento['id'] == id_accion]
        ultimo_avance = seguimientos_actuales['porc_avance'].iloc[-1] if not seguimientos_actuales.empty else 0
        
        col_avance, col_estado = st.columns(2)
        with col_avance:
            porc_avance = st.slider(f"% de avance (actual: {ultimo_avance:.0f}%)", 
                                    min_value=0.0, max_value=100.0, 
                                    value=float(ultimo_avance), step=5.0,
                                    key="avance_slider")
        with col_estado:
            if porc_avance >= 100:
                estado_default = "Finalizado"
            else:
                estado_default = "Proceso"
            estado = st.selectbox("Estado", ["Proceso", "Finalizado"], 
                                 index=0 if estado_default == "Proceso" else 1,
                                 key="estado_select")
        
        evidencia = st.text_area("Evidencia (opcional)", placeholder="URL, documento, foto, etc...")
        
        submitted = st.form_submit_button("✅ Registrar avance")
        
        if submitted:
            if not seguimiento_texto:
                st.error("❌ Por favor ingresa una descripción del seguimiento")
            else:
                nuevo_seguimiento = pd.DataFrame([{
                    'id': id_accion,
                    'seguimiento': seguimiento_texto,
                    'fecha_seguimiento': pd.Timestamp(fecha_seguimiento).normalize(),
                    'porc_avance': porc_avance,
                    'estado': estado,
                    'evidencia': evidencia if evidencia else ""
                }])
                
                st.session_state.df_seguimiento = pd.concat([st.session_state.df_seguimiento, nuevo_seguimiento], ignore_index=True)
                save_data(st.session_state.df_plan, st.session_state.df_seguimiento)
                st.success(f"✅ Seguimiento registrado exitosamente! Avance: {porc_avance:.0f}%")
                st.rerun()
    
    # Mostrar historial de seguimientos
    st.subheader("📜 Historial de seguimientos")
    seguimientos_accion = st.session_state.df_seguimiento[st.session_state.df_seguimiento['id'] == id_accion]
    
    if not seguimientos_accion.empty:
        # Calcular estado actual
        estado_actual = actualizar_estado_accion(id_accion, st.session_state.df_seguimiento)
        avance_total = calcular_avance_total(id_accion, st.session_state.df_seguimiento)
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Estado actual", estado_actual)
        col2.metric("Avance total", f"{avance_total:.0f}%")
        col3.metric("N° de seguimientos", len(seguimientos_accion))
        
        # Mostrar tabla de seguimientos ordenada por fecha - Convertir a Timestamp para ordenar
        seguimientos_accion = seguimientos_accion.copy()
        seguimientos_accion['fecha_seguimiento'] = pd.to_datetime(seguimientos_accion['fecha_seguimiento'])
        seguimientos_ordenados = seguimientos_accion.sort_values('fecha_seguimiento', ascending=False)
        
        columnas_seguimiento = ['fecha_seguimiento', 'seguimiento', 'porc_avance', 'estado', 'evidencia']
        columnas_seguimiento = [col for col in columnas_seguimiento if col in seguimientos_ordenados.columns]
        
        df_seguimiento_display = seguimientos_ordenados[columnas_seguimiento].copy()
        if 'fecha_seguimiento' in df_seguimiento_display.columns:
            df_seguimiento_display['fecha_seguimiento'] = df_seguimiento_display['fecha_seguimiento'].dt.date
        if 'porc_avance' in df_seguimiento_display.columns:
            df_seguimiento_display['porc_avance'] = df_seguimiento_display['porc_avance'].round(0).astype('int64')
        
        st.dataframe(df_seguimiento_display, use_container_width=True)
        
        # Gráfico de progreso
        if len(seguimientos_ordenados) > 1:
            st.subheader("📈 Progreso de avance")
            progreso = seguimientos_ordenados.sort_values('fecha_seguimiento')[['fecha_seguimiento', 'porc_avance']].copy()
            progreso['fecha_seguimiento'] = pd.to_datetime(progreso['fecha_seguimiento'])
            st.line_chart(progreso.set_index('fecha_seguimiento'))
    else:
        st.info("No hay seguimientos registrados para esta acción aún")
    
    # Cierre de acción
    st.subheader("🔒 Cierre de acción")
    estado_actual = actualizar_estado_accion(id_accion, st.session_state.df_seguimiento)
    avance_total = calcular_avance_total(id_accion, st.session_state.df_seguimiento)
    
    if estado_actual != "Cerrado" and avance_total >= 100:
        if st.button("🔒 Marcar acción como Cerrada", type="primary", key="cerrar_accion"):
            ultimo_seguimiento = st.session_state.df_seguimiento[
                st.session_state.df_seguimiento['id'] == id_accion
            ].sort_values('fecha_seguimiento', ascending=False).iloc[0]
            
            if ultimo_seguimiento['estado'] != "Finalizado" or ultimo_seguimiento['porc_avance'] < 100:
                seguimiento_cierre = pd.DataFrame([{
                    'id': id_accion,
                    'seguimiento': "**CIERRE DE ACCIÓN** - Se completa el plan de acción",
                    'fecha_seguimiento': pd.Timestamp.now().normalize(),
                    'porc_avance': 100,
                    'estado': "Finalizado",
                    'evidencia': "Acción completada"
                }])
                st.session_state.df_seguimiento = pd.concat([st.session_state.df_seguimiento, seguimiento_cierre], ignore_index=True)
                save_data(st.session_state.df_plan, st.session_state.df_seguimiento)
                st.success("✅ Acción marcada como CERRADA exitosamente!")
                st.rerun()
            else:
                st.success("✅ La acción ya está cerrada!")
    elif estado_actual == "Cerrado":
        st.success("🎉 Esta acción ya está CERRADA")
    else:
        st.info(f"Para cerrar la acción, debe alcanzar 100% de avance. Avance actual: {avance_total:.0f}%")

# ==============================
# INTERFAZ PRINCIPAL (TABS)
# ==============================
st.title("🏭 Portal Kaizen - Gestión de Mejora Continua")

tab1, tab2, tab3 = st.tabs(["👤 Mis Acciones", "📝 Plan Kaizen (Admin)", "📊 Seguimiento"])

with tab1:
    pestania_mis_acciones()

with tab2:
    pestania_plan_kaizen()

with tab3:
    pestania_seguimiento()

# ==============================
# FOOTER
# ==============================
st.markdown("---")
st.caption("Portal Kaizen | Gestión de mejora continua | v2.0")