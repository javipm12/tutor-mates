"""
Aplicación Privada para Docentes de Matemáticas de Secundaria (3º ESO)
Orientada a la Gestión Administrativa, Planificación Curricular y Diseño DUA
Cumplimiento estricto del RGPD (Datos Anónimos)
"""

import os
import io
import re
import streamlit as st
import pandas as pd
from google import genai
from google.genai import types

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

import tempfile

try:
    import pypandoc
except ImportError:
    pypandoc = None

def create_docx_from_markdown(markdown_text: str, title: str = "Prueba de Evaluación / Actividad de Aula", curso: str = "3º ESO") -> bytes:
    """
    Convierte texto Markdown con fórmulas LaTeX a un documento Word (.docx) nativo
    usando Pandoc (pypandoc). Todas las fórmulas LaTeX se convierten en ecuaciones
    OMML nativas y editables de Microsoft Word.
    """
    if pypandoc is None:
        raise ImportError("La librería pypandoc no está disponible. Asegúrate de incluirla en requirements.txt y tener pandoc en packages.txt.")

    # 1. Cabecera formal de instituto en formato Markdown estructurado
    cabecera_instituto = f"""| **DEPARTAMENTO DE MATEMÁTICAS** | **Curso:** {curso} |
| :--- | :--- |
| **{title.upper()}** | **Fecha:** __________________ |
| **Nombre del alumno/a:** ____________________________________________________ | **Calificación:** _________ |

---

"""

    # 2. Normalización de fórmulas LaTeX para reconocimiento perfecto en Pandoc
    t = markdown_text

    # Convertir delimitadores alternativos \[ ... \] a $$ ... $$ y \( ... \) a $ ... $
    t = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', t, flags=re.DOTALL)
    t = re.sub(r'\\\((.*?)\\\)', r'$\1$', t)

    # Limpiar espacios pegados dentro de $ ... $ en líneas simples (sin romper los bloques $$...$$)
    # para que la extensión tex_math_dollars de Pandoc reconozca el 100% de las ecuaciones
    def _trim_inline_dollar(match):
        inner = match.group(1).strip()
        return f"${inner}$"

    t = re.sub(r'(?<!\$)\$([^\$\n]+?)\$(?!\$)', _trim_inline_dollar, t)

    full_markdown = cabecera_instituto + t

    # 3. Conversión a .docx mediante archivo temporal
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    temp_path = temp_file.name
    temp_file.close()

    try:
        pypandoc.convert_text(
            full_markdown,
            'docx',
            format='markdown+tex_math_dollars',
            outputfile=temp_path
        )
        with open(temp_path, "rb") as f:
            docx_bytes = f.read()
        return docx_bytes
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="TutorMates DUA | Gestión Docente Secundaria",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para una interfaz sobria, profesional y docente
st.markdown("""
<style>
    .reportview-container {
        background-color: #f8fafc;
    }
    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .badge-alert {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: 600;
    }
    .badge-ok {
        background-color: #dcfce7;
        color: #166534;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: 600;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 18px;
        border-radius: 6px 6px 0 0;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 2. CONTROL DE ACCESO PRIVADO DOCENTE (PROTECCIÓN RGPD Y SEGURIDAD)
# -----------------------------------------------------------------------------
def check_authentication():
    """Valida la contraseña docente contra st.secrets['TEACHER_PASSWORD']."""
    # Intentar obtener la contraseña configurada en secrets
    try:
        configured_password = st.secrets["TEACHER_PASSWORD"]
    except Exception:
        configured_password = os.environ.get("TEACHER_PASSWORD", None)

    if not configured_password:
        st.error("🔒 **Error de configuración de seguridad:** No se ha definido `TEACHER_PASSWORD` en los Secrets de Streamlit.")
        st.info("Configura la variable `TEACHER_PASSWORD` en los ajustes de Streamlit Cloud (Settings > Secrets) o en `.streamlit/secrets.toml`.")
        st.stop()

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown("### 🔐 Acceso Exclusivo para Profesorado")
            st.caption("Herramienta de uso interno docente. Acceso restringido para garantizar la confidencialidad y la normativa de protección de datos (RGPD).")
            
            with st.form("login_form"):
                password_input = st.text_input("Introduce la clave docente de acceso:", type="password", placeholder="••••••••")
                submit_button = st.form_submit_button("Entrar al Cuaderno Docente", use_container_width=True)

                if submit_button:
                    if password_input == configured_password:
                        st.session_state["authenticated"] = True
                        st.success("Acceso verificado correctamente.")
                        st.rerun()
                    else:
                        st.error("Contraseña incorrecta. Acceso denegado.")
            
            st.markdown("""
            ---
            ℹ️ *Si eres alumno/a, no tienes autorización para acceder a este portal docente.*
            """)
        st.stop()

check_authentication()


# -----------------------------------------------------------------------------
# 3. CLIENTE GEMINI Y SELECCIÓN DINÁMICA DE MODELOS (MODELS.LIST)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_gemini_client():
    """Inicializa el cliente oficial de Google GenAI usando st.secrets."""
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", None)

    if not api_key:
        st.error("🔑 No se encontró `GEMINI_API_KEY` en los Secrets de Streamlit.")
        st.info("Añade `GEMINI_API_KEY = 'tu_api_key'` en los Secrets de Streamlit Cloud.")
        st.stop()
    
    return genai.Client(api_key=api_key)

client = get_gemini_client()

@st.cache_data(ttl=1800)
def discover_gemini_models(_api_key: str) -> list[str]:
    """
    Consulta dinámicamente los modelos disponibles en la API del usuario mediante models.list()
    y selecciona automáticamente modelos Flash vigentes que soporten generateContent.
    Esto previene errores 404 por modelos descatalogados.
    """
    try:
        temp_client = genai.Client(api_key=_api_key)
        all_models = list(temp_client.models.list())
        
        flash_candidates = []
        other_candidates = []
        
        for m in all_models:
            name = getattr(m, "name", "") or ""
            clean_name = name.replace("models/", "")
            
            # Verificar si soporta generación de contenido
            actions = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None)
            supports_generate = True
            if actions:
                supports_generate = any("generateContent" in str(act) for act in actions)
                
            if supports_generate:
                if "flash" in clean_name.lower():
                    flash_candidates.append(clean_name)
                elif "gemini" in clean_name.lower() and "pro" in clean_name.lower():
                    other_candidates.append(clean_name)
                    
        # Orden de preferencia para modelos Flash vigentes
        preference_order = [
            "gemini-2.5-flash",
            "gemini-2.0-flash",
            "gemini-2.0-flash-001",
            "gemini-flash",
            "gemini-2.5-flash-lite",
            "gemini-2.0-flash-lite"
        ]
        
        sorted_flash = []
        for pref in preference_order:
            for cand in flash_candidates:
                if (cand == pref or cand.startswith(pref)) and cand not in sorted_flash:
                    sorted_flash.append(cand)
                    
        # Añadir cualquier otro modelo flash encontrado
        for cand in flash_candidates:
            if cand not in sorted_flash:
                sorted_flash.append(cand)
                
        if sorted_flash:
            return sorted_flash
        if other_candidates:
            return other_candidates
            
    except Exception:
        # Fallback de respaldo si models.list() no estuviese disponible
        pass
        
    return ["gemini-2.5-flash", "gemini-2.0-flash"]

def call_gemini(prompt: str, system_instruction: str = None) -> str:
    """Ejecuta una petición usando el modelo Flash óptimo detectado en la API."""
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", "")
        
    candidate_models = discover_gemini_models(api_key)
    last_error = None
    
    for model_name in candidate_models:
        try:
            config = types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.7,
            ) if system_instruction else types.GenerateContentConfig(temperature=0.7)
            
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=config
            )
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = e
            continue
            
    return f"⚠️ Error al conectar con el servicio de IA: {str(last_error)}"


# -----------------------------------------------------------------------------
# 4. CARGA Y GESTIÓN DEL DATASET ANONIMIZADO
# -----------------------------------------------------------------------------
CSV_PATH = "clase_3eso.csv"

def init_student_data():
    """Carga o inicializa los datos de clase con códigos anónimos."""
    if "df_alumnos" not in st.session_state:
        if os.path.exists(CSV_PATH):
            df = pd.read_csv(CSV_PATH)
            # Asegurar tipos adecuados
            df["Asistencia_Hoy"] = df["Asistencia_Hoy"].astype(bool)
            df["Faltas_Acumuladas"] = df["Faltas_Acumuladas"].astype(int)
            df["Nota_Media"] = df["Nota_Media"].astype(float)
            st.session_state["df_alumnos"] = df
        else:
            # Fallback seguro en memoria si no estuviera el archivo
            data = {
                "ID_Alumno": [f"ALU_{i:02d}" for i in range(1, 16)],
                "Perfil_DUA": [
                    "Estándar", "Refuerzo_Calculo", "Altas_Capacidades", "TDAH_Adaptacion", 
                    "Discalculia_Leve", "Estándar", "Refuerzo_Calculo", "Estándar", 
                    "TDAH_Adaptacion", "Altas_Capacidades", "Estándar", "Discalculia_Leve", 
                    "Refuerzo_Calculo", "Estándar", "Estándar"
                ],
                "Faltas_Acumuladas": [1, 5, 0, 2, 3, 2, 6, 0, 1, 1, 3, 2, 4, 1, 5],
                "Nota_Media": [6.5, 3.8, 9.6, 5.2, 4.3, 7.1, 3.2, 8.0, 6.0, 9.1, 5.8, 4.7, 4.0, 7.4, 4.5],
                "Asistencia_Hoy": [True, False, True, True, True, True, False, True, True, True, True, True, True, True, False],
                "Observaciones_Docente": [
                    "Evolución positiva en álgebra y cálculo elemental.",
                    "Dificultad acusada con fracciones y jerarquía. Absentismo reiterado viernes.",
                    "Resuelve tareas rápidamente; requiere retos de profundización.",
                    "Requiere enunciados fragmentados y temporizador de foco.",
                    "Precisa apoyo visual y tablas de referencia para la operativa básica.",
                    "Buen nivel de abstracción; trabajo autónomo y constante.",
                    "Bajo rendimiento ligado a faltas continuadas; pendiente recuperación.",
                    "Metódico y ordenado en la exposición y justificación de procedimientos.",
                    "Mejora notablemente al estructurar la resolución en pasos numerados.",
                    "Gran capacidad de deducción lógica; excelente tutor entre iguales.",
                    "Comprensión adecuada; afianzar despejes en ecuaciones.",
                    "Confusión habitual con signos y posiciones numéricas; responde con guía gráfica.",
                    "En umbral de aprobado; buena predisposición en cooperativo guiado.",
                    "Progreso sólido en geometría analítica y álgebra.",
                    "Faltas no justificadas a primera hora; riesgo de descolgarse del ritmo."
                ]
            }
            st.session_state["df_alumnos"] = pd.DataFrame(data)

init_student_data()


# -----------------------------------------------------------------------------
# 5. BARRA LATERAL (SIDEBAR) DOCENTE
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<div style='font-size: 3.4rem; line-height: 1.1; margin-bottom: 0.2rem;'>📐</div>", unsafe_allow_html=True)
    st.title("TutorMates 3º ESO")
    st.markdown("**Cuaderno Digital de Matemáticas**")
    st.caption("Adaptado al marco DUA y normativa RGPD")
    
    # Modelo IA detectado dinámicamente
    try:
        _api_key_sidebar = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        _api_key_sidebar = os.environ.get("GEMINI_API_KEY", "")
        
    _modelos_detectados = discover_gemini_models(_api_key_sidebar)
    _modelo_mostrado = _modelos_detectados[0] if _modelos_detectados else "gemini-2.5-flash"
    st.caption(f"⚡ **Modelo IA activo:** `{_modelo_mostrado}`")
    
    st.divider()
    
    st.markdown("### 🛡️ Cumplimiento RGPD")
    st.success("✅ **Anonimización activa:** Los datos enviados a la IA utilizan únicamente códigos (`ALU_XX`). Ningún dato de carácter personal identificable sale del navegador.")
    
    st.divider()
    
    st.markdown("### 📊 Estado de la Sesión")
    df_current = st.session_state["df_alumnos"]
    total_alumnos = len(df_current)
    presentes = df_current["Asistencia_Hoy"].sum()
    ausentes = total_alumnos - presentes
    alertas_faltas = len(df_current[df_current["Faltas_Acumuladas"] > 4])
    
    st.metric("Total Alumnos", total_alumnos)
    st.metric("Presentes Hoy", f"{presentes} / {total_alumnos}")
    st.metric("Alumnos con Alerta (>4 faltas)", alertas_faltas, delta_color="inverse")
    
    st.divider()
    if st.button("🚪 Cerrar Sesión Docente", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()


# -----------------------------------------------------------------------------
# 6. ENCABEZADO PRINCIPAL
# -----------------------------------------------------------------------------
st.title("📐 Panel de Gestión Docente y Diversidad DUA")
st.markdown("Herramienta docente privada para automatizar la gestión de aula, diseño de pruebas multinivel y seguimiento pedagógico en **3º de Educación Secundaria Obligatoria**.")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Pasar Lista y Burocracia",
    "📅 Planificación de Aula",
    "🎯 Diseñador de Ejercicios y Exámenes DUA",
    "🤖 Asistente Docente con Contexto del Aula",
    "📜 Normativa y Plan de Convivencia"
])


# =============================================================================
# PESTAÑA 1: 📋 PASAR LISTA Y BUROCRACIA
# =============================================================================
with tab1:
    st.header("📋 Registro de Asistencia y Gestión de Absentismo")
    st.markdown("Control diario de asistencia y detección temprana de absentismo escolar conforme al protocolo educativo oficial.")
    
    # Alertas automáticas para faltas > 4
    df_alertas = st.session_state["df_alumnos"][st.session_state["df_alumnos"]["Faltas_Acumuladas"] > 4]
    
    if not df_alertas.empty:
        codigos_alerta = ", ".join([f"**{row['ID_Alumno']}** ({row['Faltas_Acumuladas']} faltas)" for _, row in df_alertas.iterrows()])
        st.error(f"⚠️ **Alerta Protocolo Absentismo (>4 faltas acumuladas):** Se detectan {len(df_alertas)} alumno(s) en situación crítica de absentismo no justificado: {codigos_alerta}. Requiere comunicación inmediata a Jefatura de Estudios / Comisión de Absentismo.")
    else:
        st.success("✅ No hay ningún alumno que supere el umbral de alerta de absentismo (>4 faltas).")
    
    # Mini-Dashboard Estadístico del Aula
    df_actual = st.session_state["df_alumnos"]
    total_alumnos_dash = len(df_actual)
    nota_media_grupo = float(df_actual["Nota_Media"].mean()) if total_alumnos_dash > 0 else 0.0
    alumnos_riesgo = df_actual[(df_actual["Nota_Media"] < 5.0) | (df_actual["Faltas_Acumuladas"] > 4)]
    num_riesgo = len(alumnos_riesgo)
    pct_riesgo = (num_riesgo / total_alumnos_dash * 100) if total_alumnos_dash > 0 else 0.0
    total_faltas_acum = int(df_actual["Faltas_Acumuladas"].sum())

    st.markdown("### 📊 Mini-Dashboard Estadístico del Aula")
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric(
            label="Nota Media del Grupo",
            value=f"{nota_media_grupo:.2f} / 10",
            delta="Aprobado General" if nota_media_grupo >= 5.0 else "Alerta Rendimiento",
            delta_color="normal" if nota_media_grupo >= 5.0 else "inverse"
        )
    with col_m2:
        st.metric(
            label="% Alumnos en Riesgo (Nota < 5 o > 4 faltas)",
            value=f"{pct_riesgo:.1f}%",
            delta=f"{num_riesgo} de {total_alumnos_dash} alumnos",
            delta_color="inverse" if num_riesgo > 0 else "normal"
        )
    with col_m3:
        st.metric(
            label="Total Faltas Acumuladas",
            value=f"{total_faltas_acum} faltas",
            delta=f"Media: {total_faltas_acum/total_alumnos_dash:.1f} por alumno" if total_alumnos_dash > 0 else None,
            delta_color="off"
        )

    # Gráficos visuales nativos de Streamlit: correlación absentismo y rendimiento
    col_ch1, col_ch2 = st.columns([1, 1])
    with col_ch1:
        st.markdown("**📉 Correlación: Faltas Acumuladas vs. Nota Media**")
        st.caption("Visualiza el impacto directo de las ausencias en las calificaciones por Perfil DUA.")
        st.scatter_chart(
            df_actual,
            x="Faltas_Acumuladas",
            y="Nota_Media",
            color="Perfil_DUA",
            height=260
        )
    with col_ch2:
        st.markdown("**📊 Faltas Acumuladas por Código ALU_XX**")
        st.caption("Comparativa individual frente al umbral crítico de absentismo (>4 faltas).")
        df_faltas_chart = df_actual.set_index("ID_Alumno")[["Faltas_Acumuladas"]]
        st.bar_chart(
            df_faltas_chart,
            color="#ef4444",
            height=260
        )

    st.divider()
    st.subheader("Tabla Interactiva de Clase (3º ESO)")
    st.caption("Modifica la asistencia de hoy o las observaciones directamente en la tabla. Los cambios se reflejarán inmediatamente.")
    
    # Configuración de columnas del data editor
    edited_df = st.data_editor(
        st.session_state["df_alumnos"],
        column_config={
            "ID_Alumno": st.column_config.TextColumn("Código Alumno", disabled=True),
            "Perfil_DUA": st.column_config.SelectboxColumn(
                "Perfil DUA",
                options=["Estándar", "Refuerzo_Calculo", "TDAH_Adaptacion", "Altas_Capacidades", "Discalculia_Leve"],
                required=True
            ),
            "Faltas_Acumuladas": st.column_config.NumberColumn("Faltas Totales", min_value=0, max_value=50, step=1),
            "Nota_Media": st.column_config.NumberColumn("Nota Media", min_value=0.0, max_value=10.0, step=0.1, format="%.1f"),
            "Asistencia_Hoy": st.column_config.CheckboxColumn("¿Asiste Hoy?", default=True),
            "Observaciones_Docente": st.column_config.TextColumn("Observaciones Docente (Sin Nombres)", width="large")
        },
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="editor_alumnos"
    )
    
    # Actualizar session state si hubo cambios
    st.session_state["df_alumnos"] = edited_df

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        # Descarga del CSV actualizado
        csv_data = edited_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="💾 Descargar Registro Actualizado (CSV)",
            data=csv_data,
            file_name="clase_3eso_actualizada.csv",
            mime="text/csv",
            use_container_width=True
        )
    with col_btn2:
        if st.button("🔄 Marcar a todos como Presentes hoy", use_container_width=True):
            st.session_state["df_alumnos"]["Asistencia_Hoy"] = True
            st.rerun()

    st.markdown("---")
    st.subheader("📄 Generador Automático de Partes y Notificaciones Oficiales con IA")
    st.markdown("Redacta borradores administrativos oficiales listos para enviar a Jefatura de Estudios o a las Familias, garantizando el anonimato del alumno mediante su código identificador.")
    
    col_dest, col_alu = st.columns([1, 1])
    with col_dest:
        destinatario = st.selectbox(
            "Destinatario del documento oficial:",
            ["Jefatura de Estudios / Comisión de Absentismo", "Familia / Tutores Legales (Tutoría Oficial)"]
        )
    with col_alu:
        # Seleccionar preferentemente alumnos con alerta, o cualquiera de la lista
        alumnos_options = edited_df["ID_Alumno"].tolist()
        alumno_sel = st.selectbox("Selecciona el alumno por su código anónimo:", alumnos_options, index=1 if len(alumnos_options) > 1 else 0)

    datos_alumno = edited_df[edited_df["ID_Alumno"] == alumno_sel].iloc[0]

    if st.button("✨ Redactar Borrador con IA", type="primary"):
        with st.spinner("Redactando comunicación formal bajo normativa educativa..."):
            prompt_burocracia = f"""
            Actúa como un profesor de Matemáticas y tutor de 3º de ESO en un Instituto de Educación Secundaria en España.
            Redacta un documento formal de tipo: "{destinatario}".
            
            DATOS OFICIALES DEL ALUMNO (ESTRICTAMENTE ANONIMIZADOS BAJO RGPD):
            - Código del Alumno: {datos_alumno['ID_Alumno']}
            - Faltas de asistencia acumuladas: {datos_alumno['Faltas_Acumuladas']} faltas
            - Asistencia en la sesión de hoy: {'Asistió' if datos_alumno['Asistencia_Hoy'] else 'Ausente'}
            - Calificación media actual: {datos_alumno['Nota_Media']}/10
            - Perfil de atención a la diversidad (DUA): {datos_alumno['Perfil_DUA']}
            - Observaciones del cuaderno docente: {datos_alumno['Observaciones_Docente']}
            
            REGLAS CRÍTICAS DE REDACCIÓN:
            1. En cumplimiento estricto del RGPD, utiliza ÚNICAMENTE el identificador '{datos_alumno['ID_Alumno']}'. NO inventes nombres ni apellidos reales.
            2. Tono riguroso, respetuoso, formal y pedagógico.
            3. Menciona la preocupación por el impacto de las faltas reiteradas en la adquisición de las competencias matemáticas de 3º ESO.
            4. Si el destinatario es Jefatura de Estudios, referencia el inicio del protocolo ordinario de absentismo escolar y solicitud de intervención del departamento de orientación.
            5. Si el destinatario es la Familia, cita con delicadeza la importancia del acompañamiento familiar, la justificación documental de las ausencias y convoca a una entrevista presencial o telemática de tutoría.
            6. Estructura con fecha simulada, asunto oficial, exposición de hechos objetivos y fórmula de despedida.
            """
            
            system_instruction = "Eres un asistente docente experto en legislación educativa secundaria (LOMLOE), protocolos de absentismo escolar y redacción administrativa en España, cumpliendo estrictamente con el RGPD."
            respuesta_borrador = call_gemini(prompt_burocracia, system_instruction)
            
            st.markdown("### 📋 Borrador Generado")
            st.text_area("Copia el texto del documento:", value=respuesta_borrador, height=350)

    st.markdown("---")
    st.subheader("👨‍👩‍👧‍👦 Ficha Preparatoria para Entrevistas con Familias (Tutoría Oficial)")
    st.markdown("Genera un informe sintético y conciso de todas las características del alumno/a, junto con una guía práctica de explicaciones para los padres y propuesta de compromisos mutuos.")

    col_fam1, col_fam2 = st.columns([1, 1])
    with col_fam1:
        alumnos_lista = edited_df["ID_Alumno"].tolist()
        def format_alumno_label(cod):
            row = edited_df[edited_df["ID_Alumno"] == cod].iloc[0]
            alerta = "⚠️ Alerta" if row["Faltas_Acumuladas"] > 4 or row["Nota_Media"] < 5.0 else "✅ Normal"
            return f"{cod} — Nota: {row['Nota_Media']:.1f} | {row['Faltas_Acumuladas']} faltas | {row['Perfil_DUA']} ({alerta})"

        alumno_fam_sel = st.selectbox(
            "Selecciona el alumno/a para la entrevista de tutoría:",
            alumnos_lista,
            format_func=format_alumno_label,
            key="sel_alumno_entrevista"
        )
        datos_fam = edited_df[edited_df["ID_Alumno"] == alumno_fam_sel].iloc[0]

    with col_fam2:
        motivo_entrevista = st.selectbox(
            "Motivo principal o enfoque prioritario de la reunión:",
            [
                "Seguimiento trimestral ordinario (Evolución global y hábitos)",
                "Preocupación por bajo rendimiento o riesgo de evaluación negativa",
                "Absentismo escolar reiterado o acumulación de faltas",
                "Dificultades de atención, organización o desmotivación en clase",
                "Atención a la diversidad / Adaptación DUA y pautas para el hogar",
                "Orientación académica y altas capacidades / motivación extra"
            ],
            key="sel_motivo_entrevista"
        )

    apuntes_tutor = st.text_input(
        "Apuntes o hechos recientes observados por el docente (opcional):",
        placeholder="Ej: Buena actitud cuando trabaja en equipo, pero le cuesta entregar tareas a tiempo y despista...",
        key="input_apuntes_tutor"
    )

    # Métricas clave del alumno seleccionado
    col_m_alu1, col_m_alu2, col_m_alu3, col_m_alu4 = st.columns(4)
    with col_m_alu1:
        st.metric("Código Alumno", datos_fam["ID_Alumno"])
    with col_m_alu2:
        st.metric(
            "Nota Media",
            f"{datos_fam['Nota_Media']:.1f} / 10",
            delta="Aprobado" if datos_fam["Nota_Media"] >= 5.0 else "Suspenso",
            delta_color="normal" if datos_fam["Nota_Media"] >= 5.0 else "inverse"
        )
    with col_m_alu3:
        st.metric(
            "Faltas Acumuladas",
            f"{datos_fam['Faltas_Acumuladas']} faltas",
            delta="Protocolo Absentismo" if datos_fam["Faltas_Acumuladas"] > 4 else "Bajo control",
            delta_color="inverse" if datos_fam["Faltas_Acumuladas"] > 4 else "normal"
        )
    with col_m_alu4:
        st.metric("Perfil DUA", datos_fam["Perfil_DUA"])

    if st.button("📋 Generar Ficha Completa para la Entrevista con Familias", type="primary", key="btn_generar_ficha_fam"):
        with st.spinner("Sintetizando perfil del alumno y preparando argumentos y explicaciones para la familia..."):
            prompt_ficha = f"""
            Actúa como un profesor tutor experimentado de Matemáticas de 3º de ESO en España, especialista en orientación educativa y atención a la diversidad (DUA).
            Genera una FICHA PREPARATORIA COMPLETA Y ESTRUCTURADA PARA UNA ENTREVISTA DE TUTORÍA CON LA FAMILIA DEL ALUMNO.
            
            DATOS OFICIALES DEL ALUMNO (RGPD ESTRICTO - IDENTIFICADO SOLO COMO {datos_fam['ID_Alumno']}):
            - Código identificador: {datos_fam['ID_Alumno']}
            - Calificación media actual: {datos_fam['Nota_Media']} / 10
            - Faltas de asistencia acumuladas: {datos_fam['Faltas_Acumuladas']} faltas
            - Perfil de atención a la diversidad (DUA): {datos_fam['Perfil_DUA']}
            - Observaciones del cuaderno docente: {datos_fam['Observaciones_Docente']}
            - Motivo principal de la reunión: {motivo_entrevista}
            - Apuntes específicos aportados por el tutor: {apuntes_tutor if apuntes_tutor.strip() else 'Sin apuntes adicionales'}

            DIRECTRICES CRÍTICAS:
            1. PRIVACIDAD RGPD: Usa EXCLUSIVAMENTE el código '{datos_fam['ID_Alumno']}'. NO inventes nombres ni apellidos reales.
            2. TONO Y ESTILO: Empático, constructivo, institucional y colaborativo. No uses un tono acusatorio; la familia debe percibir que el centro busca soluciones conjuntas.
            3. TRADUCCIÓN PEDAGÓGICA CLARA: Convierte los tecnicismos (DUA, descriptores, competencias) en explicaciones sencillas y comprensibles para padres no docentes.

            ESTRUCTURA OBLIGATORIA DE LA RESPUESTA EN MARKDOWN:

            # 📋 FICHA PREPARATORIA DE ENTREVISTA CON FAMILIA — {datos_fam['ID_Alumno']}

            ## 1. 📊 INFORME CONCISO DE CARACTERÍSTICAS DEL ALUMNO (Radiografía para el Docente)
            - **Diagnóstico General:** Estado global del alumno en 3 líneas concisas (favorable, atención preventiva o intervención urgente).
            - **Rendimiento Académico en Matemáticas:** Valoración cualitativa de su nota ({datos_fam['Nota_Media']}/10), señalando qué contenidos o competencias domina y en qué aspectos presenta mayores dificultades.
            - **Asistencia y Puntualidad:** Análisis del impacto de sus {datos_fam['Faltas_Acumuladas']} faltas en el seguimiento de la materia.
            - **Perfil DUA y Estilo de Aprendizaje:** Qué implicaciones prácticas tiene su perfil ({datos_fam['Perfil_DUA']}) y qué medidas metodológicas se aplican en el aula.
            - **Observaciones de Aula:** Síntesis de las notas del cuaderno del profesor.

            ## 2. 🗣️ GUÍA DE EXPLICACIONES PARA LOS PADRES (Traducción Pedagógica Cercana)
            - **Apertura de la Reunión (Puntos Fuertes):** 2 o 3 fortalezas, cualidades o actitudes positivas del alumno/a que el tutor debe destacar al inicio para abrir la conversación en un clima de confianza.
            - **Cómo explicar su nivel en Matemáticas:** Argumentos claros, serenos y sin tecnicismos para explicar a los padres cómo va su hijo/a, evitando alarmismos pero con total honestidad.
            - **Cómo explicar su perfil DUA a la familia:** Pautas sencillas y desestigmatizantes para explicar cómo procesa la información y qué apoyos recibe en clase.
            - **Actitud, convivencia y hábitos de trabajo:** Cómo transmitir observaciones sobre su constancia, autonomía, cuidado del material y relación con sus compañeros.

            ## 3. 🤝 COMPROMISOS COMPARTIDOS (Pacto de Colaboración Familia - Centro)
            - **Compromisos Concretos para la Familia en Casa:** 3 pautas realistas, medibles y viables para la familia (ej. horario fijo de 30 min sin pantallas, supervisión semanal de libreta/classroom, rutina de descanso).
            - **Compromisos del Profesor/Tutor en el Aula:** 2 medidas concretas de seguimiento pedagógico que el docente se compromete a realizar (ej. supervisión de notas guiadas, apoyo entre iguales, comprobación semanal).

            ## 4. ❓ PREGUNTAS ESTRATÉGICAS PARA INDAGAR CON LA FAMILIA
            - 3 preguntas clave, abiertas y respetuosas para conocer mejor la dinámica del alumno en casa (hábitos de estudio, nivel de frustración ante las tareas, descanso nocturno y motivación).
            """

            system_instruction = "Eres un tutor y orientador de Educación Secundaria en España, experto en comunicación asertiva con familias, marco DUA y legislación educativa LOMLOE."
            ficha_generada = call_gemini(prompt_ficha, system_instruction)
            st.session_state["ultima_ficha_familia"] = ficha_generada
            st.session_state["ultimo_alumno_ficha"] = datos_fam["ID_Alumno"]

    if st.session_state.get("ultima_ficha_familia"):
        st.markdown("---")
        st.markdown("### 📄 Ficha de Entrevista Generada")
        st.markdown(st.session_state["ultima_ficha_familia"])

        col_exp1, col_exp2 = st.columns([1, 1])
        with col_exp1:
            try:
                docx_ficha = create_docx_from_markdown(
                    st.session_state["ultima_ficha_familia"],
                    title=f"Ficha Entrevista Familia - {st.session_state.get('ultimo_alumno_ficha', 'Alumno')}",
                    curso="3º ESO"
                )
                st.download_button(
                    label="⬇️ Descargar Ficha en Word (.docx)",
                    data=docx_ficha,
                    file_name=f"ficha_entrevista_{st.session_state.get('ultimo_alumno_ficha', 'alu').lower()}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="btn_descarga_docx_ficha"
                )
            except Exception as e:
                st.caption(f"ℹ️ Exportación a Word no disponible: {str(e)}")
        with col_exp2:
            st.download_button(
                label="⬇️ Descargar Ficha en Texto (.txt)",
                data=st.session_state["ultima_ficha_familia"],
                file_name=f"ficha_entrevista_{st.session_state.get('ultimo_alumno_ficha', 'alu').lower()}.txt",
                mime="text/plain",
                use_container_width=True,
                key="btn_descarga_txt_ficha"
            )


# =============================================================================
# PESTAÑA 2: 📅 PLANIFICACIÓN CURRICULAR Y GESTIÓN DE AULA
# =============================================================================
with tab2:
    st.header("📅 Planificación Curricular y Gestión de Aula")
    st.markdown("Herramientas para la secuenciación de unidades didácticas y resolución inmediata de sustituciones o guardias de aula.")
    
    subtab_temario, subtab_guardia, subtab_reto = st.tabs([
        "📅 Distribución de Horas y Temario",
        "🚨 Modo Guardia Express",
        "🏆 Reto Matemático Semanal (Gamificación)"
    ])
    
    # -------------------------------------------------------------------------
    # SUB-PESTAÑA 1: DISTRIBUCIÓN DE HORAS Y TEMARIO
    # -------------------------------------------------------------------------
    with subtab_temario:
        st.subheader("Secuenciación Didáctica y Temporalización Realista")
        st.caption("Diseña una programación sesión por sesión para 3º de ESO, considerando festivos, repasos cooperativos y sesiones de evaluación formativa.")
        
        with st.form("form_temporalizacion"):
            col_t1, col_t2 = st.columns([2, 1])
            with col_t1:
                tema_matematicas = st.text_input(
                    "Unidad Didáctica o Tema:",
                    value="Ecuaciones de 1er y 2º Grado y Sistemas de Ecuaciones Lineales"
                )
            with col_t2:
                nivel_grupo = st.selectbox(
                    "Nivel general del grupo:",
                    [
                        "Heterogéneo con necesidad de refuerzo",
                        "Estándar curricular 3º ESO",
                        "Avanzado con buen ritmo de trabajo",
                        "Grupo con alta dispersión y desmotivación inicial"
                    ]
                )
                
            col_h1, col_h2, col_h3 = st.columns(3)
            with col_h1:
                horas_totales = st.number_input("Sesiones totales estimadas:", min_value=4, max_value=40, value=14, step=1)
            with col_h2:
                dias_perdidos = st.number_input("Festivos / Huelgas / Salidas:", min_value=0, max_value=10, value=2, step=1)
            with col_h3:
                frecuencia = st.selectbox("Frecuencia lectiva:", ["3 sesiones semanales", "4 sesiones semanales"])
                
            sesiones_efectivas = max(1, horas_totales - dias_perdidos)
            st.info(f"💡 **Sesiones lectivas netas de aula disponibles:** {sesiones_efectivas} sesiones (calculado restando {dias_perdidos} días no lectivos de {horas_totales} totales).")
            
            btn_planificar = st.form_submit_button("🗓️ Generar Secuenciación Didáctica Completa", type="primary")

        if btn_planificar:
            with st.spinner("Diseñando la programación sesión por sesión..."):
                prompt_plan = f"""
                Diseña una temporalización didáctica realista y detallada para el departamento de Matemáticas de 3º de ESO.
                
                PARÁMETROS DE LA PROGRAMACIÓN:
                - Unidad Temática: {tema_matematicas}
                - Sesiones reales lectivas: {sesiones_efectivas} sesiones (ya se han descontado {dias_perdidos} festivos/días perdidos).
                - Nivel del grupo: {nivel_grupo}
                - Frecuencia semanal: {frecuencia}
                
                ESTRUCTURA OBLIGATORIA QUE DEBE TENER LA RESPUESTA:
                1. **Tabla Resumen de la Unidad**: Objetivos competenciales clave y criterios de evaluación según currículo oficial de 3º ESO.
                2. **Desglose Sesión por Sesión (Sesión 1 a Sesión {sesiones_efectivas})**:
                   Para cada sesión, incluye obligatoriamente:
                   - *Título y Contenido específico*.
                   - *Estructura de la sesión (50 min)*: 10 min activación/repaso inicial, 30 min práctica guiada/autónoma, 10 min comprobación/cierre (ticket de salida).
                   - *Medida de Atención a la Diversidad (DUA)*: Ajuste específico para alumnos de refuerzo o TDAH.
                3. **Sesiones Especiales Obligatorias**:
                   - Debe incluir al menos 1 o 2 sesiones finales dedicadas expresamente a resolución de dudas y simulacro de examen en parejas de tutoría entre iguales.
                   - Debe incluir la sesión de Evaluación Sumativa (Prueba escrita/competencial).
                   - Debe incluir una sesión posterior de Retroalimentación, coevaluación y corrección activa de errores.
                4. **Estrategia ante imprevistos**: Recomendación de qué actividades compactar si se pierde una sesión adicional.
                """
                
                system_instruction = "Eres un jefe de departamento de Matemáticas de Educación Secundaria en España, experto en diseño curricular LOMLOE, metodología DUA y planificación docente efectiva."
                plan_generado = call_gemini(prompt_plan, system_instruction)
                st.session_state["ultimo_plan_temario"] = plan_generado

        if st.session_state.get("ultimo_plan_temario"):
            st.markdown("---")
            st.markdown("### 📅 Propuesta de Temporalización Curricular")
            st.markdown(st.session_state["ultimo_plan_temario"])

    # -------------------------------------------------------------------------
    # SUB-PESTAÑA 2: 🚨 MODO GUARDIA EXPRESS
    # -------------------------------------------------------------------------
    with subtab_guardia:
        st.subheader("🚨 Generador de Fichas de Guardia de Emergencia")
        st.markdown("Genera una ficha autocontenida para **50 minutos de guardia** con ejercicios graduados y una **hoja de soluciones detalladas paso a paso**, pensada para que cualquier profesor de guardia (de Historia, Inglés, Educación Física, etc.) pueda atender y corregir la clase con solvencia.")

        with st.form("form_guardia"):
            col_g1, col_g2 = st.columns([1, 2])
            with col_g1:
                curso_guardia = st.selectbox("Curso / Nivel:", ["3º ESO", "1º ESO", "2º ESO", "4º ESO"], index=0)
                duracion_guardia = st.selectbox("Duración de la guardia:", ["50 minutos (Sesión estándar)", "45 minutos", "55 minutos"])
            with col_g2:
                tema_guardia = st.text_input("Tema actual que tocaba dar hoy:", value="Ecuaciones de 1er y 2º grado y problemas numéricos")
                nivel_guardia = st.selectbox("Nivel general del grupo:", ["Heterogéneo / Estándar", "Refuerzo prioritario (necesita guía paso a paso)", "Avanzado / Ritmo alto"])
                
            btn_guardia = st.form_submit_button("⚡ Generar Ficha de Guardia de Emergencia", type="primary")

        if btn_guardia:
            with st.spinner("Diseñando ficha de guardia autocontenida con guía para profesor no especialista..."):
                prompt_guardia = f"""
                Actúa como un profesor de Matemáticas de Secundaria experto en didáctica escolar.
                Diseña una FICHA DE GUARDIA DE EMERGENCIA completa, 100% autocontenida y estructurada para una clase de {duracion_guardia} en {curso_guardia}.
                
                PARÁMETROS:
                - Curso: {curso_guardia}
                - Tema previsto para hoy: {tema_guardia}
                - Duración: {duracion_guardia}
                - Nivel del grupo: {nivel_guardia}
                
                OBJETIVO PRINCIPAL:
                Esta ficha debe permitir que un profesor/a de guardia de CUALQUIER OTRA ESPECIALIDAD (Historia, Educación Física, Lengua, etc.) que no sea de Matemáticas pueda dinamizar la clase, mantener el orden de trabajo y escribir o comprobar las soluciones en la pizarra con total seguridad y sin dudas.
                
                ESTRUCTURA OBLIGATORIA (MARKDOWN LIMPIO):
                # FICHA DE GUARDIA DE MATEMÁTICAS ({curso_guardia})
                ## Tema: {tema_guardia}
                
                ### 📌 1. Instrucciones para el Profesor/a de Guardia (5 minutos)
                - Pautas claras de entrada: organización del aula, ritmo de trabajo individual o en parejas y qué hacer si surgen dudas.
                
                ### 🚀 2. Actividad de Activación / Calentamiento (5-10 minutos)
                - 2 preguntas de cálculo mental o repaso conceptual rápido para centrar la atención en la pizarra.
                
                ### 📝 3. Ficha de Trabajo para el Alumnado (4 Ejercicios Graduados) (30 minutos)
                - **Ejercicio 1 (Nivel Básico / Algorítmico)**: Ejercicio guiado con coeficientes sencillos.
                - **Ejercicio 2 (Nivel Medio / Comprensión)**: Aplicación directa del procedimiento.
                - **Ejercicio 3 (Nivel Aplicación / Problema Cotidiano)**: Situación práctica sencilla y motivadora.
                - **Ejercicio 4 (Reto / Ampliación)**: Para quienes terminen rápido, evitando distracciones.
                *(Escribe todas las fórmulas matemáticas en formato LaTeX limpio: $...$ y $$...$$)*.
                
                ---
                
                ### 🔑 4. Hoja de Soluciones Razonadas para la Pizarra (Uso Docente) (10 minutos)
                - Solución paso a paso de cada uno de los 4 ejercicios.
                - Explicación en lenguaje sencillo indicando exactamente qué escribir en la pizarra y qué errores comunes corregir.
                """
                system_instruction_guardia = "Eres un formador de docentes de Matemáticas en España. Generas materiales de aula estructurados, claros y comprensibles para cualquier profesor de guardia."
                ficha_generada = call_gemini(prompt_guardia, system_instruction_guardia)
                st.session_state["ultima_ficha_guardia"] = ficha_generada
                st.session_state["ultimo_tema_guardia"] = tema_guardia
                st.session_state["ultimo_curso_guardia"] = curso_guardia

        if st.session_state.get("ultima_ficha_guardia"):
            st.markdown("---")
            st.markdown("### 📄 Ficha de Guardia Generada")
            st.markdown(st.session_state["ultima_ficha_guardia"])
            
            # Botón de Descarga en Word (.docx)
            try:
                docx_guardia_file = create_docx_from_markdown(
                    st.session_state["ultima_ficha_guardia"],
                    title=f"Ficha de Guardia - {st.session_state.get('ultimo_tema_guardia', 'Matemáticas')}",
                    curso=st.session_state.get("ultimo_curso_guardia", "3º ESO")
                )
                st.download_button(
                    label="⬇️ Descargar Ficha de Guardia en Word (.docx)",
                    data=docx_guardia_file,
                    file_name=f"ficha_guardia_{st.session_state.get('ultimo_tema_guardia', 'mates').lower().replace(' ', '_')[:25]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="btn_descarga_docx_guardia"
                )
            except Exception as e:
                st.error(f"Error al exportar ficha a Word: {str(e)}")

    # -------------------------------------------------------------------------
    # SUB-PESTAÑA 3: 🏆 RETO MATEMÁTICO SEMANAL (GAMIFICACIÓN)
    # -------------------------------------------------------------------------
    with subtab_reto:
        st.subheader("🏆 Reto Matemático Semanal / Escape Room de Aula")
        st.caption("Fomenta la motivación, el pensamiento lateral y el estímulo para Altas Capacidades mediante desafíos lúdicos con pistas escalonadas y narrativa envolvente.")

        with st.form("form_reto_semanal"):
            col_r1, col_r2 = st.columns([2, 1])
            with col_r1:
                tema_reto = st.text_input(
                    "Contenido o bloque matemático:",
                    value="Ecuaciones, acertijos algebraicos y lógica proporcional",
                    key="input_tema_reto"
                )
            with col_r2:
                narrativa_reto = st.selectbox(
                    "Ambientación / Formato del reto:",
                    [
                        "🕵️ Misterio detectivesco (Descifrar código o clave secreta)",
                        "🔐 Escape Room de aula (3 candados o cerraduras matemáticas)",
                        "🚀 Misión espacial / Misión científica contrarreloj",
                        "🏛️ Acertijo histórico / Paradoja matemática clásica",
                        "🍕 Dilema cotidiano / Reto de economía doméstica y lógica"
                    ],
                    key="sel_narrativa_reto"
                )

            col_rn1, col_rn2 = st.columns([1, 1])
            with col_rn1:
                dificultad_reto = st.selectbox(
                    "Nivel de desafío pedagógico:",
                    [
                        "Inclusivo multinivel (Accesible con ingenio, engancha a todo el grupo)",
                        "Estándar 3º ESO (Currículo ordinario con giro de pensamiento lateral)",
                        "Reto de profundización / Altas Capacidades (Deducción abstracta y generalización)"
                    ],
                    key="sel_dificultad_reto"
                )
            with col_rn2:
                modalidad_reto = st.selectbox(
                    "Dinámica de aplicación en el aula:",
                    [
                        "Póster semanal en corcho (Respuestas en buzón físico hasta el viernes)",
                        "Actividad cooperativa de inicio o cierre de sesión (15-20 min)",
                        "Mini Escape Room por equipos (Competición con compra de pistas)"
                    ],
                    key="sel_modalidad_reto"
                )

            col_rc1, col_rc2 = st.columns(2)
            with col_rc1:
                incluir_pistas = st.checkbox("Incluir 3 pistas escalonadas (Sistema de sobres o desbloqueo gradual)", value=True, key="chk_pistas_reto")
            with col_rc2:
                incluir_solucionario = st.checkbox("Incluir solucionario docente detallado y variantes de ampliación", value=True, key="chk_sol_reto")

            btn_generar_reto = st.form_submit_button("✨ Generar Reto Gamificado Completo", type="primary")

        if btn_generar_reto:
            with st.spinner("Diseñando reto matemático, narrativa y sistema de pistas..."):
                prompt_reto = f"""
                Actúa como un profesor de Matemáticas de Secundaria especialista en Gamificación Educativa, Aprendizaje Basado en Juegos (ABJ) y atención a la diversidad (DUA) en España.
                Diseña un RETO MATEMÁTICO SEMANAL / ESCAPE ROOM MOTIVADOR para alumnos de 3º de ESO (14-15 años).

                PARÁMETROS DEL RETO:
                - Contenido Matemático: {tema_reto}
                - Ambientación Narrativa: {narrativa_reto}
                - Nivel de Dificultad: {dificultad_reto}
                - Dinámica de Aula: {modalidad_reto}
                - Incluir 3 Pistas Graduadas: {'SÍ' if incluir_pistas else 'NO'}
                - Incluir Solucionario Docente: {'SÍ' if incluir_solucionario else 'NO'}

                REGLAS DE FORMATO Y CONTENIDO (RIGUROSAS):
                1. FORMATO LATEX: Escribe TODAS las fórmulas matemáticas y expresiones utilizando LaTeX limpio ($...$ en línea y $$...$$ en bloque).
                2. NARRATIVA ENVOLVENTE: El problema no debe ser una cuenta aburrida, sino un misterio o reto donde las matemáticas sean la herramienta esencial para resolverlo.
                3. ENFOQUE DUA E INCLUSIVO: La premisa debe ser atractiva y fácil de comprender, pero con profundidad suficiente para premiar la perseverancia y el razonamiento deductivo.

                ESTRUCTURA OBLIGATORIA EN MARKDOWN:

                # 🏆 RETO MATEMÁTICO DE LA SEMANA: [TÍTULO IMPACTANTE]

                ## 📜 1. La Misión / Enigma (Cartel para el Alumnado)
                - **Contexto Narrativo:** Breve introducción inmersiva (3-4 líneas).
                - **El Desafío:** Planteamiento visual y matemático claro del problema con fórmulas en LaTeX.
                - **La Pregunta Final:** Qué deben entregar exactamente (ej. un código de 3 cifras, el valor de una incógnita oculta o la justificación lógica).

                ## 💡 2. Sistema de Pistas Escalonadas (Para desbloquear si hay bloqueo)
                *(Diseñadas para entregar en sobres o previa consulta)*
                - **Pista 1 (Activación / Enfoque inicial):** Una pista sutil para orientar por dónde empezar a organizar los datos sin desvelar cuentas.
                - **Pista 2 (Andamiaje matemático):** Una pista intermedia que muestra la relación o ecuación clave.
                - **Pista 3 (Empujón final):** Casi la resolución, para evitar la frustración de alumnos con dificultades.

                ## 🔑 3. Guía Docente y Solucionario Razonado (Uso Exclusivo del Profesor)
                - **Solución Matemática Completa:** Resolución algebraica o numérica paso a paso con LaTeX.
                - **Errores Típicos Anticipados:** Qué fallos comunes de razonamiento cometerán los alumnos y cómo guiarlos con preguntas socráticas.
                - **Ampliación para Altas Capacidades:** Una pregunta o variante adicional para los alumnos que lo resuelvan muy rápido.

                ## 🎮 4. Reglas del Juego y Dinámica de Puntuación Sugerida
                - Puntuación recomendada (ej: 100 pts sin pistas, 75 pts con Pista 1, 50 pts con Pista 2...).
                - Cómo gestionarlo en el aula (buzón físico, corcho o plataforma virtual).
                """

                system_instruction_reto = "Eres un formador de docentes en innovación educativa y gamificación matemática para Educación Secundaria en España. Generas actividades rigurosas en contenido matemático pero sumamente atractivas y lúdicas."
                reto_generado = call_gemini(prompt_reto, system_instruction_reto)
                st.session_state["ultimo_reto_gamificacion"] = reto_generado
                st.session_state["ultimo_tema_reto"] = tema_reto

        if st.session_state.get("ultimo_reto_gamificacion"):
            st.markdown("---")
            st.markdown("### 📄 Reto Gamificado Generado")
            st.markdown(st.session_state["ultimo_reto_gamificacion"])

            col_exp_r1, col_exp_r2 = st.columns([1, 1])
            with col_exp_r1:
                try:
                    docx_reto_file = create_docx_from_markdown(
                        st.session_state["ultimo_reto_gamificacion"],
                        title=f"Reto Semanal - {st.session_state.get('ultimo_tema_reto', 'Matemáticas')}",
                        curso="3º ESO"
                    )
                    st.download_button(
                        label="⬇️ Descargar Cartel y Ficha en Word (.docx)",
                        data=docx_reto_file,
                        file_name=f"reto_semanal_{st.session_state.get('ultimo_tema_reto', 'mates').lower().replace(' ', '_')[:25]}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="btn_descarga_docx_reto"
                    )
                except Exception as e:
                    st.caption(f"ℹ️ Exportación a Word no disponible: {str(e)}")
            with col_exp_r2:
                st.download_button(
                    label="⬇️ Descargar Reto en Texto (.txt)",
                    data=st.session_state["ultimo_reto_gamificacion"],
                    file_name=f"reto_semanal_{st.session_state.get('ultimo_tema_reto', 'mates').lower().replace(' ', '_')[:25]}.txt",
                    mime="text/plain",
                    use_container_width=True,
                    key="btn_descarga_txt_reto"
                )


# =============================================================================
# PESTAÑA 3: 🎯 DISEÑADOR DE EJERCICIOS Y EXÁMENES DUA
# =============================================================================
with tab3:
    st.header("🎯 Diseñador de Pruebas y Tareas Multinivel DUA")
    st.markdown("Genera actividades de aula o exámenes adaptados a la diversidad de 3º de ESO, con fórmulas en formato LaTeX limpio y rúbricas de corrección docente.")
    
    col_dua1, col_dua2 = st.columns([1, 1])
    with col_dua1:
        tema_examen = st.text_input("Contenido matemático específico:", value="Sistemas de dos ecuaciones lineales con dos incógnitas")
        tipo_prueba = st.selectbox(
            "Tipo de instrumento de evaluación:",
            [
                "Prueba Escrita Formal (Examen Parcial o Trimestral)",
                "Hoja de Ejercicios Prácticos para el Aula (Trabajo por Estaciones)",
                "Prueba Competencial Basada en Situaciones Cotidianas"
            ]
        )
    with col_dua2:
        perfiles_seleccionados = st.multiselect(
            "Selecciona los perfiles DUA para los que generar versiones adaptadas:",
            [
                "Refuerzo / Nivel Inicial (Andamiaje, coeficientes enteros, números sencillos)",
                "Estándar / Nivel Medio (Currículo ordinario 3º ESO)",
                "TDAH (Enunciados breves, datos en negrita, tareas divididas en pasos)",
                "Altas Capacidades (Retos de deducción, generalización algebraica, pensamiento lateral)",
                "Discalculia Leve (Apoyo con esquemas, sin distractores aritméticos complejos)"
            ],
            default=["Estándar / Nivel Medio (Currículo ordinario 3º ESO)"]
        )
        
    num_ejercicios = st.slider("Número de ejercicios por versión:", min_value=2, max_value=6, value=3)
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        incluir_soluciones = st.checkbox("Incluir resoluciones completas paso a paso (Uso exclusivo docente)", value=True)
    with col_c2:
        incluir_rubrica = st.checkbox("Incluir criterios de calificación y baremo de puntuación", value=True)
        
    if st.button("🚀 Generar Material Adaptado DUA con LaTeX", type="primary"):
        if not perfiles_seleccionados:
            st.warning("Por favor, selecciona al menos un perfil de diversidad DUA.")
        else:
            with st.spinner("Generando ejercicios multinivel con tipografía LaTeX..."):
                pautas_dict = {
                    "Refuerzo / Nivel Inicial": "Problemas con coeficientes enteros pequeños, pasos guiados con andamiaje visual (\"Paso 1: Despeja la incógnita $x$...\"), sin fracciones complejas.",
                    "Estándar / Nivel Medio": "Currículo ordinario de 3º de ESO con nivel de exigencia formal estándar y resolución autónoma.",
                    "TDAH": "Enunciados sumamente concisos, palabras clave destacadas en **negrita**, sin texto accesorio de relleno, divididos en sub-apartados claros (a, b, c) para reducir la sobrecarga ejecutiva.",
                    "Altas Capacidades": "Ejercicios no rutinarios de deducción lógica, pensamiento lateral, modelización algebraica no estándar o discusión en función de un parámetro.",
                    "Discalculia Leve": "Reducción de la carga de memoria operativa numérica (números sencillos, apoyo visual estructurado), facilitando el planteamiento conceptual sin distractores aritméticos complejos."
                }

                nombres_perfiles_limpios = [p.split("(")[0].strip() for p in perfiles_seleccionados]
                directrices_perfiles_prompt = "\n".join([
                    f"   - **Versión {n}:** {pautas_dict.get(n, 'Adaptación conforme a pautas DUA.')}"
                    for n in nombres_perfiles_limpios
                ])
                encabezados_esperados = "\n".join([f"   - `### Versión: {n}`" for n in nombres_perfiles_limpios])
                lista_nombres_str = ", ".join([f"'{n}'" for n in nombres_perfiles_limpios])

                prompt_dua = f"""
                Actúa como diseñador didáctico de Matemáticas de Secundaria especialista en el Diseño Universal para el Aprendizaje (DUA).
                Genera el siguiente material educativo adaptado:
                
                PARÁMETROS:
                - Tema de Matemáticas: {tema_examen}
                - Tipo de Instrumento: {tipo_prueba}
                - Nivel: 3º de ESO (alumnos de 14-15 años en España)
                - Número de ejercicios por versión: {num_ejercicios} ejercicios
                - NÚMERO TOTAL DE VERSIONES A GENERAR: EXACTAMENTE {len(nombres_perfiles_limpios)} ({lista_nombres_str})
                - Incluir soluciones: {'SÍ' if incluir_soluciones else 'NO'}
                - Incluir rúbrica/criterios: {'SÍ' if incluir_rubrica else 'NO'}
                
                REGLAS CRÍTICAS DE EXCLUSIVIDAD Y FORMATO (OBLIGATORIAS):
                1. EXCLUSIVIDAD TOTAL DE VERSIONES (¡MUY IMPORTANTE!):
                   - Debes generar versiones ÚNICA Y EXCLUSIVAMENTE para los {len(nombres_perfiles_limpios)} perfiles solicitados: {lista_nombres_str}.
                   - ESTÁ TERMINANTEMENTE PROHIBIDO generar versiones para cualquier otro perfil DUA que no esté en la lista anterior. NO agregues versiones que el profesor no haya pedido.
                   - Si el docente seleccionó 1 solo perfil, tu respuesta debe contener ÚNICAMENTE 1 versión. Si seleccionó 2, EXACTAMENTE 2 versiones.

                2. PAUTAS DE ADAPTACIÓN (APLICAR SOLO A LAS VERSIONES SOLICITADAS):
{directrices_perfiles_prompt}

                3. FORMATO LATEX: Escribe TODAS las fórmulas matemáticas y expresiones algebraicas utilizando notación LaTeX limpia:
                   - Fórmulas en línea con un solo signo de dólar, ejemplo: $2x + y = 7$ o $x = \\frac{{-b \\pm \\sqrt{{b^2 - 4ac}}}}{{2a}}$.
                   - Fórmulas en bloque con doble dólar:
                     $$ \\begin{{cases}} 2x + 3y = 12 \\\\ x - y = 1 \\end{{cases}} $$

                4. ESTRUCTURA DE LA RESPUESTA:
                   - Incluye ÚNICAMENTE los siguientes encabezados de versión:
{encabezados_esperados}
                   - NO incluyas encabezados de versiones de perfiles que no fueron solicitados.
                   - Al final, si está solicitado, incluye `### Solucionario Razonado Docente` (únicamente con las resoluciones de las versiones generadas arriba) y `### Criterios de Evaluación y Calificación`.
                """
                
                system_instruction = "Eres un catedrático de Matemáticas de Secundaria y formador de profesorado en DUA. Tus salidas son pedagógicamente impecables y usan LaTeX matemáticamente riguroso."
                material_dua = call_gemini(prompt_dua, system_instruction)
                st.session_state["ultimo_material_dua"] = material_dua
                st.session_state["ultimo_tema_dua"] = tema_examen

    if st.session_state.get("ultimo_material_dua"):
        st.markdown("---")
        st.markdown("### 📋 Material DUA Generado")
        st.markdown(st.session_state["ultimo_material_dua"])
        
        # Botón de Descarga en Word (.docx) con cabecera formal
        try:
            docx_dua_file = create_docx_from_markdown(
                st.session_state["ultimo_material_dua"],
                title=f"Prueba / Actividades DUA - {st.session_state.get('ultimo_tema_dua', 'Matemáticas')}",
                curso="3º ESO"
            )
            st.download_button(
                label="⬇️ Descargar Material DUA en Word (.docx)",
                data=docx_dua_file,
                file_name=f"material_dua_{st.session_state.get('ultimo_tema_dua', 'mates').lower().replace(' ', '_')[:25]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="btn_descarga_docx_dua"
            )
        except Exception as e:
            st.error(f"Error al exportar material DUA a Word: {str(e)}")


# =============================================================================
# PESTAÑA 4: 🤖 ASISTENTE DOCENTE CON CONTEXTO DEL AULA
# =============================================================================
with tab4:
    st.header("🤖 Asistente Docente IA (Contextualizado con 3º ESO)")
    st.markdown("Consulta cualquier duda pedagógica, organizativa o de agrupamientos. La IA analiza los datos anonimizados de la clase para ofrecerte respuestas prácticas.")
    
    # Contexto del aula preparado para el agente
    df_actual = st.session_state["df_alumnos"]
    resumen_alumnos_texto = df_actual.to_string(index=False)
    
    system_prompt_chat = f"""
    Eres el Asistente Pedagógico de Inteligencia Artificial para el profesor de Matemáticas de 3º de ESO.
    Tienes acceso permanente al registro de datos de la clase actual, el cual cumple estrictamente con el RGPD (todos los alumnos están identificados exclusivamente por códigos como ALU_01, ALU_02, etc.).
    
    TABLA OFICIAL DE DATOS DEL AULA (3º ESO):
    {resumen_alumnos_texto}
    
    TUS DIRECTRICES PEDAGÓGICAS Y LEGALES:
    1. PRIVACIDAD RGPD: Nunca pidas ni uses nombres reales. Trabaja siempre con los códigos ALU_XX.
    2. TUTORÍA ENTRE IGUALES: Cuando te pidan agrupamientos o parejas cooperativas, empareja estratégicamente a alumnos con perfiles complementarios (por ejemplo, ALU_03 o ALU_10 de Altas Capacidades/alto rendimiento con ALU_01, ALU_13 o perfiles que se beneficien de una tutoría paciente y estructurada), teniendo en cuenta las notas y observaciones.
    3. ALUMNOS EN RIESGO: Si el docente pregunta por alumnos en riesgo o suspensos, identifica de inmediato a aquellos con Nota_Media < 5.0 y faltas elevadas (como ALU_02, ALU_07, ALU_12, ALU_13, ALU_15) y propone estrategias concretas de rescate académico.
    4. TONO: Profesional, conciso, empático con la labor docente, basado en evidencia pedagógica y matemáticas de secundaria en España.
    """
    
    # Inicializar historial de chat
    if "chat_messages" not in st.session_state:
        st.session_state["chat_messages"] = [
            {
                "role": "assistant",
                "content": "👋 Hola, colega docente. He cargado el registro de los 15 alumnos de 3º de ESO con sus perfiles DUA, notas y faltas. ¿En qué te puedo ayudar hoy? Puedes preguntarme sobre agrupamientos para tutoría entre iguales, alumnos en riesgo de evaluación negativa o adaptaciones metodológicas."
            }
        ]

    # Sugerencias rápidas de preguntas
    st.markdown("**Sugerencias rápidas de consulta:**")
    col_q1, col_q2, col_q3 = st.columns(3)
    pregunta_rapida = None
    with col_q1:
        if st.button("👥 Agrupar por tutoría entre iguales", use_container_width=True):
            pregunta_rapida = "¿Cómo agrupo a mis alumnos por parejas de tutoría entre iguales según la tabla para trabajar sistemas de ecuaciones?"
    with col_q2:
        if st.button("⚠️ Resumen de alumnos en riesgo", use_container_width=True):
            pregunta_rapida = "Hazme un resumen de los alumnos en riesgo de suspenso según sus notas y faltas, y propón un plan de choque."
    with col_q3:
        if st.button("🎯 Adaptaciones para ALU_04 y ALU_05", use_container_width=True):
            pregunta_rapida = "¿Qué pautas concretas debo seguir en clase para atender a ALU_04 (TDAH) y ALU_05 (Discalculia) en la unidad de álgebra?"

    # Renderizar historial de mensajes
    for msg in st.session_state["chat_messages"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Captura de input (por botón rápido o campo de chat)
    user_input = st.chat_input("Escribe tu consulta pedagógica sobre el grupo...")
    prompt_to_send = pregunta_rapida if pregunta_rapida else user_input

    if prompt_to_send:
        # Añadir mensaje del usuario
        st.session_state["chat_messages"].append({"role": "user", "content": prompt_to_send})
        with st.chat_message("user"):
            st.markdown(prompt_to_send)

        # Generar respuesta con IA
        with st.chat_message("assistant"):
            with st.spinner("Analizando datos del aula y reflexionando respuesta pedagógica..."):
                # Construir historial para contexto
                historial_conversacion = ""
                for m in st.session_state["chat_messages"][-6:]: # últimos 6 mensajes
                    historial_conversacion += f"{m['role'].upper()}: {m['content']}\n"
                    
                full_prompt = f"""
                HISTORIAL RECIENTE:
                {historial_conversacion}
                
                NUEVA PREGUNTA DEL PROFESOR:
                {prompt_to_send}
                """
                
                respuesta_ia = call_gemini(full_prompt, system_instruction=system_prompt_chat)
                st.markdown(respuesta_ia)
                st.session_state["chat_messages"].append({"role": "assistant", "content": respuesta_ia})


# =============================================================================
# GESTIÓN DE DOCUMENTOS NORMATIVOS (PDF) CON CACHÉ Y PYPDF
# =============================================================================
DOCS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "documentos_centro")
if not os.path.exists(DOCS_FOLDER):
    DOCS_FOLDER = "documentos_centro"
    os.makedirs(DOCS_FOLDER, exist_ok=True)

def _get_docs_signature(folder_path: str) -> tuple:
    """Genera una firma con los nombres y fechas de modificación de los PDF para auto-invalidar la caché."""
    if not os.path.exists(folder_path):
        return ()
    try:
        files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith(".pdf"))
        return tuple((f, os.path.getmtime(os.path.join(folder_path, f))) for f in files)
    except Exception:
        return ()

@st.cache_data
def load_preloaded_documents(folder_path: str, files_signature: tuple = ()):
    """
    Lee automáticamente todos los archivos .pdf dentro de folder_path usando pypdf.
    Guarda el número de página de cada texto extraído y un resumen de páginas leídas.
    Usa @st.cache_data para una carga instantánea, auto-invalidándose si se añaden o modifican PDFs.
    """
    docs_summary = []
    pages_data = []

    if not os.path.exists(folder_path) or PdfReader is None:
        return docs_summary, pages_data

    try:
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        pdf_files.sort()
    except Exception:
        return docs_summary, pages_data

    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        try:
            reader = PdfReader(file_path)
            num_pages = len(reader.pages)
            for idx, page in enumerate(reader.pages):
                text = (page.extract_text() or "").strip()
                pages_data.append({
                    "documento": filename,
                    "pagina": idx + 1,
                    "texto": text if text else "[Página sin texto seleccionable / imagen escaneada]"
                })
            docs_summary.append({
                "nombre": filename,
                "paginas": num_pages,
                "error": None
            })
        except Exception as e:
            docs_summary.append({
                "nombre": filename,
                "paginas": 0,
                "error": str(e)
            })

    return docs_summary, pages_data

def build_normativa_context(pages_data: list[dict]) -> str:
    """Construye un bloque ordenado de texto identificando claramente cada documento y página."""
    if not pages_data:
        return "No se dispone de documentos normativos en este momento."
    
    chunks = []
    for item in pages_data:
        chunks.append(
            f"--- INICIO PÁGINA ---\n"
            f"DOCUMENTO: {item['documento']} | PÁGINA: {item['pagina']}\n"
            f"CONTENIDO:\n{item['texto']}\n"
            f"--- FIN PÁGINA ---"
        )
    return "\n\n".join(chunks)

def parse_extra_pdf(uploaded_file):
    """Extrae páginas de un PDF subido temporalmente a través de st.file_uploader."""
    if PdfReader is None:
        return [], 0, "Librería pypdf no disponible."
    try:
        reader = PdfReader(uploaded_file)
        total = len(reader.pages)
        pages = []
        for idx, page in enumerate(reader.pages):
            text = (page.extract_text() or "").strip()
            pages.append({
                "documento": f"[Temporal] {uploaded_file.name}",
                "pagina": idx + 1,
                "texto": text if text else "[Página sin texto seleccionable]"
            })
        return pages, total, None
    except Exception as e:
        return [], 0, str(e)


# =============================================================================
# PESTAÑA 5: 📜 NORMATIVA Y PLAN DE CONVIVENCIA
# =============================================================================
with tab5:
    st.header("📜 Normativa y Plan de Convivencia")
    st.markdown("Asistente oficial para resolver dudas sobre el Reglamento de Régimen Interior (RRI), Plan de Convivencia, protocolos y deberes docentes con citación exacta de fuentes.")

    if PdfReader is None:
        st.error("⚠️ La librería `pypdf` no está instalada en el entorno. Por favor, ejecuta `pip install pypdf` para habilitar la lectura de documentos PDF.")
        st.stop()

    # Carga automática con @st.cache_data y auto-invalidación de caché por firma de archivos
    docs_sig = _get_docs_signature(DOCS_FOLDER)
    docs_summary, preloaded_pages = load_preloaded_documents(DOCS_FOLDER, files_signature=docs_sig)

    # 1. Resumen limpio en verde arriba indicando los documentos precargados
    if docs_summary:
        valid_docs = [d for d in docs_summary if not d["error"]]
        error_docs = [d for d in docs_summary if d["error"]]
        total_docs = len(valid_docs)
        total_paginas = sum(d["paginas"] for d in valid_docs)

        doc_badges = "".join([
            f"<div style='margin: 4px 0; color: #166534; font-size: 0.95rem;'>"
            f"📄 <strong>{d['nombre']}</strong> &nbsp; "
            f"<span style='background-color: #bbf7d0; color: #14532d; padding: 2px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85rem;'>{d['paginas']} páginas leídas</span>"
            f"</div>"
            for d in valid_docs
        ])

        col_badge, col_btn_reload = st.columns([5, 1])
        with col_badge:
            st.markdown(f"""
            <div style="background-color: #f0fdf4; border: 1px solid #86efac; border-left: 6px solid #22c55e; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px;">
                <div style="color: #166534; font-weight: 700; font-size: 1.05rem; margin-bottom: 6px; display: flex; align-items: center; gap: 6px;">
                    <span>✅</span> <span>Documentos de Normativa Precargados ({total_docs} archivo{'s' if total_docs != 1 else ''} • {total_paginas} páginas)</span>
                </div>
                {doc_badges}
            </div>
            """, unsafe_allow_html=True)
        with col_btn_reload:
            if st.button("🔄 Recargar", key="btn_reload_docs", help="Fuerza la relectura de los PDFs de la carpeta documentos_centro"):
                load_preloaded_documents.clear()
                st.rerun()

        if error_docs:
            for err in error_docs:
                st.warning(f"⚠️ Error al leer `{err['nombre']}`: {err['error']}")
    else:
        st.warning(f"⚠️ No se encontraron archivos `.pdf` en la carpeta `{DOCS_FOLDER}`. Deposita los documentos de normativa en esa carpeta para precargarlos automáticamente.")

    # Expander opcional para documento extra temporal
    with st.expander("➕ Añadir documento extra temporal"):
        st.caption("Sube un PDF adicional solo para esta sesión de consulta (no sustituye los documentos precargados del centro).")
        extra_pdf = st.file_uploader("Subir documento PDF adicional:", type=["pdf"], key="uploader_normativa_extra")
        extra_pages = []
        if extra_pdf:
            extra_pages, extra_pages_count, err_extra = parse_extra_pdf(extra_pdf)
            if err_extra:
                st.error(f"Error al leer el archivo temporal: {err_extra}")
            else:
                st.info(f"📎 Documento temporal añadido a la consulta: **{extra_pdf.name}** ({extra_pages_count} páginas).")

    # Conjunto completo de páginas activas
    all_active_pages = preloaded_pages + (extra_pages if extra_pdf else [])

    # 2. Chat listo directamente debajo para preguntar al instante
    st.markdown("### 💬 Consulta Inmediata de Normativa")
    st.caption("Formula cualquier pregunta sobre convivencia, sanciones, guardia de aula o uso de dispositivos. La IA responderá exclusivamente con los documentos indicando página exacta.")

    # Botones rápidos de consulta
    col_nq1, col_nq2, col_nq3, col_nq4 = st.columns(4)
    pregunta_rapida_norm = None
    with col_nq1:
        if st.button("📱 Móviles", use_container_width=True, key="btn_norm_movil"):
            pregunta_rapida_norm = "¿Cuál es la normativa sobre el uso de teléfonos móviles y dispositivos electrónicos en el centro y las aulas?"
    with col_nq2:
        if st.button("🚨 Faltas y sanciones", use_container_width=True, key="btn_norm_graves"):
            pregunta_rapida_norm = "¿Qué conductas se tipifican como faltas graves o muy graves y qué sanciones y medidas correctoras contempla el centro?"
    with col_nq3:
        if st.button("🚪 Retrasos y guardias", use_container_width=True, key="btn_norm_guardias"):
            pregunta_rapida_norm = "¿Cuál es el protocolo de actuación ante el retraso de un alumno y qué obligaciones tiene el profesorado de guardia?"
    with col_nq4:
        if st.button("🤝 Plan de Convivencia", use_container_width=True, key="btn_norm_convivencia"):
            pregunta_rapida_norm = "¿Cuáles son las medidas preventivas, protocolos de mediación escolar y normas que establece el Plan de Convivencia?"

    # Inicializar historial de chat de normativa
    if "normativa_chat_messages" not in st.session_state:
        st.session_state["normativa_chat_messages"] = [
            {
                "role": "assistant",
                "content": "👋 Hola, colega docente. He precargado en memoria la normativa del centro con sus páginas indexadas. ¿Qué duda de régimen interno o convivencia necesitas consultar? Responderé citando siempre el documento y la página correspondiente."
            }
        ]

    # Renderizar historial de mensajes
    for msg in st.session_state["normativa_chat_messages"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Input de consulta
    user_input_norm = st.chat_input("Escribe tu consulta sobre la normativa del centro...", key="chat_input_normativa")
    prompt_norm_to_send = pregunta_rapida_norm if pregunta_rapida_norm else user_input_norm

    if prompt_norm_to_send:
        if not all_active_pages:
            st.warning("⚠️ No hay documentos disponibles para consultar. Añade archivos PDF en la carpeta `documentos_centro/` o sube un PDF temporal.")
        else:
            st.session_state["normativa_chat_messages"].append({"role": "user", "content": prompt_norm_to_send})
            with st.chat_message("user"):
                st.markdown(prompt_norm_to_send)

            with st.chat_message("assistant"):
                with st.spinner("Buscando en la normativa del centro y localizando páginas..."):
                    contexto_normativa = build_normativa_context(all_active_pages)

                    system_instruction_normativa = f"""
                    Eres el Asistente Oficial Especializado en Normativa de Régimen Interior, Organización y Plan de Convivencia de este Centro Educativo.
                    Tu labor es orientar y resolver dudas docentes con máximo rigor institucional y legal.

                    DOCUMENTOS NORMATIVOS OFICIALES DISPONIBLES (ORGANIZADOS POR DOCUMENTO Y PÁGINA):
                    {contexto_normativa}

                    REGLAS ESTRICTAS DE RESPUESTA (CUMPLIMIENTO OBLIGATORIO):
                    1. EXCLUSIVIDAD DOCUMENTAL ABSOLUTA:
                       - Debes responder ÚNICA Y EXCLUSIVAMENTE basándote en la información que conste de manera explícita en los documentos normativos aportados arriba.
                       - NO utilices conocimientos previos generales, leyes no citadas en los textos, suposiciones ni inventes procedimientos bajo ningún concepto.
                    2. CITACIÓN OBLIGATORIA DE DOCUMENTO Y PÁGINA:
                       - En CADA afirmación, regla, sanción o medida que menciones, DEBES CITAR SIEMPRE el nombre exacto del documento y el número de página donde se encuentra la norma.
                       - Formato de cita obligatorio: `[Nombre_Documento.pdf, pág. X]` o `[Nombre_Documento.pdf, Página X]`.
                    3. INFORMACIÓN NO DISPONIBLE:
                       - Si la consulta del profesor NO aparece en los documentos precargados o no hay información suficiente en ellos para responder, DEBES RESPONDER TAXATIVAMENTE:
                         "Esa información no consta en la normativa del centro."
                       - No formules conjeturas ni ofrezcas respuestas hipotéticas si la norma no consta en los documentos facilitados.
                    4. TONO Y FORMATO:
                       - Tono formal, riguroso, institucional y de apoyo pedagógico.
                       - Estructura las respuestas con claridad (puntos clave, listas con viñetas) indicando las referencias documentales exactas.
                    """

                    historial_norm = ""
                    for m in st.session_state["normativa_chat_messages"][-6:]:
                        historial_norm += f"{m['role'].upper()}: {m['content']}\n"

                    full_prompt_norm = f"""
                    HISTORIAL DE CONSULTAS RECIENTES:
                    {historial_norm}

                    CONSULTA DEL PROFESOR:
                    {prompt_norm_to_send}
                    """

                    respuesta_norm = call_gemini(full_prompt_norm, system_instruction=system_instruction_normativa)
                    st.markdown(respuesta_norm)
                    st.session_state["normativa_chat_messages"].append({"role": "assistant", "content": respuesta_norm})
