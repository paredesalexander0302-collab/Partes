import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime, timedelta
import PyPDF2  

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
# --- IDs DE DESTINO ---
ID_RESULTADOS = "1daklixp_VDJiTLDVzWFkczT9UQIFqTse2bJbeM2mO_g" # El ID del Excel donde se guardará la productividad
ID_CARPETA_PARTES = "1V6htHEebTWGFym0ffNDCzEmLTiHmIZhU" # El ID de la carpeta donde se guardarán los PDFs
# Reemplaza los textos con los IDs reales de cada subcarpeta en tu Drive
CARPETAS_POR_SUBTIPO = {
    "Antidelincuencial": "10QU99LacLrIzdOiWeSd0Ij2n7d_uZyB6",
    "Eje vial": "1c9W1W2OI3eNra9lW3vJ9jQBTWyBsjyYg",
    "Convoy": "1X-A3V0DpIvJzULLW04Spa1nYIgnGDlw0",
    "Mega operativo": "1KTiegLx4BTAMMXHUXEv5CrtixUti8NUJ",
    "Pandora": "1KmHhAoWhX3Uo5NUMSFT6NDjmLMEzbWlh",
    "Autoridad competente": "1kF_nX8jrbrrblcI-DCuNvzOcZCrqCa81",
    # NOTA: Si todos los POAI van a la misma carpeta "POAI", pega el mismo ID en estos cuatro:
    "POAI Camex": "16uZ4Yrz8OAs-gXSC4_ULLL5WbQl27RRP",
    "POAI Violencia": "1mf07xBwDVqxUHc2fLNx9Lkvg6DLEthh6",
    "POAI Robo de vehículos": "1arBoFN8HvLpEDH_QPSozD_nHKXF0f-VX",
    "POAI Robo a personas": "1HqYyrJqNAuSWfPIAYG8b-hdqU3CRFikz"
}

# --- CONFIGURACIÓN DE LA INTERFAZ ---
st.set_page_config(page_title="Gestión de Operativos", layout="wide")
st.title("Sistema de Registro de Operativos Policiales")
# --- INICIALIZAR MEMORIA TEMPORAL ---
# Esto guarda la lista de policías agregados al operativo actual
if 'personal_participante' not in st.session_state:
    st.session_state.personal_participante = []

# --- CONEXIÓN A GOOGLE SHEETS ---
@st.cache_resource 
def conectar_google():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        # Intenta leer desde la bóveda segura de la nube
        creds_dict = st.json.loads(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    except Exception:
        # Si estás en tu computadora local, usa el archivo
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        
    # ESTA LÍNEA DEBE IR AFUERA DEL EXCEPT (Alineada con el try)
    return gspread.authorize(creds)

# ASEGÚRATE DE TENER ESTA LÍNEA JUSTO DEBAJO PARA CREAR EL CLIENTE
cliente = conectar_google()
# --- LECTURA DE LA MATRIZ ---
# Pega aquí el ID de tu matriz de planificación
ID_PLANIFICACION = "1DYO7c_rqtKdk2s5ZT8JecMsDVrZWW4ZfIDmdYUQux1o" 

@st.cache_data(ttl=60) # Se actualiza cada 60 segundos
def obtener_datos():
    hoja = cliente.open_by_key(ID_PLANIFICACION).sheet1
    datos = hoja.get_all_records()
    return pd.DataFrame(datos)

# --- MOSTRAR EN LA APP ---
st.subheader("📋 Operativos Planificados (Hoy y Mañana)")

try:
    df = obtener_datos()
    
    if not df.empty:
        # 1. Crear una columna interna convirtiendo los textos de fecha a un formato que Python entienda
        df['FECHA_INTERNA'] = pd.to_datetime(df['FECHA'], errors='coerce').dt.date
        
        # 2. Obtener la fecha actual y la de mañana
        hoy = datetime.now().date()
        manana = hoy + timedelta(days=1)
        
        # 3. Filtrar la tabla para que solo queden las filas de hoy y mañana
        df_filtrado = df[(df['FECHA_INTERNA'] == hoy) | (df['FECHA_INTERNA'] == manana)]
        
        # 4. Eliminar la columna interna para que no se vea en la pantalla
        df_mostrar = df_filtrado.drop(columns=['FECHA_INTERNA'])
        
        # Verificar si hay operativos programados para estas fechas antes de mostrar la tabla
        if not df_mostrar.empty:
            st.dataframe(
                df_mostrar,
                use_container_width=True,
                column_config={
                    "UBICACIÓN GPS DEL OPERATIVO": st.column_config.LinkColumn(
                        "Mapa del Operativo",
                        help="Abrir en Google Maps",
                        display_text="📍 Ver ubicación"
                    ),
                    "ENLACE ANTI/EJE": st.column_config.LinkColumn(
                        "Formulario Anti/Eje",
                        help="Abrir formulario Survey123",
                        display_text="🔗 Abrir Anti/Eje"
                    ),
                    "ENLACE CONVOY": st.column_config.LinkColumn(
                        "Formulario Convoy",
                        help="Abrir formulario Survey123",
                        display_text="🔗 Abrir Convoy"
                    ),
                    "ENLACE CHECK POINT": st.column_config.LinkColumn(
                        "Formulario Check Point",
                        help="Abrir formulario Survey123",
                        display_text="🔗 Abrir Check Point"
                    ),
                    "ACTIVIDADES": st.column_config.LinkColumn( 
                        "Documento de Actividades",
                        help="Abrir documento adjunto",
                        display_text="📄 Ver Documento"
                    )
                }
            )
        else:
            st.info("No hay operativos planificados para el día de hoy ni para mañana.")
            
    else:
        st.warning("La matriz está vacía. Añade datos en Google Sheets.")
        
except Exception as e:
    st.error(f"Hubo un error al procesar los datos: {e}")
   # --- SECCIÓN DE REGISTRO DE OPERATIVOS ---
st.divider()
st.header("📝 Registro de Nuevo Operativo")

# --- 1. CLASIFICACIÓN ---
# --- 1. CLASIFICACIÓN ---
st.subheader("1. Clasificación")
col_tipo, col_subtipo = st.columns(2)

with col_tipo:
    tipo_operativo = st.selectbox("Tipo de Operativo", ["Ordinario", "Extraordinario"])

with col_subtipo:
    if tipo_operativo == "Ordinario":
        # AGREGAMOS "Convoy" A ESTA LISTA
        sub_tipo = st.selectbox("Clase de Operativo", [
            "Antidelincuencial", "Eje vial", "Mega operativo", "Convoy", 
            "POAI Camex", "POAI Violencia", "POAI Robo de vehículos", "POAI Robo a personas"
        ])
    else:
        sub_tipo = st.selectbox("Clase de Operativo", ["Pandora", "Autoridad competente"])

# --- CONEXIÓN PARA LA MATRIZ DE PERSONAL ---
# RECUERDA: Mantén aquí tu ID real del archivo de personal
ID_PERSONAL = "1PCqbjJ4CoDrk9Y9z_aFxHEJTBS2L-Sj-H7eNV4aK6w4" 

@st.cache_data(ttl=600)
def obtener_personal():
    try:
        hoja = cliente.open_by_key(ID_PERSONAL).sheet1
        datos = hoja.get_all_records()
        df = pd.DataFrame(datos)
        if not df.empty:
            # Convertimos a texto y rellenamos con ceros a la izquierda hasta completar los 10 dígitos
            df['CEDULA'] = df['CEDULA'].astype(str).str.zfill(10)
        return df
    except Exception as e:
        st.error(f"Error técnico al conectar con Google Sheets (Personal): {e}")
        return pd.DataFrame()

# --- 2. PERSONAL POLICIAL Y DESPLIEGUE ---
st.markdown("---")
st.subheader("👥 2. Personal Policial Participante y Despliegue")
st.info("Ingrese la cédula y asigne la unidad a la que pertenece el servidor en este operativo.")

df_personal = obtener_personal()

# Dividimos en 5 columnas
col_ced, col_tipo, col_uni, col_comp, col_btn = st.columns([2, 1.5, 2, 1, 1])

with col_ced:
    cedula_input = st.text_input("Número de cédula:", "")

with col_tipo:
    # Separamos las 3 opciones principales
    tipo_asignacion = st.selectbox("Tipo de Asignación:", ["Circuito", "Check Point", "GOM"])

with col_uni:
    # Lógica condicional para el lugar según lo elegido
    if tipo_asignacion == "Circuito":
        opciones_unidad = ["San Miguel", "Chimbo", "La Magdalena", "San Pablo", "Balsapamba", "Santiago", "Telimbela"]
        unidad_seleccionada = st.selectbox("Lugar:", opciones_unidad)
    elif tipo_asignacion == "Check Point":
        # Permite escribir el lugar exacto del check point
        unidad_seleccionada = st.text_input("Referencia del Check Point:", "CHECK POINT")
    else:
        opciones_unidad = ["GOM 1", "GOM 2", "GOM 3", "GOM 4", "GOM 5", "GOM 6", "GOM 7", "GOM 8"]
        unidad_seleccionada = st.selectbox("Unidad Motorizada:", opciones_unidad)

with col_comp:
    # Si es Circuito o Check Point, pide compañía. Si es GOM, lo bloquea en N/A.
    if tipo_asignacion in ["Circuito", "Check Point"]:
        compania_seleccionada = st.selectbox("Compañía:", ["1", "2", "3"])
    else:
        st.text_input("Compañía:", "N/A", disabled=True)
        compania_seleccionada = "N/A"

with col_btn:
    st.write("") 
    st.write("")
    btn_buscar = st.button("🔍 Agregar")

# Lógica de búsqueda
if btn_buscar and cedula_input:
    if not df_personal.empty:
        cedula_limpia = cedula_input.strip()
        servidor = df_personal[df_personal['CEDULA'] == cedula_limpia]
        
        if not servidor.empty:
            nombre = servidor.iloc[0].get('NOMBRES_COMPLETOS', 'Desconocido')
            grado = servidor.iloc[0].get('GRADO', '')
            
            ya_existe = any(p['CEDULA'] == cedula_limpia for p in st.session_state.personal_participante)
            
            if not ya_existe:
                st.session_state.personal_participante.append({
                    "CEDULA": cedula_limpia,
                    "GRADO": grado,
                    "NOMBRES_COMPLETOS": nombre,
                    "COMPAÑÍA": compania_seleccionada,
                    "UNIDAD": unidad_seleccionada 
                })
                if compania_seleccionada == "N/A":
                    st.success(f"✅ Agregado: {grado} {nombre} ({unidad_seleccionada})")
                else:
                    st.success(f"✅ Agregado: {grado} {nombre} ({unidad_seleccionada} - Cía {compania_seleccionada})")
            else:
                st.warning("⚠️ Este servidor ya fue agregado.")
        else:
            st.error("❌ Cédula no encontrada en la matriz.")
    else:
        st.error("No se pudo cargar la base de datos de personal.")

# Mostrar la tabla en tiempo real
if st.session_state.personal_participante:
    st.write("**Servidores policiales registrados en el operativo:**")
    st.dataframe(pd.DataFrame(st.session_state.personal_participante), use_container_width=True)
    if st.button("🗑️ Limpiar lista de personal"):
        st.session_state.personal_participante = []
        st.rerun()
        
# --- 3. PRODUCTIVIDAD DEL OPERATIVO ---
st.markdown("---")
st.subheader("📊 3. Productividad del Operativo")
st.info("Ingrese los resultados numéricos obtenidos durante el operativo (deje en 0 si no hubo novedades en ese rubro).")

# --- BLOQUE PRINCIPAL (REGISTROS) ---
st.write("**Controles Generales**")
col_reg1, col_reg2, col_reg3 = st.columns(3)

with col_reg1:
    personas_reg = st.number_input("Personas Registradas", min_value=0, step=1)
with col_reg2:
    vehiculos_reg = st.number_input("Vehículos Registrados", min_value=0, step=1)
with col_reg3:
    motos_reg = st.number_input("Motos Registradas", min_value=0, step=1)

# Separador visual
st.divider() 

# --- BLOQUE SECUNDARIO (NOVEDADES Y RESULTADOS) ---
st.write("**Novedades y Retenciones**")
col_nov1, col_nov2, col_nov3 = st.columns(3)

with col_nov1:
    armas_fuego = st.number_input("Armas de Fuego", min_value=0, step=1)
    detenidos_bol = st.number_input("Detenidos Boletas", min_value=0, step=1)
    vehiculos_ret = st.number_input("Vehículos Retenidos", min_value=0, step=1)

with col_nov2:
    armas_blancas = st.number_input("Armas Blancas", min_value=0, step=1)
    detenidos_vif = st.number_input("Detenidos VIF", min_value=0, step=1)
    vehiculos_recup = st.number_input("Vehículos Recuperados", min_value=0, step=1)

with col_nov3:
    polarizados = st.number_input("Polarizados", min_value=0, step=1)
    detenidos_del = st.number_input("Detenidos Delitos", min_value=0, step=1)
    clausura_bares = st.number_input("Clausura Bares", min_value=0, step=1)
    
    # --- 4. NOVEDADES DEL OPERATIVO ---
st.markdown("---")
st.subheader("📝 4. Novedades del Operativo")

# Cuadro de texto amplio para describir detalles
novedades = st.text_area(
    "Detalle de las novedades suscitadas:", 
    height=150, 
    placeholder="Ej: Durante el operativo no se registraron novedades de importancia... o detalle de las circunstancias de una aprehensión."
)
# --- 5. CARGA DE PARTE Y VERIFICACIÓN DE FIRMAS ---
st.markdown("---")
st.subheader("📄 5. Carga de Parte Policial y Verificación")

parte_pdf = st.file_uploader("Sube el parte policial firmado digitalmente (PDF):", type=["pdf"])

if parte_pdf is not None:
    try:
        # Leemos el archivo PDF en memoria
        lector = PyPDF2.PdfReader(parte_pdf)
        firmas_encontradas = 0
        
        # Buscamos los campos de firma digital
        for pagina in lector.pages:
            if '/Annots' in pagina:
                for anotacion in pagina['/Annots']:
                    obj_anotacion = anotacion.get_object()
                    if obj_anotacion.get('/FT') == '/Sig':
                        firmas_encontradas += 1
                        
        # Calculamos cuántos policías hay en la lista del Paso 2
        num_policias = len(st.session_state.personal_participante)
        
        col_res1, col_res2 = st.columns(2)
        col_res1.info(f"👮‍♂️ Policías registrados: **{num_policias}**")
        col_res2.info(f"✍️ Firmas detectadas: **{firmas_encontradas}**")
        
        # Lógica de validación
        if num_policias == 0:
            st.warning("⚠️ No has registrado ningún servidor policial en el Paso 2.")
        elif firmas_encontradas == num_policias:
            st.success("✅ ¡Verificación exitosa! El número de firmas coincide perfectamente.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Botón DEFINITIVO de guardado
            if st.button("💾 Guardar Operativo y Subir a Drive", type="primary", use_container_width=True):
                with st.spinner("Subiendo PDF a Drive y guardando datos en la matriz..."):
                    try:
                        import base64
                        import requests
                        
                        # --- 1. SUBIR A GOOGLE DRIVE VÍA APPS SCRIPT ---
                        # !!! PEGA AQUÍ LA URL QUE OBTUVISTE EN GOOGLE APPS SCRIPT !!!
                        URL_WEB_APP = "https://script.google.com/macros/s/AKfycbzY7Z4BtIdOpJf2NIhe2rsbGymsAIW_dX80jPHouX1ltCJGEraw1RB4JshE-mU-TgD-/exec" 
                        
                        fecha_hora_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
                        nombre_archivo = f"PARTE_{sub_tipo}_{fecha_hora_actual}.pdf"
                        
                        # Convertir el PDF a código de texto (Base64) para enviarlo por internet
                        pdf_bytes = parte_pdf.getvalue()
                        pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
                        
                        # 1. Identificar a qué subcarpeta debe ir. 
                        # Si por algún error no encuentra el subtipo en el mapa, usará la carpeta raíz por defecto.
                        carpeta_destino = CARPETAS_POR_SUBTIPO.get(sub_tipo, ID_CARPETA_PARTES)
                        
                        # 2. Paquete de datos que enviaremos al puente de Google
                        payload = {
                            "folderId": carpeta_destino,
                            "filename": nombre_archivo,
                            "fileData": pdf_base64
                        }
                        
                        # Enviar el archivo
                        respuesta = requests.post(URL_WEB_APP, data=payload)
                        datos_respuesta = respuesta.json()
                        
                        if datos_respuesta.get("status") == "success":
                            enlace_pdf = datos_respuesta.get("url")
                            
# --- 2. CONECTAR Y GUARDAR EN GOOGLE SHEETS ---
                            hoja_resultados = cliente.open_by_key(ID_RESULTADOS).sheet1
                            
                            # Formato de nombres de los servidores
                            nombres_personal = " | ".join([f"{p.get('GRADO', '')} {p.get('APELLIDOS Y NOMBRES', p.get('NOMBRES_COMPLETOS', 'Desconocido'))} ({p.get('UNIDAD', '')})" for p in st.session_state.personal_participante])
                            
                            # Extraer automáticamente las compañías, circuitos y GOMs del personal registrado
                            companias_usadas = ", ".join(list(set([str(p.get('COMPAÑÍA', '')) for p in st.session_state.personal_participante if p.get('COMPAÑÍA', '') != 'N/A'])))
                            circuitos_usados = ", ".join(list(set([str(p.get('UNIDAD', '')) for p in st.session_state.personal_participante if not str(p.get('UNIDAD', '')).startswith('GOM')])))
                            goms_usados = ", ".join(list(set([str(p.get('UNIDAD', '')) for p in st.session_state.personal_participante if str(p.get('UNIDAD', '')).startswith('GOM')])))
                            
                            # ESTA ES LA ESTRUCTURA EXACTA DE TU EXCEL (De la A a la V)
                            nueva_fila = [
                                datetime.now().strftime("%d/%m/%Y %H:%M"), # A. Fecha
                                companias_usadas,                          # B. Compañías
                                circuitos_usados,                          # C. Circuitos
                                goms_usados,                               # D. GOM
                                tipo_operativo,                            # E. Tipo
                                sub_tipo,                                  # F. Subtipo
                                "SI",                                      # G. Ejecutado
                                nombres_personal,                          # H. Servidores_Policiales
                                personas_reg,                              # I. Personas_Registradas
                                vehiculos_reg,                             # J. Vehiculos_Registrados
                                motos_reg,                                 # K. Motos_Registradas
                                armas_fuego,                               # L. Armas_Fuego
                                detenidos_bol,                             # M. Detenidos_Boletas
                                detenidos_vif,                             # N. Detenidos_VIF
                                polarizados,                               # O. Polarizados
                                vehiculos_ret,                             # P. Vehiculos_Retenidos
                                armas_blancas,                             # Q. Armas_Blancas
                                vehiculos_recup,                           # R. Vehiculos_Recuperados
                                detenidos_del,                             # S. Detenidos_Delitos
                                clausura_bares,                            # T. Clausura_Bares
                                novedades,                                 # U. Novedades
                                enlace_pdf                                 # V. Archivo_Parte
                            ]
                            
                            hoja_resultados.append_row(nueva_fila)
                            
                            st.success("🎉 ¡Operativo registrado y Parte Policial archivado exitosamente!")
                            st.session_state.personal_participante = [] 
                            
                        else:
                            st.error(f"❌ Error al subir el PDF a Drive: {datos_respuesta.get('message')}")
                            
                    except Exception as e:
                        st.error(f"Hubo un error en el proceso interno de guardado: {e}")
                        
        else:
            st.error("❌ Discrepancia detectada: El número de firmas en el documento no coincide con el personal registrado. Por favor, verifica el documento.")
            
    except Exception as e:
        st.error(f"Error al analizar el documento PDF: {e}")