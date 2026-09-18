import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configuración de la página
st.set_page_config(
    page_title="Portal de Garantías - Sucursales",
    page_icon="📦",
    layout="centered"
)

# Título principal
st.title("📦 Portal Único de Gestión de Garantías")
st.markdown("Herramienta estandarizada para radicar garantías ante marcas y proveedores según **Ley 1480**.")
st.markdown("---")

# Carga de archivos Excel
@st.cache_data
def cargar_matrices():
    try:
        df_marcas = pd.read_excel("matrix marcas.xlsx")
    except Exception as e:
        df_marcas = pd.DataFrame()
        st.error(f"Error al cargar 'matrix marcas.xlsx': {e}")
        
    try:
        df_sucursales = pd.read_excel("CORREOS SUCURSALES.xlsx")
    except Exception as e:
        df_sucursales = pd.DataFrame()
        st.error(f"Error al cargar 'CORREOS SUCURSALES.xlsx': {e}")
        
    return df_marcas, df_sucursales

df_marcas, df_sucursales = cargar_matrices()

# Inicialización de estado de sesión
if "confirmar_sin_foto" not in st.session_state:
    st.session_state["confirmar_sin_foto"] = False

if df_marcas.empty:
    st.warning("⚠️ No se pudo cargar la matriz de marcas. Verifica que el archivo 'matrix marcas.xlsx' esté en la misma carpeta.")
else:
    # = PASO 1: DATOS DE CONTROL INTERNO =
    st.subheader("1. Datos de Control Interno")
    
    if not df_sucursales.empty and "SUCURSAL" in df_sucursales.columns:
        lista_sucursales = df_sucursales["SUCURSAL"].dropna().unique()
    else:
        lista_sucursales = [f"Sucursal {i}" for i in range(100, 150)]
        
    sucursal_sel = st.selectbox("Selecciona la Sucursal de origen *", ["-- Seleccione una Sucursal --"] + list(lista_sucursales))
    
    # Campo Obligatorio de Número de Caso
    numero_caso = st.text_input("Número de Caso SIESA / CRM *", placeholder="Ej: CASO-10575")
    
    if not numero_caso.strip():
        st.warning("⚠️ El Número de Caso SIESA / CRM es obligatorio para radicar cualquier solicitud.")

    st.markdown("---")

    # = PASO 2: SELECCIÓN DE MARCA =
    st.subheader("2. Selección de Marca")
    
    col_marcas = "MARCA" if "MARCA" in df_marcas.columns else df_marcas.columns[0]
    lista_marcas = df_marcas[col_marcas].dropna().unique()
    
    marca_sel = st.selectbox("Selecciona la Marca del Producto *", ["-- Seleccione una Marca --"] + list(lista_marcas))

    if marca_sel != "-- Seleccione una Marca --":
        info_marca = df_marcas[df_marcas[col_marcas] == marca_sel].iloc[0]
        
        canal = info_marca.get("CANAL DE ATENCION", "Correo")
        correo_marca = info_marca.get("CORREO", "")
        telefono_marca = info_marca.get("TELEFONO / WHATSAPP", "")
        linea_marca = info_marca.get("LINEA 018000", "")

        st.info(f"📍 **Canal de Atención Registrado:** {canal}")

        # CANAL: WHATSAPP / TELEFÓNICO
        if "whatsapp" in str(canal).lower() or "linea" in str(canal).lower() or "telef" in str(canal).lower():
            st.success("📱 **Esta marca gestiona garantías mediante línea de atención o WhatsApp.**")
            
            if pd.notna(telefono_marca) and str(telefono_marca).strip() != "":
                st.write(f"👉 **Número de Contacto / WhatsApp:** {telefono_marca}")
                num_wa = "".join(filter(str.isdigit, str(telefono_marca)))
                if num_wa:
                    st.markdown(f"[💬 Abrir Chat de WhatsApp Directo](https://wa.me/{num_wa})", unsafe_allow_html=True)
            
            if pd.notna(linea_marca) and str(linea_marca).strip() != "":
                st.write(f"📞 **Línea 01-8000:** {linea_marca}")

        # CANAL: CORREO ELECTRÓNICO
        else:
            st.markdown("---")
            st.subheader("3. Formulario de Radicación de Garantía")
            
            nombre_cliente = st.text_input("Nombre Completo del Cliente *")
            cedula_cliente = st.text_input("Cédula / NIT del Cliente *")
            telefono_cliente = st.text_input("Teléfono de Contacto del Cliente *")
            direccion_cliente = st.text_input("Dirección de Residencia del Cliente *")
            
            st.markdown("---")
            referencia = st.text_input("Referencia o Modelo del Producto *")
            serial = st.text_input("Número de Serial del Producto")
            falla = st.text_area("Descripción Detallada de la Falla Reclamada *")

            st.markdown("---")
            st.subheader("4. Evidencia Fotográfica")
            st.caption("Adjunte la fotografía de la placa con el serial o el sticker visible.")
            foto_serial = st.file_uploader("Subir foto de la Placa / Serial (JPG, PNG)", type=["jpg", "jpeg", "png"])

            st.markdown("---")

            # LÓGICA DE PROCESAMIENTO Y ENVÍO DE CORREO
            def procesar_envio(con_foto=True):
                # 1. Obtener correo de la sucursal seleccionada
                correo_sucursal = ""
                if not df_sucursales.empty and "SUCURSAL" in df_sucursales.columns:
                    match_suc = df_sucursales[df_sucursales["SUCURSAL"] == sucursal_sel]
                    if not match_suc.empty:
                        col_correo = [c for c in match_suc.columns if "CORREO" in str(c).upper() or "EMAIL" in str(c).upper()]
                        if col_correo:
                            correo_sucursal = str(match_suc.iloc[0][col_correo[0]]).strip()

                # 2. Construir lista de correos en Copia (CC)
                correos_cc = [
                    "gestiondegarantias@lagobo.com",
                    "garantiasoportunidades@lagobo.com",
                    "garantiasuplementaria@lagobo.com"
                ]
                if correo_sucursal and "@" in correo_sucursal:
                    correos_cc.append(correo_sucursal)

                # 3. Construir Asunto en MAYÚSCULAS Obligatorio
                caso_clean = numero_caso.strip().upper()
                cliente_clean = nombre_cliente.strip().upper()
                marca_clean = str(marca_sel).strip().upper()
                asunto_final = f"CASO: {caso_clean} - GARANTÍA {cliente_clean} - {marca_clean}"

                # Visualización exitosa en pantalla
                st.success(f"✅ ¡Garantía radicada con éxito!")
                st.markdown(f"**📌 Asunto Generado:** `{asunto_final}`")
                st.markdown(f"**📩 Destinatario Principal (Marca):** `{correo_marca}`")
                st.markdown(f"**📧 En Copia (CC):** `{', '.join(correos_cc)}`")

                if not con_foto:
                    st.info("ℹ️ Solicitud enviada sin foto de serial según confirmación de la sucursal.")
                else:
                    st.info("📷 Foto del serial adjuntada correctamente.")

                st.balloons()
                st.session_state["confirmar_sin_foto"] = False

            # BOTÓN DE ENVÍO Y VALIDACIONES
            if st.button("🚀 ENVIAR SOLICITUD AUTOMÁTICA"):
                if sucursal_sel == "-- Seleccione una Sucursal --":
                    st.error("❌ ERROR OBLIGATORIO: Debes seleccionar la sucursal de origen.")
                    st.session_state["confirmar_sin_foto"] = False
                elif not numero_caso.strip():
                    st.error("❌ ERROR OBLIGATORIO: Debe ingresar el Número de Caso SIESA / CRM para poder radicar.")
                    st.session_state["confirmar_sin_foto"] = False
                elif not nombre_cliente.strip() or not cedula_cliente.strip() or not referencia.strip() or not falla.strip():
                    st.error("❌ ERROR: Completa todos los campos obligatorios del cliente y del producto.")
                    st.session_state["confirmar_sin_foto"] = False
                elif foto_serial is None:
                    st.session_state["confirmar_sin_foto"] = True
                else:
                    procesar_envio(con_foto=True)

            # DESPLEGABLE / ALERTA SI NO HAY FOTO DE SERIAL
            if st.session_state["confirmar_sin_foto"]:
                st.warning("⚠️ **ATENCIÓN: No has adjuntado la fotografía del serial.**")
                st.write("Si el producto no tiene placa de serial o no es posible tomarle foto, puedes enviar el correo de todas formas para no varar la atención del cliente.")
                
                col_si, col_no = st.columns(2)
                with col_si:
                    if st.button("⚠️ Confirmar y Enviar SIN foto"):
                        procesar_envio(con_foto=False)
                with col_no:
                    if st.button("❌ Cancelar para subir foto"):
                        st.session_state["confirmar_sin_foto"] = False
                        st.rerun()