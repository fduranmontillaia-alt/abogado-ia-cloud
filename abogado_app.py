import streamlit as st
from groq import Groq
from docx import Document
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email sophistication import encoders

# Configuración del título principal de la aplicación web
st.title("⚖️ Franc - Experto en Derecho Inmobiliario y Mercantil")
st.write("Asistente legal inteligente para la redacción de contratos blindados en Venezuela.")

# ------------------------------------------------------------------
# FASE 0: RECOPILACIÓN REQUERIDA DE INFORMACIÓN (MANDATO INICIAL)
# ------------------------------------------------------------------
st.subheader("📋 Fase 0: Datos del Requerimiento Legal")
st.write("Por favor, suministre los siguientes datos para iniciar. Copie, pegue y rellene este formato en la casilla de abajo:")

formato_ejemplo = """- Tipo de documento/requerimiento: [Ej. Contrato de arrendamiento comercial]
- Datos de las partes: [Ej. Arrendador: Persona jurídica C.A. / Arrendatario: Persona natural]
- Ubicación y características del inmueble: [Ej. Local comercial de 80mt2 en Chacao, Caracas]
- Moneda y condiciones financieras: [Ej. Canon de $1.000 mensuales, pagaderos en divisas, 3 meses de depósito]
- Duración de la relación: [Ej. Contrato por 2 años]"""

st.code(formato_ejemplo, language="text")

datos_usuario = st.text_area("Introduzca aquí los datos recolectados:", height=250)

# ------------------------------------------------------------------
# FUNCIÓN DEL CEREBRO LEGAL (CONEXIÓN CON GROQ CLOUD)
# ------------------------------------------------------------------
def consultar_abogado_ia(datos_fase0):
    try:
        client = Groq(api_key=st.secrets["GROQ_API_KEY"])
        
        system_prompt = """ROLE: Eres un abogado experto en Derecho Inmobiliario y Mercantil en Venezuela, con más de 15 años de experiencia en asesoría corporativa, redacción de contratos complejos y litigios estratégicos. Tu conocimiento abarca el Código Civil venezolano, el Código de Comercio, el Decreto con Rango, Valor y Fuerza de Ley de Regulación del Arrendamiento Inmobiliario para el Uso Comercial (2014), la Sundde, y los criterios vinculantes de la Sala de Casación Civil y Sala Constitucional del Tribunal Supremo de Justicia (TSJ).

TASK: Tu objetivo es brindar asesoría legal de alto nivel, redactar contratos blindados y analizar casos complejos sobre arrendamientos de locales comerciales, compraventas de inmuebles, procedimientos de desalojo y la regulación de bienes inmuebles de uso comercial, corporativo o industrial en Venezuela.

PROTOCOLO DE VERACIDAD Y VALIDACIÓN INTERNA: Basa cada afirmación en la legislación venezolana vigente y jurisprudencia real del TSJ. Cita la ley, el artículo específico o el número de sentencia de forma clara. Si una norma es ambigua o un criterio judicial no está consolidado, di explícitamente la duda. Prohibido inventar leyes, artículos o gacetas.

OUTPUT FORMAT: Para Redacción de Contratos: Entrega la estructura completa del documento legal con títulos claros, cláusulas numeradas, y añade notas entre corchetes explicando el porqué de las cláusulas críticas (como la cláusula de valor de referencia en divisas o jurisdicción)."""

        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Aquí tienes los datos recolectados de la Fase 0:\n\n{datos_fase0}\n\nPor favor, procede a generar el documento legal blindado aplicando estrictamente tu protocolo."}
            ],
            temperature=0.3
        )
        return completion.choices[0].message.content
    except Exception as e:
        return f"Error en el cerebro legal de IA: {e}"

# ------------------------------------------------------------------
# FUNCIÓN PARA CREAR EL ARCHIVO WORD (.DOCX) EN MEMORIA
# ------------------------------------------------------------------
def crear_documento_word(texto_legal):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    
    for linea in texto_legal.split('\n'):
        doc.add_paragraph(linea)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------------------------------------------------------
# FUNCIÓN PARA ENVIAR EL CORREO ELECTRÓNICO CON EL ARCHIVO ADJUNTO
# ------------------------------------------------------------------
def enviar_correo_con_adjunto(correo_destino, texto_legal, archivo_bytes):
    try:
        # Extraemos las credenciales guardadas en la caja fuerte de Streamlit
        remitente = st.secrets["EMAIL_REMITENTE"]
        password = st.secrets["EMAIL_PASSWORD"]
        
        # Estructuramos el esqueleto del mensaje de correo
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = correo_destino
        msg['Subject'] = "⚖️ Su Documento Legal Blindado - Abogado IA"
        
        cuerpo = "Hola. Adjunto a este correo encontrará el documento legal en formato Word (.docx) generado por el sistema automatizado de Inteligencia Artificial."
        msg.attach(MIMEText(cuerpo, 'plain'))
        
        # Configuramos el archivo adjunto en binario
        adjunto = MIMEBase('application', 'octet-stream')
        adjunto.set_payload(archivo_bytes.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header('Content-Disposition', 'attachment', filename="documento_legal_blindado.docx")
        msg.attach(adjunto)
        
        # Conexión cifrada y segura con los servidores SMTP de Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, correo_destino, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error al enviar el correo electrónico: {e}")
        return False

# Botón para activar el procesamiento
if st.button("Generar Documento Legal"):
    if datos_usuario.strip() == "":
        st.warning("⚠️ Por favor, debe introducir los datos solicitados en la Fase 0 antes de continuar.")
    else:
        with st.spinner("⚖️ El Abogado de IA está analizando las leyes venezolanas y redactando el documento..."):
            documento_redactado = consultar_abogado_ia(datos_usuario)
            st.session_state["documento_resultado"] = documento_redactado
            st.success("✨ ¡Documento Legal Generado con éxito!")

# Si el documento ya existe en memoria, habilitamos las herramientas de entrega
if "documento_resultado" in st.session_state:
    st.markdown("### 📄 Previsualización del Documento:")
    st.write(st.session_state["documento_resultado"])
    
    # Fabricamos el Word en la memoria intermedia
    archivo_word_descarga = crear_documento_word(st.session_state["documento_resultado"])
    archivo_word_correo = crear_documento_word(st.session_state["documento_resultado"])
    
    # Opción A: Descarga Local
    st.download_button(
        label="📥 Descargar Documento en formato Word (.docx)",
        data=archivo_word_descarga,
        file_name="documento_legal_blindado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
    st.markdown("---")
    
    # Opción B: Envío por Email Automatizado
    st.markdown("### 📧 Enviar Documento Directo al Correo")
    email_cliente = st.text_input("Escriba el correo electrónico del destinatario:")
    
    if st.button("Enviar por Correo Electrónico"):
        if email_cliente.strip() == "" or "@" not in email_cliente:
            st.error("⚠️ Por favor, introduzca una dirección de correo electrónico válida.")
        else:
            with st.spinner("🚀 Enviando correo con el archivo adjunto..."):
                exito = enviar_correo_con_adjunto(email_cliente, st.session_state["documento_resultado"], archivo_word_correo)
                if exito:
                    st.success(f"📬 ¡Correo enviado con éxito a {email_cliente}!")