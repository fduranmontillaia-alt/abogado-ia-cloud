import streamlit as st
from groq import Groq
from docx import Document
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configuración del título principal de la aplicación web
st.title("⚖️ Franc - Experto en Derecho Inmobiliario y Mercantil")
st.write("Asistente legal inteligente para la redacción de contratos blindados en Venezuela.")

# ------------------------------------------------------------------
# FASE 0: RECOPILACIÓN REQUERIDA DE INFORMACIÓN (MANDATO INICIAL OBLIGATORIO)
# ------------------------------------------------------------------
st.subheader("📋 Fase 0: Datos del Requerimiento Legal")
st.write("Por favor, suministre los siguientes datos para iniciar (copie, pegue y rellene este formato):")

formato_ejemplo = """- Tipo de documento/requerimiento: [Ej. Contrato de arrendamiento comercial / Análisis de riesgo de desalojo]
- Datos de las partes: [Ej. Arrendador (Nombre, Cédula/RIF, Domicilio) / Arrendatario (Empresa, Registro Mercantil, Representante)]
- Ubicación y características del inmueble: [Ej. Local comercial de 80mt2 en Mérida, Edo. Mérida. Especificar linderos o estado de pintura/pisos si aplica]
- Uso específico del comercio: [Ej. Carnicería, venta de repuestos, oficinas]
- Moneda y condiciones financieras: [Ej. Canon en USD mensuales, porcentaje para reparaciones menores, indemnización por ocupación indebida]
- Garantías y Pagos: [Ej. Cantidad de meses de depósito, método de pago tasa BCV]
- Duración de la relación y Fecha: [Ej. Contrato por 1 año, Fecha de inicio y fecha de firma]"""

st.code(formato_ejemplo, language="text")

datos_usuario = st.text_area("Introduzca aquí los datos recolectados:", height=250)

# ------------------------------------------------------------------
# FUNCIÓN DEL CEREBRO LEGAL (CONEXIÓN CON GROQ CLOUD)
# ------------------------------------------------------------------
def consultar_abogado_ia(datos_fase0):
    try:
        client = Groq(api_key=st.secrets["GROQ_API_KEY"])
        
        system_prompt = """FASE 0: RECOPILACIÓN REQUERIDA DE INFORMACIÓN (MANDATO INICIAL OBLIGATORIO)
No asumas datos ni comiences a redactar hasta que el usuario te provee las variables requeridas.

ROLE:
Eres un abogado experto en Derecho Inmobiliario y Mercantil en Venezuela, con más de 15 años de experiencia en asesoría corporativa, redacción de contratos complejos y litigios estratégicos. Tu conocimiento abarca el Código Civil venezolano, el Código de Comercio, el Decreto con Rango, Valor y Fuerza de Ley de Regulación del Arrendamiento Inmobiliario para el Uso Comercial (2014), la Sundde, y los criterios vinculantes de la Sala de Casación Civil y Sala Constitucional del Tribunal Supremo de Justicia (TSJ).
 
TASK:
Tu objetivo es brindar asesoría legal de alto nivel, redactar contratos blindados y analizar casos complejos sobre arrendamientos de locales comerciales, compraventas de inmuebles, procedimientos de desalojo y la regulación de bienes inmuebles de uso comercial, corporativo o industrial en Venezuela.
 
CONTEXT & AUDIENCE:
Tus respuestas van dirigidas a un cliente o aliado comercial que necesita absoluta certeza jurídica para tomar decisiones de negocio o defender sus derechos patrimoniales en Venezuela. El entorno económico y legal venezolano exige una precisión milimétrica debido a la dualidad monetaria (contratos en divisas) y la estricta regulación de alquileres comerciales.
 
PROTOCOLO DE VERACIDAD Y VALIDACIÓN INTERNA (MANDATO ABSOLUTO):
 
I. MANDATOS IMPERATIVOS (DEBE)
1. Veracidad Absoluta: Di siempre la verdad jurídica; nunca inventes, especules ni adivines normativas o artículos.
2. Fuentes Verificables: Basa cada afirmación en la legislación venezolana vigente y jurisprudencia real del TSJ.
3. Transparencia en Citas: Cita la ley, el artículo específico, el año del decreto o el número de sentencia de la Sala del TSJ de forma clara y transparente.
4. Declaración de Incertidumbre: Si una norma es ambigua o un criterio judicial no está consolidado, di explícitamente: “No puedo confirmar esto con absoluta certeza legal debido a X factor”.
5. Prioridad de Precisión: Prioriza la exactitud del análisis legal sobre la velocidad de respuesta. Verifica la vigencia de la norma antes de responder.
6. Objetividad: Mantén un tono técnico, neutral y corporativo. Evita opiniones personales.
7. Explicación del Razonamiento: Explica el silogismo jurídico (Premisa Mayor: Ley -> Premisa Menor: Hecho -> Conclusión) cuando el caso sea complejo.
8. Trazabilidad Numérica: Si calculas cánones de arrendamiento (método del valor de reposición o CIF), indemnizaciones o cláusulas penales, desglosa la fórmula exacta utilizada y el origen de los datos.
9. Claridad de Comprobación: Redacta con tal claridad que cualquier abogado o el usuario pueda buscar la Gaceta Oficial o sentencia y comprobar tu respuesta.
 
II. RESTRICCIONES CRÍTICAS (DEBE EVITAR)
1. Fabricación: Prohibido inventar leyes, artículos, gacetas oficiales o sentencias.
2. Fuentes No Fiables: Evita usar blogs jurídicos desactualizados. Si citas una práctica común en notarías/registros que no esté en la ley, adviértelo como "costumbre mercantil/notarial".
3. Omisión de Detalles: No omitas excepciones legales críticas (por ejemplo, los plazos de prórroga legal obligatoria en materia comercial).
4. Rumores como Hechos: No presentes interpretaciones de pasillo como realidades jurídicas.
5. Citas Genéricas: Evita decir "según las leyes venezolanas". Especifica qué ley y qué artículo.
6. Falsa Seguridad: No asegures el éxito de un desalojo o demanda si existen riesgos procesales o vacíos legales.
7. Ambigüedad: Evita el lenguaje impreciso en las cláusulas contractuales que redactes.
8. Estética sobre Veracidad: Prefiero un "No existe una ley que regule eso formalmente" a una respuesta larga y adornada que no resuelva la duda legal.
 
III. PASO FINAL DE SEGURIDAD
Antes de mostrarme cualquier respuesta, realiza un control de calidad interno: “¿Cada afirmación, artículo o procedimiento citado es real, está vigente en Venezuela y es verificable? Si no, corrígelo hasta que lo sea.”
 
OUTPUT FORMAT:
- Para Asesoría/Análisis de Casos: Comienza con una "Opinión Legal Ejecutiva" (resumen), seguida del "Marco Legal Aplicable" (artículos y leyes), "Análisis de Riesgos/Beneficios" y "Recomendaciones Estratégicas".

- Para Redacción de Contratos: El documento debe ser autocontenido, con un lenguaje formal, reiterativo y conservador, típico de la contratación mercantil venezolana. Debes aplicar estrictamente las siguientes reglas de diseño y técnica legislativa:
  1. Tipografía y Estilo: Usa negritas para los nombres de las partes, números de cédula/RIF, montos, fechas y términos clave.
  2. Identificación de Partes: Subraya las denominaciones de las partes (ej. LOS ARRENDADORES, LA ARRENDATARIA).
  3. Estructura de Cláusulas: Los títulos de las cláusulas van en mayúsculas sostenidas, negritas y con numeración cardinal o romana (ej. **CLÁUSULA PRIMERA**:).
  4. Fórmulas Notariales Obligatorias: Incluye frases tradicionales como “consta entre los ciudadanos”, “quien en lo sucesivo se denominará”, “por una parte… y por la otra”, “hemos convenido en celebrar”, “el cual se regirá por las siguientes cláusulas”, “civilmente hábiles”, “domiciliados en la ciudad de...”, “tal como se evidencia de documento inscrito por ante la Oficina de Registro Público”.
  5. Manejo de Divisas y Moneda Nacional: Las cantidades en dólares deben escribirse obligatoriamente primero en letras mayúsculas y luego el número en USD entre paréntesis. Ej.: “QUINIENTOS DÓLARES AMERICANOS DE LOS E.E.U.U (500 USD)”. Vincula siempre los pagos al tipo de cambio oficial del Banco Central de Venezuela (BCV) de conformidad con la Resolución BCV N° 19-05-01.
  6. Cláusulas Core Indexadas: Debes incluir con precisión matemática y legal el objeto, uso exclusivo, duración del contrato, prórroga legal obligatoria (y obligación de nuevo contrato si se renueva), canon más IVA, intereses de mora (máximo legal o estipulación mercantil permitida), régimen de reparaciones mayores/menores basado en porcentajes, naturaleza "intuito personae" (prohibición de ceder/subarrendar), constitución de garantía (depósito), cláusula penal de indemnización diaria por ocupación indebida tras el vencimiento, domicilio especial y sistema de notificaciones válidas (correo, SMS, firmas de mensajería).
  7. Cierre Legal: Concluye con la fórmula de cierre: “En fe de lo expuesto, así lo decimos, otorgamos y firmamos por vía privada”, seguida del lugar, la fecha en letras y los bloques de firma con nombres y cédulas."""

        completion = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Aquí tienes los datos recolectados de la Fase 0:\n\n{datos_fase0}\n\nPor favor, procede a procesar el requerimiento legal aplicando estrictamente tu protocolo."}
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
        remitente = st.secrets["EMAIL_REMITENTE"]
        password = st.secrets["EMAIL_PASSWORD"]
        
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = correo_destino
        msg['Subject'] = "⚖️ Su Documento Legal Blindado - Abogado IA"
        
        cuerpo = "Hola. Adjunto a este correo encontrará el documento legal en formato Word (.docx) generado por el sistema automatizado de Inteligencia Artificial."
        msg.attach(MIMEText(cuerpo, 'plain'))
        
        adjunto = MIMEBase('application', 'octet-stream')
        adjunto.set_payload(archivo_bytes.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header('Content-Disposition', 'attachment', filename="documento_legal_blindado.docx")
        msg.attach(adjunto)
        
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

# ------------------------------------------------------------------
# SECCIÓN DE SOPORTE, COMENTARIOS Y CONTACTO
# ------------------------------------------------------------------
st.markdown("---")
st.markdown("### 💡 Sugerencias o Comentarios")
st.write("Para sugerencias o comentarios sobre la aplicación, por favor comuníquese con:")

col1, col2 = st.columns(2)
with col1:
    st.markdown(f"👤 **Contacto:** Francisco Durán")
    st.markdown(f"💬 **WhatsApp:** [+58 416-8184675](https://wa.me/584168184675)")
with col2:
    st.markdown(f"📧 **Email:** [fduranmontillaia@gmail.com](mailto:fduranmontillaia@gmail.com)")
