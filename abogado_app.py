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
st.title("⚖️ Franc - Asistente Legal Corporativo Experto")
st.write("Sistema inteligente dual para la redacción y análisis de documentos jurídicos blindados en Venezuela.")

# TEXTO CONSTANTE DEL AVISO LEGAL REQUERIDO
AVISO_LEGAL_TEXTO = (
    "⚠️ AVISO LEGAL IMPORTANTE: Este documento ha sido generado mediante Inteligencia Artificial "
    "y constituye únicamente un borrador o modelo de referencia. No sustituye el asesoramiento legal "
    "profesional. Debido a la constante actualización de normativas y criterios discrecionales del SAREN "
    "y tribunales venezolanos, este texto debe ser revisado, adaptado y visado obligatoriamente por un "
    "abogado colegiado (con Inpreabogado) antes de su firma, autenticación o presentación ante cualquier "
    "Notaría o Registro Público/Mercantil. La plataforma y sus desarrolladores no asumen ninguna "
    "responsabilidad por rechazos, daños, perjuicios o consecuencias legales derivadas del uso directo de este documento."
)

# ------------------------------------------------------------------
# SECCIÓN SECTORIAL: SELECCIÓN DEL EXPERTO IA
# ------------------------------------------------------------------
st.markdown("### 🗂️ Selección de la Especialidad Legal")
experto_seleccionado = st.selectbox(
    "Seleccione el perfil del Abogado Consultor que necesita:",
    [
        "Derecho Inmobiliario y Mercantil (Alquileres, Ventas, Desalojos)",
        "Derecho Mercantil Avanzado (Constitución de Sociedades Anónimas C.A.)"
    ]
)

st.markdown("---")

# ------------------------------------------------------------------
# FASE 0: DINÁMICA SEGÚN EXPERTO SELECCIONADO
# ------------------------------------------------------------------
st.subheader("📋 Fase 0: Datos del Requerimiento Legal")
st.write("Por favor, suministre los siguientes datos para iniciar (copie, pegue y rellene este formato):")

formato_inmueble = """- Tipo de documento/requerimiento: [Ej. Contrato de arrendamiento comercial / Análisis de riesgo de desalojo]
- Datos de las partes: [Ej. Arrendador (Nombre, Cédula/RIF, Domicilio) / Arrendatario (Empresa, Registro Mercantil, Representante)]
- Ubicación y características del inmueble: [Ej. Local comercial de 80mt2 en Mérida, Edo. Mérida. Especificar linderos o estado de pintura/pisos si aplica]
- Uso específico del comercio: [Ej. Carnicería, venta de repuestos, oficinas]
- Moneda y condiciones financieras: [Ej. Canon en USD mensuales, porcentaje para reparaciones menores, indemnización por ocupación indebida]
- Garantías y Pagos: [Ej. Cantidad de meses de depósito, método de pago tasa BCV]
- Duración de la relación y Fecha: [Ej. Contrato por 1 año, Fecha de inicio y fecha de firma]"""

formato_sociedades = """- Oficina de Registro Mercantil de Destino: [Ej. Registro Mercantil de la Circunscripción Judicial del Estado Mérida]
- Denominación Comercial solicitada: [Ej. Inversiones Alfa, C.A. / Corporación Beta, C.A. (Indicar si ya tiene reserva de nombre)]
- Objeto Social: [Ej. Explotación del ramo de restaurantes, compra y venta de víveres, importación de repuestos. Describir actividad principal y conexas]
- Domicilio Social: [Ej. Ciudad de Mérida, Estado Mérida, República Bolivariana de Venezuela]
- Capital Social y Acciones: [Ej. Capital de 500.000,00 bolívares, representado en 500 acciones de valor nominal de 1.000,00 bolívares cada una]
- Identificación de los Accionistas y su Suscripción: [Ej. Accionista A (Nombre, Cédula, 60% de acciones) y Accionista B (Nombre, Cédula, 40% de acciones)]
- Forma de Pago del Capital: [Ej. Mediante inventario de bienes muebles según balance de apertura / Depósito bancario en efectivo]
- Administración y Firma de la Compañía: [Ej. Junta Directiva (Presidente y Vicepresidente), firma conjunta/separada, duración de 5 años]
- Comisario Principal: [Ej. Nombre, Cédula, Licenciado en Contaduría o Administración, con su respectivo Nro. de CPC o CL]"""

if experto_seleccionado == "Derecho Inmobiliario y Mercantil (Alquileres, Ventas, Desalojos)":
    st.code(formato_inmueble, language="text")
else:
    st.code(formato_sociedades, language="text")

datos_usuario = st.text_area("Introduzca aquí los datos recolectados:", height=250)

# ------------------------------------------------------------------
# PROMPTS DE SISTEMA ASIGNADOS POR CASO
# ------------------------------------------------------------------
prompt_inmuebles = """FASE 0: RECOPILACIÓN REQUERIDA DE INFORMACIÓN (MANDATO INICIAL OBLIGATORIO)
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
8. Estética sobre Veracidad: Prefiero un
