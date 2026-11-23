import os
import re
import unicodedata
import docx
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict
from openai import OpenAI

API_KEY = os.environ.get("GROQ_API_KEY", "").strip()

app = FastAPI(title="Asistente Virtual - Escuela Candelaria Pérez")


def normalizar_texto(texto: str) -> str:
    """Elimina tildes y pasa a minúsculas para mejorar la coincidencia de búsqueda."""
    texto = texto.lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def extraer_bloques_estructurados(ruta_archivo: str, nombre_doc: str) -> List[Dict]:
    """
    Extrae párrafos y tablas respetando el orden exacto de aparición en el Word
    y agrupa el contenido por Artículos, Títulos, Protocolos y Anexos completos.
    """
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")

    doc = docx.Document(ruta_archivo)
    elementos = []

    # Recorre párrafos y tablas en el orden exacto en que aparecen en el documento
    for bloque in doc.element.body:
        tag = bloque.tag.split("}")[-1]
        if tag == "p":
            texto = "".join(bloque.itertext()).strip()
            if texto:
                elementos.append(("p", texto))
        elif tag == "tbl":
            filas_texto = []
            for tr in bloque.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"):
                celdas = []
                for tc in tr.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
                    celda_txt = " ".join("".join(tc.itertext()).split())
                    if celda_txt:
                        celdas.append(celda_txt)
                if celdas:
                    filas_texto.append(" | ".join(celdas))
            if filas_texto:
                elementos.append(("tbl", "\n".join(filas_texto)))

    # Patrones de inicio de sección estructural
    patron_inicio = re.compile(
        r"^(T[ÍI]TULO\s+[IVXLC]+|Art[íi]culo\s+\d+|PROTOCOLO\s+|ANEXO\s+\d+|\d+\)\s*PROTOCOLO|"
        r"[IVXLC]+\.\s+[A-ZÁÉÍÓÚÑ]{4,}|DEL\s+[A-ZÁÉÍÓÚÑ\s]{5,}|DE\s+LA[S]?\s+[A-ZÁÉÍÓÚÑ\s]{5,})",
        re.IGNORECASE
    )

    secciones = []
    titulo_actual = f"{nombre_doc} - Introducción / Disposiciones Iniciales"
    buffer_lineas = []
    indice_encontrado = False

    for tipo, texto in elementos:
        # Omitir las líneas punteadas del índice inicial para no contaminar búsquedas
        if re.search(r"\.{5,}\s*\d+$", texto):
            indice_encontrado = True
            continue

        es_encabezado = (tipo == "p" and len(texto) < 180 and patron_inicio.match(texto))

        if es_encabezado and buffer_lineas:
            contenido_seccion = "\n".join(buffer_lineas).strip()
            if len(contenido_seccion) > 40:
                secciones.append({
                    "doc": nombre_doc,
                    "titulo": titulo_actual,
                    "contenido": f"[{nombre_doc} | {titulo_actual}]\n{contenido_seccion}",
                    "norm": normalizar_texto(f"{titulo_actual} {contenido_seccion}")
                })
            titulo_actual = texto
            buffer_lineas = [texto]
        else:
            if es_encabezado and not buffer_lineas:
                titulo_actual = texto
            buffer_lineas.append(texto)

    if buffer_lineas:
        contenido_seccion = "\n".join(buffer_lineas).strip()
        if len(contenido_seccion) > 40:
            secciones.append({
                "doc": nombre_doc,
                "titulo": titulo_actual,
                "contenido": f"[{nombre_doc} | {titulo_actual}]\n{contenido_seccion}",
                "norm": normalizar_texto(f"{titulo_actual} {contenido_seccion}")
            })

    # Si alguna sección (ej. tabla larga de faltas o protocolo) supera 3200 caracteres,
    # se subdivide en partes conservando el encabezado para no perder ningún ítem
    secciones_finales = []
    for sec in secciones:
        txt = sec["contenido"]
        if len(txt) <= 3200:
            secciones_finales.append(sec)
        else:
            lineas = txt.split("\n")
            encabezado = lineas[0]
            parte = [encabezado]
            largo = len(encabezado)
            num_parte = 1
            for linea in lineas[1:]:
                if largo + len(linea) > 3000 and len(parte) > 1:
                    bloque_txt = "\n".join(parte)
                    secciones_finales.append({
                        "doc": sec["doc"],
                        "titulo": f"{sec['titulo']} (Parte {num_parte})",
                        "contenido": bloque_txt,
                        "norm": normalizar_texto(bloque_txt)
                    })
                    num_parte += 1
                    parte = [f"{encabezado} (Continuación Parte {num_parte})", linea]
                    largo = len(parte[0]) + len(linea)
                else:
                    parte.append(linea)
                    largo += len(linea)
            if len(parte) > 1:
                bloque_txt = "\n".join(parte)
                secciones_finales.append({
                    "doc": sec["doc"],
                    "titulo": f"{sec['titulo']} (Parte {num_parte})",
                    "contenido": bloque_txt,
                    "norm": normalizar_texto(bloque_txt)
                })

    return secciones_finales


SECCIONES_DOCS = (
    extraer_bloques_estructurados(
        "REGLAMENTO DE EVALUACIÓN, PROMOCIÓN Y CALIFICACIONES abril 2026 CON INDICE.docx",
        "REGLAMENTO DE EVALUACIÓN 2026"
    )
    + extraer_bloques_estructurados(
        "RICE 2026 (ACTUALIZADO AGOSTO).docx",
        "RICE CONVIVENCIA EDUCATIVA 2026"
    )
)

STOPWORDS = {
    "como", "para", "pero", "porque", "sobre", "entre", "hasta", "desde",
    "donde", "cuando", "quien", "cual", "cuales", "este", "esta", "estos",
    "estas", "hacer", "debe", "debo", "puede", "puedo", "tiene", "tienen",
    "alumno", "alumnos", "estudiante", "estudiantes", "escuela", "reglamento",
    "dice", "sobre", "acerca", "favor", "necesito", "saber", "cuales", "que"
}

SINONIMOS = {
    "pelea": ["agresion", "fisica", "maltrato", "violencia", "golpe"],
    "pelean": ["agresion", "fisica", "maltrato", "violencia", "golpe"],
    "golpe": ["agresion", "fisica", "maltrato", "lesion"],
    "celular": ["telefono", "movil", "electronico", "21801", "dispositivo"],
    "telefono": ["celular", "movil", "electronico", "21801", "dispositivo"],
    "nota": ["calificacion", "promedio", "evaluacion", "40", "70", "exigencia"],
    "notas": ["calificaciones", "promedio", "evaluacion", "semestral", "final"],
    "repitencia": ["promocion", "reprobado", "asistencia", "85", "riesgo"],
    "repetir": ["promocion", "repitencia", "reprobado", "85"],
    "falta": ["inasistencia", "justificar", "certificado", "medico", "48"],
    "falto": ["inasistencia", "justificar", "certificado", "medico", "48"],
    "copia": ["copiando", "adulterar", "evaluacion", "grave", "torpedo"],
    "copiando": ["copia", "adulterar", "evaluacion", "grave", "torpedo"],
    "accidente": ["seguro", "escolar", "313", "salud", "ambulancia", "declaracion"],
    "embarazo": ["embarazada", "maternidad", "lactancia", "prenatal", "postparto"],
    "embarazada": ["embarazo", "maternidad", "lactancia", "prenatal", "postparto"],
    "pie": ["adecuacion", "curricular", "paci", "dua", "necesidades", "83"],
    "tea": ["autista", "acompanamiento", "emocional", "conductual", "gatillantes"],
    "autismo": ["autista", "acompanamiento", "emocional", "conductual", "gatillantes"],
    "droga": ["alcohol", "sustancias", "vaper", "cigarro", "gravisima"],
    "alcohol": ["droga", "sustancias", "ebriedad", "gravisima"],
    "apelar": ["apelacion", "reconsideracion", "48", "horas", "sancion"],
    "expulsion": ["caducidad", "matricula", "condicionalidad", "superintendencia"],
    "bano": ["avd", "vida", "diaria", "higiene", "muda", "ropa"],
    "uniforme": ["vestuario", "buzo", "pantalon", "falda", "corbata", "presentacion"],
    "horario": ["jornada", "entrada", "salida", "0830", "1555", "1330", "atraso"]
}


def buscar_contexto_relevante(consulta: str, max_caracteres: int = 21000) -> str:
    """
    Recupera la mayor cantidad posible de artículos y protocolos exactos
    manteniendo el total dentro de la ventana segura de tokens de Groq.
    """
    consulta_norm = normalizar_texto(consulta)
    tokens_base = [
        p for p in re.findall(r"[a-z0-9]+", consulta_norm)
        if len(p) >= 2 and p not in STOPWORDS
    ]

    # Detectar si el usuario preguntó explícitamente por un número de artículo
    numeros_articulo = re.findall(r"articulo\s+(\d+)", consulta_norm)

    palabras_expandidas = set(tokens_base)
    for token in tokens_base:
        if token in SINONIMOS:
            palabras_expandidas.update(SINONIMOS[token])

    puntuados = []
    for idx, sec in enumerate(SECCIONES_DOCS):
        puntaje = 0
        norm_txt = sec["norm"]
        titulo_norm = normalizar_texto(sec["titulo"])

        # Coincidencia directa con número de artículo pedido
        for num in numeros_articulo:
            if f"articulo {num}" in titulo_norm or f"articulo {num}." in titulo_norm or f"articulo {num}:" in titulo_norm:
                puntaje += 100

        # Coincidencia de términos principales y sinónimos
        for palabra in palabras_expandidas:
            es_original = palabra in tokens_base
            peso = 4 if es_original else 2
            if len(palabra) >= 6:
                peso += 2

            apariciones_titulo = titulo_norm.count(palabra)
            apariciones_cuerpo = norm_txt.count(palabra)

            if apariciones_titulo > 0:
                puntaje += apariciones_titulo * (peso * 3)
            if apariciones_cuerpo > 0:
                puntaje += min(apariciones_cuerpo, 6) * peso

        if puntaje > 0:
            puntuados.append((puntaje, idx, sec["contenido"]))

    puntuados.sort(key=lambda x: x[0], reverse=True)

    seleccionados = []
    total_chars = 0
    for puntaje, idx, contenido in puntuados:
        if total_chars + len(contenido) <= max_caracteres:
            seleccionados.append((idx, contenido))
            total_chars += len(contenido)
        if len(seleccionados) >= 12:
            break

    if not seleccionados:
        # Si es una pregunta muy general, entrega disposiciones generales de ambos reglamentos
        fallback = [s["contenido"] for s in SECCIONES_DOCS[:6]]
        return "\n\n========================================\n\n".join(fallback)

    # Ordena los artículos según su orden real en el reglamento
    seleccionados.sort(key=lambda x: x[0])
    return "\n\n========================================\n\n".join(c for _, c in seleccionados)


def limpiar_formato_sin_asteriscos(texto: str) -> str:
    """Elimina cualquier asterisco o símbolo Markdown residual en la respuesta."""
    texto = re.sub(r"\*\*(.*?)\*\*", r"\1", texto)
    texto = re.sub(r"\*(.*?)\*", r"\1", texto)
    texto = texto.replace("*", "")
    texto = re.sub(r"^#{1,6}\s*", "", texto, flags=re.MULTILINE)
    return texto.strip()


client = OpenAI(api_key=API_KEY, base_url="https://api.groq.com/openai/v1") if API_KEY else None
MODELO = "openai/gpt-oss-120b"


class Mensaje(BaseModel):
    role: str
    text: str


class SolicitudChat(BaseModel):
    mensaje: str
    historial: List[Mensaje] = []


@app.post("/api/chat")
async def procesar_chat(solicitud: SolicitudChat):
    if not API_KEY or client is None:
        raise HTTPException(status_code=500, detail="Falta configurar la variable GROQ_API_KEY en el servidor.")
    try:
        contexto_filtrado = buscar_contexto_relevante(solicitud.mensaje)

        instrucciones = f"""
Eres el asistente virtual oficial de la Escuela G-416 Candelaria Pérez (RBD 3517-3, localidad de Alquihue, San Javier, Chile).

REGLAS ESTRICTAS DE FORMATO (OBLIGATORIAS):
- PROHIBIDO USAR ASTERISCOS (*) EN CUALQUIER PARTE DE LA RESPUESTA. No uses **negritas**, ni *cursivas*, ni viñetas con asterisco.
- PROHIBIDO USAR ALMOHADILLAS (#) O SÍMBOLOS MARKDOWN.
- Escribe en texto plano limpio, ordenado y muy fácil de leer. Para destacar títulos o secciones usa MAYÚSCULAS seguidas de dos puntos (ejemplo: ARTÍCULO 40 - INASISTENCIA SIN JUSTIFICACIÓN:). Para enumerar pasos o requisitos usa números (1., 2., 3.) o guiones simples (-).

REGLAS DE EXACTITUD Y COMPLETITUD (OBLIGATORIAS):
1. RESPUESTA EXHAUSTIVA Y DETALLADA: Nunca resumas ni recortes información importante. Entrega TODOS los detalles, pasos, plazos exactos (horas, días hábiles, fechas), porcentajes de exigencia o asistencia, responsables por cargo, sanciones específicas (con y sin sanción anterior), excepciones y documentos asociados que aparezcan en los extractos oficiales.
2. CITA LITERAL Y EXACTA: Siempre indica el documento exacto (Reglamento de Evaluación 2026 o RICE 2026), el número de Artículo, Título, Protocolo o Anexo correspondiente. Si el texto tiene tablas (como cantidad mínima de notas o faltas graves/gravísimas), incluye todos los datos aplicables al caso consultado.
3. NORMATIVA CHILENA COMPLEMENTARIA: Si la consulta involucra leyes mencionadas en los documentos (Decreto 67/2018, Decreto 83/2015, Decreto 170, Ley 21.801 sobre celulares, Ley Karin Decreto 2227, Ley Aula Segura, Seguro Escolar DS 313, Ordinario 05/2610 AVD) o pide redactar una carta formal de apelación o justificación, entrega la orientación completa respetando al pie de la letra los plazos y reglas internas de la Escuela Candelaria Pérez.
4. HONESTIDAD DOCUMENTAL: No inventes artículos, porcentajes ni normas que no existan en los documentos proporcionados.

=== BASE DOCUMENTAL OFICIAL DE LA ESCUELA CANDELARIA PÉREZ (2026) ===
{contexto_filtrado}
"""
        mensajes = [{"role": "system", "content": instrucciones}]

        for msg in solicitud.historial[-4:]:
            rol = "assistant" if msg.role == "model" else msg.role
            mensajes.append({"role": rol, "content": msg.text})

        mensajes.append({"role": "user", "content": solicitud.mensaje})

        respuesta = client.chat.completions.create(
            model=MODELO,
            messages=mensajes,
            temperature=0.15,
            max_tokens=2048,
        )
        texto_limpio = limpiar_formato_sin_asteriscos(respuesta.choices[0].message.content)
        return {"respuesta": texto_limpio}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def inicio():
    return FileResponse("static/index.html")