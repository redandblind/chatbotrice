import os
import re
import docx
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
from openai import OpenAI

# Pega aquí tu clave de Groq (empieza con gsk_...)
API_KEY = os.environ.get("GROQ_API_KEY", "").strip()

app = FastAPI(title="Asistente Virtual - Escuela Candelaria Pérez")

def cargar_fragmentos_docx(ruta_archivo: str, etiqueta: str) -> List[str]:
    """Divide el documento Word en bloques manejables conservando el contexto."""
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_archivo}")
    doc = docx.Document(ruta_archivo)
    lineas = []
    for p in doc.paragraphs:
        if p.text.strip():
            lineas.append(p.text.strip())
    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = [celda.text.strip().replace('\n', ' ') for celda in fila.cells]
            lineas.append(" | ".join(celdas))

    fragmentos = []
    bloque_actual = []
    largo_actual = 0
    for linea in lineas:
        bloque_actual.append(linea)
        largo_actual += len(linea)
        if largo_actual >= 1400:
            fragmentos.append(f"[{etiqueta}]\n" + "\n".join(bloque_actual))
            # Solapamiento de las últimas 2 líneas para no cortar artículos
            bloque_actual = bloque_actual[-2:]
            largo_actual = sum(len(x) for x in bloque_actual)
    if bloque_actual:
        fragmentos.append(f"[{etiqueta}]\n" + "\n".join(bloque_actual))
    return fragmentos

FRAGMENTOS_DOCS = (
    cargar_fragmentos_docx("REGLAMENTO DE EVALUACIÓN, PROMOCIÓN Y CALIFICACIONES abril 2026 CON INDICE.docx", "REGLAMENTO DE EVALUACIÓN")
    + cargar_fragmentos_docx("RICE 2026 (ACTUALIZADO AGOSTO).docx", "RICE CONVIVENCIA ESCOLAR")
)

STOPWORDS = {
    "como", "para", "pero", "porque", "sobre", "entre", "hasta", "desde",
    "donde", "cuando", "quien", "cual", "cuales", "este", "esta", "estos",
    "estas", "hacer", "debe", "debo", "puede", "puedo", "tiene", "tienen",
    "alumno", "alumnos", "estudiante", "estudiantes", "escuela", "reglamento"
}

def buscar_contexto_relevante(consulta: str, max_fragmentos: int = 8) -> str:
    """Selecciona los fragmentos más relevantes para mantenerse bajo el límite de 8000 tokens."""
    palabras = [
        p for p in re.findall(r"[a-záéíóúñ0-9]+", consulta.lower())
        if len(p) >= 3 and p not in STOPWORDS
    ]
    if not palabras:
        return "\n\n---\n\n".join(FRAGMENTOS_DOCS[:6])

    puntuados = []
    for idx, frag in enumerate(FRAGMENTOS_DOCS):
        texto_lower = frag.lower()
        puntaje = sum(texto_lower.count(p) * (2 if len(p) > 5 else 1) for p in palabras)
        if puntaje > 0:
            puntuados.append((puntaje, idx, frag))

    puntuados.sort(key=lambda x: x[0], reverse=True)
    seleccionados = puntuados[:max_fragmentos]

    if not seleccionados:
        return "\n\n---\n\n".join(FRAGMENTOS_DOCS[:6])

    # Ordena cronológicamente según su aparición en el documento
    seleccionados.sort(key=lambda x: x[1])
    return "\n\n---\n\n".join(frag for _, _, frag in seleccionados)

client = OpenAI(api_key=API_KEY, base_url="https://api.groq.com/openai/v1")
MODELO = "openai/gpt-oss-120b"

class Mensaje(BaseModel):
    role: str
    text: str

class SolicitudChat(BaseModel):
    mensaje: str
    historial: List[Mensaje] = []

@app.post("/api/chat")
async def procesar_chat(solicitud: SolicitudChat):
    if not API_KEY or API_KEY == "PEGA_AQUI_TU_CLAVE_GSK":
        raise HTTPException(status_code=500, detail="Falta pegar tu API Key en main.py.")
    try:
        contexto_filtrado = buscar_contexto_relevante(solicitud.mensaje)

        instrucciones = f"""
Eres el asistente virtual oficial de la Escuela G-416 Candelaria Pérez (RBD 3517-3, localidad de Alquihue, San Javier, Chile).

JERARQUÍA DE CONOCIMIENTO Y CRITERIO:
1. FUENTE PRINCIPAL (OBLIGATORIA): Responde priorizando y citando textualmente los artículos y protocolos entregados en los extractos oficiales más abajo (ej. "Según el Artículo 40 del Reglamento de Evaluación..." o "De acuerdo al Protocolo del RICE...").
2. CRITERIO COMPLEMENTARIO DE INTERNET / NORMATIVA CHILENA: Si el usuario consulta sobre leyes chilenas mencionadas (Decreto 67, Decreto 83, Ley Aula Segura, Ley Karin, Ley 21.801 sobre celulares, Seguro Escolar DS 313) o pide redactar una carta de justificación/apelación, puedes usar tu criterio general respetando siempre las reglas exactas de la Escuela Candelaria Pérez.
3. LÍMITE ESTRICTO: Nunca inventes sanciones, fechas, porcentajes ni normas internas que no aparezcan en los extractos.

=== EXTRACTOS RELEVANTES DE LOS REGLAMENTOS OFICIALES (2026) ===
{contexto_filtrado}
"""
        mensajes = [{"role": "system", "content": instrucciones}]

        # Conserva los últimos 4 mensajes del historial para no exceder el límite de tokens
        for msg in solicitud.historial[-4:]:
            rol = "assistant" if msg.role == "model" else msg.role
            mensajes.append({"role": rol, "content": msg.text})

        mensajes.append({"role": "user", "content": solicitud.mensaje})

        respuesta = client.chat.completions.create(
            model=MODELO,
            messages=mensajes,
            temperature=0.25,
        )
        return {"respuesta": respuesta.choices[0].message.content}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def inicio():
    return FileResponse("static/index.html")