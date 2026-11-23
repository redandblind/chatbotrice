# Asistente Virtual - Escuela G-416 Candelaria Pérez

Chatbot web basado en la API de Gemini que responde consultas sobre el **Reglamento de Evaluación, Promoción y Calificaciones (2026)** y el **Reglamento Interno de Convivencia Educativa - RICE (2026)**, guardando el historial de conversación de forma privada en el navegador (`localStorage`) de cada teléfono o computador.

## Archivos incluidos
- `main.py`: Servidor FastAPI que lee ambos reglamentos `.docx` y consulta a Gemini.
- `static/index.html`: Interfaz web responsiva con historial local por dispositivo.
- `requirements.txt`: Dependencias de Python para despliegue en la nube.
- `REGLAMENTO DE EVALUACIÓN, PROMOCIÓN Y CALIFICACIONES abril 2026 CON INDICE.docx`
- `RICE 2026 (ACTUALIZADO AGOSTO).docx`

## Prueba local (PowerShell)
```powershell
pip install -r requirements.txt
$env:GEMINI_API_KEY="TU_API_KEY_DE_GEMINI"
uvicorn main:app --host 0.0.0.0 --port 8000
```
Abre `http://localhost:8000` en tu navegador.

## Despliegue en Internet (Render.com)
1. Sube todos los archivos de esta carpeta a un repositorio de GitHub.
2. En Render.com crea un **New Web Service** conectado a tu repositorio.
3. Configura:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Environment Variable:** `GEMINI_API_KEY` = *(tu clave de Google AI Studio)*
