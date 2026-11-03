# AULA IA – versión para publicar en Streamlit Community Cloud

Aplicación web para generar materiales de Ciencia y Tecnología de 1.º a 5.º de secundaria.

## Archivos
- `app.py`: aplicación principal.
- `requirements.txt`: dependencias.
- `plantillas/plantilla_referencia.docx`: documento institucional de referencia.

## Publicación
1. Crea un repositorio en GitHub (puede llamarse `aula-ia`).
2. Sube el contenido de esta carpeta a la raíz del repositorio, incluyendo `app.py`, `requirements.txt` y la carpeta `plantillas`.
3. Entra a https://share.streamlit.io/ e inicia sesión con GitHub.
4. Crea una app nueva, selecciona el repositorio y la rama `main`; como archivo principal indica `app.py`.
5. Antes de generar materiales, abre la configuración de la app en Streamlit y añade en **Secrets**:

```toml
OPENAI_API_KEY = "PEGA_AQUI_TU_CLAVE_API"
OPENAI_MODEL = "gpt-4.1-mini"
```

6. Guarda los secretos y reinicia la app.

## Seguridad
- Nunca subas una clave API a GitHub ni la escribas en `app.py`.
- Configura la clave únicamente en Streamlit > App settings > Secrets.
- No publiques datos personales de estudiantes ni documentos confidenciales.

## Alcance
Esta versión reproduce la organización general de la plantilla institucional en documentos Word nuevos; no garantiza una copia exacta de todos los elementos gráficos o saltos de página del original. El docente debe revisar la precisión curricular antes de usar los materiales.
