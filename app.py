
import os, json, io, zipfile, tempfile
from datetime import date
import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="AULA IA", page_icon="🧪", layout="wide")

NAVY = "17365D"
BLUE = "1F4E79"
LIGHT = "D9EAF7"
GRAY = "F2F2F2"

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, color=None, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(str(text or ""))
    r.bold = bold
    r.font.name = "Arial"
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def setup_doc(title, subtitle="Ciencia y Tecnología | Educación Secundaria"):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)
    styles["Normal"].paragraph_format.space_after = Pt(4)
    for stylename in ["Title", "Heading 1", "Heading 2"]:
        styles[stylename].font.name = "Arial"
    styles["Heading 1"].font.color.rgb = RGBColor.from_string(NAVY)
    styles["Heading 1"].font.size = Pt(14)
    styles["Heading 2"].font.color.rgb = RGBColor.from_string(BLUE)
    styles["Heading 2"].font.size = Pt(11)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("AULA IA")
    r.bold = True; r.font.name = "Arial"; r.font.size = Pt(12)
    r.font.color.rgb = RGBColor.from_string(BLUE)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title.upper())
    r.bold = True; r.font.name = "Arial"; r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string(NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle)
    r.italic = True; r.font.size = Pt(9)
    doc.add_paragraph()
    return doc

def add_bullets(doc, items):
    for item in items or []:
        doc.add_paragraph(str(item), style="List Bullet")

def add_numbered(doc, items):
    for item in items or []:
        doc.add_paragraph(str(item), style="List Number")

def add_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, True, "FFFFFF", 8)
        shade(table.rows[0].cells[i], NAVY)
    for row in rows or []:
        cells = table.add_row().cells
        for i, value in enumerate(row[:len(headers)]):
            set_cell_text(cells[i], value, size=8)
    doc.add_paragraph()
    return table

def add_section(doc, heading, content):
    doc.add_heading(heading, level=1)
    if isinstance(content, str):
        doc.add_paragraph(content)
    elif isinstance(content, list):
        add_bullets(doc, content)
    elif isinstance(content, dict):
        for k, v in content.items():
            doc.add_heading(str(k), level=2)
            if isinstance(v, list): add_bullets(doc, v)
            else: doc.add_paragraph(str(v))

def build_session(data, meta):
    doc = setup_doc(data.get("titulo", meta["tema"]))
    doc.add_heading("Datos informativos", level=1)
    add_table(doc, ["Campo", "Información"], [
        ["Grado", meta["grado"]], ["Área", "Ciencia y Tecnología"],
        ["Duración", meta["duracion"]], ["Código", meta["codigo"]],
        ["Fecha", str(date.today())]
    ])
    doc.add_heading("Propósito de aprendizaje", level=1)
    doc.add_paragraph(data.get("proposito", ""))
    doc.add_heading("Situación significativa", level=1)
    doc.add_paragraph(data.get("situacion_significativa", ""))
    doc.add_heading("Competencias, capacidades y desempeños", level=1)
    comp_rows = []
    for c in data.get("competencias", []):
        comp_rows.append([c.get("competencia",""), "\n".join(c.get("capacidades",[])), "\n".join(c.get("desempenos",[]))])
    add_table(doc, ["Competencia", "Capacidades", "Desempeños precisados"], comp_rows)
    doc.add_heading("Enfoques transversales y valores maristas", level=1)
    add_table(doc, ["Enfoque", "Valor marista", "Actitudes observables"], [
        [x.get("enfoque",""), x.get("valor_marista",""), x.get("actitudes","")]
        for x in data.get("enfoques_valores", [])
    ])
    doc.add_heading("Evaluación de los aprendizajes", level=1)
    add_table(doc, ["Criterios de evaluación", "Evidencias", "Instrumentos"], [
        [x.get("criterio",""), x.get("evidencia",""), x.get("instrumento","")]
        for x in data.get("evaluacion", [])
    ])
    doc.add_heading("Secuencia didáctica", level=1)
    for i, act in enumerate(data.get("actividades", []), 1):
        doc.add_heading(f"ACTIVIDAD {i}: {act.get('titulo','')}", level=2)
        add_table(doc, ["Momento", "Instrucciones / estrategias", "Recursos"], [
            [m.get("momento",""), m.get("instrucciones",""), "\n".join(m.get("recursos",[]))]
            for m in act.get("momentos", [])
        ])
    add_section(doc, "Atención a la diversidad y retroalimentación", data.get("atencion_diversidad", []))
    add_section(doc, "Metacognición y compromiso", data.get("metacognicion", []))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Documento generado con AULA IA. Revisar la correspondencia curricular antes de su aplicación.")
    r.italic = True; r.font.size = Pt(8); r.font.color.rgb = RGBColor(100,100,100)
    return doc

def build_worksheet(data, meta):
    doc = setup_doc("Ficha de trabajo: " + meta["tema"], f"{meta['grado']} | Ciencia y Tecnología")
    doc.add_paragraph(f"Nombre: ____________________________________    Fecha: ______________")
    doc.add_paragraph(f"Propósito: {data.get('proposito','')}")
    add_section(doc, "Situación o reto", data.get("situacion",""))
    add_section(doc, "Recordamos y comprendemos", data.get("lectura",""))
    doc.add_heading("Actividades", level=1)
    for i, q in enumerate(data.get("preguntas", []), 1):
        doc.add_paragraph(f"{i}. {q}")
        doc.add_paragraph("Respuesta: ______________________________________________________________")
        doc.add_paragraph("________________________________________________________________________")
    add_section(doc, "Reto de aplicación", data.get("reto",""))
    add_section(doc, "Autoevaluación", data.get("autoevaluacion", []))
    return doc

def build_instrument(title, headers, rows, meta):
    doc = setup_doc(title, f"{meta['grado']} | Ciencia y Tecnología | {meta['tema']}")
    doc.add_paragraph("Estudiante/equipo: __________________________________    Fecha: ______________")
    add_table(doc, headers, rows)
    return doc

def extract_json(text):
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n",1)[1].rsplit("```",1)[0]
    return json.loads(text)

def generate_with_ai(meta, curriculum_text, api_key, model):
    client = OpenAI(api_key=api_key)
    prompt = f"""
Eres especialista en planificación curricular peruana de Ciencia y Tecnología de secundaria.
Genera un paquete pedagógico completo en español, contextualizado al Callao y al contexto/valores maristas.
Datos: grado={meta['grado']}; tema={meta['tema']}; duración={meta['duracion']}; código={meta['codigo']}.
Orientaciones: aprendizaje por competencias; indagación científica; explicación científica; diseño de soluciones tecnológicas cuando corresponda; trabajo colaborativo; actividades prácticas y seguras; aula invertida/recursos digitales cuando aporten valor.
IMPORTANTE: no inventes citas ni afirmes que un desempeño es textual oficial si no aparece en el material curricular entregado. Si el material no contiene información suficiente, redacta desempeños precisados y márcalos como "precisados por el docente". Usa las tres competencias oficiales de CyT cuando sean pertinentes.
Valores maristas sugeridos: respeto, espíritu de familia, sencillez, presencia y solidaridad; enfoque de orientación al bien común y ambiental cuando corresponda.
Material curricular de referencia (si está disponible):
{curriculum_text[:18000]}
Devuelve SOLO JSON válido con estas claves:
{{
"titulo":"...",
"proposito":"...",
"situacion_significativa":"...",
"competencias":[{{"competencia":"nombre","capacidades":["..."],"desempenos":["..."]}}],
"enfoques_valores":[{{"enfoque":"...","valor_marista":"...","actitudes":"..."}}],
"evaluacion":[{{"criterio":"...","evidencia":"...","instrumento":"..."}}],
"actividades":[{{"titulo":"...","momentos":[{{"momento":"Inicio","instrucciones":"...","recursos":["..."]}},{{"momento":"Desarrollo","instrucciones":"...","recursos":["..."]}},{{"momento":"Cierre","instrucciones":"...","recursos":["..."]}}]}}],
"atencion_diversidad":["..."],
"metacognicion":["..."],
"ficha":{{"proposito":"...","situacion":"...","lectura":"...","preguntas":["..."],"reto":"...","autoevaluacion":["..."]}},
"lista_cotejo":{{"criterios":["..."]}},
"rubrica":{{"criterios":["..."],"niveles":["Inicio","En proceso","Logrado","Destacado"],"descriptores":[["...","...","...","..."]]}},
"practica_laboratorio":{{"titulo":"...","objetivo":"...","materiales":["..."],"seguridad":["..."],"procedimiento":["..."],"registro_datos":["..."],"analisis":["..."],"conclusion":"..."}},
"recursos_digitales":["..."],
"proyecto_reto":{{"problema":"...","producto":"...","etapas":["..."]}}
}}
Asegura que los ejercicios tengan respuestas científicamente coherentes y que los procedimientos experimentales sean seguros para estudiantes. No propongas manipular ácidos concentrados, sustancias tóxicas ni muestras ambientales desconocidas sin supervisión.
"""
    response = client.chat.completions.create(
        model=model,
        response_format={"type":"json_object"},
        messages=[
            {"role":"system","content":"Responde con un JSON válido, completo y estructurado. No incluyas Markdown."},
            {"role":"user","content":prompt}
        ],
        temperature=0.4
    )
    return extract_json(response.choices[0].message.content)

def make_package(data, meta):
    files = {}
    docs = {
        f"01_Sesion_{meta['codigo']}.docx": build_session(data, meta),
        f"02_Ficha_{meta['codigo']}.docx": build_worksheet(data.get("ficha",{}), meta),
        f"03_Lista_Cotejo_{meta['codigo']}.docx": build_instrument("Lista de cotejo", ["Criterio", "Sí", "En proceso", "No", "Observaciones"],
            [[x, "☐", "☐", "☐", ""] for x in data.get("lista_cotejo",{}).get("criterios",[])], meta),
        f"04_Rubrica_{meta['codigo']}.docx": build_instrument("Rúbrica analítica", ["Criterio", "Inicio", "En proceso", "Logrado", "Destacado"],
            [[crit] + (desc[:4] if len(desc)>=4 else desc+[""]*(4-len(desc))) for crit,desc in zip(data.get("rubrica",{}).get("criterios",[]), data.get("rubrica",{}).get("descriptores",[]))], meta),
        f"05_Practica_Laboratorio_{meta['codigo']}.docx": build_instrument("Práctica de laboratorio",
            ["Componente", "Desarrollo"], [[k, "\n".join(v) if isinstance(v,list) else str(v)] for k,v in data.get("practica_laboratorio",{}).items()], meta),
        f"06_Proyecto_Reto_{meta['codigo']}.docx": build_instrument("Proyecto o reto tecnológico", ["Componente", "Descripción"],
            [[k, "\n".join(v) if isinstance(v,list) else str(v)] for k,v in data.get("proyecto_reto",{}).items()], meta),
    }
    for name, doc in docs.items():
        b = io.BytesIO(); doc.save(b); files[name] = b.getvalue()
    # Resumen digital complementario
    summary = {
        "recursos_digitales": data.get("recursos_digitales", []),
        "nota_curricular": "Verifica los desempeños precisados con el Programa Curricular de Educación Secundaria vigente antes de publicar."
    }
    files[f"07_Recursos_y_notas_{meta['codigo']}.json"] = json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items(): archive.writestr(name, content)
    return z.getvalue(), files

st.markdown("<h1 style='color:#17365D;margin-bottom:0'>🧪 AULA IA</h1>", unsafe_allow_html=True)
st.markdown("**Generador integral de experiencias de aprendizaje | Ciencia y Tecnología – Secundaria**")
st.caption("Planificación contextualizada al Callao y al enfoque marista · Exportación editable a Word")
with st.sidebar:
    st.header("Configuración")
    # En Streamlit Community Cloud, configura OPENAI_API_KEY en App settings > Secrets.
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))
        model = st.secrets.get("OPENAI_MODEL", os.getenv("OPENAI_MODEL", "gpt-4.1-mini"))
    except Exception:
        api_key = os.getenv("OPENAI_API_KEY", "")
        model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    st.markdown("---")
    if api_key:
        st.success("Conexión con IA configurada.")
    else:
        st.warning("Falta configurar la clave de OpenAI en los secretos de Streamlit.")

c1, c2 = st.columns(2)
with c1:
    grado = st.selectbox("Grado", [f"{i}.º de secundaria" for i in range(1,6)], index=2)
    tema = st.text_input("Tema de aprendizaje", value="Ácidos oxácidos e hidrácidos")
    codigo = st.text_input("Código de sesión", value="B3_S2")
with c2:
    duracion = st.text_input("Duración estimada", value="2 sesiones de 90 minutos")
    contexto = st.selectbox("Contextualización", ["Callao – realidad ambiental e industrial", "Callao – comunidad y vida cotidiana", "Contexto nacional peruano", "Personalizar en el tema"])
    foco = st.multiselect("Énfasis metodológico", ["Indagación científica", "Aprendizaje basado en problemas", "Laboratorio y experimentación", "Diseño de soluciones tecnológicas", "Trabajo colaborativo", "Aula invertida", "Gamificación"], default=["Indagación científica","Laboratorio y experimentación","Trabajo colaborativo"])
st.markdown("### Referencia curricular")
curr_file = st.file_uploader("Carga el Programa Curricular MINEDU en PDF (opcional, recomendado para precisión curricular)", type=["pdf"])
curriculum_text = ""
if curr_file:
    try:
        from pypdf import PdfReader
        reader = PdfReader(curr_file)
        curriculum_text = "\n".join(page.extract_text() or "" for page in reader.pages)
        st.success(f"Referencia cargada: {len(curriculum_text):,} caracteres extraídos.")
    except Exception as e:
        st.warning(f"No se pudo leer el PDF: {e}")
st.markdown("### Material institucional")
st.caption("Se incluye como referencia el documento Word que compartiste. La salida utiliza una estructura editable basada en sus secciones.")
st.markdown("---")
if st.button("✨ Generar paquete pedagógico completo", type="primary", use_container_width=True):
    if not tema.strip():
        st.error("Escribe un tema de aprendizaje.")
    elif not api_key:
        st.error("El administrador debe configurar OPENAI_API_KEY en App settings > Secrets de Streamlit.")
    else:
        meta = {"grado":grado, "tema":tema.strip(), "codigo":codigo.strip() or "SESION", "duracion":duracion.strip()}
        meta["contexto"] = contexto
        meta["foco"] = foco
        with st.spinner("Diseñando la sesión, fichas, instrumentos y recursos..."):
            try:
                data = generate_with_ai(meta, curriculum_text, api_key, model)
                package, files = make_package(data, meta)
                st.session_state["package"] = package
                st.session_state["files"] = files
                st.session_state["data"] = data
                st.session_state["meta"] = meta
                st.success("¡Paquete pedagógico generado!")
            except Exception as e:
                st.error(f"No se pudo generar el paquete. Revisa tu API key, modelo y conexión. Detalle: {e}")

if "package" in st.session_state:
    st.markdown("## 📦 Materiales generados")
    st.write(f"**{st.session_state['meta']['grado']} · {st.session_state['meta']['tema']}**")
    for name, content in st.session_state["files"].items():
        st.write("📄 " + name)
        st.download_button("Descargar " + name, data=content, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document" if name.endswith(".docx") else "application/json",
                           key="dl_"+name)
    st.download_button("⬇️ Descargar paquete completo (.ZIP)", data=st.session_state["package"],
                       file_name=f"AULA_IA_{st.session_state['meta']['codigo']}.zip",
                       mime="application/zip", type="primary", use_container_width=True)
    with st.expander("Vista previa del contenido generado"):
        d = st.session_state["data"]
        st.write("**Propósito:**", d.get("proposito",""))
        st.write("**Situación significativa:**", d.get("situacion_significativa",""))
        st.write("**Actividades:**")
        for a in d.get("actividades",[]): st.write("•", a.get("titulo",""))

st.markdown("---")
st.caption("AULA IA – Prototipo inicial. El docente debe revisar la pertinencia curricular, los desempeños, las actividades y la seguridad de las experiencias antes de aplicarlas.")
