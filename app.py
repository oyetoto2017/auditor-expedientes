import os
import re
import io
import json
import time
import hashlib
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from dateparser import parse
from PIL import Image
from google import genai
from google.genai import types

# =====================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# =====================================================================
st.set_page_config(
    page_title="Auditor Jurídico IA | Análisis de Expedientes y Contradicciones",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #0f172a 0%, #1e3a8a 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #f8fafc;
        border-left: 4px solid #1e3a8a;
        padding: 1rem;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# ESTADO DE SESIÓN (MEMORIA VOLÁTIL SEGURA)
# =====================================================================
if "expediente_docs" not in st.session_state:
    st.session_state.expediente_docs = []
if "actuaciones_df" not in st.session_state:
    st.session_state.actuaciones_df = pd.DataFrame()
if "catalogo_df" not in st.session_state:
    st.session_state.catalogo_df = pd.DataFrame()
if "informe_md" not in st.session_state:
    st.session_state.informe_md = ""
if "historial_chat" not in st.session_state:
    st.session_state.historial_chat = []

# =====================================================================
# EXPRESIONES REGULARES PARA METADATOS PROCESALES
# =====================================================================
RE_FECHA = re.compile(r'(\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b)')
RE_RADICADO = re.compile(
    r'(Radicado|Rad\.|Proceso|Expediente|No\.?\s*)\s*[:#\-]?\s*([0-9]{4,23}[-]?[0-9]{0,6})',
    re.I
)
RE_ACTOR = re.compile(
    r'(Demandante|Actor|Solicitante|Promovente|Accionante|Quejoso|Denunciante)\s*[:\-]\s*(.+)',
    re.I
)
RE_DEMANDADO = re.compile(
    r'(Demandado|Demandada|Convocado|Entidad|Accionado|Investigado|Denunciado)\s*[:\-]\s*(.+)',
    re.I
)
RE_JUEZ = re.compile(
    r'(Juez|Jueza|Magistrad[oa]|Despacho|Juzgado|Superintendencia|Procuradur[ií]a|Personer[ií]a)\s*[:\-]?\s*(.+)',
    re.I
)
RE_NUMERO_OFICIO = re.compile(
    r'(OFICIO|Resoluci[oó]n|Auto|Acta)\s+No\.?\s*([0-9A-Za-z\-]{1,15})',
    re.I
)

TIPOS_ACT = [
    'DEMANDA', 'QUEJA', 'DERECHO DE PETICION', 'CONTESTACION', 'REPLICA',
    'AUTO', 'PROVIDENCIA', 'RESOLUCION', 'SENTENCIA', 'APELACION', 'RECURSO',
    'NOTIFICACION', 'ACTA', 'OFICIO', 'INFORME', 'PERITAJE', 'LIQUIDACION',
    'AUDIENCIA', 'MEMORIAL', 'PRUEBA', 'CONTRATO', 'PROMESA'
]

# =====================================================================
# CLIENTE GEMINI CON AUTO-DETECCIÓN DE MODELOS ACTIVOS (GEMINI 3 / LATEST)
# =====================================================================
MODELOS_RESPALDO = [
    "gemini-3-flash-preview",
    "gemini-3.0-flash",
    "gemini-3-flash",
    "gemini-3.1-flash-preview",
    "gemini-3.1-flash",
    "gemini-3.1-flash-lite-preview",
    "gemini-flash-latest",
    "gemini-3-pro-preview",
    "gemini-3.0-pro",
    "gemini-3-pro",
    "gemini-3.1-pro-preview",
    "gemini-pro-latest",
    "gemini-3.8-flash",
    "gemini-3.6-flash",
    "gemini-2.5-flash",
    "gemini-2.5-pro"
]

def obtener_cliente_gemini(api_key: str):
    if not api_key:
        return None
    return genai.Client(api_key=api_key.strip())

def descubrir_modelos_activos(client):
    """
    Consulta directamente a Google qué modelos Gemini están habilitados
    para la API Key del usuario en este momento, evitando errores 404.
    """
    modelos_cuenta = []
    try:
        for m in client.models.list():
            nombre = getattr(m, "name", "")
            if not nombre:
                continue
            limpio = nombre.replace("models/", "")
            low = limpio.lower()
            if "gemini" not in low:
                continue
            excluidos = [
                "embedding", "imagen", "veo", "tts", "aqa",
                "audio", "live", "robotics", "computer-use", "learnlm"
            ]
            if any(x in low for x in excluidos):
                continue
            acciones = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None)
            if acciones and not any("generatecontent" in str(a).lower() for a in acciones):
                continue
            modelos_cuenta.append(limpio)
    except Exception:
        pass

    # Priorizar modelos Gemini 3 y Flash/Pro
    def puntaje_modelo(mod_name):
        n = mod_name.lower()
        score = 0
        if "gemini-3" in n:
            score += 100
        elif "latest" in n:
            score += 80
        elif "gemini-2.5" in n:
            score += 50
        if "flash" in n and "lite" not in n:
            score += 20
        elif "pro" in n:
            score += 15
        return -score

    modelos_cuenta_ordenados = sorted(list(dict.fromkeys(modelos_cuenta)), key=puntaje_modelo)
    # Unir los detectados en la cuenta con la lista de respaldo
    cola_final = modelos_cuenta_ordenados + [m for m in MODELOS_RESPALDO if m not in modelos_cuenta_ordenados]
    return cola_final

def llamar_gemini(client, prompt_contents, modelo_preferido="Auto-Detectar (Recomendado)", temperatura=0.1, max_reintentos=2):
    """Llama a Gemini detectando automáticamente el modelo activo en la API Key del usuario."""
    cola_modelos = descubrir_modelos_activos(client)
    if modelo_preferido and modelo_preferido != "Auto-Detectar (Recomendado)":
        cola_modelos = [modelo_preferido] + [m for m in cola_modelos if m != modelo_preferido]

    errores = []

    for modelo in cola_modelos:
        for intento in range(max_reintentos):
            try:
                resp = client.models.generate_content(
                    model=modelo,
                    contents=prompt_contents,
                    config=types.GenerateContentConfig(
                        temperature=temperatura,
                    )
                )
                if resp and resp.text:
                    return resp.text, modelo
            except Exception as e:
                msg = str(e).lower()
                errores.append(f"{modelo}: {str(e)[:120]}")
                # Si el modelo no existe o no está disponible para usuarios nuevos, saltar de inmediato al siguiente
                if "404" in msg or "not found" in msg or "no longer available" in msg or "not supported" in msg:
                    break
                # Si es límite de cuota por minuto (429), esperar un poco o probar otro modelo de la cola
                if "429" in msg or "quota" in msg or "resource_exhausted" in msg:
                    if intento == 0:
                        time.sleep(4)
                        continue
                    else:
                        break
                time.sleep(2)

    raise RuntimeError("No fue posible obtener respuesta. Últimos intentos: " + " | ".join(errores[-4:]))

# =====================================================================
# EXTRACCIÓN DOCUMENTAL AVANZADA + OCR HÍBRIDO (TESSERACT / GEMINI VISION)
# =====================================================================
def ocr_pagina_hibrido(page, client=None):
    """
    Intenta extraer texto de una página escaneada usando Tesseract local.
    Si Tesseract no está disponible, usa Gemini Vision con auto-detección de modelo.
    """
    pix = page.get_pixmap(dpi=200)
    img_bytes = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_bytes))

    # 1. Intento con Tesseract local
    try:
        import pytesseract
        texto_ocr = pytesseract.image_to_string(img, lang="spa").strip()
        if len(texto_ocr) > 25:
            return texto_ocr, "OCR-Tesseract"
    except Exception:
        pass

    # 2. Respaldo con Visión Nativa de Gemini
    if client is not None:
        try:
            texto_ia, _ = llamar_gemini(
                client=client,
                prompt_contents=[
                    img,
                    "Transcribe de manera exacta y literal todo el texto visible en esta página de documento jurídico/administrativo, incluyendo sellos, fechas, radicados y tablas. No agregues comentarios."
                ],
                temperatura=0.0,
                max_reintentos=1
            )
            if texto_ia:
                return texto_ia.strip(), "OCR-GeminiVision"
        except Exception:
            pass

    return "", "SIN_TEXTO"

def extraer_metadatos_documento(texto: str):
    meta = {
        'fechas': [],
        'radicado': 'No identificado',
        'actor': 'No identificado',
        'demandado': 'No identificado',
        'juez': 'No identificado',
        'tipo_actuacion': [],
        'numero_oficio': 'No identificado'
    }
    meta['fechas'] = list(dict.fromkeys(RE_FECHA.findall(texto)))[:15]

    rr = RE_RADICADO.search(texto)
    if rr:
        meta['radicado'] = rr.group(0).strip()[:80]

    no_of = RE_NUMERO_OFICIO.search(texto)
    if no_of:
        meta['numero_oficio'] = no_of.group(0).strip()[:60]

    da = RE_ACTOR.search(texto)
    if da:
        meta['actor'] = da.group(2).splitlines()[0].strip()[:120]

    dd = RE_DEMANDADO.search(texto)
    if dd:
        meta['demandado'] = dd.group(2).splitlines()[0].strip()[:120]

    jj = RE_JUEZ.search(texto)
    if jj:
        meta['juez'] = jj.group(0).splitlines()[0].strip()[:120]

    texto_upper = texto.upper()
    for tipo in TIPOS_ACT:
        if tipo in texto_upper:
            meta['tipo_actuacion'].append(tipo)

    return meta

def procesar_archivo_subido(uploaded_file, client=None, usar_ocr_ia=True):
    """Lee PDF, DOCX o TXT conservando trazabilidad exacta por página."""
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    hash_sha256 = hashlib.sha256(file_bytes).hexdigest()[:16]
    nombre = uploaded_file.name
    ext = Path(nombre).suffix.lower()

    paginas_estructuradas = []
    paginas_ocr = []
    paginas_vacias = []

    try:
        if ext == ".pdf":
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            total_pags = len(doc)
            for i in range(total_pags):
                page = doc[i]
                num_pag = i + 1
                txt = page.get_text("text").strip()

                if len(txt) >= 35:
                    paginas_estructuradas.append({
                        "pagina": num_pag,
                        "origen": "Digital",
                        "texto": f"[[ARCHIVO: {nombre} | PÁGINA {num_pag}]]\n{txt}"
                    })
                else:
                    cliente_ocr = client if usar_ocr_ia else None
                    txt_ocr, metodo_ocr = ocr_pagina_hibrido(page, cliente_ocr)
                    if txt_ocr:
                        paginas_ocr.append(num_pag)
                        paginas_estructuradas.append({
                            "pagina": num_pag,
                            "origen": metodo_ocr,
                            "texto": f"[[ARCHIVO: {nombre} | PÁGINA {num_pag} ({metodo_ocr})]]\n{txt_ocr}"
                        })
                    else:
                        paginas_vacias.append(num_pag)
                        paginas_estructuradas.append({
                            "pagina": num_pag,
                            "origen": "Sin Texto",
                            "texto": f"[[ARCHIVO: {nombre} | PÁGINA {num_pag} - PÁGINA ESCANEADA SIN TEXTO LEGIBLE]]"
                        })
            doc.close()

        elif ext == ".docx":
            doc = Document(io.BytesIO(file_bytes))
            parrafos = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parrafos.append(p.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    celdas = [c.text.strip().replace("\n", " ") for c in row.cells]
                    if any(celdas):
                        parrafos.append(" | ".join(celdas))

            texto_completo_docx = "\n".join(parrafos)
            tam_pag = 3000
            for idx_b, inicio in enumerate(range(0, max(1, len(texto_completo_docx)), tam_pag)):
                fragmento = texto_completo_docx[inicio:inicio + tam_pag]
                paginas_estructuradas.append({
                    "pagina": idx_b + 1,
                    "origen": "Word Nativo",
                    "texto": f"[[ARCHIVO: {nombre} | SECCIÓN/PÁGINA {idx_b + 1}]]\n{fragmento}"
                })

        elif ext in (".txt", ".rtf"):
            contenido = file_bytes.decode("utf-8", errors="ignore").strip()
            paginas_estructuradas.append({
                "pagina": 1,
                "origen": "Texto Plano",
                "texto": f"[[ARCHIVO: {nombre} | PÁGINA 1]]\n{contenido}"
            })

    except Exception as e:
        paginas_estructuradas.append({
            "pagina": 1,
            "origen": "Error",
            "texto": f"[[ARCHIVO: {nombre} | ERROR DE EXTRACCIÓN: {str(e)}]]"
        })

    texto_unificado = "\n\n".join([p["texto"] for p in paginas_estructuradas])
    meta = extraer_metadatos_documento(texto_unificado)

    return {
        "nombre": nombre,
        "hash": hash_sha256,
        "num_paginas": len(paginas_estructuradas),
        "caracteres": len(texto_unificado),
        "paginas_ocr": paginas_ocr,
        "paginas_vacias": paginas_vacias,
        "meta": meta,
        "paginas": paginas_estructuradas,
        "raw": texto_unificado
    }

def extraer_linea_tiempo_regex(items):
    """Detecta actuaciones y fechas en los documentos para construir la tabla cronológica base."""
    patron_bloque = re.compile(
        r'(?m)^(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s*(?:[-–:]\s*)?(?P<encabezado>[\w\s\(\)\.,#º°\-]{8,110})$'
    )
    patron_acto = re.compile(
        r'(?i)(auto|providencia|resoluci[oó]n|sentencia|acta|oficio|notificaci[oó]n|audiencia|memorial|radicado|queja|recurso)\s+(?:no\.?|n[o°]?)?\s*[\w\-]*\s*(?:del|de fecha)?\s*(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})'
    )

    actuaciones = []
    for it in items:
        for pag in it["paginas"]:
            txt = pag["texto"]
            num_pag = pag["pagina"]

            for m in patron_bloque.finditer(txt):
                actuaciones.append({
                    "Archivo": it["nombre"],
                    "Página": num_pag,
                    "Fecha Literal": m.group("fecha").strip(),
                    "Actuación / Referencia": m.group("encabezado").strip(),
                    "Contexto": txt[max(0, m.start() - 30):min(len(txt), m.end() + 140)].replace("\n", " ")
                })

            for m in patron_acto.finditer(txt):
                actuaciones.append({
                    "Archivo": it["nombre"],
                    "Página": num_pag,
                    "Fecha Literal": m.group("fecha").strip(),
                    "Actuación / Referencia": m.group(0).strip(),
                    "Contexto": txt[max(0, m.start() - 30):min(len(txt), m.end() + 140)].replace("\n", " ")
                })

    if not actuaciones:
        return pd.DataFrame(columns=["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"])

    def normalizar_fecha(f_str):
        try:
            d = parse(str(f_str), languages=["es"], settings={"DATE_ORDER": "DMY"})
            if d and 1980 <= d.year <= 2035:
                return d.date().isoformat()
        except Exception:
            pass
        return "9999-12-31"

    df_act = pd.DataFrame(actuaciones)
    df_act["Fecha ISO"] = df_act["Fecha Literal"].apply(normalizar_fecha)
    df_act = df_act.drop_duplicates(subset=["Archivo", "Página", "Actuación / Referencia"])
    df_act = df_act.sort_values(by=["Fecha ISO", "Archivo", "Página"]).reset_index(drop=True)
    df_act["Fecha ISO"] = df_act["Fecha ISO"].replace("9999-12-31", "Sin Fecha Exacta")

    columnas_orden = ["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"]
    return df_act[columnas_orden]

# =====================================================================
# GENERADOR PROFESIONAL DE WORD (.DOCX) CON TABLAS Y NEGRILLAS REALES
# =====================================================================
def aplicar_color_celda(cell, hex_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def agregar_texto_con_negrillas(parrafo, texto):
    partes = re.split(r'(\*\*.*?\*\*)', texto)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**') and len(parte) > 4:
            run = parrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            parrafo.add_run(parte)

def volcar_tabla_en_word(doc, filas_tabla):
    if not filas_tabla:
        return
    num_cols = max(len(f) for f in filas_tabla)
    if num_cols == 0:
        return

    tabla = doc.add_table(rows=len(filas_tabla), cols=num_cols)
    tabla.style = 'Table Grid'

    for r_idx, fila in enumerate(filas_tabla):
        for c_idx in range(num_cols):
            texto_celda = fila[c_idx] if c_idx < len(fila) else ""
            celda = tabla.cell(r_idx, c_idx)
            celda.text = ""
            p = celda.paragraphs[0]
            agregar_texto_con_negrillas(p, texto_celda)

            if r_idx == 0:
                aplicar_color_celda(celda, "1E3A8A")
                for run in p.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9.5)
            else:
                if r_idx % 2 == 1:
                    aplicar_color_celda(celda, "F8FAFC")
                for run in p.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph()

def convertir_markdown_a_docx_bytes(md_texto: str, titulo_portada="INFORME DE AUDITORÍA Y ANÁLISIS DE EXPEDIENTE"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    titulo = doc.add_heading(titulo_portada, level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub.add_run(f"Generado el {datetime.now().strftime('%d/%m/%Y a las %H:%M')} | Protocolo de Verificación y Trazabilidad Documental")
    run_sub.italic = True
    run_sub.font.size = Pt(9.5)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)
    doc.add_paragraph("_" * 72)

    lineas = md_texto.split('\n')
    buffer_tabla = []

    for linea in lineas:
        linea_limpia = linea.strip()
        es_fila_tabla = linea_limpia.startswith('|') and linea_limpia.endswith('|') and len(linea_limpia) > 2

        if es_fila_tabla:
            contenido_sin_barras = linea_limpia.replace('|', '').replace('-', '').replace(':', '').strip()
            if not contenido_sin_barras:
                continue
            celdas = [c.strip() for c in linea_limpia.split('|')[1:-1]]
            buffer_tabla.append(celdas)
            continue
        else:
            if buffer_tabla:
                volcar_tabla_en_word(doc, buffer_tabla)
                buffer_tabla = []

        if not linea_limpia or linea_limpia == "---":
            continue

        if linea_limpia.startswith('# '):
            doc.add_heading(linea_limpia[2:].replace('**', ''), level=1)
        elif linea_limpia.startswith('## '):
            doc.add_heading(linea_limpia[3:].replace('**', ''), level=2)
        elif linea_limpia.startswith('### '):
            doc.add_heading(linea_limpia[4:].replace('**', ''), level=3)
        elif linea_limpia.startswith('#### '):
            doc.add_heading(linea_limpia[5:].replace('**', ''), level=4)
        elif linea_limpia.startswith(('- ', '* ')):
            p = doc.add_paragraph(style='List Bullet')
            agregar_texto_con_negrillas(p, linea_limpia[2:])
        elif re.match(r'^\d+\.\s+', linea_limpia):
            texto_item = re.sub(r'^\d+\.\s+', '', linea_limpia)
            p = doc.add_paragraph(style='List Number')
            agregar_texto_con_negrillas(p, texto_item)
        elif linea_limpia.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.4)
            agregar_texto_con_negrillas(p, linea_limpia[2:])
            for run in p.runs:
                run.italic = True
        else:
            p = doc.add_paragraph()
            agregar_texto_con_negrillas(p, linea_limpia)

    if buffer_tabla:
        volcar_tabla_en_word(doc, buffer_tabla)

    buffer_word = io.BytesIO()
    doc.save(buffer_word)
    buffer_word.seek(0)
    return buffer_word

# =====================================================================
# CONSTRUCCIÓN DEL PROMPT DE AUDITORÍA PROFUNDA
# =====================================================================
def construir_prompt_auditoria(
    items,
    catalogo_df,
    actuaciones_df,
    tipo_proceso,
    rol_usuario,
    enfoque_especial,
    limite_chars_por_doc=120000
):
    total_docs = len(items)
    total_pags = sum(it["num_paginas"] for it in items)
    total_ocr = sum(len(it["paginas_ocr"]) for it in items)

    catalogo_md = catalogo_df.to_markdown(index=False) if not catalogo_df.empty else "Sin catálogo"
    actos_md = actuaciones_df.head(80).to_markdown(index=False) if not actuaciones_df.empty else "Sin actuaciones previas detectadas por regex"

    bloques_documentales = []
    for idx, it in enumerate(items, 1):
        texto_doc = it["raw"]
        nota_recorte = ""
        if len(texto_doc) > limite_chars_por_doc:
            mitad = limite_chars_por_doc // 2
            texto_doc = (
                texto_doc[:mitad]
                + f"\n\n[... SECCIÓN INTERMEDIA OMITIDA POR LONGITUD EXTREMA ({len(it['raw'])} chars totales) ...]\n\n"
                + texto_doc[-mitad:]
            )
            nota_recorte = f" (Documento extenso optimizado a {limite_chars_por_doc} caracteres)"

        bloque = f"""
=====================================================================
DOCUMENTO #{idx}: {it['nombre']}{nota_recorte}
- Hash SHA-256: {it['hash']}
- Páginas totales: {it['num_paginas']} (Páginas con OCR: {len(it['paginas_ocr'])})
- Radicado detectado: {it['meta']['radicado']}
- Tipo de actuación: {', '.join(it['meta']['tipo_actuacion']) or 'General'}
- Actor / Solicitante: {it['meta']['actor']}
- Demandado / Entidad: {it['meta']['demandado']}
=====================================================================
{texto_doc}
"""
        bloques_documentales.append(bloque)

    expediente_completo_str = "\n\n".join(bloques_documentales)

    prompt = f"""
Actúa como un **Abogado Litigante, Auditor Forense de Expedientes y Estratega Procesal Colombiano** de máximo nivel.

CONTEXTO ESTRATÉGICO DEL ENCARGO:
- **Naturaleza del Proceso / Actuación:** {tipo_proceso}
- **Posición / Rol que representamos:** {rol_usuario}
- **Instrucciones o puntos críticos a verificar con lupa:** {enfoque_especial or 'Detectar todas las contradicciones fácticas, cronológicas, probatorias y normativas entre los PDFs, y definir con exactitud qué se debe hacer a continuación.'}

PROTOCOLO OBLIGATORIO DE VERIFICACIÓN Y ANTI-ALUCINACIÓN:
1. **Inventario Real:** Analizas {total_docs} documentos que suman {total_pags} páginas ({total_ocr} procesadas mediante OCR).
2. **Trazabilidad Estricta por Página:** Cada hallazgo, hecho o contradicción DEBE citar explícitamente `[Archivo | Página X]` usando las marcas `[[ARCHIVO: ... | PÁGINA X]]` del texto entregado y transcribir entre comillas el fragmento probatorio clave.
3. **Cero Invenciones:** No inventes fechas, radicados, leyes ni nombres. Si un dato indispensable no aparece en el expediente, escribe en mayúsculas: `"NO CONSTA EN LOS DOCUMENTOS APORTADOS (VACÍO PROBATORIO)"`.
4. **Cruce Inter-Documental Exhaustivo:** Compara lo que se afirma en demandas, quejas, incidentes o contestaciones contra lo que dicen los oficios, resoluciones, autos, historias clínicas, contratos o pruebas en otras páginas o archivos.

INSUMO 1: CATÁLOGO GENERAL DE ARCHIVOS
{catalogo_md}

INSUMO 2: REGISTRO PRELIMINAR DE FECHAS Y ACTUACIONES
{actos_md}

INSUMO 3: CONTENIDO ÍNTEGRO PAGINADO DE LOS DOCUMENTOS DEL EXPEDIENTE
{expediente_completo_str}

---

GENERA UN INFORME JURÍDICO Y TÉCNICO EXHAUSTIVO EN FORMATO MARKDOWN CON LAS SIGUIENTES 11 SECCIONES OBLIGATORIAS:

# INFORME DE AUDITORÍA PROCESAL, CONTRADICCIONES Y PLAN DE ACCIÓN

## 1. FICHA TÉCNICA E IDENTIFICACIÓN DEL EXPEDIENTE
(Tabla Markdown con: Radicado/Expediente, Entidad/Despacho, Demandante/Accionante/Incidentante, Demandado/Convocado/Accionado, Tipo de Proceso, Estado Procesal Actual y Nivel de Riesgo/Viabilidad).

## 2. RESUMEN EJECUTIVO DEL CASO
(Síntesis clara del conflicto, qué se reclama, qué ha respondido la contraparte o el juzgado/entidad y cuál es el cuello de botella actual del expediente).

## 3. LÍNEA DE TIEMPO CRONOLÓGICA CON TRAZABILIDAD
(Tabla Markdown ordenada cronológicamente con columnas: `Fecha` | `Actuación / Documento` | `Archivo y Página` | `Resumen del Hecho o Decisión` | `Impacto Procesal`).

## 4. MATRIZ EXHAUSTIVA DE CONTRADICCIONES E INCONSISTENCIAS DETECTADAS
(¡SECCIÓN CRÍTICA! Revisa cada PDF y cruza las versiones. Presenta una Tabla Markdown detallada con columnas: `N°` | `Tema / Punto en Conflicto` | `Lo que afirma el Documento A (Cita + Archivo y Página)` | `Lo que demuestra el Documento B o la Prueba (Cita + Archivo y Página)` | `Tipo de Contradicción (Fáctica, Cronológica, Técnica o Jurídica)` | `Gravedad y Cómo Explotarla a Nuestro Favor`).
Además de la tabla, desarrolla en viñetas el análisis detallado de cada contradicción encontrada.

## 5. AUDITORÍA DE PRUEBAS: HECHOS ACREDITADOS VS. AFIRMACIONES SIN SOPORTE
(Separa en dos subsecciones o tabla comparativa:
- **Hechos plenamente probados** indicando con qué documento y página se prueban.
- **Alegaciones o excusas sin sustento probatorio** o documentos mencionados en los escritos pero que NO obran o están incompletos en los archivos subidos).

## 6. CONTROL DE ÓRDENES JUDICIALES, REQUERIMIENTOS Y OBLIGACIONES INCUMPLIDAS
(Identifica sentencias, autos, requerimientos o compromisos citados en el expediente: qué se ordenó, quién debía cumplirlo, en qué plazo y si existe o no prueba real de cumplimiento en los PDFs).

## 7. ANÁLISIS JURÍDICO Y NORMATIVO APLICABLE (COLOMBIA)
(Fundamentos constitucionales, legales, procesales y jurisprudenciales aplicables al caso según su naturaleza: Acción de Tutela, Decreto 2591 de 1991, Incidente de Desacato, CGP, CPACA, Ley 1480 de 2011, Ley 1755 de 2015, etc., explicando cómo respaldan nuestra postura).

## 8. QUÉ SE DEBE HACER: HOJA DE RUTA Y PLAN DE ACCIÓN PASO A PASO
(¡SECCIÓN CRÍTICA! Explica con total claridad y orden secuencial **qué se debe hacer ahora mismo**:
1. **Acciones Inmediatas (0 a 5 días):** Memoriales, recursos, solicitudes de sanción/desacato, tutela o pruebas sobrevinientes que se deben radicar.
2. **Pruebas adicionales que se deben anexar o solicitar por oficio.**
3. **Estrategia de bloqueo frente a las excusas o argumentos de la contraparte/entidad.**)

## 9. TABLA DE PLAN DE ACCIÓN PRIORIZADO
(Tabla Markdown con columnas: `Prioridad` | `Acción Concreta a Realizar` | `Ante Quién / Destinatario` | `Objetivo Jurídico` | `Documento o Prueba Base a Invocar`).

## 10. BORRADOR DE MEMORIAL / OFICIO / PRONUNCIAMIENTO LISTO PARA RADICAR
(Redacta el borrador formal completo —con encabezado, referencia, radicado, hechos/contradicciones denunciadas, fundamentos de derecho y peticiones concretas— listo para copiar, ajustar datos finales y radicar en nombre de **{rol_usuario}**).

## 11. CHECKLIST DE VERIFICACIÓN ANTES DE RADICAR
(Lista de chequeo práctica de anexos, firmas, términos y validaciones que el usuario debe revisar antes de enviar el documento).
"""
    return prompt

# =====================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# =====================================================================
st.markdown("""
<div class="main-header">
    <h2 style="margin:0;">⚖️ Sistema Inteligente de Auditoría de Expedientes y Demandas</h2>
    <p style="margin:0.3rem 0 0 0; opacity:0.9;">
        Detección profunda de contradicciones entre PDFs, trazabilidad por folio, OCR híbrido y generación de informes ejecutivos en Word (.docx)
    </p>
</div>
""", unsafe_allow_html=True)

# ----------------- BARRA LATERAL (CONFIGURACIÓN) -----------------
with st.sidebar:
    st.header("🔑 Configuración del Motor IA")
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=os.environ.get("GOOGLE_API_KEY", ""),
        type="password",
        help="Ingresa tu clave de Google AI Studio. Se mantiene únicamente en la memoria temporal de tu sesión."
    )

    modelo_sel = st.selectbox(
        "Modelo Gemini Principal",
        options=[
            "Auto-Detectar (Recomendado)",
            "gemini-3-flash-preview",
            "gemini-3.0-flash",
            "gemini-3.1-flash-preview",
            "gemini-flash-latest",
            "gemini-3-pro-preview",
            "gemini-3.0-pro"
        ],
        index=0,
        help="Con 'Auto-Detectar', el programa consulta automáticamente qué modelos Gemini 3 están activos en tu API Key."
    )

    usar_ocr_vision = st.toggle(
        "Activar OCR Inteligente en páginas escaneadas",
        value=True,
        help="Si un PDF tiene páginas escaneadas como imagen, usa Tesseract o Gemini Vision para leerlas."
    )

    st.divider()
    st.header("⚖️ Parámetros del Caso")

    tipo_proceso = st.selectbox(
        "Naturaleza del Expediente / Trámite",
        options=[
            "Acción de Tutela / Incidente de Desacato / Derecho a la Salud",
            "Demanda Judicial (Civil / Contencioso Administrativo / Laboral)",
            "Demanda o Queja ante Superintendencia (SIC - Protección al Consumidor / Publicidad Engañosa)",
            "Actuación ante Procuraduría / Personería / Control Urbano y Espacio Público",
            "Proceso Policivo / Querella / Propiedad Horizontal",
            "Auditoría / Interventoría de Contratos y Cumplimiento de Obligaciones"
        ]
    )

    rol_usuario = st.selectbox(
        "Tu Rol en el Expediente",
        options=[
            "Demandante / Accionante / Agente Oficioso / Quejoso",
            "Demandado / Entidad Convocada / Defensa Técnica",
            "Auditor / Analista Jurídico Imparcial"
        ]
    )

    limite_chars = st.slider(
        "Capacidad máxima de lectura por documento (caracteres)",
        min_value=20000,
        max_value=250000,
        value=120000,
        step=20000,
        help="Con 120.000 caracteres lee cerca de 50-60 páginas completas por cada archivo sin saturar el límite por minuto de cuentas gratuitas."
    )

    st.divider()
    if st.button("🗑️ Limpiar Sesión y Memoria", use_container_width=True):
        st.session_state.expediente_docs = []
        st.session_state.actuaciones_df = pd.DataFrame()
        st.session_state.catalogo_df = pd.DataFrame()
        st.session_state.informe_md = ""
        st.session_state.historial_chat = []
        st.rerun()

# ----------------- PESTAÑAS DE TRABAJO -----------------
tab_carga, tab_auditoria, tab_rag, tab_export = st.tabs([
    "📂 1. Carga e Inventario del Expediente",
    "🔍 2. Auditoría de Contradicciones y Plan de Acción",
    "💬 3. Consultor Directo al Expediente (Chat)",
    "📄 4. Descargar Informe en Word y Anexos Excel"
])

# =====================================================================
# PESTAÑA 1: CARGA, OCR E INVENTARIO DOCUMENTAL
# =====================================================================
with tab_carga:
    st.subheader("📤 Subir Documentos del Expediente o Demanda")
    archivos_subidos = st.file_uploader(
        "Selecciona todos los archivos del caso (Demandas, Contestaciones, Oficios, Autos, Pruebas, Contratos en PDF, DOCX o TXT)",
        type=["pdf", "docx", "txt", "rtf"],
        accept_multiple_files=True
    )

    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        procesar_btn = st.button("⚙️ Procesar e Indexar Expediente", type="primary", use_container_width=True)

    if procesar_btn:
        if not archivos_subidos:
            st.warning("⚠️ Por favor selecciona al menos un archivo antes de procesar.")
        else:
            cliente_tmp = obtener_cliente_gemini(api_key_input) if api_key_input else None
            docs_procesados = []
            barra = st.progress(0, text="Iniciando lectura paginada de documentos...")

            for idx, arch in enumerate(archivos_subidos):
                barra.progress(
                    (idx) / len(archivos_subidos),
                    text=f"Extrayendo texto y folios de: {arch.name} ({idx+1}/{len(archivos_subidos)})..."
                )
                item_doc = procesar_archivo_subido(arch, client=cliente_tmp, usar_ocr_ia=usar_ocr_vision)
                docs_procesados.append(item_doc)

            barra.progress(1.0, text="¡Extracción completada!")
            st.session_state.expediente_docs = docs_procesados

            filas_cat = []
            for it in docs_procesados:
                filas_cat.append({
                    "Archivo": it["nombre"],
                    "Páginas": it["num_paginas"],
                    "Págs con OCR": len(it["paginas_ocr"]),
                    "Caracteres Extraídos": it["caracteres"],
                    "Radicado Detectado": it["meta"]["radicado"],
                    "Actuaciones Detectadas": ", ".join(it["meta"]["tipo_actuacion"][:4]) or "General",
                    "Actor / Solicitante": it["meta"]["actor"],
                    "Demandado / Entidad": it["meta"]["demandado"],
                    "Oficio / Auto": it["meta"]["numero_oficio"],
                    "Hash SHA-256": it["hash"]
                })
            st.session_state.catalogo_df = pd.DataFrame(filas_cat)
            st.session_state.actuaciones_df = extraer_linea_tiempo_regex(docs_procesados)
            st.success(f"✅ Se procesaron exitosamente {len(docs_procesados)} documentos con trazabilidad página por página.")

    if st.session_state.expediente_docs:
        docs_mem = st.session_state.expediente_docs
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📁 Documentos Cargados", len(docs_mem))
        m2.metric("📄 Total Páginas / Folios", sum(d["num_paginas"] for d in docs_mem))
        m3.metric("🔍 Páginas Recuperadas con OCR", sum(len(d["paginas_ocr"]) for d in docs_mem))
        m4.metric("📝 Caracteres Totales Leídos", f"{sum(d['caracteres'] for d in docs_mem):,}")

        st.markdown("#### 📋 Catálogo Documental Detectado")
        st.dataframe(st.session_state.catalogo_df, use_container_width=True)

        with st.expander("🕒 Ver Actuaciones y Fechas Detectadas Automáticamente"):
            st.dataframe(st.session_state.actuaciones_df, use_container_width=True)

        with st.expander("🔎 Inspeccionar Texto Extraído por Archivo y Página"):
            nombres_arch = [d["nombre"] for d in docs_mem]
            arch_inspeccion = st.selectbox("Selecciona un documento para verificar su lectura:", nombres_arch)
            doc_sel = next(d for d in docs_mem if d["nombre"] == arch_inspeccion)
            st.text_area(
                f"Contenido extraído de {arch_inspeccion} ({doc_sel['caracteres']:,} caracteres):",
                value=doc_sel["raw"][:25000] + ("\n\n[... Vista previa truncada en pantalla, pero completa en memoria ...]" if doc_sel["caracteres"] > 25000 else ""),
                height=300
            )

# =====================================================================
# PESTAÑA 2: AUDITORÍA DE CONTRADICCIONES Y PLAN DE ACCIÓN
# =====================================================================
with tab_auditoria:
    st.subheader("🔍 Auditoría Forense de Contradicciones e Informe de Actuación")

    enfoque_usuario = st.text_area(
        "🎯 ¿Qué puntos específicos, promesas incumplidas, fechas o contradicciones quieres que la IA revise con especial rigor? (Opcional)",
        placeholder="Ejemplo: Verificar contradicciones entre lo ordenado en el fallo de tutela y lo que responden las EPS o el juzgado en los autos del incidente de desacato; identificar incumplimientos y redactar el memorial correspondiente.",
        height=100
    )

    generar_informe_btn = st.button("🚀 Ejecutar Auditoría Completa y Generar Informe", type="primary", use_container_width=True)

    if generar_informe_btn:
        if not api_key_input:
            st.error("❌ Ingresa tu Google Gemini API Key en la barra lateral izquierda.")
        elif not st.session_state.expediente_docs:
            st.warning("⚠️ Primero sube y procesa los documentos en la Pestaña 1.")
        else:
            client = obtener_cliente_gemini(api_key_input)
            prompt_auditoria = construir_prompt_auditoria(
                items=st.session_state.expediente_docs,
                catalogo_df=st.session_state.catalogo_df,
                actuaciones_df=st.session_state.actuaciones_df,
                tipo_proceso=tipo_proceso,
                rol_usuario=rol_usuario,
                enfoque_especial=enfoque_usuario,
                limite_chars_por_doc=limite_chars
            )

            with st.spinner("⚖️ Detectando modelo Gemini activo, cruzando los 27 documentos y redactando el informe..."):
                try:
                    texto_informe, modelo_usado = llamar_gemini(
                        client=client,
                        prompt_contents=prompt_auditoria,
                        modelo_preferido=modelo_sel,
                        temperatura=0.1
                    )
                    st.session_state.informe_md = texto_informe
                    st.success(f"✅ Auditoría completada con éxito utilizando el modelo `{modelo_usado}`. Ya puedes descargar el documento Word abajo.")
                except Exception as e:
                    st.error(f"❌ Error al generar el informe: {str(e)}")

    if st.session_state.informe_md:
        word_bytes_top = convertir_markdown_a_docx_bytes(st.session_state.informe_md)
        st.download_button(
            label="📥 Descargar Informe Completo en Word (.docx)",
            data=word_bytes_top,
            file_name=f"Informe_Auditoria_Expediente_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        st.divider()
        st.markdown(st.session_state.informe_md)

# =====================================================================
# PESTAÑA 3: CONSULTOR DIRECTO AL EXPEDIENTE (CHAT CON TRAZABILIDAD)
# =====================================================================
with tab_rag:
    st.subheader("💬 Hacer Preguntas Puntuales al Expediente")
    st.caption("El sistema consulta la totalidad de las páginas extraídas y cita el archivo y la página exacta de cada respuesta.")

    for msj in st.session_state.historial_chat:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    pregunta_chat = st.chat_input("Escribe tu pregunta sobre el expediente (ej. ¿En qué página se contradice la respuesta del accionado con las pruebas?)...")

    if pregunta_chat:
        if not api_key_input:
            st.error("❌ Configura tu API Key en la barra lateral.")
        elif not st.session_state.expediente_docs:
            st.warning("⚠️ Sube y procesa primero los documentos en la Pestaña 1.")
        else:
            st.session_state.historial_chat.append({"role": "user", "content": pregunta_chat})
            with st.chat_message("user"):
                st.markdown(pregunta_chat)

            with st.chat_message("assistant"):
                with st.spinner("Buscando en los folios del expediente..."):
                    client = obtener_cliente_gemini(api_key_input)
                    contexto_docs = "\n\n".join(
                        [f"=== DOCUMENTO: {d['nombre']} ===\n{d['raw'][:limite_chars]}" for d in st.session_state.expediente_docs]
                    )
                    prompt_pregunta = f"""
Responde como abogado auditor colombiano basándote ÚNICAMENTE en los documentos del expediente.
Cita siempre entre corchetes `[Archivo | Página X]` y transcribe entre comillas las frases clave que prueban tu respuesta.
Si el dato no existe en los archivos, indícalo expresamente como: "NO CONSTA EN LOS DOCUMENTOS".

EXPEDIENTE:
{contexto_docs}

PREGUNTA DEL USUARIO:
{pregunta_chat}
"""
                    try:
                        respuesta_chat, _ = llamar_gemini(
                            client=client,
                            prompt_contents=prompt_pregunta,
                            modelo_preferido=modelo_sel,
                            temperatura=0.1
                        )
                        st.markdown(respuesta_chat)
                        st.session_state.historial_chat.append({"role": "assistant", "content": respuesta_chat})
                    except Exception as e:
                        st.error(f"Error en la consulta: {str(e)}")

# =====================================================================
# PESTAÑA 4: CENTRO DE EXPORTACIÓN (WORD, EXCEL Y MARKDOWN)
# =====================================================================
with tab_export:
    st.subheader("📄 Descarga de Entregables Profesionales")

    col_exp1, col_exp2, col_exp3 = st.columns(3)

    with col_exp1:
        st.markdown("#### 📘 1. Informe Jurídico en Word (`.docx`)")
        st.write("Incluye tablas formateadas, matriz de contradicciones, ruta de lo que se debe hacer y borrador de memorial.")
        if st.session_state.informe_md:
            docx_file = convertir_markdown_a_docx_bytes(st.session_state.informe_md)
            st.download_button(
                label="⬇️ Descargar Informe en Word (.docx)",
                data=docx_file,
                file_name=f"01_INFORME_AUDITORIA_EXPEDIENTE_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        else:
            st.info("Genera primero el informe en la Pestaña 2.")

    with col_exp2:
        st.markdown("#### 📊 2. Catálogo y Cronología (`.xlsx`)")
        st.write("Libro de Excel con dos hojas: Inventario técnico con Hash SHA-256 y Línea de Tiempo de actuaciones.")
        if not st.session_state.catalogo_df.empty:
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
                st.session_state.catalogo_df.to_excel(writer, sheet_name="01_Catalogo_Documentos", index=False)
                st.session_state.actuaciones_df.to_excel(writer, sheet_name="02_Cronologia_Actuaciones", index=False)
            buffer_excel.seek(0)
            st.download_button(
                label="⬇️ Descargar Matrices en Excel (.xlsx)",
                data=buffer_excel,
                file_name=f"02_MATRIZ_CONTROL_EXPEDIENTE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Procesa primero los documentos en la Pestaña 1.")

    with col_exp3:
        st.markdown("#### 📝 3. Respaldo en Markdown (`.md`)")
        st.write("Archivo de texto estructurado original del informe para archivo digital o edición rápida.")
        if st.session_state.informe_md:
            st.download_button(
                label="⬇️ Descargar Informe (.md)",
                data=st.session_state.informe_md.encode("utf-8"),
                file_name=f"03_INFORME_RESPALDO_{datetime.now().strftime('%Y%m%d')}.md",
                mime="text/markdown",
                use_container_width=True
            )
        else:
            st.info("Genera primero el informe en la Pestaña 2.")
