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
from openpyxl.styles import Font, PatternFill, Alignment
from google import genai
from google.genai import types

# =====================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# =====================================================================
st.set_page_config(
    page_title="Auditor & Estratega Jurídico IA Colombia | Expedientes y Consultoría",
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
    .tool-box {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #1e3a8a;
        padding: 1.2rem;
        border-radius: 8px;
        margin-bottom: 1rem;
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
if "estrategia_md" not in st.session_state:
    st.session_state.estrategia_md = ""
if "historial_chat" not in st.session_state:
    st.session_state.historial_chat = []

# =====================================================================
# EXPRESIONES REGULARES CORREGIDAS PARA METADATOS Y FECHAS
# =====================================================================
# Prioriza años de 4 dígitos (ej. 2026) antes de años de 2 dígitos para no truncar 2026 en 20
RE_FECHA = re.compile(r'\b\d{1,2}[\/\-]\d{1,2}[\/\-](?:\d{4}|\d{2})\b')
RE_RADICADO = re.compile(
    r'(Radicado|Rad\.|Proceso|Expediente|Caso|No\.?)[ \t]*[:#\-]?[ \t]*([0-9]{4,23}(?:[-][0-9]{1,8})*)',
    re.I
)
RE_ACTOR = re.compile(
    r'(Demandante|Actor|Solicitante|Promovente|Accionante|Quejoso|Denunciante|Peticionario|Incidentante)[ \t]*[:\-][ \t]*([^\n\r]{3,120})',
    re.I
)
RE_DEMANDADO = re.compile(
    r'(Demandado|Demandada|Convocado|Entidad|Accionado|Accionada|Investigado|Denunciado|Querellado)[ \t]*[:\-][ \t]*([^\n\r]{3,120})',
    re.I
)
RE_JUEZ = re.compile(
    r'(Juez|Jueza|Magistrad[oa]|Despacho|Juzgado|Superintendencia|Procuradur[ií]a|Personer[ií]a|Fiscal[ií]a|Alcald[ií]a|Inspecci[oó]n|Secretar[ií]a)[ \t]*[:\-]?[ \t]*([^\n\r]{3,120})',
    re.I
)
RE_NUMERO_OFICIO = re.compile(
    r'(OFICIO|Resoluci[oó]n|Auto|Acta)[ \t]+No\.?[ \t]*([0-9A-Za-z\-]{1,20})',
    re.I
)

TIPOS_ACT = [
    'DEMANDA', 'TUTELA', 'DESACATO', 'QUEJA', 'DERECHO DE PETICION', 'CONTESTACION',
    'REPLICA', 'AUTO', 'PROVIDENCIA', 'RESOLUCION', 'SENTENCIA', 'APELACION',
    'RECURSO', 'REPOSICION', 'NOTIFICACION', 'ACTA', 'OFICIO', 'INFORME',
    'PERITAJE', 'LIQUIDACION', 'AUDIENCIA', 'MEMORIAL', 'PRUEBA', 'CONTRATO',
    'PROMESA', 'ESCRITURA', 'LICENCIA', 'QUERELLA', 'DENUNCIA'
]

# =====================================================================
# CLIENTE GEMINI CON AUTO-DETECCIÓN DE MODELOS ACTIVOS
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
    "gemini-2.5-flash",
    "gemini-2.5-pro"
]

def obtener_cliente_gemini(api_key: str):
    if not api_key:
        return None
    return genai.Client(api_key=api_key.strip())

def descubrir_modelos_activos(client):
    """Consulta a Google qué modelos Gemini están activos en la API Key actual."""
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
    cola_final = modelos_cuenta_ordenados + [m for m in MODELOS_RESPALDO if m not in modelos_cuenta_ordenados]
    return cola_final

def llamar_gemini(client, prompt_contents, modelo_preferido="Auto-Detectar (Recomendado)", temperatura=0.1, max_reintentos=2):
    """Ejecuta la consulta en Gemini seleccionando automáticamente el modelo habilitado."""
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
                if "404" in msg or "not found" in msg or "no longer available" in msg or "not supported" in msg:
                    break
                if "429" in msg or "quota" in msg or "resource_exhausted" in msg:
                    if intento == 0:
                        time.sleep(4)
                        continue
                    else:
                        break
                time.sleep(2)

    raise RuntimeError("No fue posible obtener respuesta. Últimos intentos: " + " | ".join(errores[-4:]))

# =====================================================================
# EXTRACCIÓN DOCUMENTAL + OCR HÍBRIDO (SIN BARRAS '|' EN LAS MARCAS)
# =====================================================================
def ocr_pagina_hibrido(page, client=None):
    pix = page.get_pixmap(dpi=200)
    img_bytes = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_bytes))

    try:
        import pytesseract
        texto_ocr = pytesseract.image_to_string(img, lang="spa").strip()
        if len(texto_ocr) > 25:
            return texto_ocr, "OCR-Tesseract"
    except Exception:
        pass

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

def limpiar_linea_meta(txt: str, max_len=110):
    limpio = re.sub(r'\s+', ' ', str(txt)).strip()
    return limpio[:max_len]

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
        meta['radicado'] = limpiar_linea_meta(rr.group(0), 70)

    no_of = RE_NUMERO_OFICIO.search(texto)
    if no_of:
        meta['numero_oficio'] = limpiar_linea_meta(no_of.group(0), 60)

    da = RE_ACTOR.search(texto)
    if da:
        meta['actor'] = limpiar_linea_meta(da.group(2), 110)

    dd = RE_DEMANDADO.search(texto)
    if dd:
        meta['demandado'] = limpiar_linea_meta(dd.group(2), 110)

    jj = RE_JUEZ.search(texto)
    if jj:
        meta['juez'] = limpiar_linea_meta(jj.group(0), 110)

    texto_upper = texto.upper()
    for tipo in TIPOS_ACT:
        if tipo in texto_upper:
            meta['tipo_actuacion'].append(tipo)

    return meta

def procesar_archivo_subido(uploaded_file, client=None, usar_ocr_ia=True):
    """
    Lee PDF, DOCX o TXT conservando trazabilidad exacta por página.
    Usa el formato [Archivo: X - Pág. Y] sin barras verticales '|' para no romper tablas Markdown/Word.
    """
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
                        "texto": f"[Archivo: {nombre} - Pág. {num_pag}]\n{txt}"
                    })
                else:
                    cliente_ocr = client if usar_ocr_ia else None
                    txt_ocr, metodo_ocr = ocr_pagina_hibrido(page, cliente_ocr)
                    if txt_ocr:
                        paginas_ocr.append(num_pag)
                        paginas_estructuradas.append({
                            "pagina": num_pag,
                            "origen": metodo_ocr,
                            "texto": f"[Archivo: {nombre} - Pág. {num_pag} ({metodo_ocr})]\n{txt_ocr}"
                        })
                    else:
                        paginas_vacias.append(num_pag)
                        paginas_estructuradas.append({
                            "pagina": num_pag,
                            "origen": "Sin Texto",
                            "texto": f"[Archivo: {nombre} - Pág. {num_pag} (SIN TEXTO LEGIBLE)]"
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
                        parrafos.append(" - ".join(celdas))

            texto_completo_docx = "\n".join(parrafos)
            tam_pag = 3000
            for idx_b, inicio in enumerate(range(0, max(1, len(texto_completo_docx)), tam_pag)):
                fragmento = texto_completo_docx[inicio:inicio + tam_pag]
                paginas_estructuradas.append({
                    "pagina": idx_b + 1,
                    "origen": "Word Nativo",
                    "texto": f"[Archivo: {nombre} - Sección {idx_b + 1}]\n{fragmento}"
                })

        elif ext in (".txt", ".rtf"):
            contenido = file_bytes.decode("utf-8", errors="ignore").strip()
            paginas_estructuradas.append({
                "pagina": 1,
                "origen": "Texto Plano",
                "texto": f"[Archivo: {nombre} - Pág. 1]\n{contenido}"
            })

    except Exception as e:
        paginas_estructuradas.append({
            "pagina": 1,
            "origen": "Error",
            "texto": f"[Archivo: {nombre} - ERROR DE EXTRACCIÓN: {str(e)}]"
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
    """Detecta actuaciones y fechas completas (con 4 dígitos de año) sin capturar saltos de línea sucios."""
    patron_bloque = re.compile(
        r'(?m)^(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-](?:\d{4}|\d{2}))[ \t]*(?:[-–:][ \t]*)?(?P<encabezado>[\w\t \(\)\.,#º°\-]{8,110})$'
    )
    patron_acto = re.compile(
        r'(?i)(auto|providencia|resoluci[oó]n|sentencia|acta|oficio|notificaci[oó]n|audiencia|memorial|radicado|queja|recurso|orden|f[oó]rmula)[ \t]+(?:no\.?|n[o°]?)?[ \t]*[\w\-]*[ \t]*(?:del|de fecha|:)?[ \t]*(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-](?:\d{4}|\d{2}))'
    )

    actuaciones = []
    for it in items:
        for pag in it["paginas"]:
            txt = pag["texto"]
            num_pag = pag["pagina"]

            for m in patron_bloque.finditer(txt):
                enc_limpio = re.sub(r'\s+', ' ', m.group("encabezado")).strip()
                ctx_limpio = re.sub(r'\s+', ' ', txt[max(0, m.start() - 30):min(len(txt), m.end() + 140)]).strip()
                actuaciones.append({
                    "Archivo": it["nombre"],
                    "Página": num_pag,
                    "Fecha Literal": m.group("fecha").strip(),
                    "Actuación / Referencia": enc_limpio,
                    "Contexto": ctx_limpio
                })

            for m in patron_acto.finditer(txt):
                enc_limpio = re.sub(r'\s+', ' ', m.group(0)).strip()
                ctx_limpio = re.sub(r'\s+', ' ', txt[max(0, m.start() - 30):min(len(txt), m.end() + 140)]).strip()
                actuaciones.append({
                    "Archivo": it["nombre"],
                    "Página": num_pag,
                    "Fecha Literal": m.group("fecha").strip(),
                    "Actuación / Referencia": enc_limpio,
                    "Contexto": ctx_limpio
                })

    if not actuaciones:
        return pd.DataFrame(columns=["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"])

    def normalizar_fecha(f_str):
        try:
            d = parse(str(f_str), languages=["es"], settings={"DATE_ORDER": "DMY"})
            if d and 1990 <= d.year <= 2035:
                return d.date().isoformat()
        except Exception:
            pass
        return "9999-12-31"

    df_act = pd.DataFrame(actuaciones)
    df_act["Fecha ISO"] = df_act["Fecha Literal"].apply(normalizar_fecha)
    df_act = df_act.drop_duplicates(subset=["Archivo", "Página", "Fecha Literal", "Actuación / Referencia"])
    df_act = df_act.sort_values(by=["Fecha ISO", "Archivo", "Página"]).reset_index(drop=True)
    df_act["Fecha ISO"] = df_act["Fecha ISO"].replace("9999-12-31", "Sin Fecha Exacta")

    columnas_orden = ["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"]
    return df_act[columnas_orden]

# =====================================================================
# GENERADOR PROFESIONAL DE WORD (.DOCX) Y EXCEL (.XLSX) BLINDADO
# =====================================================================
def aplicar_color_celda(cell, hex_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def agregar_texto_con_negrillas(parrafo, texto):
    texto_limpio = texto.replace('`', '')
    partes = re.split(r'(\*\*.*?\*\*)', texto_limpio)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**') and len(parte) > 4:
            run = parrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            parrafo.add_run(parte)

def dividir_fila_markdown_segura(linea_limpia: str):
    """
    Evita que una barra vertical '|' dentro de corchetes [[... | ...]] o [... | ...]
    parta por error una celda en dos columnas dentro de las tablas de Word.
    """
    linea_sanitizada = re.sub(
        r'\[+[^\]]*\]+',
        lambda m: m.group(0).replace('|', ' - '),
        linea_limpia
    )
    partes = linea_sanitizada.strip().split('|')
    if linea_sanitizada.startswith('|'):
        partes = partes[1:]
    if linea_sanitizada.endswith('|'):
        partes = partes[:-1]
    return [c.strip() for c in partes]

def volcar_tabla_en_word(doc, filas_tabla):
    if not filas_tabla:
        return
    num_cols = len(filas_tabla[0])
    if num_cols == 0:
        return

    tabla = doc.add_table(rows=len(filas_tabla), cols=num_cols)
    tabla.style = 'Table Grid'

    for r_idx, fila in enumerate(filas_tabla):
        # Si una fila tiene más elementos que el encabezado, unir el excedente en la última celda
        if len(fila) > num_cols:
            fila_ajustada = fila[:num_cols - 1] + [" - ".join(fila[num_cols - 1:])]
        else:
            fila_ajustada = fila + [""] * (num_cols - len(fila))

        for c_idx in range(num_cols):
            texto_celda = fila_ajustada[c_idx]
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

def convertir_markdown_a_docx_bytes(md_texto: str, titulo_portada="INFORME JURÍDICO Y PLAN DE ACCIÓN PROCESAL"):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.95)
        section.right_margin = Inches(0.95)

    titulo = doc.add_heading(titulo_portada, level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub.add_run(f"Generado el {datetime.now().strftime('%d/%m/%Y a las %H:%M')} | Auditoría Jurídica y Estratégica Bajo Normativa Colombiana")
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
            celdas = dividir_fila_markdown_segura(linea_limpia)
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

def generar_excel_formateado_bytes(catalogo_df: pd.DataFrame, actuaciones_df: pd.DataFrame):
    """Crea un archivo Excel estilizado con encabezados institucionales y columnas ajustadas."""
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
        catalogo_df.to_excel(writer, sheet_name="01_Catalogo_Documentos", index=False)
        actuaciones_df.to_excel(writer, sheet_name="02_Cronologia_Actuaciones", index=False)

        wb = writer.book
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=10)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col), default=12)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 55)
                for cell in col[1:]:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

    buffer_excel.seek(0)
    return buffer_excel

# =====================================================================
# PROMPTS JURÍDICOS MULTIJURISDICCIONALES (EXPEDIENTES Y CONSULTORÍA)
# =====================================================================
MARCO_JURIDICO_COLOMBIA = """
DOMINIO NORMATIVO Y PROCESAL COLOMBIANO OBLIGATORIO SEGÚN LA ENTIDAD:
- Constitucional y Salud: Constitución Política (Arts. 11, 23, 29, 49, 86), Decreto 2591 de 1991 (Tutela y Desacato), Ley Estatutaria 1751 de 2015, Sentencias SU-034 de 2018 y T-760 de 2008.
- Derecho de Petición y Procedimiento Administrativo: Ley 1755 de 2015 y CPACA (Ley 1437 de 2011 modificada por Ley 2080 de 2021).
- Protección al Consumidor, Constructoras e Inmobiliarias (SIC): Ley 1480 de 2011 (Estatuto del Consumidor: publicidad engañosa, efectividad de la garantía, información mínima, cláusulas abusivas), Circular Única de la SIC.
- Control Disciplinario, Urbano y Ministerio Público (Procuraduría, Personerías, Alcaldías): Ley 1952 de 2019 y Ley 2094 de 2021 (Código General Disciplinario), Ley 388 de 1997 (Ordenamiento Territorial, entregas de zonas de cesión y espacio público), Ley 1801 de 2016 (Código Nacional de Seguridad y Convivencia Ciudadana - Querellas Policivas), Ley 142 de 1994 (Servicios Públicos Domiciliarios).
- Procesal Civil, Comercial y Propiedad Horizontal: Ley 1564 de 2012 (Código General del Proceso - CGP), Código Civil, Código de Comercio, Ley 675 de 2001.
- Laboral y Función Pública / Concursos de Méritos: CST, CPTSS, Ley 909 de 2004, Decretos reglamentarios CNSC.
- Penal y Disciplinario Judicial: Ley 599 de 2000 (Código Penal: fraude en resolución judicial, prevaricato por acción u omisión, estafa, urbanización ilegal), Ley 906 de 2004, Comisión Nacional de Disciplina Judicial.
"""

def construir_prompt_auditoria(
    items,
    catalogo_df,
    actuaciones_df,
    tipo_proceso,
    entidad_especifica,
    rol_usuario,
    enfoque_especial,
    limite_chars_por_doc=120000
):
    total_docs = len(items)
    total_pags = sum(it["num_paginas"] for it in items)
    total_ocr = sum(len(it["paginas_ocr"]) for it in items)

    catalogo_md = catalogo_df.to_markdown(index=False) if not catalogo_df.empty else "Sin catálogo"
    actos_md = actuaciones_df.head(80).to_markdown(index=False) if not actuaciones_df.empty else "Sin actuaciones previas"

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
Actúa como un **Abogado Litigante Senior, Auditor Forense de Expedientes y Estratega Procesal Colombiano** experto en litigio ante cualquier entidad judicial, administrativa, superintendencia u órgano de control en Colombia.

{MARCO_JURIDICO_COLOMBIA}

CONTEXTO ESTRATÉGICO DEL CASO:
- **Jurisdicción / Tipo de Trámite:** {tipo_proceso}
- **Entidad / Despacho / Autoridad Específica:** {entidad_especifica or 'Determinar según los documentos del expediente'}
- **Posición / Rol que representamos:** {rol_usuario}
- **Enfoque Especial o Problema Planteado por el Usuario:** {enfoque_especial or 'Auditar todo el expediente, detectar contradicciones entre los PDFs, identificar incumplimientos o errores de la entidad/contraparte y definir exactamente qué hacer y qué radicar a continuación.'}

REGLAS TÉCNICAS Y DE FORMATO OBLIGATORIAS:
1. **Inventario Real:** Analizas {total_docs} documentos que suman {total_pags} páginas ({total_ocr} procesadas mediante OCR).
2. **Formato de Citación (PROHIBIDO USAR BARRAS VERTICALES `|` DENTRO DE CITAS):** Cada hallazgo, hecho o contradicción DEBE citar el archivo y la página usando ÚNICAMENTE el formato `(Archivo: nombre.pdf, Pág. X)`. NUNCA uses el símbolo `|` dentro de los corchetes o paréntesis para no desajustar las tablas de Word.
3. **Cero Invenciones:** No inventes fechas, radicados ni pruebas. Si falta un soporte clave, escribe: `"NO CONSTA EN LOS DOCUMENTOS APORTADOS (VACÍO PROBATORIO)"`.

INSUMO 1: CATÁLOGO GENERAL DE ARCHIVOS
{catalogo_md}

INSUMO 2: REGISTRO PRELIMINAR DE FECHAS Y ACTUACIONES
{actos_md}

INSUMO 3: CONTENIDO ÍNTEGRO PAGINADO DE LOS DOCUMENTOS DEL EXPEDIENTE
{expediente_completo_str}

---

GENERA EL DICTAMEN JURÍDICO COMPLETO EN FORMATO MARKDOWN CON ESTAS 12 SECCIONES OBLIGATORIAS:

# INFORME DE AUDITORÍA PROCESAL, CONTRADICCIONES Y PLAN DE ACCIÓN

## 1. FICHA TÉCNICA E IDENTIFICACIÓN DEL EXPEDIENTE
(Tabla Markdown de 2 columnas `Campo` | `Detalle` con: Radicado/Expediente, Entidad o Despacho, Demandante/Accionante/Quejoso, Demandado/Convocado/Investigado, Naturaleza del Proceso, Etapa Actual y Nivel de Riesgo/Viabilidad).

## 2. RESUMEN EJECUTIVO Y DIAGNÓSTICO PROCESAL
(Síntesis precisa del litigio o trámite, qué se reclama, qué ha hecho o respondido la contraparte/entidad y cuál es el nudo crítico actual).

## 3. LÍNEA DE TIEMPO CRONOLÓGICA CON TRAZABILIDAD
(Tabla Markdown con 5 columnas: `Fecha` | `Actuación o Documento` | `Archivo y Página (sin usar símbolo |)` | `Hecho o Decisión Clave` | `Impacto Procesal`).

## 4. MATRIZ EXHAUSTIVA DE CONTRADICCIONES E INCONSISTENCIAS DETECTADAS
(Tabla Markdown con 6 columnas: `N°` | `Punto en Conflicto` | `Versión o Afirmación A (Cita + Archivo y Pág.)` | `Prueba o Versión B que la contradice (Cita + Archivo y Pág.)` | `Tipo de Contradicción` | `Gravedad y Cómo Explotarla Procesalmente`).
Después de la tabla, explica a fondo cada contradicción y qué implicación legal tiene (falsa motivación, defecto fáctico, publicidad engañosa, fraude procesal, confesión, etc.).

## 5. AUDITORÍA DE PRUEBAS: HECHOS PROBADOS VS. VACÍOS PROBATORIOS
(Separa claramente qué hechos ya están 100% acreditados con su página exacta y qué afirmaciones de la contraparte o nuestras carecen de prueba idónea en el expediente).

## 6. CONTROL DE ÓRDENES, OBLIGACIONES Y PLAZOS INCUMPLIDOS
(Tabla Markdown con 5 columnas: `Orden / Compromiso / Norma` | `Obligado o Responsable` | `Plazo Legal o Fijado` | `Estado Actual` | `Evidencia en el Expediente`).

## 7. FUNDAMENTACIÓN JURÍDICA, LEYES Y JURISPRUDENCIA APLICABLE EN COLOMBIA
(Cita con precisión las normas colombianas exactas —artículos de leyes, códigos, decretos y sentencias de la Corte Constitucional, Corte Suprema, Consejo de Estado o Superintendencias— que gobiernan este caso y respaldan nuestra posición).

## 8. QUÉ SE DEBE HACER AHORA MISMO: RUTA PROCESAL PASO A PASO (ACTUACIONES REALIZADAS Y PENDIENTES)
(Explica como abogado litigante experto:
1. **Actuación Principal Inmediata (0 a 3 días):** Qué escrito, recurso, incidente, demanda o queja se debe radicar y por qué.
2. **Actuaciones Complementarias o Ante Otras Entidades (Plan B / Presión Institucional):** Si cabe queja disciplinaria (Procuraduría/Personería/Comisión de Disciplina Judicial), vigilancia judicial/administrativa, queja ante Superintendencia (SIC/Supersalud/Superservicios), acción popular, tutela o denuncia penal.
3. **Términos Legales y Precauciones:** Plazos de vencimiento, caducidad o requisitos de procedibilidad que debemos cuidar).

## 9. TABLA DE PLAN DE ACCIÓN PRIORIZADO
(Tabla Markdown con 5 columnas: `Prioridad` | `Acción Jurídica Concreta` | `Entidad / Destinatario` | `Objetivo Procesal` | `Sustento Normativo y Probatorio`).

## 10. BORRADOR COMPLETO DE MEMORIAL / DEMANDA / QUEJA / OFICIO LISTO PARA RADICAR
(Redacta el documento jurídico formal completo y extenso, con estructura procesal colombiana: Encabezado de la entidad/juez, Referencia/Radicado, Identificación de quien suscribe, Hechos numerados con citas del expediente, Contradicciones o Incumplimientos denunciados, Fundamentos de Derecho, Pruebas/Anexos, Notificaciones y Peticiones principales y subsidiarias).

## 11. CHECKLIST DE VERIFICACIÓN ANTES DE RADICAR
(Lista de chequeo detallada para validar anexos, firmas, correos de notificación y folios antes de enviar).
"""
    return prompt

def construir_prompt_consultor_desde_cero(
    problema_descrito,
    tipo_proceso,
    entidad_especifica,
    rol_usuario,
    documento_a_generar,
    incluir_docs_cargados,
    items,
    limite_chars_por_doc=80000
):
    contexto_expediente = "NO SE ADJUNTARON DOCUMENTOS (CONSULTA PREVENTIVA / DESDE CERO)."
    if incluir_docs_cargados and items:
        partes_docs = []
        for idx, it in enumerate(items, 1):
            partes_docs.append(f"--- DOCUMENTO ADJUNTO #{idx}: {it['nombre']} ---\n{it['raw'][:limite_chars_por_doc]}")
        contexto_expediente = "\n\n".join(partes_docs)

    prompt = f"""
Actúa como un **Abogado Litigante Colombiano Senior y Consultor Estratégico Especialista en todas las jurisdicciones y entidades de Colombia**.
El usuario necesita orientación jurídica precisa sobre **qué hacer frente a un problema jurídico o actuación que aún no ha realizado (o que está por iniciar)**, de acuerdo estrictamente con las leyes colombianas vigentes.

{MARCO_JURIDICO_COLOMBIA}

DATOS DEL PROBLEMA O CASO PLANTEADO POR EL USUARIO:
- **Área / Naturaleza del Caso:** {tipo_proceso}
- **Entidad / Autoridad / Contraparte Involucrada:** {entidad_especifica or 'Por definir según la ley colombiana'}
- **Rol del Usuario:** {rol_usuario}
- **Documento Principal que el Usuario desea obtener redactado al final:** {documento_a_generar}

DESCRIPCIÓN DETALLADA DEL PROBLEMA / INCONVENIENTE DEL USUARIO:
\"\"\"
{problema_descrito}
\"\"\"

DOCUMENTOS DE SOPORTE CARGADOS EN SESIÓN (SI APLICAN):
{contexto_expediente}

REGLAS DE FORMATO:
- NO uses barras verticales `|` dentro de paréntesis o corchetes de citas para no alterar las tablas de Word.
- Sé sumamente específico con las leyes colombianas (cita número de ley, artículos aplicables y jurisprudencia).

GENERA UNA GUÍA JURÍDICA ESTRATÉGICA Y EL DOCUMENTO LISTO PARA RADICAR CON ESTA ESTRUCTURA EN MARKDOWN:

# DICTAMEN DE ESTRATEGIA JURÍDICA Y HOJA DE RUTA (COLOMBIA)

## 1. DIAGNÓSTICO JURÍDICO DEL PROBLEMA Y VIABILIDAD
(Explica qué figura jurídica se configura en Colombia con este problema, qué derechos o normas se están vulnerando y qué tan viable es reclamar).

## 2. ENTIDADES COMPETENTES: ¿ANTE QUIÉN ACUDIR Y POR QUÉ?
(Tabla Markdown de 4 columnas: `Entidad o Despacho Competente` | `Trámite o Acción a Interponer` | `Facultad de la Entidad (Sancionatoria, Jurisdiccional, Control, etc.)` | `Plazo de Respuesta o Término Legal`).

## 3. LEYES COLOMBIANAS Y ARTÍCULOS EXACTOS QUE TE RESPALDAN
(Desglose claro de las leyes, decretos y sentencias aplicables al caso, explicando cómo usar cada artículo a favor del usuario).

## 4. REQUISITOS DE PROCEDIBILIDAD, CADUCIDAD Y ERRORES QUE DEBES EVITAR
(¡MUY IMPORTANTE! Explica si antes de demandar o quejarse debe agotar reclamo directo, derecho de petición, conciliación extrajudicial o recurso previo; cuáles son los tiempos de prescripción/caducidad y qué errores hacen que las entidades rechacen este trámite).

## 5. RUTA DE ACCIÓN PASO A PASO: QUÉ HACER DESDE HOY
(Guía numerada y cronológica:
- **Paso 1 (Preparación y Pruebas):** Qué documentos, fotos, certificados, contratos, historias clínicas o actas debe reunir.
- **Paso 2 (Actuación Principal):** Cómo y dónde radicar el escrito principal (ventanilla virtual, correo institucional, plataforma de la SIC/Supersalud/Procuraduría/Rama Judicial).
- **Paso 3 (Plan B y Escalamiento):** Qué hacer si la entidad guarda silencio, niega la petición o incumple).

## 6. MATRIZ ESTRATÉGICA DE ACTUACIONES RECOMENDADAS
(Tabla Markdown de 5 columnas: `Orden` | `Actuación a Realizar` | `Ante Qué Entidad` | `Resultado Esperado` | `Fundamento Legal`).

## 7. BORRADOR COMPLETO DE: {documento_a_generar.upper()} (LISTO PARA RADICAR)
(Redacta de forma completa, extensa, rigurosa y persuasiva el documento solicitado —**{documento_a_generar}**— adaptado al relato del usuario y a los documentos cargados si los hay. Incluye: Ciudad y fecha, Destinatario exacto, Asunto/Referencia, Identificación del accionante/peticionario, Hechos claros y numerados, Fundamentos Constitucionales y Legales, Pruebas y Anexos, Solicitudes/Peticiones Principales y Subsidiarias, Notificaciones y Firma).

## 8. CHECKLIST DE PRUEBAS Y ANEXOS NECESARIOS
(Lista de chequeo de los soportes exactos que el usuario debe adjuntar a este documento antes de enviarlo).
"""
    return prompt

# =====================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# =====================================================================
st.markdown("""
<div class="main-header">
    <h2 style="margin:0;">⚖️ Suite Jurídica IA Colombia: Auditor de Expedientes y Estratega Legal</h2>
    <p style="margin:0.3rem 0 0 0; opacity:0.9;">
        Auditoría forense de contradicciones en PDFs, consultoría para actuaciones aún no realizadas y generación de escritos en Word (.docx) para cualquier entidad
    </p>
</div>
""", unsafe_allow_html=True)

# ----------------- BARRA LATERAL (CONFIGURACIÓN UNIVERSAL) -----------------
with st.sidebar:
    st.header("🔑 Configuración del Motor IA")
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=os.environ.get("GOOGLE_API_KEY", ""),
        type="password",
        help="Ingresa tu clave de Google AI Studio. Se mantiene segura en la memoria temporal de tu sesión."
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
        index=0
    )

    usar_ocr_vision = st.toggle(
        "Activar OCR en páginas escaneadas",
        value=True
    )

    st.divider()
    st.header("🏛️ Jurisdicción y Entidad")

    tipo_proceso = st.selectbox(
        "1. Área o Naturaleza del Trámite",
        options=[
            "Acción Constitucional (Tutela / Incidente de Desacato / Cumplimiento / Acción Popular)",
            "Superintendencia de Industria y Comercio (SIC - Consumidor / Publicidad Engañosa / Constructoras / Garantías)",
            "Superintendencia de Salud / Servicios Públicos / Financiera / Transporte",
            "Órganos de Control (Procuraduría / Personería / Contraloría / Queja Disciplinaria / Omisión de Funciones)",
            "Alcaldías, Secretarías de Planeación e Inspecciones de Policía (Urbanismo / Espacio Público / Querellas / Servicios)",
            "Jurisdicción Civil, Comercial, Inmobiliaria y Propiedad Horizontal (Contratos / Promesas / Ejecutivos)",
            "Jurisdicción Contencioso Administrativa (Nulidad y Restablecimiento / Reparación Directa / Estatal)",
            "Jurisdicción Laboral, Seguridad Social y Función Pública (CNSC / Concursos de Méritos / Pensiones)",
            "Jurisdicción Penal y Fiscalía (Denuncia Penal / Fraude en Resolución / Prevaricato / Estafa)",
            "Auditoría / Interventoría de Contratos, Presupuestos y Cumplimiento de Obligaciones",
            "Otra Entidad o Caso Mixto (Especificar abajo)"
        ]
    )

    entidad_especifica = st.text_input(
        "2. Nombre de la Entidad, Juzgado o Contraparte (Opcional)",
        placeholder="Ej: SIC, Procuraduría, Personería de Madrid, Juzgado Promiscuo, Constructora, EPS..."
    )

    rol_usuario = st.selectbox(
        "3. Tu Rol en el Caso",
        options=[
            "Demandante / Accionante / Quejoso / Peticionario / Agente Oficioso / Veedor",
            "Demandado / Entidad Convocada / Investigado / Defensa Técnica",
            "Auditor / Interventor / Consultor Jurídico Imparcial"
        ]
    )

    limite_chars = st.slider(
        "Capacidad de lectura por documento (caracteres)",
        min_value=20000,
        max_value=250000,
        value=120000,
        step=20000
    )

    st.divider()
    if st.button("🗑️ Limpiar Sesión y Memoria", use_container_width=True):
        st.session_state.expediente_docs = []
        st.session_state.actuaciones_df = pd.DataFrame()
        st.session_state.catalogo_df = pd.DataFrame()
        st.session_state.informe_md = ""
        st.session_state.estrategia_md = ""
        st.session_state.historial_chat = []
        st.rerun()

# ----------------- PESTAÑAS DE LA SUITE JURÍDICA -----------------
tab_carga, tab_auditoria, tab_consultor, tab_rag, tab_export = st.tabs([
    "📂 1. Subir Expediente",
    "🔍 2. Auditar Expediente y Contradicciones",
    "🧰 3. Caja de Herramientas: ¿Qué Hacer? (Con o Sin Expediente)",
    "💬 4. Chat con el Expediente",
    "📄 5. Descargar Word y Excel"
])

# =====================================================================
# PESTAÑA 1: CARGA, OCR E INVENTARIO DOCUMENTAL
# =====================================================================
with tab_carga:
    st.subheader("📤 Subir Archivos del Expediente, Demanda, Petición o Pruebas")
    st.caption("Puedes subir expedientes judiciales completos, respuestas de entidades, oficios, contratos, promesas de compraventa, folletos publicitarios o historias clínicas.")
    archivos_subidos = st.file_uploader(
        "Selecciona tus archivos (PDF, DOCX, TXT o RTF)",
        type=["pdf", "docx", "txt", "rtf"],
        accept_multiple_files=True
    )

    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        procesar_btn = st.button("⚙️ Procesar e Indexar Documentos", type="primary", use_container_width=True)

    if procesar_btn:
        if not archivos_subidos:
            st.warning("⚠️ Selecciona al menos un archivo para procesar, o ve directo a la Pestaña 3 si deseas hacer una consulta jurídica desde cero sin documentos.")
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

        with st.expander("🕒 Ver Línea de Tiempo Preliminar (Fechas de 4 dígitos corregidas)"):
            st.dataframe(st.session_state.actuaciones_df, use_container_width=True)

# =====================================================================
# PESTAÑA 2: AUDITORÍA DE EXPEDIENTES Y CONTRADICCIONES
# =====================================================================
with tab_auditoria:
    st.subheader("🔍 Auditoría Forense de Contradicciones e Informe de Actuación")
    st.write("Cruza todos los PDFs cargados en la Pestaña 1, detecta contradicciones fácticas, técnicas y normativas, y redacta el memorial o actuación que debes presentar.")

    enfoque_usuario = st.text_area(
        "🎯 Instrucciones específicas o puntos que quieres revisar con lupa en estos PDFs (Opcional):",
        placeholder="Ejemplo: Revisar las contradicciones entre la promesa/publicidad inicial y las respuestas de la entidad o constructora; verificar incumplimientos de autos u oficios; proponer qué hacer y redactar el memorial o recurso.",
        height=100
    )

    generar_informe_btn = st.button("🚀 Ejecutar Auditoría Completa del Expediente", type="primary", use_container_width=True)

    if generar_informe_btn:
        if not api_key_input:
            st.error("❌ Ingresa tu Google Gemini API Key en la barra lateral izquierda.")
        elif not st.session_state.expediente_docs:
            st.warning("⚠️ Primero sube y procesa los documentos en la Pestaña 1 (o usa la Pestaña 3 si aún no tienes expediente).")
        else:
            client = obtener_cliente_gemini(api_key_input)
            prompt_auditoria = construir_prompt_auditoria(
                items=st.session_state.expediente_docs,
                catalogo_df=st.session_state.catalogo_df,
                actuaciones_df=st.session_state.actuaciones_df,
                tipo_proceso=tipo_proceso,
                entidad_especifica=entidad_especifica,
                rol_usuario=rol_usuario,
                enfoque_especial=enfoque_usuario,
                limite_chars_por_doc=limite_chars
            )

            with st.spinner("⚖️ Cruzando todos los documentos del expediente, verificando leyes colombianas y redactando el informe..."):
                try:
                    texto_informe, modelo_usado = llamar_gemini(
                        client=client,
                        prompt_contents=prompt_auditoria,
                        modelo_preferido=modelo_sel,
                        temperatura=0.1
                    )
                    st.session_state.informe_md = texto_informe
                    st.success(f"✅ Auditoría completada con `{modelo_usado}`. Las tablas de Word ya están blindadas para no desajustar columnas.")
                except Exception as e:
                    st.error(f"❌ Error al generar el informe: {str(e)}")

    if st.session_state.informe_md:
        word_bytes_top = convertir_markdown_a_docx_bytes(
            st.session_state.informe_md,
            titulo_portada="INFORME DE AUDITORÍA PROCESAL Y PLAN DE ACCIÓN"
        )
        st.download_button(
            label="📥 Descargar Informe de Auditoría en Word (.docx)",
            data=word_bytes_top,
            file_name=f"Informe_Auditoria_Expediente_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        st.divider()
        st.markdown(st.session_state.informe_md)

# =====================================================================
# PESTAÑA 3: CAJA DE HERRAMIENTAS - CONSULTOR Y ESTRATEGA DESDE CERO
# =====================================================================
with tab_consultor:
    st.markdown("""
    <div class="tool-box">
        <h3 style="margin-top:0; color:#1e3a8a;">🧰 Caja de Herramientas: Consultor Jurídico y Generador de Actuaciones Desde Cero</h3>
        <p style="margin-bottom:0;">
            ¿Tienes un inconveniente jurídico, administrativo, inmobiliario, de salud, servicios públicos o ante cualquier entidad y <b>quieres saber qué hacer según las leyes colombianas o redactar una actuación que aún no has realizado</b>?
            Describe tu caso aquí abajo. Funciona <b>incluso sin subir documentos</b>, o puedes combinarlo con los PDFs que hayas subido en la Pestaña 1.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_c1, col_c2 = st.columns([2, 1])

    with col_c1:
        problema_descrito = st.text_area(
            "✍️ Describe aquí tu inconveniente, caso o lo que necesitas lograr (incluye hechos, fechas aproximadas, entidad o persona involucrada y qué te han dicho):",
            placeholder="Ejemplo 1: Compré un lote en un proyecto urbanístico donde prometieron servicios públicos de acueducto, energía y zonas comunes, pero han pasado años y no han cumplido ni entregado las obras urbanísticas. ¿Qué acciones debo iniciar ante la SIC, Alcaldía o Procuraduría/Personería y en qué orden?\n\nEjemplo 2: Una entidad pública o juzgado no ha respondido mi solicitud o tomó una decisión ignorando una prueba clave. ¿Qué recurso, queja disciplinaria o acción procede?",
            height=190
        )

    with col_c2:
        documento_a_generar = st.selectbox(
            "📝 ¿Qué documento o actuación quieres que te redacte listo para radicar?",
            options=[
                "Diagnóstico Estratégico + Borrador del Escrito Más Efectivo",
                "Derecho de Petición Blindado (Ley 1755 de 2015)",
                "Demanda o Queja ante Superintendencia (SIC / Supersalud / Superservicios)",
                "Queja Disciplinaria o Solicitud de Vigilancia (Procuraduría / Personería / Comisión Disciplina)",
                "Acción de Tutela con Solicitud de Medida Provisional",
                "Incidente de Desacato / Memorial de Impulso Procesal",
                "Recurso de Reposición y en Subsidio de Apelación",
                "Reclamación Directa Previa / Requerimiento Prejudicial",
                "Querella Policiva / Solicitud de Verificación Urbanística ante Alcaldía/Planeación",
                "Denuncia Penal formal ante la Fiscalía General de la Nación"
            ]
        )

        incluir_docs_cargados = st.checkbox(
            "📎 Apoyarse también en los PDFs subidos en la Pestaña 1 (si subiste alguno)",
            value=bool(st.session_state.expediente_docs)
        )

    btn_consultar_problema = st.button("⚖️ Analizar Problema con Leyes Colombianas y Redactar Documento", type="primary", use_container_width=True)

    if btn_consultar_problema:
        if not api_key_input:
            st.error("❌ Ingresa tu Google Gemini API Key en la barra lateral izquierda.")
        elif not problema_descrito or len(problema_descrito.strip()) < 15:
            st.warning("⚠️ Por favor describe tu problema o caso en el cuadro de texto para poder darte una estrategia precisa.")
        else:
            client = obtener_cliente_gemini(api_key_input)
            prompt_estrategia = construir_prompt_consultor_desde_cero(
                problema_descrito=problema_descrito,
                tipo_proceso=tipo_proceso,
                entidad_especifica=entidad_especifica,
                rol_usuario=rol_usuario,
                documento_a_generar=documento_a_generar,
                incluir_docs_cargados=incluir_docs_cargados,
                items=st.session_state.expediente_docs,
                limite_chars_por_doc=limite_chars
            )

            with st.spinner("⚖️ Consultando legislación colombiana, diseñando hoja de ruta procesal y redactando tu escrito..."):
                try:
                    texto_estrategia, modelo_usado = llamar_gemini(
                        client=client,
                        prompt_contents=prompt_estrategia,
                        modelo_preferido=modelo_sel,
                        temperatura=0.15
                    )
                    st.session_state.estrategia_md = texto_estrategia
                    st.success(f"✅ Estrategia jurídica y borrador generados con éxito (`{modelo_usado}`). Descárgalo en Word aquí abajo.")
                except Exception as e:
                    st.error(f"❌ Error al generar la consultoría: {str(e)}")

    if st.session_state.estrategia_md:
        word_estrategia = convertir_markdown_a_docx_bytes(
            st.session_state.estrategia_md,
            titulo_portada="CONCEPTO JURÍDICO ESTRATÉGICO Y BORRADOR DE ACTUACIÓN"
        )
        st.download_button(
            label="📥 Descargar Estrategia y Borrador en Word (.docx)",
            data=word_estrategia,
            file_name=f"Estrategia_Juridica_Colombia_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        st.divider()
        st.markdown(st.session_state.estrategia_md)

# =====================================================================
# PESTAÑA 4: CONSULTOR DIRECTO AL EXPEDIENTE (CHAT CON TRAZABILIDAD)
# =====================================================================
with tab_rag:
    st.subheader("💬 Chat Jurídico Interactivo (Sobre tus PDFs o sobre Leyes Colombianas)")
    st.caption("Haz preguntas puntuales sobre cualquier folio de los PDFs cargados o consulta dudas procesales adicionales.")

    for msj in st.session_state.historial_chat:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    pregunta_chat = st.chat_input("Escribe tu pregunta (ej. ¿Qué artículo de la Ley 1480 o del CGP aplica aquí, o en qué página está tal oficio?)...")

    if pregunta_chat:
        if not api_key_input:
            st.error("❌ Configura tu API Key en la barra lateral.")
        else:
            st.session_state.historial_chat.append({"role": "user", "content": pregunta_chat})
            with st.chat_message("user"):
                st.markdown(pregunta_chat)

            with st.chat_message("assistant"):
                with st.spinner("Analizando consulta jurídica..."):
                    client = obtener_cliente_gemini(api_key_input)
                    if st.session_state.expediente_docs:
                        contexto_docs = "\n\n".join(
                            [f"=== DOCUMENTO: {d['nombre']} ===\n{d['raw'][:limite_chars]}" for d in st.session_state.expediente_docs]
                        )
                    else:
                        contexto_docs = "Sin documentos subidos en esta sesión. Responde conforme a la legislación y jurisprudencia colombiana vigente."

                    prompt_pregunta = f"""
Actúa como abogado litigante colombiano experto.
{MARCO_JURIDICO_COLOMBIA}

Si la respuesta proviene de los documentos cargados, cita siempre `(Archivo: nombre.pdf, Pág. X)` y transcribe entre comillas el fragmento probatorio.
Si es una consulta estratégica o legal sobre Colombia, cita las normas exactas (Leyes, Decretos, Códigos y Sentencias) y explica paso a paso qué hacer.

CONTEXTO DOCUMENTAL:
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
# PESTAÑA 5: CENTRO DE EXPORTACIÓN (WORD, EXCEL Y MARKDOWN)
# =====================================================================
with tab_export:
    st.subheader("📄 Centro de Descarga de Entregables Profesionales")

    col_exp1, col_exp2, col_exp3 = st.columns(3)

    with col_exp1:
        st.markdown("#### 📘 1. Auditoría de Expediente (`.docx`)")
        st.write("Informe completo con tablas blindadas, matriz de contradicciones y memorial del expediente.")
        if st.session_state.informe_md:
            docx_file = convertir_markdown_a_docx_bytes(
                st.session_state.informe_md,
                titulo_portada="INFORME DE AUDITORÍA PROCESAL Y PLAN DE ACCIÓN"
            )
            st.download_button(
                label="⬇️ Descargar Auditoría en Word (.docx)",
                data=docx_file,
                file_name=f"01_INFORME_AUDITORIA_EXPEDIENTE_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        else:
            st.info("Genera primero la auditoría en la Pestaña 2.")

    with col_exp2:
        st.markdown("#### 🧭 2. Estrategia Desde Cero (`.docx`)")
        st.write("Dictamen de la Caja de Herramientas con ruta legal paso a paso y borrador para actuaciones nuevas.")
        if st.session_state.estrategia_md:
            docx_est = convertir_markdown_a_docx_bytes(
                st.session_state.estrategia_md,
                titulo_portada="CONCEPTO JURÍDICO ESTRATÉGICO Y BORRADOR DE ACTUACIÓN"
            )
            st.download_button(
                label="⬇️ Descargar Estrategia en Word (.docx)",
                data=docx_est,
                file_name=f"02_ESTRATEGIA_JURIDICA_COLOMBIA_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
        else:
            st.info("Usa la Caja de Herramientas en la Pestaña 3 para habilitar esta descarga.")

    with col_exp3:
        st.markdown("#### 📊 3. Matriz de Control en Excel (`.xlsx`)")
        st.write("Libro Excel estilizado con Catálogo de Documentos (SHA-256) y Cronología de Actuaciones corregida.")
        if not st.session_state.catalogo_df.empty:
            excel_bytes = generar_excel_formateado_bytes(
                st.session_state.catalogo_df,
                st.session_state.actuaciones_df
            )
            st.download_button(
                label="⬇️ Descargar Matriz en Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"03_MATRIZ_CONTROL_EXPEDIENTE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Procesa documentos en la Pestaña 1 para habilitar el Excel.")
