import os
import re
import io
import gc
import time
import hashlib
from datetime import datetime, date, timedelta
from pathlib import Path

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
from google import genai
from google.genai import types

try:
    import streamlit.components.v1 as components
except Exception:
    components = None

try:
    from dateparser import parse as parse_fecha_es
except Exception:
    parse_fecha_es = None

try:
    from openpyxl.styles import Font, PatternFill, Alignment
    ESTILOS_EXCEL_OK = True
except Exception:
    ESTILOS_EXCEL_OK = False

# =====================================================================
# CONFIGURACIÓN DE PÁGINA Y ESCUDO ANTI-TRADUCTOR (EVITA ERROR removeChild)
# =====================================================================
st.set_page_config(
    page_title="Bufete Litigante IA Colombia | Auditoría y Estrategia",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
if components is not None:
    try:
        components.html(
            """
            <script>
                try {
                    const doc = window.parent.document;
                    doc.documentElement.setAttribute('lang', 'es');
                    doc.documentElement.setAttribute('translate', 'no');
                    doc.documentElement.classList.add('notranslate');
                    if (doc.body) {
                        doc.body.setAttribute('translate', 'no');
                        doc.body.classList.add('notranslate');
                    }
                    if (!doc.querySelector('meta[name="google"]')) {
                        const m = doc.createElement('meta');
                        m.name = 'google';
                        m.content = 'notranslate';
                        doc.head.appendChild(m);
                    }
                } catch (e) {}
            </script>
            """,
            height=0,
            width=0
        )
    except Exception:
        pass

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 65%, #1d4ed8 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 0.9rem;
        border-bottom: 4px solid #f59e0b;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.14);
    }
    .status-bar {
        display: flex;
        flex-wrap: wrap;
        gap: 0.65rem;
        margin-bottom: 1.1rem;
    }
    .status-pill {
        background-color: #f1f5f9;
        border: 1px solid #cbd5e1;
        color: #1e293b;
        padding: 0.32rem 0.85rem;
        border-radius: 999px;
        font-size: 0.83rem;
        font-weight: 600;
    }
    .status-pill-ok {
        background-color: #ecfdf5;
        border: 1px solid #6ee7b7;
        color: #065f46;
    }
    .status-pill-warn {
        background-color: #fffbeb;
        border: 1px solid #fcd34d;
        color: #92400e;
    }
    .tool-box {
        background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid #cbd5e1;
        border-left: 5px solid #1e3a8a;
        padding: 1.15rem;
        border-radius: 10px;
        margin-bottom: 1.1rem;
    }
    .gold-box {
        background-color: #fffbeb;
        border: 1px solid #fde68a;
        border-left: 5px solid #d97706;
        padding: 1.05rem;
        border-radius: 10px;
        color: #78350f;
        margin-bottom: 1rem;
    }
    div[data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-top: 3px solid #1e3a8a;
        padding: 0.75rem 1rem;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# ESTADO DE SESIÓN ULTRA-LIGERO EN MEMORIA RAM (EVITA CAÍDAS 'Oh, no')
# =====================================================================
def inicializar_sesion():
    if "expediente_docs" not in st.session_state:
        st.session_state.expediente_docs = []
    if "cache_hashes" not in st.session_state:
        st.session_state.cache_hashes = {}
    if "actuaciones_df" not in st.session_state:
        st.session_state.actuaciones_df = pd.DataFrame()
    if "catalogo_df" not in st.session_state:
        st.session_state.catalogo_df = pd.DataFrame()
    if "informe_md" not in st.session_state:
        st.session_state.informe_md = ""
    if "estrategia_md" not in st.session_state:
        st.session_state.estrategia_md = ""
    if "blindaje_md" not in st.session_state:
        st.session_state.blindaje_md = ""
    if "historial_chat" not in st.session_state:
        st.session_state.historial_chat = []
    if "modelos_cache" not in st.session_state:
        st.session_state.modelos_cache = {}

inicializar_sesion()

def obtener_api_key_inicial():
    try:
        if "GOOGLE_API_KEY" in st.secrets:
            return str(st.secrets["GOOGLE_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GOOGLE_API_KEY", "").strip()

def df_a_markdown_seguro(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "Sin registros detectados."
    cols = [str(c) for c in df.columns]
    header = "| " + " | ".join(cols) + " |"
    sep = "| " + " | ".join(["---"] * len(cols)) + " |"
    filas = [header, sep]
    for _, row in df.iterrows():
        celdas = [re.sub(r'\s+', ' ', str(row[c]).replace("|", "-")).strip() for c in df.columns]
        filas.append("| " + " | ".join(celdas) + " |")
    return "\n".join(filas)

# =====================================================================
# EXPRESIONES REGULARES PARA METADATOS Y FECHAS
# =====================================================================
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
# CALCULADORA OFICIAL DE TÉRMINOS Y FESTIVOS EN COLOMBIA (LEY EMILIANI)
# =====================================================================
def calcular_pascua(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)

def trasladar_a_lunes(d: date) -> date:
    if d.weekday() == 0:
        return d
    return d + timedelta(days=(7 - d.weekday()))

def obtener_festivos_colombia(year: int) -> set:
    fijos = [
        date(year, 1, 1),
        date(year, 5, 1),
        date(year, 7, 20),
        date(year, 8, 7),
        date(year, 12, 8),
        date(year, 12, 25),
    ]
    emiliani = [
        trasladar_a_lunes(date(year, 1, 6)),
        trasladar_a_lunes(date(year, 3, 19)),
        trasladar_a_lunes(date(year, 6, 29)),
        trasladar_a_lunes(date(year, 8, 15)),
        trasladar_a_lunes(date(year, 10, 12)),
        trasladar_a_lunes(date(year, 11, 1)),
        trasladar_a_lunes(date(year, 11, 11)),
    ]
    dom_pascua = calcular_pascua(year)
    religiosos = [
        dom_pascua - timedelta(days=3),
        dom_pascua - timedelta(days=2),
        trasladar_a_lunes(dom_pascua + timedelta(days=39)),
        trasladar_a_lunes(dom_pascua + timedelta(days=60)),
        trasladar_a_lunes(dom_pascua + timedelta(days=68)),
    ]
    return set(fijos + emiliani + religiosos)

def sumar_dias_habiles_colombia(fecha_inicio: date, dias_habiles: int):
    actual = fecha_inicio
    contados = 0
    festivos_omitidos = []
    fines_semana_omitidos = 0
    ruta_dias = []

    while contados < dias_habiles:
        actual += timedelta(days=1)
        festivos_anio = obtener_festivos_colombia(actual.year)
        if actual.weekday() >= 5:
            fines_semana_omitidos += 1
            continue
        if actual in festivos_anio:
            festivos_omitidos.append(actual.strftime("%d/%m/%Y"))
            continue
        contados += 1
        ruta_dias.append({"Día Hábil N°": contados, "Fecha": actual.strftime("%d/%m/%Y")})

    return actual, fines_semana_omitidos, festivos_omitidos, pd.DataFrame(ruta_dias)

# =====================================================================
# CLIENTE GEMINI CON AUTO-DETECCIÓN LIGERA EN SESIÓN
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

def obtener_cola_modelos(client, api_key: str):
    key_fp = hashlib.sha256(api_key.strip().encode("utf-8")).hexdigest()[:12]
    if key_fp in st.session_state.modelos_cache:
        return list(st.session_state.modelos_cache[key_fp])

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

    ordenados = sorted(list(dict.fromkeys(modelos_cuenta)), key=puntaje_modelo)
    cola = ordenados + [m for m in MODELOS_RESPALDO if m not in ordenados]
    st.session_state.modelos_cache[key_fp] = cola
    return list(cola)

def llamar_gemini(client, api_key: str, prompt_contents, modelo_preferido="Auto-Detectar (Recomendado)", temperatura=0.1, max_reintentos=2):
    cola_modelos = obtener_cola_modelos(client, api_key)
    if modelo_preferido and modelo_preferido != "Auto-Detectar (Recomendado)":
        cola_modelos = [modelo_preferido] + [m for m in cola_modelos if m != modelo_preferido]

    errores = []
    for modelo in cola_modelos:
        for intento in range(max_reintentos):
            try:
                resp = client.models.generate_content(
                    model=modelo,
                    contents=prompt_contents,
                    config=types.GenerateContentConfig(temperature=temperatura)
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
                        time.sleep(3)
                        continue
                    break
                time.sleep(1.5)

    raise RuntimeError("No fue posible obtener respuesta. Últimos intentos: " + " | ".join(errores[-4:]))

# =====================================================================
# LECTURA ULTRA-LIGERA EN RAM Y AHORRO DE CPU
# =====================================================================
PATRON_BLOQUE_FECHA = re.compile(
    r'(?m)^(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-](?:\d{4}|\d{2}))[ \t]*(?:[-–:][ \t]*)?(?P<encabezado>[\w\t \(\)\.,#º°\-]{8,110})$'
)
PATRON_ACTO_FECHA = re.compile(
    r'(?i)(auto|providencia|resoluci[oó]n|sentencia|acta|oficio|notificaci[oó]n|audiencia|memorial|radicado|queja|recurso|orden|f[oó]rmula)[ \t]+(?:no\.?|n[o°]?)?[ \t]*[\w\-]*[ \t]*(?:del|de fecha|:)?[ \t]*(?P<fecha>\d{1,2}[\/\-]\d{1,2}[\/\-](?:\d{4}|\d{2}))'
)

def ocr_pagina_ligero(page, api_key: str, client=None, modo_ocr="Nube Gemini Vision (0% CPU Local - Recomendado)"):
    if modo_ocr.startswith("Desactivado"):
        return "", "Omitido"

    pix = page.get_pixmap(dpi=110, alpha=False)
    img_bytes = pix.tobytes("jpeg")
    del pix
    img = Image.open(io.BytesIO(img_bytes))

    if modo_ocr.startswith("Nube") and client is not None and api_key:
        try:
            texto_ia, _ = llamar_gemini(
                client=client,
                api_key=api_key,
                prompt_contents=[
                    img,
                    "Transcribe de forma exacta y literal todo el texto visible en esta página jurídica/administrativa, incluyendo sellos, fechas, radicados y tablas. No agregues comentarios."
                ],
                temperatura=0.0,
                max_reintentos=1
            )
            img.close()
            del img_bytes
            if texto_ia:
                return texto_ia.strip(), "OCR-GeminiCloud"
        except Exception:
            pass

    try:
        import pytesseract
        texto_ocr = pytesseract.image_to_string(img, lang="spa").strip()
        img.close()
        del img_bytes
        if len(texto_ocr) > 20:
            return texto_ocr, "OCR-Tesseract"
    except Exception:
        try:
            img.close()
        except Exception:
            pass

    return "", "SIN_TEXTO"

def limpiar_linea_meta(txt: str, max_len=110):
    return re.sub(r'\s+', ' ', str(txt)).strip()[:max_len]

def extraer_metadatos_documento(texto: str):
    muestra = texto[:60000]
    meta = {
        'radicado': 'No identificado',
        'actor': 'No identificado',
        'demandado': 'No identificado',
        'juez': 'No identificado',
        'tipo_actuacion': [],
        'numero_oficio': 'No identificado'
    }
    rr = RE_RADICADO.search(muestra)
    if rr:
        meta['radicado'] = limpiar_linea_meta(rr.group(0), 70)
    no_of = RE_NUMERO_OFICIO.search(muestra)
    if no_of:
        meta['numero_oficio'] = limpiar_linea_meta(no_of.group(0), 60)
    da = RE_ACTOR.search(muestra)
    if da:
        meta['actor'] = limpiar_linea_meta(da.group(2), 110)
    dd = RE_DEMANDADO.search(muestra)
    if dd:
        meta['demandado'] = limpiar_linea_meta(dd.group(2), 110)
    jj = RE_JUEZ.search(muestra)
    if jj:
        meta['juez'] = limpiar_linea_meta(jj.group(0), 110)

    texto_upper = muestra.upper()
    for tipo in TIPOS_ACT:
        if tipo in texto_upper:
            meta['tipo_actuacion'].append(tipo)
    return meta

def extraer_fechas_de_pagina(nombre_archivo: str, num_pag: int, txt_pag: str) -> list:
    hallazgos = []
    for m in PATRON_BLOQUE_FECHA.finditer(txt_pag):
        hallazgos.append({
            "Archivo": nombre_archivo,
            "Página": num_pag,
            "Fecha Literal": m.group("fecha").strip(),
            "Actuación / Referencia": re.sub(r'\s+', ' ', m.group("encabezado")).strip(),
            "Contexto": re.sub(r'\s+', ' ', txt_pag[max(0, m.start() - 25):min(len(txt_pag), m.end() + 120)]).strip()
        })
    for m in PATRON_ACTO_FECHA.finditer(txt_pag):
        hallazgos.append({
            "Archivo": nombre_archivo,
            "Página": num_pag,
            "Fecha Literal": m.group("fecha").strip(),
            "Actuación / Referencia": re.sub(r'\s+', ' ', m.group(0)).strip(),
            "Contexto": re.sub(r'\s+', ' ', txt_pag[max(0, m.start() - 25):min(len(txt_pag), m.end() + 120)]).strip()
        })
    return hallazgos

def procesar_archivo_ligero(uploaded_file, api_key: str, client=None, modo_ocr="Nube Gemini Vision (0% CPU Local - Recomendado)"):
    file_bytes = uploaded_file.getvalue()
    hash_sha256 = hashlib.sha256(file_bytes).hexdigest()[:16]
    clave_cache = f"{hash_sha256}_{modo_ocr}"

    if clave_cache in st.session_state.cache_hashes:
        del file_bytes
        return st.session_state.cache_hashes[clave_cache]

    nombre = uploaded_file.name
    ext = Path(nombre).suffix.lower()
    bloques_texto = []
    actuaciones_doc = []
    num_paginas = 0
    pags_ocr = 0
    pags_sin_texto = 0

    try:
        if ext == ".pdf":
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            num_paginas = len(doc)
            for i in range(num_paginas):
                page = doc[i]
                num_pag = i + 1
                txt = page.get_text("text", sort=True).strip()
                if len(txt) >= 35:
                    bloque = f"[Archivo: {nombre} - Pág. {num_pag}]\n{txt}"
                else:
                    txt_ocr, met_ocr = ocr_pagina_ligero(page, api_key=api_key, client=client, modo_ocr=modo_ocr)
                    if txt_ocr:
                        pags_ocr += 1
                        bloque = f"[Archivo: {nombre} - Pág. {num_pag} ({met_ocr})]\n{txt_ocr}"
                    else:
                        pags_sin_texto += 1
                        bloque = f"[Archivo: {nombre} - Pág. {num_pag} (SIN TEXTO EXTRAÍDO)]"
                bloques_texto.append(bloque)
                actuaciones_doc.extend(extraer_fechas_de_pagina(nombre, num_pag, bloque))
            doc.close()

        elif ext == ".docx":
            doc = Document(io.BytesIO(file_bytes))
            parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    celdas = [c.text.strip().replace("\n", " ") for c in row.cells]
                    if any(celdas):
                        parrafos.append(" - ".join(celdas))
            texto_docx = "\n".join(parrafos)
            tam_pag = 3000
            secciones = [texto_docx[k:k + tam_pag] for k in range(0, max(1, len(texto_docx)), tam_pag)]
            num_paginas = len(secciones)
            for idx_s, frag in enumerate(secciones, 1):
                bloque = f"[Archivo: {nombre} - Sección {idx_s}]\n{frag}"
                bloques_texto.append(bloque)
                actuaciones_doc.extend(extraer_fechas_de_pagina(nombre, idx_s, bloque))

        elif ext in (".txt", ".rtf"):
            contenido = file_bytes.decode("utf-8", errors="ignore").strip()
            num_paginas = 1
            bloque = f"[Archivo: {nombre} - Pág. 1]\n{contenido}"
            bloques_texto.append(bloque)
            actuaciones_doc.extend(extraer_fechas_de_pagina(nombre, 1, bloque))

    except Exception as e:
        num_paginas = 1
        bloques_texto.append(f"[Archivo: {nombre} - ERROR DE LECTURA: {str(e)}]")

    del file_bytes
    texto_unificado = "\n\n".join(bloques_texto)
    del bloques_texto
    meta = extraer_metadatos_documento(texto_unificado)

    item = {
        "nombre": nombre,
        "hash": hash_sha256,
        "num_paginas": num_paginas,
        "pags_ocr": pags_ocr,
        "pags_sin_texto": pags_sin_texto,
        "caracteres": len(texto_unificado),
        "meta": meta,
        "actuaciones": actuaciones_doc,
        "raw": texto_unificado
    }

    if len(st.session_state.cache_hashes) >= 35:
        st.session_state.cache_hashes.clear()
    st.session_state.cache_hashes[clave_cache] = item
    return item

def consolidar_cronologia_df(items: list) -> pd.DataFrame:
    todas = []
    for it in items:
        todas.extend(it.get("actuaciones", []))
    if not todas:
        return pd.DataFrame(columns=["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"])

    def normalizar_fecha(f_str):
        texto_f = str(f_str).strip()
        if parse_fecha_es is not None:
            try:
                d = parse_fecha_es(texto_f, languages=["es"], settings={"DATE_ORDER": "DMY"})
                if d and 1990 <= d.year <= 2035:
                    return d.date().isoformat()
            except Exception:
                pass
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                d = datetime.strptime(texto_f, fmt).date()
                if 1990 <= d.year <= 2035:
                    return d.isoformat()
            except Exception:
                pass
        return "9999-12-31"

    df_act = pd.DataFrame(todas)
    df_act["Fecha ISO"] = df_act["Fecha Literal"].apply(normalizar_fecha)
    df_act = df_act.drop_duplicates(subset=["Archivo", "Página", "Fecha Literal", "Actuación / Referencia"])
    df_act = df_act.sort_values(by=["Fecha ISO", "Archivo", "Página"]).reset_index(drop=True)
    df_act["Fecha ISO"] = df_act["Fecha ISO"].replace("9999-12-31", "Sin Fecha Exacta")
    return df_act[["Fecha ISO", "Fecha Literal", "Archivo", "Página", "Actuación / Referencia", "Contexto"]]

# =====================================================================
# GENERADOR DE WORD (.DOCX) Y EXCEL (.XLSX) BAJO DEMANDA
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
        if len(fila) > num_cols:
            fila_ajustada = fila[:num_cols - 1] + [" - ".join(fila[num_cols - 1:])]
        else:
            fila_ajustada = fila + [""] * (num_cols - len(fila))

        for c_idx in range(num_cols):
            celda = tabla.cell(r_idx, c_idx)
            celda.text = ""
            p = celda.paragraphs[0]
            agregar_texto_con_negrillas(p, fila_ajustada[c_idx])

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

def extraer_solo_borrador_radicable(md_texto: str) -> str:
    patron_inicio = re.compile(r'(?im)^##\s*(?:10|7)\.\s*BORRADOR.*$')
    m_ini = patron_inicio.search(md_texto)
    if not m_ini:
        return ""
    resto = md_texto[m_ini.end():].strip()
    patron_fin = re.compile(r'(?im)^##\s*(?:11|8)\.\s*CHECKLIST.*$')
    m_fin = patron_fin.search(resto)
    if m_fin:
        resto = resto[:m_fin.start()].strip()
    return resto

def convertir_markdown_a_docx_bytes(md_texto: str, titulo_portada="INFORME JURÍDICO Y PLAN DE ACCIÓN PROCESAL", incluir_portada=True) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    if incluir_portada:
        titulo = doc.add_heading(titulo_portada, level=0)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = sub.add_run(f"Generado el {datetime.now().strftime('%d/%m/%Y a las %H:%M')} | Bufete Litigante IA Colombia")
        run_sub.italic = True
        run_sub.font.size = Pt(9.5)
        run_sub.font.color.rgb = RGBColor(100, 116, 139)
        doc.add_paragraph("_" * 70)

    lineas = md_texto.split('\n')
    buffer_tabla = []

    for linea in lineas:
        linea_limpia = linea.strip()
        es_fila_tabla = linea_limpia.startswith('|') and linea_limpia.endswith('|') and len(linea_limpia) > 2

        if es_fila_tabla:
            contenido_sin_barras = linea_limpia.replace('|', '').replace('-', '').replace(':', '').strip()
            if not contenido_sin_barras:
                continue
            buffer_tabla.append(dividir_fila_markdown_segura(linea_limpia))
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
    return buffer_word.getvalue()

def generar_excel_formateado_bytes(catalogo_df: pd.DataFrame, actuaciones_df: pd.DataFrame) -> bytes:
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
        catalogo_df.to_excel(writer, sheet_name="01_Catalogo_Documentos", index=False)
        actuaciones_df.to_excel(writer, sheet_name="02_Cronologia_Actuaciones", index=False)

        if ESTILOS_EXCEL_OK:
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
                    ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 14), 55)
                    for cell in col[1:]:
                        cell.alignment = Alignment(vertical="top", wrap_text=True)

    return buffer_excel.getvalue()

# =====================================================================
# PROMPTS JURÍDICOS MULTIJURISDICCIONALES COLOMBIA
# =====================================================================
MARCO_JURIDICO_COLOMBIA = """
DOMINIO NORMATIVO Y PROCESAL COLOMBIANO OBLIGATORIO SEGÚN LA ENTIDAD:
- Constitucional y Salud: Constitución Política (Arts. 11, 23, 29, 49, 86), Decreto 2591 de 1991 (Tutela, Medida Provisional Art. 7, Impugnación Art. 31, Cumplimiento Art. 27 y Desacato Art. 52), Ley Estatutaria 1751 de 2015, Sentencias SU-034 de 2018 y T-760 de 2008.
- Derecho de Petición y Procedimiento Administrativo: Ley 1755 de 2015 (15 días hábiles petición general, 10 días documentos/información, 30 días consultas) y CPACA (Ley 1437 de 2011 modificada por Ley 2080 de 2021).
- Protección al Consumidor, Constructoras e Inmobiliarias (SIC): Ley 1480 de 2011 (Estatuto del Consumidor: publicidad engañosa Arts. 29 y 30, efectividad de la garantía inmobiliaria y de bienes Arts. 7 a 16, información mínima Art. 23, cláusulas abusivas Arts. 42 y 43, acción jurisdiccional Art. 58), Circular Única de la SIC.
- Control Disciplinario, Urbano y Ministerio Público (Procuraduría, Personerías, Alcaldías): Ley 1952 de 2019 y Ley 2094 de 2021 (Código General Disciplinario), Ley 388 de 1997 y Decreto 1077 de 2015 (Ordenamiento Territorial, licencias urbanísticas, entregas de zonas de cesión y espacio público), Ley 1801 de 2016 (Código Nacional de Seguridad y Convivencia Ciudadana - Querellas Policivas), Ley 142 de 1994 (Servicios Públicos Domiciliarios y silencio administrativo positivo).
- Procesal Civil, Comercial y Propiedad Horizontal: Ley 1564 de 2012 (Código General del Proceso - CGP), Ley 2213 de 2022 (Justicia Digital y traslado previo de escritos), Código Civil, Código de Comercio, Ley 675 de 2001.
- Laboral y Función Pública / Concursos de Méritos: CST, CPTSS, Ley 909 de 2004, Decretos reglamentarios CNSC.
- Penal y Disciplinario Judicial: Ley 599 de 2000 (Código Penal: fraude a resolución judicial Art. 454, fraude procesal Art. 453, prevaricato por acción u omisión Arts. 413-414, estafa Art. 246, urbanización ilegal Art. 318), Ley 906 de 2004, Comisión Nacional de Disciplina Judicial.
"""

def formatear_datos_firmante(datos_firmante: dict) -> str:
    if not any(datos_firmante.values()):
        return "El usuario no ingresó datos personales fijos; deja espacios claros entre corchetes solo donde falte cédula, correo o teléfono."
    return (
        f"- Nombre del Firmante: {datos_firmante.get('nombre') or '[Por completar]'}\n"
        f"- Identificación / C.C. / T.P.: {datos_firmante.get('documento') or '[Por completar]'}\n"
        f"- Municipio / Ciudad: {datos_firmante.get('ciudad') or '[Por completar]'}\n"
        f"- Correo de Notificación Judicial: {datos_firmante.get('correo') or '[Por completar]'}\n"
        f"- Teléfono / Celular: {datos_firmante.get('telefono') or '[Por completar]'}"
    )

def construir_prompt_auditoria(
    items,
    catalogo_df,
    actuaciones_df,
    tipo_proceso,
    entidad_especifica,
    rol_usuario,
    enfoque_especial,
    datos_firmante,
    limite_chars_por_doc=120000
):
    total_docs = len(items)
    total_pags = sum(it["num_paginas"] for it in items)
    total_ocr = sum(it["pags_ocr"] for it in items)

    catalogo_md = df_a_markdown_seguro(catalogo_df)
    actos_md = df_a_markdown_seguro(actuaciones_df.head(80))

    bloques_documentales = []
    for idx, it in enumerate(items, 1):
        texto_doc = it["raw"]
        nota_recorte = ""
        if len(texto_doc) > limite_chars_por_doc:
            mitad = limite_chars_por_doc // 2
            texto_doc = (
                texto_doc[:mitad]
                + f"\n\n[... SECCIÓN INTERMEDIA OMITIDA POR EXTENSIÓN ({len(it['raw'])} chars totales) ...]\n\n"
                + texto_doc[-mitad:]
            )
            nota_recorte = f" (Optimizado a {limite_chars_por_doc} caracteres)"

        bloque = f"""
=====================================================================
DOCUMENTO #{idx}: {it['nombre']}{nota_recorte}
- Hash SHA-256: {it['hash']}
- Páginas totales: {it['num_paginas']} (Páginas con OCR: {it['pags_ocr']})
- Radicado detectado: {it['meta']['radicado']}
- Tipo de actuación: {', '.join(it['meta']['tipo_actuacion']) or 'General'}
- Actor / Solicitante: {it['meta']['actor']}
- Demandado / Entidad: {it['meta']['demandado']}
=====================================================================
{texto_doc}
"""
        bloques_documentales.append(bloque)

    expediente_completo_str = "\n\n".join(bloques_documentales)

    return f"""
Actúa como un **Abogado Litigante Senior, Auditor Forense de Expedientes y Estratega Procesal Colombiano** experto ante cualquier despacho judicial, superintendencia, alcaldía u órgano de control en Colombia.

{MARCO_JURIDICO_COLOMBIA}

CONTEXTO ESTRATÉGICO DEL CASO:
- **Jurisdicción / Tipo de Trámite:** {tipo_proceso}
- **Entidad / Despacho / Autoridad Específica:** {entidad_especifica or 'Determinar según los documentos del expediente'}
- **Posición / Rol que representamos:** {rol_usuario}
- **Enfoque Especial Solicitado:** {enfoque_especial or 'Auditar todo el expediente, detectar contradicciones entre los PDFs, identificar errores o incumplimientos y definir qué hacer y qué radicar a continuación.'}

DATOS DEL SUSCRIPTOR PARA EL BORRADOR DE LA SECCIÓN 10:
{formatear_datos_firmante(datos_firmante)}

REGLAS TÉCNICAS Y DE FORMATO OBLIGATORIAS:
1. **Inventario Real:** Analizas {total_docs} documentos que suman {total_pags} páginas ({total_ocr} procesadas mediante OCR).
2. **Formato de Citación (PROHIBIDO USAR BARRAS VERTICALES `|` DENTRO DE CITAS):** Cita siempre usando ÚNICAMENTE `(Archivo: nombre.pdf, Pág. X)`. NUNCA uses el símbolo `|` dentro de paréntesis o corchetes para no alterar las tablas de Word.
3. **Cero Invenciones:** No inventes fechas, radicados ni pruebas. Si falta un soporte clave, escribe: `"NO CONSTA EN LOS DOCUMENTOS APORTADOS (VACÍO PROBATORIO)"`.

INSUMO 1: CATÁLOGO GENERAL DE ARCHIVOS
{catalogo_md}

INSUMO 2: REGISTRO PRELIMINAR DE FECHAS Y ACTUACIONES
{actos_md}

INSUMO 3: CONTENIDO ÍNTEGRO PAGINADO DE LOS DOCUMENTOS DEL EXPEDIENTE
{expediente_completo_str}

---

GENERA EL DICTAMEN JURÍDICO COMPLETO EN FORMATO MARKDOWN CON ESTAS 11 SECCIONES EXACTAS:

# INFORME DE AUDITORÍA PROCESAL, CONTRADICCIONES Y PLAN DE ACCIÓN

## 1. FICHA TÉCNICA E IDENTIFICACIÓN DEL EXPEDIENTE
(Tabla Markdown de 2 columnas `Campo` | `Detalle` con: Radicado/Expediente, Entidad o Despacho, Demandante/Accionante/Quejoso, Demandado/Convocado/Investigado, Naturaleza del Proceso, Etapa Actual y Nivel de Riesgo/Viabilidad).

## 2. RESUMEN EJECUTIVO Y DIAGNÓSTICO PROCESAL
(Síntesis precisa del litigio o trámite, qué se reclama, qué ha respondido la contraparte/entidad y cuál es el nudo crítico actual).

## 3. LÍNEA DE TIEMPO CRONOLÓGICA CON TRAZABILIDAD
(Tabla Markdown con 5 columnas: `Fecha` | `Actuación o Documento` | `Archivo y Página` | `Hecho o Decisión Clave` | `Impacto Procesal`).

## 4. MATRIZ EXHAUSTIVA DE CONTRADICCIONES E INCONSISTENCIAS DETECTADAS
(Tabla Markdown con 6 columnas: `N°` | `Punto en Conflicto` | `Versión o Afirmación A (Cita + Archivo y Pág.)` | `Prueba o Versión B que la contradice (Cita + Archivo y Pág.)` | `Tipo de Contradicción` | `Gravedad y Cómo Explotarla Procesalmente`).
Después de la tabla, desarrolla a fondo cada contradicción y su consecuencia jurídica (confesión, defecto fáctico, falsa motivación, publicidad engañosa, fraude procesal, etc.).

## 5. AUDITORÍA DE PRUEBAS: HECHOS PROBADOS VS. VACÍOS PROBATORIOS
(Separa hechos plenamente probados con su página exacta frente a afirmaciones sin soporte probatorio).

## 6. CONTROL DE ÓRDENES, OBLIGACIONES Y PLAZOS INCUMPLIDOS
(Tabla Markdown con 5 columnas: `Orden / Compromiso / Norma` | `Obligado o Responsable` | `Plazo Legal o Fijado` | `Estado Actual` | `Evidencia en el Expediente`).

## 7. FUNDAMENTACIÓN JURÍDICA, LEYES Y JURISPRUDENCIA APLICABLE EN COLOMBIA
(Cita artículos exactos de leyes, códigos, decretos y sentencias aplicables en Colombia).

## 8. QUÉ SE DEBE HACER AHORA MISMO: RUTA PROCESAL PASO A PASO
(1. Actuación Principal Inmediata de 0 a 3 días; 2. Actuaciones Complementarias o de Presión Institucional ante Superintendencias, Procuraduría, Personería, Fiscalía o Comisión de Disciplina; 3. Términos legales y precauciones procesales como envío simultáneo Ley 2213 de 2022).

## 9. TABLA DE PLAN DE ACCIÓN PRIORIZADO
(Tabla Markdown con 5 columnas: `Prioridad` | `Acción Jurídica Concreta` | `Entidad / Destinatario` | `Objetivo Procesal` | `Sustento Normativo y Probatorio`).

## 10. BORRADOR COMPLETO DE MEMORIAL / DEMANDA / QUEJA / OFICIO LISTO PARA RADICAR
(IMPORTANTE: Redacta aquí ÚNICAMENTE el texto formal, limpio, extenso y completo del documento listo para radicar ante la autoridad, sin comentarios internos de auditoría, incluyendo Encabezado, Referencia, Identificación del suscriptor, Hechos numerados, Contradicciones/Incumplimientos probados, Fundamentos de Derecho, Pruebas/Anexos, Peticiones principales y subsidiarias, Notificaciones y Firma).

## 11. CHECKLIST DE VERIFICACIÓN ANTES DE RADICAR
(Lista de chequeo práctica de anexos, firmas, traslado previo y correos antes de radicar).
"""

def construir_prompt_consultor_desde_cero(
    problema_descrito,
    tipo_proceso,
    entidad_especifica,
    rol_usuario,
    documento_a_generar,
    incluir_docs_cargados,
    items,
    datos_firmante,
    limite_chars_por_doc=80000
):
    contexto_expediente = "NO SE ADJUNTARON DOCUMENTOS (CONSULTA PREVENTIVA / DESDE CERO)."
    if incluir_docs_cargados and items:
        partes_docs = []
        for idx, it in enumerate(items, 1):
            partes_docs.append(f"--- DOCUMENTO ADJUNTO #{idx}: {it['nombre']} ---\n{it['raw'][:limite_chars_por_doc]}")
        contexto_expediente = "\n\n".join(partes_docs)

    return f"""
Actúa como un **Abogado Litigante Colombiano Senior y Consultor Estratégico Especialista en todas las jurisdicciones y entidades de Colombia**.
Orienta al usuario con precisión quirúrgica sobre **qué hacer frente a su problema jurídico o actuación por iniciar**, conforme a las leyes colombianas vigentes.

{MARCO_JURIDICO_COLOMBIA}

DATOS DEL PROBLEMA PLANTEADO:
- **Área / Naturaleza del Caso:** {tipo_proceso}
- **Entidad / Autoridad / Contraparte Involucrada:** {entidad_especifica or 'Determinar según ley colombiana'}
- **Rol del Usuario:** {rol_usuario}
- **Documento Principal a Redactar en la Sección 7:** {documento_a_generar}

DATOS DEL SUSCRIPTOR PARA FIRMAR EL DOCUMENTO:
{formatear_datos_firmante(datos_firmante)}

DESCRIPCIÓN DEL PROBLEMA DEL USUARIO:
\"\"\"
{problema_descrito}
\"\"\"

DOCUMENTOS DE SOPORTE EN SESIÓN (SI APLICAN):
{contexto_expediente}

REGLAS DE FORMATO:
- NO uses barras verticales `|` dentro de paréntesis o corchetes de citas.
- Cita leyes, artículos y sentencias exactas de Colombia.

GENERA LA ESTRATEGIA Y EL ESCRITO LISTO PARA RADICAR CON ESTA ESTRUCTURA EN MARKDOWN:

# DICTAMEN DE ESTRATEGIA JURÍDICA Y HOJA DE RUTA (COLOMBIA)

## 1. DIAGNÓSTICO JURÍDICO DEL PROBLEMA Y VIABILIDAD
## 2. ENTIDADES COMPETENTES: ¿ANTE QUIÉN ACUDIR Y POR QUÉ?
(Tabla Markdown de 4 columnas: `Entidad o Despacho Competente` | `Trámite o Acción a Interponer` | `Facultad de la Entidad` | `Plazo Legal de Respuesta`).
## 3. LEYES COLOMBIANAS Y ARTÍCULOS EXACTOS QUE TE RESPALDAN
## 4. REQUISITOS DE PROCEDIBILIDAD, CADUCIDAD Y ERRORES A EVITAR
## 5. RUTA DE ACCIÓN PASO A PASO: QUÉ HACER DESDE HOY
## 6. MATRIZ ESTRATÉGICA DE ACTUACIONES RECOMENDADAS
(Tabla Markdown de 5 columnas: `Orden` | `Actuación a Realizar` | `Ante Qué Entidad` | `Resultado Esperado` | `Fundamento Legal`).
## 7. BORRADOR COMPLETO DE: {documento_a_generar.upper()}
(IMPORTANTE: Redacta aquí ÚNICAMENTE el documento jurídico formal completo y listo para firmar y radicar, con todos los requisitos de ley colombiana).
## 8. CHECKLIST DE PRUEBAS Y ANEXOS NECESARIOS
"""

def construir_prompt_abogado_diablo(borrador_usuario, tipo_proceso, entidad_especifica, datos_firmante):
    return f"""
Actúa simultáneamente en dos fases como:
1) **Juez Estricto / Abogado Senior de la Contraparte en Colombia ("Abogado del Diablo")** que examina el escrito buscando cualquier fisura para inadmitirlo, rechazarlo o ganarle con excepciones.
2) **Magistrado / Litigante Maestro de Blindaje** que corrige cada fisura encontrada y entrega una versión blindada, implacable y lista para radicar.

{MARCO_JURIDICO_COLOMBIA}

CONTEXTO:
- Trámite / Jurisdicción: {tipo_proceso}
- Entidad / Contraparte: {entidad_especifica or 'Según escrito'}
- Datos del Suscriptor:
{formatear_datos_firmante(datos_firmante)}

ESCRITO A SOMETER A PRUEBA DE FUEGO:
\"\"\"
{borrador_usuario}
\"\"\"

ENTREGA EL RESULTADO CON ESTA ESTRUCTURA EN MARKDOWN (SIN USAR `|` DENTRO DE CITAS):

# AUDITORÍA DE BLINDAJE ("ABOGADO DEL DIABLO") Y VERSIÓN REFORZADA

## 1. ATAQUES QUE HARÍA LA CONTRAPARTE O CAUSALES DE INADMISIÓN DEL JUEZ
(Tabla Markdown de 4 columnas: `N°` | `Debilidad o Riesgo Detectado en el Escrito` | `Cómo lo Atacaría la Contraparte o el Juez` | `Solución Jurídica Aplicada para Blindarlo`).

## 2. PRUEBAS O REQUISITOS FORMALES QUE FALTABAN MENCIONAR
(Revisión de juramento estimatorio, medida provisional, agotamiento de reclamación directa/procedibilidad, traslado Ley 2213 de 2022, legitimación por activa/pasiva o competencia).

## 7. BORRADOR COMPLETO DE: ESCRITO 100% BLINDADO Y CORREGIDO LISTO PARA RADICAR
(Redacta aquí la versión final perfeccionada, contundente y blindada del documento, lista para firmar y radicar).

## 8. CHECKLIST FINAL DE RADICACIÓN SEGURA
"""

# =====================================================================
# BARRA LATERAL: CONFIGURACIÓN, ANTI-CPU Y DATOS DEL FIRMANTE
# =====================================================================
with st.sidebar:
    st.header("🔑 Motor IA & Rendimiento")
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=obtener_api_key_inicial(),
        type="password",
        help="Pega tu clave de Google AI Studio. Si la guardas en 'Secrets' de Streamlit como GOOGLE_API_KEY, cargará sola siempre."
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

    modo_ocr_sel = st.selectbox(
        "⚡ Lectura de PDF / OCR",
        options=[
            "Desactivado (PDF digital - más rápido)",
            "Tesseract Local Ligero (120 DPI)",
            "Nube Gemini Vision (OCR para PDF escaneado)"
        ],
        index=0,
        help="La lectura normal de PDF digital es rápida. El OCR de páginas escaneadas tarda más; actívalo solo cuando el PDF no tenga texto seleccionable."
    )

    st.divider()
    st.header("🏛️ Jurisdicción y Caso")

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
        "2. Entidad, Juzgado o Contraparte",
        placeholder="Ej: SIC, Procuraduría, Personería, Juzgado, Constructora, EPS..."
    )

    rol_usuario = st.selectbox(
        "3. Tu Rol Procesal",
        options=[
            "Demandante / Accionante / Quejoso / Peticionario / Agente Oficioso / Veedor",
            "Demandado / Entidad Convocada / Investigado / Defensa Técnica",
            "Auditor / Interventor / Consultor Jurídico Imparcial"
        ]
    )

    with st.expander("👤 Mis Datos para Firmar Escritos (Opcional)"):
        firmante_nombre = st.text_input("Nombre Completo", placeholder="Ej: Juan Pérez Gómez")
        firmante_doc = st.text_input("Cédula / NIT / T.P.", placeholder="Ej: C.C. 1.000.000.000")
        firmante_ciudad = st.text_input("Ciudad y Departamento", placeholder="Ej: Bogotá D.C.")
        firmante_correo = st.text_input("Correo de Notificaciones", placeholder="Ej: correo@ejemplo.com")
        firmante_tel = st.text_input("Celular / Teléfono", placeholder="Ej: 300 000 0000")

    datos_firmante_dict = {
        "nombre": firmante_nombre.strip(),
        "documento": firmante_doc.strip(),
        "ciudad": firmante_ciudad.strip(),
        "correo": firmante_correo.strip(),
        "telefono": firmante_tel.strip()
    }

    limite_chars = st.slider(
        "Capacidad de lectura por archivo (caracteres)",
        min_value=20000,
        max_value=250000,
        value=120000,
        step=20000
    )

    st.divider()
    if st.button("🗑️ Limpiar Sesión y Memoria RAM", use_container_width=True):
        st.session_state.expediente_docs = []
        st.session_state.cache_hashes = {}
        st.session_state.actuaciones_df = pd.DataFrame()
        st.session_state.catalogo_df = pd.DataFrame()
        st.session_state.informe_md = ""
        st.session_state.estrategia_md = ""
        st.session_state.blindaje_md = ""
        st.session_state.historial_chat = []
        gc.collect()

# =====================================================================
# CABECERA EJECUTIVA Y CONTENEDOR DE ESTADO
# =====================================================================
st.markdown("""
<div class="main-header" translate="no">
    <h2 style="margin:0; font-weight:800;">⚖️ Bufete Litigante IA Colombia — Suite Forense & Estrategia Legal</h2>
    <p style="margin:0.35rem 0 0 0; opacity:0.92; font-size:0.98rem;">
        Auditoría cruzada de contradicciones en expedientes, redacción jurídica multijurisdiccional, simulador de contraparte y control de términos hábiles
    </p>
</div>
""", unsafe_allow_html=True)

banner_estado = st.empty()

def actualizar_banner_estado():
    n_arch = len(st.session_state.expediente_docs)
    n_fol = sum(d["num_paginas"] for d in st.session_state.expediente_docs)
    pill_ia = (
        '<span class="status-pill status-pill-ok">🟢 Motor Gemini Conectado</span>'
        if api_key_input
        else '<span class="status-pill status-pill-warn">🟠 Ingresa tu API Key en el panel izquierdo</span>'
    )
    banner_estado.markdown(
        f"""
        <div class="status-bar" translate="no">
            {pill_ia}
            <span class="status-pill status-pill-ok">⚡ Escudo Anti-RAM / Anti-Traductor Activo</span>
            <span class="status-pill">📁 Expediente en Memoria: {n_arch} archivos ({n_fol} folios)</span>
        </div>
        """,
        unsafe_allow_html=True
    )

actualizar_banner_estado()

# ----------------- SEIS PESTAÑAS DE LITIGIO INTEGRAL -----------------
tab_carga, tab_auditoria, tab_consultor, tab_blindaje, tab_rag, tab_export = st.tabs([
    "📂 1. Expediente Digital",
    "🔍 2. Auditar Contradicciones",
    "🧰 3. ¿Qué Hacer? (Desde Cero)",
    "⚔️ 4. Blindaje y Términos Legales",
    "💬 5. Sala de Interrogatorio (Chat)",
    "📄 6. Bóveda Word y Excel"
])

# =====================================================================
# PESTAÑA 1: CARGA ULTRA-ESTABLE (SIN st.rerun() PARA EVITAR removeChild)
# =====================================================================
with tab_carga:
    st.subheader("📤 Cargar Documentos del Expediente, Demanda, Oficios o Pruebas")
    st.caption("Los PDF digitales se leen directamente y rápido. Si un PDF está escaneado como imagen, selecciona Tesseract o Gemini Vision en la barra lateral para activar OCR; el OCR puede tardar más.")

    archivos_subidos = st.file_uploader(
        "Arrastra aquí tus archivos (PDF, DOCX, TXT o RTF)",
        type=["pdf", "docx", "txt", "rtf"],
        accept_multiple_files=True
    )

    col_b1, col_b2 = st.columns([1, 2])
    with col_b1:
        procesar_btn = st.button("⚙️ Procesar e Indexar Expediente", type="primary", use_container_width=True)

    if procesar_btn:
        if not archivos_subidos:
            st.warning("⚠️ Selecciona al menos un archivo para procesar, o ve directo a la Pestaña 3 si deseas consultar un caso desde cero.")
        else:
            if modo_ocr_sel.startswith("Nube") and not api_key_input:
                st.info("💡 Consejo: Ingresa tu API Key en la barra izquierda antes de procesar PDFs escaneados para que el OCR corra en la nube de Google con 0% de uso de CPU local.")

            cliente_tmp = obtener_cliente_gemini(api_key_input) if api_key_input else None
            docs_procesados = []
            zona_progreso = st.empty()

            for idx, arch in enumerate(archivos_subidos):
                zona_progreso.info(f"⏳ Leyendo documento {idx + 1}/{len(archivos_subidos)}: **{arch.name}**. Los PDF digitales suelen procesarse rápido; los escaneados requieren OCR.")
                item_doc = procesar_archivo_ligero(
                    arch,
                    api_key=api_key_input,
                    client=cliente_tmp,
                    modo_ocr=modo_ocr_sel
                )
                docs_procesados.append(item_doc)

            zona_progreso.empty()
            gc.collect()
            st.session_state.expediente_docs = docs_procesados

            filas_cat = []
            for it in docs_procesados:
                filas_cat.append({
                    "Archivo": it["nombre"],
                    "Páginas": it["num_paginas"],
                    "Págs OCR": it["pags_ocr"],
                    "Págs sin texto": it.get("pags_sin_texto", 0),
                    "Caracteres": it["caracteres"],
                    "Radicado Detectado": it["meta"]["radicado"],
                    "Actuaciones": ", ".join(it["meta"]["tipo_actuacion"][:4]) or "General",
                    "Actor / Solicitante": it["meta"]["actor"],
                    "Demandado / Entidad": it["meta"]["demandado"],
                    "Oficio / Auto": it["meta"]["numero_oficio"],
                    "Hash SHA-256": it["hash"]
                })
            st.session_state.catalogo_df = pd.DataFrame(filas_cat)
            st.session_state.actuaciones_df = consolidar_cronologia_df(docs_procesados)
            actualizar_banner_estado()
            total_sin_texto = sum(d.get("pags_sin_texto", 0) for d in docs_procesados)
            total_ocr = sum(d.get("pags_ocr", 0) for d in docs_procesados)
            st.success(
                f"✅ Lectura terminada: {len(docs_procesados)} documentos, "
                f"{sum(d['num_paginas'] for d in docs_procesados)} páginas y {total_ocr} páginas procesadas con OCR."
            )
            if total_sin_texto:
                st.warning(
                    f"⚠️ {total_sin_texto} página(s) no contienen texto extraíble. "
                    "Si son páginas escaneadas, selecciona 'Tesseract Local Ligero' o "
                    "'Nube Gemini Vision' en la barra lateral y vuelve a procesar los archivos."
                )
            st.info("Ahora puedes pasar a la Pestaña 2 para ejecutar la auditoría.")

    if st.session_state.expediente_docs:
        docs_mem = st.session_state.expediente_docs
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📁 Archivos en Bóveda", len(docs_mem))
        m2.metric("📄 Total Páginas / Folios", sum(d["num_paginas"] for d in docs_mem))
        m3.metric("🔍 Folios Leídos con OCR", sum(d.get("pags_ocr", 0) for d in docs_mem))
        m4.metric("📝 Caracteres Extraídos", f"{sum(d['caracteres'] for d in docs_mem):,}")
        paginas_sin_texto_total = sum(d.get("pags_sin_texto", 0) for d in docs_mem)
        if paginas_sin_texto_total:
            st.warning(
                f"Hay {paginas_sin_texto_total} página(s) sin texto extraíble. "
                "Si contienen documentos escaneados, vuelve a procesar con OCR activado."
            )

        st.markdown("#### 📋 Inventario Técnico del Expediente")
        st.dataframe(st.session_state.catalogo_df, use_container_width=True)

        col_exp_a, col_exp_b = st.columns(2)
        with col_exp_a:
            with st.expander("🕒 Línea de Tiempo Detectada Automáticamente"):
                st.dataframe(st.session_state.actuaciones_df, use_container_width=True)
        with col_exp_b:
            with st.expander("🔎 Visor de Folios Extraídos por Documento"):
                nombres_arch = [d["nombre"] for d in docs_mem]
                arch_sel = st.selectbox("Selecciona un archivo:", nombres_arch)
                doc_ver = next(d for d in docs_mem if d["nombre"] == arch_sel)
                st.text_area(
                    f"Vista previa de {arch_sel}:",
                    value=doc_ver["raw"][:20000],
                    height=240
                )

# =====================================================================
# PESTAÑA 2: AUDITORÍA FORENSE DE CONTRADICCIONES + DOBLE WORD
# =====================================================================
with tab_auditoria:
    st.subheader("🔍 Auditoría Forense de Contradicciones, Pruebas y Plan de Acción")
    st.write("Cruza todos los folios de tus PDFs para descubrir contradicciones fácticas o errores judiciales/administrativos, indicándote qué hacer y redactando el memorial.")

    enfoque_usuario = st.text_area(
        "🎯 ¿Qué contradicciones, incumplimientos o puntos críticos quieres que tu Abogado IA revise con lupa? (Opcional):",
        placeholder="Ejemplo: Cruzar lo prometido en publicidad o en la sentencia frente a lo que respondió la constructora, EPS, alcaldía o juzgado en los oficios recientes; detectar vacíos probatorios y redactar el memorial listo para radicar.",
        height=95
    )

    if st.button("🚀 Ejecutar Auditoría Forense del Expediente", type="primary", use_container_width=True):
        if not api_key_input:
            st.error("❌ Ingresa tu Google Gemini API Key en la barra lateral izquierda.")
        elif not st.session_state.expediente_docs:
            st.warning("⚠️ Primero carga tus documentos en la Pestaña 1 (o usa la Pestaña 3 si aún no tienes documentos).")
        else:
            with st.spinner("⚖️ Auditando folios en la nube de Gemini, cruzando pruebas y redactando informe + memorial..."):
                try:
                    client = obtener_cliente_gemini(api_key_input)
                    prompt_aud = construir_prompt_auditoria(
                        items=st.session_state.expediente_docs,
                        catalogo_df=st.session_state.catalogo_df,
                        actuaciones_df=st.session_state.actuaciones_df,
                        tipo_proceso=tipo_proceso,
                        entidad_especifica=entidad_especifica,
                        rol_usuario=rol_usuario,
                        enfoque_especial=enfoque_usuario,
                        datos_firmante=datos_firmante_dict,
                        limite_chars_por_doc=limite_chars
                    )
                    texto_informe, modelo_usado = llamar_gemini(
                        client=client,
                        api_key=api_key_input,
                        prompt_contents=prompt_aud,
                        modelo_preferido=modelo_sel,
                        temperatura=0.1
                    )
                    st.session_state.informe_md = texto_informe
                    gc.collect()
                    st.success(f"✅ Auditoría completada con `{modelo_usado}`.")
                except Exception as e:
                    st.error(f"❌ Error al generar la auditoría: {str(e)}")

    if st.session_state.informe_md:
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            word_completo = convertir_markdown_a_docx_bytes(
                st.session_state.informe_md,
                titulo_portada="INFORME DE AUDITORÍA PROCESAL Y PLAN DE ACCIÓN",
                incluir_portada=True
            )
            st.download_button(
                label="📘 Descargar Informe Completo de Auditoría (.docx)",
                data=word_completo,
                file_name=f"Informe_Auditoria_Completo_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with col_dl2:
            solo_memorial_md = extraer_solo_borrador_radicable(st.session_state.informe_md)
            if solo_memorial_md:
                word_memorial = convertir_markdown_a_docx_bytes(
                    solo_memorial_md,
                    titulo_portada="",
                    incluir_portada=False
                )
                st.download_button(
                    label="⚖️ Descargar SOLO el Memorial Limpio para Radicar (.docx)",
                    data=word_memorial,
                    file_name=f"Memorial_Listo_Para_Radicar_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        st.divider()
        st.markdown(st.session_state.informe_md)

# =====================================================================
# PESTAÑA 3: CONSULTOR Y ESTRATEGA DESDE CERO (CON O SIN EXPEDIENTE)
# =====================================================================
with tab_consultor:
    st.markdown("""
    <div class="tool-box">
        <h3 style="margin-top:0; color:#1e3a8a;">🧰 Tu Abogado Estratega Desde Cero (Para Actuaciones Nuevas o Problemas Legales)</h3>
        <p style="margin-bottom:0;">
            Cuéntale tu caso o inconveniente en lenguaje natural. El sistema analiza qué leyes colombianas te protegen, ante qué entidades acudir, qué plazos cuidar y <b>te entrega redactado tanto el Plan Estratégico como el Escrito Limpio en Word listo para radicar</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_c1, col_c2 = st.columns([2, 1])

    with col_c1:
        problema_descrito = st.text_area(
            "✍️ Describe detalladamente tu problema jurídico, administrativo, contractual, inmobiliario o judicial:",
            placeholder="Ejemplo: Compré un lote o inmueble donde ofrecieron servicios públicos y obras urbanísticas que no han entregado; o una entidad/juzgado vulneró mis derechos o guarda silencio frente a una petición. Describe fechas aproximadas, nombres de entidades/personas y qué buscas lograr.",
            height=195
        )

    with col_c2:
        documento_a_generar = st.selectbox(
            "📝 Escrito a redactar listo para radicar:",
            options=[
                "Derecho de Petición Blindado (Ley 1755 de 2015)",
                "Demanda Jurisdiccional o Queja ante la SIC (Ley 1480 de 2011 - Consumidor / Publicidad Engañosa)",
                "Queja Disciplinaria o Vigilancia (Procuraduría / Personería / Comisión de Disciplina Judicial)",
                "Acción de Tutela con Medida Provisional Urgente",
                "Impugnación de Fallo de Tutela (Dentro de los 3 días)",
                "Incidente de Desacato / Solicitud de Cumplimiento de Sentencia",
                "Recurso de Reposición y en Subsidio de Apelación",
                "Acción Popular o Acción de Cumplimiento (Espacio Público / Urbanismo / Comunidad)",
                "Reclamación Directa Previa / Requerimiento Prejudicial",
                "Contestación de Demanda / Escrito de Defensa y Excepciones",
                "Denuncia Penal Formal ante la Fiscalía General de la Nación"
            ]
        )

        incluir_docs_cargados = st.checkbox(
            "📎 Cruzar también con los documentos cargados en la Pestaña 1",
            value=bool(st.session_state.expediente_docs)
        )

    if st.button("⚖️ Diseñar Estrategia Legal Colombiana y Redactar Escrito", type="primary", use_container_width=True):
        if not api_key_input:
            st.error("❌ Ingresa tu Google Gemini API Key en la barra lateral izquierda.")
        elif not problema_descrito or len(problema_descrito.strip()) < 15:
            st.warning("⚠️ Describe tu caso o inconveniente con al menos un párrafo breve para darte una asesoría exacta.")
        else:
            with st.spinner("⚖️ Analizando legislación y jurisprudencia colombiana y redactando tu documento..."):
                try:
                    client = obtener_cliente_gemini(api_key_input)
                    prompt_est = construir_prompt_consultor_desde_cero(
                        problema_descrito=problema_descrito,
                        tipo_proceso=tipo_proceso,
                        entidad_especifica=entidad_especifica,
                        rol_usuario=rol_usuario,
                        documento_a_generar=documento_a_generar,
                        incluir_docs_cargados=incluir_docs_cargados,
                        items=st.session_state.expediente_docs,
                        datos_firmante=datos_firmante_dict,
                        limite_chars_por_doc=limite_chars
                    )
                    texto_estrategia, modelo_usado = llamar_gemini(
                        client=client,
                        api_key=api_key_input,
                        prompt_contents=prompt_est,
                        modelo_preferido=modelo_sel,
                        temperatura=0.15
                    )
                    st.session_state.estrategia_md = texto_estrategia
                    gc.collect()
                    st.success(f"✅ Estrategia y escrito generados con `{modelo_usado}`.")
                except Exception as e:
                    st.error(f"❌ Error al generar la estrategia: {str(e)}")

    if st.session_state.estrategia_md:
        col_ed1, col_ed2 = st.columns(2)
        with col_ed1:
            word_est_full = convertir_markdown_a_docx_bytes(
                st.session_state.estrategia_md,
                titulo_portada="DICTAMEN DE ESTRATEGIA JURÍDICA Y HOJA DE RUTA",
                incluir_portada=True
            )
            st.download_button(
                label="🧭 Descargar Estrategia + Borrador Completo (.docx)",
                data=word_est_full,
                file_name=f"Estrategia_Juridica_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with col_ed2:
            solo_escrito_est = extraer_solo_borrador_radicable(st.session_state.estrategia_md)
            if solo_escrito_est:
                word_est_limpio = convertir_markdown_a_docx_bytes(
                    solo_escrito_est,
                    titulo_portada="",
                    incluir_portada=False
                )
                st.download_button(
                    label="⚖️ Descargar SOLO el Escrito Limpio para Radicar (.docx)",
                    data=word_est_limpio,
                    file_name=f"Escrito_Radicable_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        st.divider()
        st.markdown(st.session_state.estrategia_md)

# =====================================================================
# PESTAÑA 4: SIMULADOR "ABOGADO DEL DIABLO" + CALCULADORA DE TÉRMINOS COLOMBIA
# =====================================================================
with tab_blindaje:
    col_bl1, col_bl2 = st.columns([3, 2])

    with col_bl1:
        st.subheader("⚔️ Simulador 'Abogado del Diablo' (Blindaje de Escritos)")
        st.caption("Pon a prueba cualquier borrador antes de radicarlo: la IA actúa como el Juez estricto o el abogado de tu contraparte para detectar debilidades y entregarte una versión 100% blindada.")

        fuente_borrador = st.radio(
            "¿Qué texto deseas someter a blindaje?",
            options=[
                "Pegar mi propio borrador o texto manualmente",
                "Usar el Memorial generado en la Pestaña 2 (Auditoría)",
                "Usar el Escrito generado en la Pestaña 3 (Consultor)"
            ],
            horizontal=True
        )

        texto_inicial_blindaje = ""
        if fuente_borrador.startswith("Usar el Memorial") and st.session_state.informe_md:
            texto_inicial_blindaje = extraer_solo_borrador_radicable(st.session_state.informe_md) or st.session_state.informe_md
        elif fuente_borrador.startswith("Usar el Escrito") and st.session_state.estrategia_md:
            texto_inicial_blindaje = extraer_solo_borrador_radicable(st.session_state.estrategia_md) or st.session_state.estrategia_md

        borrador_a_blindar = st.text_area(
            "Escrito o memorial a auditar y blindar:",
            value=texto_inicial_blindaje,
            height=220,
            placeholder="Pega aquí tu derecho de petición, demanda, queja, recurso o memorial para que el Abogado del Diablo busque huecos legales y lo reescriba blindado..."
        )

        if st.button("🛡️ Atacar Debilidades y Generar Escrito Blindado", type="primary", use_container_width=True):
            if not api_key_input:
                st.error("❌ Ingresa tu API Key en la barra lateral.")
            elif not borrador_a_blindar or len(borrador_a_blindar.strip()) < 30:
                st.warning("⚠️ Ingresa o selecciona un borrador válido para blindar.")
            else:
                with st.spinner("⚔️ Simulando ataques de la contraparte y blindando tu documento..."):
                    try:
                        client = obtener_cliente_gemini(api_key_input)
                        prompt_diablo = construir_prompt_abogado_diablo(
                            borrador_usuario=borrador_a_blindar,
                            tipo_proceso=tipo_proceso,
                            entidad_especifica=entidad_especifica,
                            datos_firmante=datos_firmante_dict
                        )
                        res_blindaje, mod_b = llamar_gemini(
                            client=client,
                            api_key=api_key_input,
                            prompt_contents=prompt_diablo,
                            modelo_preferido=modelo_sel,
                            temperatura=0.1
                        )
                        st.session_state.blindaje_md = res_blindaje
                        gc.collect()
                        st.success(f"✅ Escrito blindado con éxito (`{mod_b}`).")
                    except Exception as e:
                        st.error(f"❌ Error en el blindaje: {str(e)}")

    with col_bl2:
        st.subheader("📅 Calculadora de Términos Hábiles (Colombia)")
        st.caption("Descuenta automáticamente sábados, domingos y los 18 festivos oficiales de Colombia (Ley Emiliani y Semana Santa).")

        termino_preset = st.selectbox(
            "Tipo de Término Legal en Colombia:",
            options=[
                "Derecho de Petición General — 15 días hábiles (Art. 14 Ley 1755/2015)",
                "Petición de Documentos y Copias — 10 días hábiles (Si no responden: Silencio Positivo)",
                "Petición de Consulta — 30 días hábiles (Art. 14 Ley 1755/2015)",
                "Reclamación Directa al Productor/Proveedor SIC — 15 días hábiles (Ley 1480/2011)",
                "Ejecutoria / Recurso de Reposición contra Auto Judicial — 3 días hábiles (CGP)",
                "Impugnación de Sentencia de Tutela — 3 días hábiles (Art. 31 Dto. 2591/1991)",
                "Recurso de Reposición/Apelación Acto Administrativo — 10 días hábiles (Art. 76 CPACA)",
                "Trámite de Incidente de Desacato — 10 días hábiles (Sentencia C-367/2014)",
                "Contestación Demanda Civil Verbal / SIC — 20 días hábiles",
                "Personalizado (Elegir número exacto de días hábiles)"
            ]
        )

        mapa_dias = {"15 días": 15, "10 días": 10, "30 días": 30, "3 días": 3, "20 días": 20}
        dias_defecto = 15
        for clave_d, val_d in mapa_dias.items():
            if clave_d in termino_preset:
                dias_defecto = val_d
                break

        fecha_notif = st.date_input("Fecha de Radicación o Notificación Efectiva:", value=date.today())
        dias_calcular = st.number_input("Días hábiles del término:", min_value=1, max_value=180, value=dias_defecto)

        fecha_venc, fines_sem, festivos_cruzados, df_ruta_dias = sumar_dias_habiles_colombia(fecha_notif, int(dias_calcular))
        dias_restantes_hoy = (fecha_venc - date.today()).days

        st.markdown(f"""
        <div class="gold-box">
            <h4 style="margin:0 0 0.4rem 0;">📌 Fecha Exacta de Vencimiento: {fecha_venc.strftime('%d/%m/%Y')}</h4>
            <p style="margin:0; font-size:0.9rem;">
                • <b>Días calendario faltantes desde hoy:</b> {dias_restantes_hoy} día(s)<br>
                • <b>Sábados y domingos descontados:</b> {fines_sem}<br>
                • <b>Festivos de Colombia descontados ({len(festivos_cruzados)}):</b> {', '.join(festivos_cruzados) if festivos_cruzados else 'Ninguno en este rango'}
            </p>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("📆 Ver desglose día hábil por día hábil"):
            st.dataframe(df_ruta_dias, use_container_width=True, height=210)

    if st.session_state.blindaje_md:
        st.divider()
        c_bl_d1, c_bl_d2 = st.columns(2)
        with c_bl_d1:
            w_blindaje_full = convertir_markdown_a_docx_bytes(
                st.session_state.blindaje_md,
                titulo_portada="DICTAMEN DE BLINDAJE Y ESCRITO REFORZADO",
                incluir_portada=True
            )
            st.download_button(
                label="🛡️ Descargar Auditoría de Blindaje + Escrito (.docx)",
                data=w_blindaje_full,
                file_name=f"Blindaje_Completo_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with c_bl_d2:
            solo_blindado = extraer_solo_borrador_radicable(st.session_state.blindaje_md)
            if solo_blindado:
                w_blindaje_clean = convertir_markdown_a_docx_bytes(
                    solo_blindado,
                    titulo_portada="",
                    incluir_portada=False
                )
                st.download_button(
                    label="⚖️ Descargar SOLO el Escrito Blindado Limpio (.docx)",
                    data=w_blindaje_clean,
                    file_name=f"Escrito_Blindado_Radicable_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        st.markdown(st.session_state.blindaje_md)

# =====================================================================
# PESTAÑA 5: SALA DE INTERROGATORIO AL EXPEDIENTE (CON BOTONES TÁCTICOS)
# =====================================================================
with tab_rag:
    st.subheader("💬 Sala de Interrogatorio al Expediente y Consultas Rápidas")
    st.caption("Usa los botones tácticos de un clic o escribe cualquier pregunta sobre tus PDFs o sobre procedimiento colombiano.")

    q_col1, q_col2, q_col3, q_col4 = st.columns(4)
    pregunta_rapida = None
    if q_col1.button("🎯 Confesiones de la contraparte", use_container_width=True):
        pregunta_rapida = "Busca en todos los documentos cargados qué admisiones, confesiones espontáneas o contradicciones ha cometido la contraparte o la entidad y dime en qué archivo y página exacta están."
    if q_col2.button("⏳ ¿Hay plazos vencidos o silencio?", use_container_width=True):
        pregunta_rapida = "Verifica todas las fechas del expediente y dime si se configuró vencimiento de términos judiciales, incumplimiento de plazos o silencio administrativo positivo/negativo según la ley colombiana."
    if q_col3.button("🛡️ ¿Qué pruebas nos faltan?", use_container_width=True):
        pregunta_rapida = "Dime con rigor de juez estricto qué vacíos probatorios tenemos hoy y qué pruebas documentales o por oficio debemos pedir o adjuntar de inmediato."
    if q_col4.button("🎙️ Guion para audiencia / cita", use_container_width=True):
        pregunta_rapida = "Prepárame un guion ejecutivo, contundente y ordenado de 5 puntos clave con citas de folios para exponer verbalmente este caso en una audiencia, inspección o reunión."

    for msj in st.session_state.historial_chat:
        with st.chat_message(msj["role"]):
            st.markdown(msj["content"])

    pregunta_manual = st.chat_input("Escribe tu pregunta sobre cualquier folio o ley colombiana...")
    pregunta_activa = pregunta_rapida or pregunta_manual

    if pregunta_activa:
        if not api_key_input:
            st.error("❌ Configura tu API Key en la barra lateral.")
        else:
            st.session_state.historial_chat.append({"role": "user", "content": pregunta_activa})
            with st.chat_message("user"):
                st.markdown(pregunta_activa)

            with st.chat_message("assistant"):
                with st.spinner("Consultando folios y normativa colombiana..."):
                    try:
                        client = obtener_cliente_gemini(api_key_input)
                        if st.session_state.expediente_docs:
                            contexto_docs = "\n\n".join(
                                [f"=== DOCUMENTO: {d['nombre']} ===\n{d['raw'][:limite_chars]}" for d in st.session_state.expediente_docs]
                            )
                        else:
                            contexto_docs = "Sin documentos subidos en esta sesión. Responde conforme a la legislación colombiana vigente."

                        prompt_pregunta = f"""
Actúa como abogado litigante colombiano senior.
{MARCO_JURIDICO_COLOMBIA}

Si la respuesta proviene de los documentos cargados, cita siempre `(Archivo: nombre.pdf, Pág. X)` (sin usar barras verticales) y transcribe entre comillas el fragmento probatorio.
Si es una consulta legal sobre Colombia, cita las normas exactas (Leyes, Artículos y Sentencias) y explica paso a paso qué hacer.

CONTEXTO DOCUMENTAL:
{contexto_docs}

PREGUNTA DEL USUARIO:
{pregunta_activa}
"""
                        respuesta_chat, _ = llamar_gemini(
                            client=client,
                            api_key=api_key_input,
                            prompt_contents=prompt_pregunta,
                            modelo_preferido=modelo_sel,
                            temperatura=0.1
                        )
                        st.markdown(respuesta_chat)
                        st.session_state.historial_chat.append({"role": "assistant", "content": respuesta_chat})
                    except Exception as e:
                        st.error(f"Error en la consulta: {str(e)}")

# =====================================================================
# PESTAÑA 6: BÓVEDA CENTRAL DE EXPORTACIÓN (WORD, ESCRITOS LIMPIOS Y EXCEL)
# =====================================================================
with tab_export:
    st.subheader("📄 Bóveda Central de Entregables Profesionales")
    st.caption("Descarga tus dictámenes completos, escritos limpios listos para firmar y matrices de control en Excel.")

    col_e1, col_e2, col_e3, col_e4 = st.columns(4)

    with col_e1:
        st.markdown("#### 📘 1. Auditoría Expediente")
        st.write("Informe completo + Memorial separado listo para firmar.")
        if st.session_state.informe_md:
            st.download_button(
                label="⬇️ Informe Completo (.docx)",
                data=convertir_markdown_a_docx_bytes(st.session_state.informe_md, "INFORME DE AUDITORÍA PROCESAL", True),
                file_name=f"01_INFORME_AUDITORIA_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
            mem_limpio = extraer_solo_borrador_radicable(st.session_state.informe_md)
            if mem_limpio:
                st.download_button(
                    label="⚖️ Solo Memorial Limpio (.docx)",
                    data=convertir_markdown_a_docx_bytes(mem_limpio, "", False),
                    file_name=f"01B_MEMORIAL_RADICABLE_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.info("Disponible tras ejecutar la Pestaña 2.")

    with col_e2:
        st.markdown("#### 🧭 2. Estrategia Desde Cero")
        st.write("Concepto jurídico y escrito limpio para casos nuevos.")
        if st.session_state.estrategia_md:
            st.download_button(
                label="⬇️ Estrategia Completa (.docx)",
                data=convertir_markdown_a_docx_bytes(st.session_state.estrategia_md, "ESTRATEGIA JURÍDICA COLOMBIA", True),
                file_name=f"02_ESTRATEGIA_JURIDICA_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
            esc_limpio = extraer_solo_borrador_radicable(st.session_state.estrategia_md)
            if esc_limpio:
                st.download_button(
                    label="⚖️ Solo Escrito Limpio (.docx)",
                    data=convertir_markdown_a_docx_bytes(esc_limpio, "", False),
                    file_name=f"02B_ESCRITO_RADICABLE_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.info("Disponible tras usar la Pestaña 3.")

    with col_e3:
        st.markdown("#### 🛡️ 3. Escrito Blindado")
        st.write("Versión reforzada por el Simulador 'Abogado del Diablo'.")
        if st.session_state.blindaje_md:
            st.download_button(
                label="⬇️ Blindaje Completo (.docx)",
                data=convertir_markdown_a_docx_bytes(st.session_state.blindaje_md, "DICTAMEN DE BLINDAJE PROCESAL", True),
                file_name=f"03_BLINDAJE_COMPLETO_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
            bl_limpio = extraer_solo_borrador_radicable(st.session_state.blindaje_md)
            if bl_limpio:
                st.download_button(
                    label="⚖️ Solo Escrito Blindado (.docx)",
                    data=convertir_markdown_a_docx_bytes(bl_limpio, "", False),
                    file_name=f"03B_ESCRITO_BLINDADO_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.info("Disponible tras usar la Pestaña 4.")

    with col_e4:
        st.markdown("#### 📊 4. Matriz en Excel")
        st.write("Inventario SHA-256 y Cronología de Actuaciones.")
        if not st.session_state.catalogo_df.empty:
            st.download_button(
                label="⬇️ Descargar Excel (.xlsx)",
                data=generar_excel_formateado_bytes(st.session_state.catalogo_df, st.session_state.actuaciones_df),
                file_name=f"04_MATRIZ_CONTROL_EXPEDIENTE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Disponible tras cargar PDFs en la Pestaña 1.")
