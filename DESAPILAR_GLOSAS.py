import re
import sys
import os
import shutil
from pathlib import Path
import io

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

try:
    import fitz
except Exception:
    fitz = None

try:
    import pytesseract
    import shutil as _shutil
    _tess = _shutil.which("tesseract") or r"C:\Users\jhernandez\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = _tess
    if not Path(_tess).exists():
        print(f"[ADVERTENCIA] Tesseract no encontrado en: {_tess}. OCR desactivado.")
        pytesseract = None
except Exception:
    pytesseract = None

try:
    from PIL import Image
except Exception:
    Image = None

APP_NAME      = "DESAPILAR_GLOSAS"
REPORT_NAME   = "REPORTE_GLOSAS.xlsx"
MANUAL_FOLDER = "PROCESAR MANUAL"
PROCESSED_FOLDER = "PROCESADOS"


# =========================
# IDENTIDAD
# =========================
def datos_identidad():
    return {
        "sistema":          "DESAPILAR GLOSAS",
        "propietario":      "Salud-Net",
        "desarrollado_por": "DESARROLLO E INNOVACIÓN SALUD NET",
        "version":          "v2.1",
        "licencia":         "Uso interno autorizado",
    }


def mostrar_marca_agua():
    datos = datos_identidad()
    try:
        os.system(f'title {datos["sistema"]} - {datos["propietario"]}')
    except Exception:
        pass
    print("\n" + "=" * 68)
    print(f'{datos["sistema"]:^68}')
    print("=" * 68)
    print(f'PROPIETARIO : {datos["propietario"]}')
    print(f'DESARROLLADO: {datos["desarrollado_por"]}')
    print(f'VERSION     : {datos["version"]}')
    print(f'LICENCIA    : {datos["licencia"]}')
    print("=" * 68 + "\n")


def script_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR  = script_dir()
INPUT_DIR = BASE_DIR


def safe_mkdir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


# =========================
# TEXTO
# =========================
def clean_digits(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))


def clean_nit(value: str) -> str:
    digits = clean_digits(value)
    return digits[:9] if len(digits) >= 9 else digits


def normalize(text: str) -> str:
    text = str(text or "").upper()
    for a, b in {"Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ñ":"N"}.items():
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip()


def normalize_keep_lines(text: str) -> str:
    text = str(text or "").upper()
    for a, b in {"Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","Ñ":"N"}.items():
        text = text.replace(a, b)
    text = text.replace("\r", "")
    return re.sub(r"[ \t]+", " ", text)


def normalize_ocr_token(value: str) -> str:
    t = normalize(value).replace(" ", "")
    for a, b in [("O","0"),("I","1"),("L","1"),("S","5"),("B","8"),("Z","2")]:
        t = t.replace(a, b)
    return t


# =========================
# FILTRO DE RUIDO
# Excluye números de teléfono, páginas, órdenes de pago, radicaciones, etc.
# =========================
_NOISE_CTX = re.compile(
    r"(?:TELEFONO|TELEFONOS|TEL|FAX|EXTENSION|EXT|"
    r"LLAMENOS|LINEA\s+DE\s+ATENCION|"
    r"ORDEN\s+DE\s+PAGO|CHEQUE|TRANSFERENCIA|"
    r"SINIESTRO\s+NO|RADICACION|RIQ|CMVIQ|RDL)"
    r"\s*[:\-]?\s*(\d+)",
    re.IGNORECASE,
)
_PAGE_CTX = re.compile(r"P[AÁ]GINA\s+(\d+)\s*(?:/|DE)\s*(\d+)", re.IGNORECASE)
_TEL_BARE = re.compile(r"(?:TEL|TELEFONO|FAX|LLAME|COMUNICARSE)[^\d]{0,10}(\d{7,10})", re.IGNORECASE)


def build_noise_numbers(text: str) -> set:
    noise = set()
    text_n = normalize(text)
    for m in _NOISE_CTX.finditer(text_n):
        noise.add(clean_digits(m.group(1)))
    for m in _PAGE_CTX.finditer(text_n):
        noise.add(m.group(1))
        noise.add(m.group(2))
    for m in _TEL_BARE.finditer(text_n):
        noise.add(m.group(1))
    return noise


def looks_like_invoice(token: str) -> bool:
    """Acepta 4-10 dígitos con o sin F inicial."""
    t = normalize_ocr_token(token)
    return bool(re.fullmatch(r"(F)?\d{4,10}", t))


def looks_like_date(token: str) -> bool:
    return bool(re.fullmatch(r"\d{2}-\d{2}-\d{4}", token))


# =========================
# ARM INVOICE
# =========================
def arm_invoice(raw_invoice: str, prefixes: list) -> str:
    raw = normalize_ocr_token(raw_invoice)
    for w in ("N°","NO.","NO","FACTURA","RECLAMACION","RECLAMACION",":",".",","):
        raw = raw.replace(normalize_ocr_token(w), "")
    raw = raw.strip()
    raw_digits = clean_digits(raw)
    pref_strings = [str(p).strip() for p in (prefixes or []) if str(p).strip()]
    raw_has_f = raw.startswith("F")

    # Intento 1: coincidencia por cadena normalizada
    same, other = [], []
    for pref in pref_strings:
        p_norm = normalize_ocr_token(pref).replace("-","").replace(" ","")
        if not p_norm: continue
        (same if p_norm.startswith("F") == raw_has_f else other).append((pref, p_norm))

    for group in (same, other):
        for pref, p_norm in sorted(group, key=lambda x: len(x[1]), reverse=True):
            if raw.startswith(p_norm):
                rest = raw[len(p_norm):]
                return (f"{pref}-{rest}" if rest else pref).replace("--", "-")

    # Intento 2: coincidencia por dígitos
    same, other = [], []
    for pref in pref_strings:
        p_norm   = normalize_ocr_token(pref).replace("-","").replace(" ","")
        p_digits = clean_digits(p_norm)
        if not p_digits: continue
        (same if p_norm.startswith("F") == raw_has_f else other).append((pref, p_digits))

    for group in (same, other):
        for pref, p_digits in sorted(group, key=lambda x: len(x[1]), reverse=True):
            if raw_digits.startswith(p_digits):
                rest = raw_digits[len(p_digits):]
                return (f"{pref}-{rest}" if rest else pref).replace("--", "-")

    # Fallback hardcoded
    if raw.startswith("F71") and len(raw_digits) > 2:
        return f"F71-{raw_digits[2:]}"
    if raw_digits.startswith("71") and len(raw_digits) > 2:
        return f"71-{raw_digits[2:]}"
    if raw_digits.startswith("20") and len(raw_digits) > 2:
        return f"20-{raw_digits[2:]}"

    return raw


def filter_invoices_by_prefixes(invoices: list, prefixes: list) -> list:
    if not invoices:
        return []
    norm_prefs, digit_prefs = [], []
    for p in (prefixes or []):
        pn = normalize_ocr_token(str(p)).replace("-","").replace(" ","")
        pd = clean_digits(pn)
        if pn: norm_prefs.append(pn)
        if pd: digit_prefs.append(pd)
    if not norm_prefs and not digit_prefs:
        return invoices
    out, seen = [], set()
    for inv in invoices:
        inv_n = normalize_ocr_token(inv)
        inv_d = clean_digits(inv_n)
        ok = any(inv_n.startswith(p) for p in norm_prefs) or \
             any(inv_d.startswith(p) for p in digit_prefs)
        if ok and inv_n not in seen:
            seen.add(inv_n)
            out.append(inv_n)
    return out


# =========================
# MATRIX
# =========================
def _normalizar_dir(texto: str) -> str:
    """
    Normaliza una dirección para comparación robusta:
    quita tildes, colapsa espacios alrededor de guiones y puntos.
    """
    t = normalize(texto)
    t = re.sub(r"\s*-\s*", "-", t)
    t = re.sub(r"\s*\.\s*", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def load_matrix():
    """
    Carga MATRIZ_IPS.xlsx (NIT + nombre + prefijos de factura).
    Luego enriquece cada entrada con equivalentes y direcciones
    desde NOMBRES_Y_NIT_EQUIVALENTES.xlsx si está disponible.
    Los prefijos SIEMPRE vienen de MATRIZ_IPS — es la fuente autoritativa.
    """
    matrix_path = BASE_DIR / "MATRIZ_IPS.xlsx"
    if not matrix_path.exists():
        print(f"[ERROR] No se encontró: {matrix_path}")
        return {}
    try:
        df = pd.read_excel(matrix_path)
    except Exception as e:
        print(f"[ERROR] No se pudo leer la matriz: {e}")
        return {}

    nit_col = name_col = None
    for c in df.columns:
        cl = normalize(c)
        if cl == "NIT": nit_col = c
        if "IPS NOMBRE COMPLETO" in cl: name_col = c
    if nit_col is None or name_col is None:
        print("[ERROR] Faltan columnas NIT o IPS NOMBRE COMPLETO en la matriz.")
        return {}

    matrix = {}
    for _, row in df.iterrows():
        nit = clean_nit(row.get(nit_col, ""))
        if not nit: continue
        name  = str(row.get(name_col, "")).strip()
        prefs = [str(row.get(c,"")).strip()
                 for c in df.columns
                 if "PREFIJO" in normalize(c)
                 and str(row.get(c,"")).strip()
                 and str(row.get(c,"")).strip().lower() != "nan"]
        matrix[nit] = {
            "name"      : name,
            "prefixes"  : prefs,
            "equivs"    : [normalize(name)],  # al menos el nombre propio
            "dirs"      : [],
        }

    # ── Enriquecer con NOMBRES_Y_NIT_EQUIVALENTES si existe ─────────────────
    nombres_path = BASE_DIR / "NOMBRES Y NIT EQUIVALENTES.xlsx"
    if not nombres_path.exists():
        # Intentar variante sin espacios o con guion bajo
        for alt in ["NOMBRES_Y_NIT_EQUIVALENTES.xlsx", "NOMBRES Y NIT EQUIVALENTES.xlsx"]:
            cand = BASE_DIR / alt
            if cand.exists():
                nombres_path = cand
                break

    if nombres_path.exists():
        try:
            df2 = pd.read_excel(nombres_path, sheet_name="ESTRUCTURA")
            df2.columns = df2.columns.str.strip().str.upper()
            col_eq  = [c for c in df2.columns if c.startswith("EQUIVALENTE")]
            col_dir = [c for c in df2.columns if c.startswith("DIR")]
            col_sed = [c for c in df2.columns if c.startswith("NOMBRE_SEDE")]
            enriq = 0
            for _, row in df2.iterrows():
                nit2 = clean_nit(str(row.get("NIT", "")).strip().split(".")[0])
                if not nit2 or nit2 not in matrix:
                    continue
                equivs = list(matrix[nit2]["equivs"])
                for c in col_eq + col_sed:
                    val = row.get(c)
                    if pd.notna(val) and str(val).strip() and str(val).strip().lower() != "nan":
                        en = normalize(str(val).strip())
                        if en and en not in equivs:
                            equivs.append(en)
                dirs = []
                for c in col_dir:
                    val = row.get(c)
                    if pd.notna(val) and str(val).strip() and str(val).strip().lower() != "nan":
                        dirs.append(_normalizar_dir(str(val).strip()))
                matrix[nit2]["equivs"] = equivs
                matrix[nit2]["dirs"]   = dirs
                enriq += 1
            print(f"[OK] Enriquecido con NOMBRES_Y_NIT_EQUIVALENTES: {enriq} IPS actualizadas.")
        except Exception as e:
            print(f"[AVISO] No se pudo enriquecer con NOMBRES_Y_NIT_EQUIVALENTES: {e}")
    else:
        print("[AVISO] NOMBRES_Y_NIT_EQUIVALENTES no encontrado — sin equivalentes ni direcciones.")

    if not matrix:
        print("[ERROR] Matriz cargada vacía.")
    else:
        print(f"[OK] Matriz cargada: {len(matrix)} registros.")
    return matrix


MATRIX = load_matrix()


def _normalizar_dir(texto: str) -> str:
    """
    Normaliza una dirección para comparación:
    - Quita tildes (via normalize)
    - Colapsa espacios alrededor de guiones y puntos
    - Elimina puntos sueltos
    - Colapsa espacios múltiples
    """
    t = normalize(texto)
    t = re.sub(r"\s*-\s*", "-", t)
    t = re.sub(r"\s*\.\s*", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _quitar_espacios(texto: str) -> str:
    """Elimina todos los espacios — permite comparar VALLESALUD vs VALLE SALUD."""
    return texto.replace(" ", "")


def matrix_lookup_from_text(text: str):
    """
    Busca la IPS en MATRIX con el siguiente orden de prioridad:
    1. NIT como número delimitado en el texto del PDF
    2. Nombre completo normalizado (subcadena exacta, luego sin espacios)
    3. Equivalentes y nombres de sede (subcadena exacta, luego sin espacios;
       el más largo que coincida gana)
    4. Direcciones normalizadas (sin puntuaciones ni espacios variables)
    Retorna (nit, name, prefixes).

    La comparación sin espacios resuelve casos como VALLESALUD vs VALLE SALUD
    donde el nombre en el Excel está pegado pero en el PDF viene separado o
    viceversa, sin afectar la precisión para nombres que no tienen ese patrón.
    """
    text_n      = normalize(text)
    text_n_nsp  = _quitar_espacios(text_n)   # versión sin espacios para fallback
    text_dir    = _normalizar_dir(text)

    # ── Paso 1: NIT delimitado ───────────────────────────────────────────────
    for nit, info in MATRIX.items():
        if nit and re.search(r"\b" + re.escape(nit) + r"\b", text_n):
            return nit, info["name"], info["prefixes"]

    # ── Paso 2: nombre completo normalizado ──────────────────────────────────
    best = (None, None, [], 0)
    for nit, info in MATRIX.items():
        name_n     = normalize(info["name"])
        name_n_nsp = _quitar_espacios(name_n)
        if not name_n:
            continue
        # 2a: subcadena exacta
        if name_n in text_n and len(name_n) > best[3]:
            best = (nit, info["name"], info["prefixes"], len(name_n))
        # 2b: sin espacios (VALLESALUD == VALLE SALUD)
        elif len(name_n_nsp) > 4 and name_n_nsp in text_n_nsp and len(name_n_nsp) > best[3]:
            best = (nit, info["name"], info["prefixes"], len(name_n_nsp))
    if best[0]:
        return best[0], best[1], best[2]

    # ── Paso 3: equivalentes (el más largo que coincida gana) ────────────────
    best = (None, None, [], 0)
    for nit, info in MATRIX.items():
        for equiv in info.get("equivs", []):
            if not equiv or len(equiv) <= 4:
                continue
            equiv_nsp = _quitar_espacios(equiv)
            # 3a: subcadena exacta
            if equiv in text_n and len(equiv) > best[3]:
                best = (nit, info["name"], info["prefixes"], len(equiv))
            # 3b: sin espacios
            elif len(equiv_nsp) > 4 and equiv_nsp in text_n_nsp and len(equiv_nsp) > best[3]:
                best = (nit, info["name"], info["prefixes"], len(equiv_nsp))
    if best[0]:
        return best[0], best[1], best[2]

    # ── Paso 4: dirección normalizada ────────────────────────────────────────
    for nit, info in MATRIX.items():
        for dir_n in info.get("dirs", []):
            if dir_n and len(dir_n) > 8 and dir_n in text_dir:
                return nit, info["name"], info["prefixes"]

    return None, None, []


# =========================
# PDF TEXT
# =========================
def _text_is_garbage(txt: str) -> bool:
    if not txt: return True
    sample  = txt[:2000]
    if re.search(r"(\/\d+\s+){10,}", sample): return True
    cleaned = re.sub(r"[^A-Z0-9ÁÉÍÓÚÑ ]", "", normalize(sample))
    if len(cleaned.strip()) < 80: return True
    if len(re.findall(r"[A-ZÁÉÍÓÚÑ]", cleaned)) < 20: return True
    return False


def _ocr_page(pdf_path: Path, page_index: int) -> str:
    if fitz is None or pytesseract is None or Image is None:
        return ""
    try:
        doc  = fitz.open(str(pdf_path))
        page = doc.load_page(page_index)
        pix  = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        img  = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="spa")
        doc.close()
        return text or ""
    except Exception:
        return ""


def page_texts(pdf_path: Path) -> list:
    reader = PdfReader(str(pdf_path))
    pages  = []
    for idx, page in enumerate(reader.pages):
        try:    txt = page.extract_text() or ""
        except: txt = ""
        if _text_is_garbage(txt):
            ocr = _ocr_page(pdf_path, idx)
            if ocr.strip(): txt = ocr
        pages.append(txt)
    return pages


def page_texts_force_ocr(pdf_path: Path) -> list:
    reader = PdfReader(str(pdf_path))
    return [_ocr_page(pdf_path, i) or "" for i in range(len(reader.pages))]


def full_text(pages: list) -> str:
    return "\n".join(pages)


# =========================
# DETECCIÓN DE FORMATO
# =========================
INSURER_CANON = {
    "AXA COLPATRIA SEGUROS S.A.":             ["AXA COLPATRIA", "AXA"],
    "SEGUROS COMERCIALES BOLIVAR S.A.":       ["SEGUROS BOLIVAR", "ENVIO DOCUMENTOS A TERCEROS"],
    "LA EQUIDAD SEGUROS GENERALES S.A.":      ["LA EQUIDAD", "EQUIDAD SEGUROS"],
    "SEGUROS DEL ESTADO S.A.":                ["SEGUROS DEL ESTADO"],
    "HDI SEGUROS COLOMBIA S.A.":              ["HDI"],
    "MUNDIAL DE SEGUROS S.A.":                ["SEGUROS MUNDIAL", "SEGUROSMUNDIAL.COM.CO", "MUNDIAL"],
    "LA PREVISORA S.A.":                      ["LA PREVISORA", "PREVISORA SEGUROS"],
    "ASEGURADORA SOLIDARIA DE COLOMBIA S.A.": ["ASEGURADORA SOLIDARIA", "SOLIDARIA"],
    "SEGUROS GENERALES SURAMERICANA S.A.":    ["SURAMERICANA", "SURA"],
}


def is_axa_table_format(pages: list) -> bool:
    """
    Detector flexible: PyPDF2 extrae 'No.reclamo' sin espacio,
    pypdf extrae 'No. reclamo' con espacio. Usamos regex en ambos casos.
    """
    if not pages: return False
    doc = normalize(" ".join(pages))
    return (
        bool(re.search(r"NO\.?\s*RECLAMO",     doc)) and
        bool(re.search(r"FECHA\s+LIQUIDACION",  doc)) and
        bool(re.search(r"NO\.?\s*FACTURA",      doc)) and
        bool(re.search(r"VALOR\s+FACTURA",       doc))
    )


def is_axa_letter_format(pages: list) -> bool:
    if not pages: return False
    fn = normalize(normalize_keep_lines(pages[0]))
    return "AXA COLPATRIA" in fn and "LIQUIDACION DE RECLAMACIONES SOAT" in fn


def is_envio_documentos_terceros_format(pages: list) -> bool:
    if not pages: return False
    text = normalize(full_text(pages))
    return (
        "ENVIO DOCUMENTOS A TERCEROS" in text and
        bool(re.search(r"NUM\.?\s*FACTURA", text)) and
        bool(re.search(r"VAL\.?\s*FACTURA", text))
    )


def is_bolivar_rgc_format(pages: list) -> bool:
    """Formato RGC Activa (tabla multilínea, marcador Nro. glosas)."""
    if not pages: return False
    text = normalize(full_text(pages[:3]))
    return (
        "ENVIO DOCUMENTOS A TERCEROS" in text and
        ("RGC ACTIVA" in text or
         bool(re.search(r"NRO\.?\s*GLOSAS", text)) or
         "VALOR TOTAL DE GLOSA" in text)
    )


def is_sura_format(pages: list) -> bool:
    if not pages: return False
    text = normalize(full_text(pages[:3]))
    return (
        ("SURAMERICANA" in text or "SURA" in text) and
        "LIQUIDACION DE SINIESTROS SOAT" in text and
        "LIQUIDACION SINIESTRO" in text
    )


def detect_insurer(pages: list) -> str | None:
    if not pages: return None

    # AXA tabla: el formato de columnas es único — no necesita el logo en texto
    if is_axa_table_format(pages):
        return "AXA COLPATRIA SEGUROS S.A."

    # AXA carta: requiere texto explícito
    if is_axa_letter_format(pages):
        return "AXA COLPATRIA SEGUROS S.A."

    # SURA: detector dedicado (logo también puede ser imagen)
    if is_sura_format(pages):
        return "SEGUROS GENERALES SURAMERICANA S.A."

    # Score por triggers para el resto
    header_text = normalize(" ".join(
        [l.strip() for l in normalize_keep_lines(pages[0]).splitlines() if l.strip()][:20]
    ))
    full_text_n = normalize("\n".join(pages[:2] + pages[-2:] + pages))

    best = (None, 0)
    for canon, triggers in INSURER_CANON.items():
        score = 0
        for trig in triggers:
            tn = normalize(trig)
            if tn in header_text: score += len(tn) * 3
            if tn in full_text_n:  score += len(tn)
        if score > best[1]:
            best = (canon, score)
    return best[0]


# =========================
# EXTRACCIÓN — helpers RGC
# =========================
def _rgc_find_invoice_starts(pages: list) -> list:
    """
    Detecta inicios de factura en documentos Bolívar RGC Activa.
    El encabezado de columnas y el número de factura siempre están
    en la misma página física.
    """
    header_re = re.compile(
        r"Item\s+Num\.?\s*\nFactura\s+Val\.?\s*Bruto"
        r"(?:.*?)"
        r"C[oó]digo\s+de\s+barras\s*\n(\d{1,3})\s*\n(\d{6,10})",
        re.IGNORECASE | re.DOTALL,
    )
    header_alt = re.compile(
        r"Item\s+Num\.?\s*\nFactura\b[^\n]*\n(\d{1,3})\n(\d{6,10})\n",
        re.IGNORECASE,
    )
    starts, seen = [], set()
    for idx, page in enumerate(pages):
        for pattern in (header_re, header_alt):
            for m in pattern.finditer(page):
                inv = normalize_ocr_token(m.group(2))
                if looks_like_invoice(inv) and inv not in seen:
                    seen.add(inv)
                    starts.append((inv, idx))
    return starts


# =========================
# EXTRACCIÓN — por aseguradora
# =========================
def extract_previsora_single(pages: list) -> list:
    text  = normalize(full_text(pages))
    noise = build_noise_numbers(text)
    for p in [r"NRO\s+FACTURA\s*[:\-]?\s*(F?\d{6,10})",
              r"FACTURA\s*/\s*RADICADO\s*[:\-]?\s*(F?\d{6,10})",
              r"RECLAMACION\s*/\s*RADICADO\s*[:\-]?\s*(F?\d{6,10})",
              r"RECLAMACION\s+NO\.?\s*(F?\d{6,10})"]:
        m = re.search(p, text)
        if m:
            inv = normalize_ocr_token(m.group(1))
            if looks_like_invoice(inv) and clean_digits(inv) not in noise:
                return [inv]
    return []


def extract_solidaria_single(pages: list) -> list:
    text  = normalize(full_text(pages))
    noise = build_noise_numbers(text)
    for p in [r"NUMERO DE LA FACTURA DE LA RECLAMACION\s*[:\-]?\s*(F?\d{6,10})",
              r"FACTURA DE LA RECLAMACION\s*[:\-]?\s*(F?\d{6,10})"]:
        m = re.search(p, text)
        if m:
            inv = normalize_ocr_token(m.group(1))
            if looks_like_invoice(inv) and clean_digits(inv) not in noise:
                return [inv]
    return []


def extract_hdi_multiple(pages: list) -> list:
    text  = normalize(full_text(pages))
    noise = build_noise_numbers(text)
    out, seen = [], set()
    for n in re.findall(r"\bF?\d{6,10}\b", text):
        inv = normalize_ocr_token(n)
        if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
            seen.add(inv); out.append(inv)
    return out[:50]


def extract_envio_terceros_invoices(pages: list) -> list:
    if not pages: return []

    # RGC: usar detector de encabezados reales
    raw_starts = _rgc_find_invoice_starts(pages)
    if raw_starts:
        out, seen = [], set()
        for inv, _ in raw_starts:
            if inv not in seen:
                seen.add(inv); out.append(inv)
        return out

    # Formato clásico (línea completa: ítem + factura + valor)
    text     = normalize_keep_lines(full_text(pages))
    invoices = []
    seen     = set()
    line_re  = re.compile(r"^\s*(\d{1,3})\s+(F?\d{6,10})\s+\$?\d[\d\.\,]*", re.MULTILINE)
    for m in line_re.finditer(text):
        inv = normalize_ocr_token(m.group(2))
        if looks_like_invoice(inv) and inv not in seen:
            seen.add(inv); invoices.append(inv)
    if invoices: return invoices

    # Fallback colapsado
    collapsed = normalize(text)
    noise     = build_noise_numbers(collapsed)
    for m in re.finditer(r"\b(\d{1,3})\s+(F?\d{6,10})\s+\$?\d[\d\.\,]*", collapsed):
        inv = normalize_ocr_token(m.group(2))
        if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
            seen.add(inv); invoices.append(inv)
    return invoices


def extract_envio_terceros_invoice_starts(pages: list) -> list:
    raw = _rgc_find_invoice_starts(pages)
    if raw: return raw

    starts  = []
    seen    = set()
    line_re = re.compile(r"^\s*(\d{1,3})\s+(F?\d{6,10})\s+\$?\d[\d\.\,]*", re.MULTILINE)
    for idx, page in enumerate(pages):
        text  = normalize_keep_lines(page)
        noise = build_noise_numbers(normalize(page))
        for m in line_re.finditer(text):
            inv = normalize_ocr_token(m.group(2))
            if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                seen.add(inv); starts.append((inv, idx))
    return starts


def _axa_detect_factura_column_x(page, tolerance: float = 30.0):
    """
    Detecta el rango X (x_min, x_max) de la columna 'No. factura' en la
    página usando las coordenadas del encabezado de la tabla AXA.
    Retorna (x_min, x_max) o None si no se encuentra.
    """
    words = page.get_text("words")  # (x0,y0,x1,y1,text,block,line,word)
    # Buscar la palabra "factura" que venga precedida de "no." o "no" en la misma línea
    header_candidates = []
    # Agrupar palabras por línea (mismo bloque, misma línea)
    line_groups = {}
    for w in words:
        x0, y0, x1, y1, txt, bn, ln, wn = w
        line_groups.setdefault((bn, ln), []).append((x0, y0, x1, y1, txt.upper()))
    for key, wlist in line_groups.items():
        wlist_sorted = sorted(wlist, key=lambda w: w[0])
        texts = [w[4] for w in wlist_sorted]
        joined = " ".join(texts)
        # Detectar encabezado "No. factura" o "No.factura" o "NO FACTURA"
        if re.search(r"NO\.?\s*FACTURA", joined):
            for i, (x0, y0, x1, y1, txt) in enumerate(wlist_sorted):
                if "FACTURA" in txt:
                    header_candidates.append((x0, x1, y0))
    if not header_candidates:
        return None
    # Tomar el primer encabezado encontrado (página 1 generalmente)
    col_x0, col_x1, header_y = header_candidates[0]
    # Ampliar tolerancia lateral
    return (col_x0 - tolerance, col_x1 + tolerance)


def _axa_extract_factura_by_column(pdf_path: Path) -> list:
    """
    Estrategia PRINCIPAL para AXA tabla: detecta la columna 'No. factura'
    por coordenadas X en el encabezado y lee ÚNICAMENTE los valores que
    caen bajo esa columna en las filas de datos.
    Retorna lista de números de factura únicos.
    """
    if fitz is None or not pdf_path or not pdf_path.exists():
        return []
    try:
        doc = fitz.open(str(pdf_path))
        factura_col = None   # (x_min, x_max) detectado del encabezado
        header_y    = None   # coordenada Y del encabezado para filtrar filas
        invoices    = []
        seen        = set()

        for page_idx, page in enumerate(doc):
            words = page.get_text("words")

            # ── 1. Detectar columna en esta página si aún no se encontró ──
            if factura_col is None:
                col = _axa_detect_factura_column_x(page, tolerance=30.0)
                if col:
                    factura_col = col
                    # Guardar Y del encabezado para no leer su propio valor
                    line_groups = {}
                    for w in words:
                        x0, y0, x1, y1, txt, bn, ln, wn = w
                        line_groups.setdefault((bn, ln), []).append(
                            (x0, y0, x1, y1, txt.upper())
                        )
                    for key, wlist in line_groups.items():
                        joined = " ".join(w[4] for w in sorted(wlist, key=lambda w: w[0]))
                        if re.search(r"NO\.?\s*FACTURA", joined):
                            header_y = min(w[1] for w in wlist)
                            break

            if factura_col is None:
                continue  # Esta página no tiene encabezado aún; seguir buscando

            x_min, x_max = factura_col

            # ── 2. Recolectar palabras que caen dentro del rango X de la columna ──
            for w in words:
                x0, y0, x1, y1, txt, bn, ln, wn = w
                # Ignorar la fila del encabezado
                if header_y is not None and abs(y0 - header_y) < 5:
                    continue
                # Centroide X del token debe estar dentro del rango de la columna
                cx = (x0 + x1) / 2.0
                if not (x_min <= cx <= x_max):
                    continue
                token = normalize_ocr_token(txt)
                if not looks_like_invoice(token):
                    continue
                # Validación adicional: facturas AXA SOAT tienen 6-8 dígitos
                # Los IDs de accidentado tienen 9-10 dígitos → los descartamos
                digits = clean_digits(token)
                if len(digits) >= 9:
                    continue
                if token not in seen:
                    seen.add(token)
                    invoices.append(token)

        doc.close()
        return invoices
    except Exception:
        return []


def extract_axa_table_rows(pages: list, pdf_path: Path | None = None) -> list:
    """
    Extrae filas de la tabla AXA.
    ESTRATEGIA PRINCIPAL: detección posicional de la columna 'No. factura'
    mediante coordenadas X (fitz). Solo lee tokens que caen bajo esa columna,
    eliminando ambigüedad con Id. accidentado o Valor factura.
    FALLBACK: regex secuencial con validación de longitud de dígitos.
    """
    rows = []
    seen = set()

    # ── Regex fallback (PyPDF2 / texto plano) ──
    # NOTA: factura limitada a 6-8 dígitos para no confundir con cédulas (9-10)
    row_re = re.compile(
        r"^\s*"
        r"(?P<reclamo>\d{6,12})\s+"
        r"(?P<fecha>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<poliza>\d{8,14})\s+"
        r"(?P<ident>\d{7,12})\s+"
        r"(?P<factura>F?\d{6,8})\s+"        # ← máx 8 dígitos: excluye cédulas de 9-10
        r"(?P<valor>\d{4,12})\s+"
        r"(?P<codigo>\d{3,10})(?=\s|[A-Z\$]|\Z)"
    )

    def try_add(page_idx: int, line_n: str):
        if not line_n: return
        if re.search(r"NO\.?\s*RECLAMO", line_n) and re.search(r"NO\.?\s*FACTURA", line_n): return
        if line_n.startswith("PAGINA "): return
        m = row_re.match(line_n)
        if not m: return
        reclamo = clean_digits(m.group("reclamo"))
        fecha   = m.group("fecha")
        poliza  = clean_digits(m.group("poliza"))
        ident   = clean_digits(m.group("ident"))
        factura = normalize_ocr_token(m.group("factura"))
        valor   = clean_digits(m.group("valor"))
        codigo  = clean_digits(m.group("codigo"))
        if len(reclamo) < 6 or len(poliza) < 8 or len(ident) < 7: return
        if not looks_like_invoice(factura): return
        if len(valor) < 4 or len(codigo) < 3: return
        if clean_digits(factura) in (ident, valor): return
        if len(clean_digits(factura)) >= 9: return   # guardia extra: no cédulas
        key = (page_idx, reclamo, fecha, poliza, ident, factura, valor, codigo)
        if key in seen: return
        seen.add(key)
        rows.append({"page": page_idx, "reclamo": reclamo, "fecha": fecha,
                     "poliza": poliza, "ident": ident, "factura": factura,
                     "valor": valor, "codigo": codigo})

    # ── Estrategia 1: extracción posicional por columna (fitz) ──
    col_invoices = _axa_extract_factura_by_column(pdf_path)
    if col_invoices:
        # Reconstruir estructura mínima de rows compatible con el resto del pipeline
        # usando la lista de facturas detectadas posicionalmente
        seen_inv = set()
        for inv in col_invoices:
            if inv not in seen_inv:
                seen_inv.add(inv)
                rows.append({"page": 0, "reclamo": "", "fecha": "",
                             "poliza": "", "ident": "", "factura": inv,
                             "valor": "", "codigo": ""})
        return rows

    # ── Estrategia 2: fitz línea a línea (fallback layout) ──
    fitz_lines = []
    if fitz is not None and pdf_path and pdf_path.exists():
        try:
            doc = fitz.open(str(pdf_path))
            for page_idx, page in enumerate(doc):
                words  = page.get_text("words")
                groups = {}
                for w in words:
                    x0, y0, x1, y1, txt, bn, ln, wn = w
                    groups.setdefault((bn, ln), []).append((wn, x0, txt))
                for key in sorted(groups):
                    parts = [t for _, _, t in sorted(groups[key], key=lambda x: (x[0], x[1]))]
                    line  = " ".join(parts)
                    if line.strip():
                        fitz_lines.append((page_idx, normalize(line)))
            doc.close()
        except Exception:
            fitz_lines = []

    if fitz_lines:
        for page_idx, line_n in fitz_lines:
            try_add(page_idx, line_n)
        if rows: return rows

    # ── Estrategia 3: PyPDF2 texto plano ──
    for page_idx, page in enumerate(pages):
        for raw_line in normalize_keep_lines(page).splitlines():
            try_add(page_idx, normalize(raw_line.strip()))
    return rows


def extract_axa_table_invoices(pages: list, pdf_path: Path | None = None) -> list:
    rows = extract_axa_table_rows(pages, pdf_path=pdf_path)
    seen, out = set(), []
    for row in rows:
        inv = row["factura"]
        if inv not in seen:
            seen.add(inv); out.append(inv)
    return out


def restore_axa_invoice_family(pages: list, invoices: list) -> list:
    if not invoices: return []
    text = normalize(full_text(pages))
    out, seen = [], set()
    for inv in invoices:
        inv_n    = normalize_ocr_token(inv)
        inv_d    = clean_digits(inv_n)
        repaired = f"F{inv_d}" if (inv_d and re.search(rf"\bF{re.escape(inv_d)}\b", text)) \
                   else (inv_n if inv_n.startswith("F") else inv_d)
        repaired = normalize_ocr_token(repaired)
        if repaired not in seen:
            seen.add(repaired); out.append(repaired)
    return out


def extract_axa_single_objection(pages: list) -> list:
    if is_axa_table_format(pages): return []
    invoices = []
    seen     = set()
    noise    = build_noise_numbers(normalize(full_text(pages)))
    for page in pages:
        for raw_line in normalize_keep_lines(page).splitlines():
            line   = normalize(raw_line)
            tokens = [normalize_ocr_token(t) for t in re.findall(r"[A-Z0-9\-]+", line)]
            if len(tokens) < 5: continue
            for i in range(len(tokens)-4):
                poliza,ident,factura,valor,codigo = tokens[i:i+5]
                if len(clean_digits(poliza)) < 6: continue
                if len(clean_digits(ident))  < 7: continue
                if not looks_like_invoice(factura): continue
                if len(clean_digits(valor))  < 4: continue
                if len(clean_digits(codigo)) < 3: continue
                if clean_digits(factura) in (clean_digits(ident),clean_digits(valor)): continue
                if clean_digits(factura) in noise: continue
                if factura not in seen:
                    seen.add(factura); invoices.append(factura)
    return invoices


def extract_axa_letter_invoice(pages: list) -> list:
    """
    Soporta:
    1. Campo 'Factura: XXXXX' en encabezado de liquidación individual
    2. Fila de encabezado AXA carta: "No. XXXX No. XXXX Siniestro XXXX No. FACTURA"
       El número de factura es SIEMPRE el último 'No. XXXXXXX' de esa línea.
    3. 'NO. XXXXXXX' en cartas clásicas AXA (fallback)
    """
    if not pages: return []
    found = []
    seen  = set()
    noise = build_noise_numbers(normalize(full_text(pages)))

    # Capa 1: campo explícito FACTURA: XXXXX
    for page in pages[:5]:
        for line in normalize_keep_lines(page).splitlines()[:50]:
            m = re.search(r"\bFACTURA\s*:?\s*(F?\d{6,10})\b", normalize(line))
            if m:
                inv = normalize_ocr_token(m.group(1))
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); found.append(inv)
    if found: return found

    # Capa 2: fila de encabezado AXA carta
    # La línea tiene "NO." y "SINIESTRO" y al menos 3 números.
    # El número de factura es SIEMPRE el último token de esa línea.
    # PyPDF2 puede extraer la línea como:
    #   "NO. 22844858 NO. 159731 SINIESTRO 945937 NO. F7193080"
    # o partida en dos líneas consecutivas cuando el PDF tiene columnas.
    # Se busca en las primeras 30 líneas de la primera página.
    lines_p0 = normalize_keep_lines(pages[0]).splitlines()[:35]
    for idx, line in enumerate(lines_p0):
        line_n = normalize(line)
        # Combinar con línea siguiente por si PyPDF2 partió el encabezado
        if idx + 1 < len(lines_p0):
            line_n_ext = normalize(line + " " + lines_p0[idx + 1])
        else:
            line_n_ext = line_n

        for candidate in [line_n_ext, line_n]:
            has_no  = bool(re.search(r"\bNO\.?\b", candidate))
            has_sin = "SINIESTRO" in candidate
            if not (has_no and has_sin):
                continue
            # Capturar todos los tokens alfanuméricos que parezcan facturas
            # (dígitos solos O letra+dígitos como F7193080)
            tokens = re.findall(r"\b([A-Z]?\d{4,10})\b", candidate)
            if len(tokens) >= 3:
                inv = normalize_ocr_token(tokens[-1])
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); found.append(inv)
                    return found
            break  # si la combinada tampoco sirvió, no reintentar con la sola
    if found: return found

    # Capa 3: fallback — cualquier No. XXXXXXX en primeras líneas
    for line in normalize_keep_lines(pages[0]).splitlines()[:25]:
        line_n = normalize(line)
        for p in [r"\bNO\.\s*(F?\d{6,10})\b", r"\bNO\s+(F?\d{6,10})\b"]:
            m = re.search(p, line_n)
            if m:
                inv = normalize_ocr_token(m.group(1))
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); found.append(inv)
    return found


def extract_mundial_detail_reclamations(pages: list) -> list:
    if not pages: return []
    out   = []
    seen  = set()
    noise = build_noise_numbers(normalize(full_text(pages)))
    search = pages[1:] if len(pages) > 1 else pages
    for page in search:
        txt = normalize(page)
        for p in [r"NUMERO\s+DE\s+RECLAMACION\s*[:\s]\s*(F?\d{4,10})\b"]:
            for m in re.finditer(p, txt):
                inv = normalize_ocr_token(m.group(1))
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); out.append(inv)
    return out


def mundial_page_has_reclamation(page_text: str, invoice: str) -> bool:
    t   = normalize(page_text)
    inv = normalize_ocr_token(invoice)
    return bool(re.search(
        rf"NUMERO\s+DE\s+RECLAMACION\s*[:\s]\s*{re.escape(inv)}\b", t
    ))


def extract_mundial_summary_first_page(pages: list) -> list:
    if not pages: return []
    first = normalize_keep_lines(pages[0])
    noise = build_noise_numbers(normalize(first))
    in_summary    = False
    summary_lines = []
    for line in [l.strip() for l in first.splitlines() if l.strip()]:
        ln = normalize(line)
        if re.search(r"NUMERO DE RECLAMACION(ES)?", ln):
            in_summary = True; continue
        if in_summary and re.search(r"ADJUNTO ENCONTRARA|SI DESEA COMUNICARSE", ln):
            break
        if in_summary:
            summary_lines.append(ln)

    row_re = re.compile(
        r"^\s*(F?\d{4,10})"
        r"(?:\s+\d{4,10})?"
        r"(?:\s+\$?\d[\d\.\,]*)?\s*$"
    )
    candidates, seen = [], set()
    for ln in summary_lines:
        m = row_re.match(ln)
        if m:
            inv = normalize_ocr_token(m.group(1))
            if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                seen.add(inv); candidates.append(inv)

    details    = extract_mundial_detail_reclamations(pages)
    detail_set = {normalize_ocr_token(x) for x in details}
    filtered   = [inv for inv in candidates if inv in detail_set]
    return filtered or details or candidates


def extract_estado_multiple(pages: list) -> list:
    text  = normalize(full_text(pages))
    noise = build_noise_numbers(text)
    out, seen = [], set()
    for m in re.finditer(r"FACTURA\s*:\s*(F?\d{6,10})", text):
        inv = normalize_ocr_token(m.group(1))
        if inv not in seen and clean_digits(inv) not in noise:
            seen.add(inv); out.append(inv)
    return out


def extract_sura_invoices(pages: list) -> list:
    out   = []
    seen  = set()
    noise = build_noise_numbers(normalize(full_text(pages)))
    if pages:
        m = re.search(r"FACTURA\s*\(S\)\s*:?\s*([\d\s;,]+)", normalize(pages[0]))
        if m:
            for t in re.split(r"[;,\s]+", m.group(1)):
                t = t.strip()
                if not t: continue
                inv = normalize_ocr_token(t)
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); out.append(inv)
    if out: return out
    for page in pages:
        for m in re.finditer(r"\bFACTURA\s*:\s*(F?\d{6,10})\b", normalize(page)):
            inv = normalize_ocr_token(m.group(1))
            if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                seen.add(inv); out.append(inv)
    return out


def extract_bolivar_invoice_starts(pages: list) -> list:
    starts, seen = [], set()
    for idx, page in enumerate(pages):
        txt   = normalize_keep_lines(page)
        noise = build_noise_numbers(normalize(page))
        for line in txt.splitlines():
            ln = normalize(line)
            m  = re.match(r"^\s*\d+\s+(F?\d{6,10})\b", ln)
            if m:
                inv = normalize_ocr_token(m.group(1))
                if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                    seen.add(inv); starts.append((inv, idx))
                continue
            if any(k in ln for k in ["VAL FACTURA","DOCUMENTO","PACIENTE","CODIGO DE BARRAS"]):
                for n in re.findall(r"\bF?\d{6,10}\b", ln):
                    inv = normalize_ocr_token(n)
                    if looks_like_invoice(inv) and inv not in seen and clean_digits(inv) not in noise:
                        seen.add(inv); starts.append((inv, idx))
    return starts


def extract_bolivar_invoices(pages: list) -> list:
    out   = []
    seen  = set()
    noise = build_noise_numbers(normalize(full_text(pages)))
    def add(value: str):
        d = clean_digits(value)
        if len(d) >= 6 and d not in seen and d not in noise:
            seen.add(d); out.append(d)
    for page in pages:
        t      = normalize_keep_lines(page)
        t_norm = normalize(page)
        for m in re.findall(r"NUM\.?\s*FACTURA\s+(\d{6,10})", t_norm):
            add(m)
        for line in t.splitlines():
            ln = normalize(line)
            m  = re.match(r"^\s*\d+\s+(F?\d{6,10})\b", ln)
            if m:
                add(m.group(1)); continue
            if any(k in ln for k in ["VAL FACTURA","DOCUMENTO","PACIENTE","CODIGO DE BARRAS"]):
                for n in re.findall(r"\b\d{6,10}\b", ln):
                    if not any(n.startswith(x) for x in ["900","470","839"]):
                        add(n)
    return out


def extract_single_generic(pages: list) -> list:
    """Solo acepta números precedidos de palabras clave de factura."""
    text  = normalize(full_text(pages))
    noise = build_noise_numbers(text)
    for p in [
        r"N[°O]\.?\s*FACTURA\s+RECLAMACI[OÓ]N\s*(F?\d{6,10})",
        r"NUMERO\s+DE\s+RECLAMACION\s*[:\-]?\s*(F?\d{6,10})",
        r"RECLAMACION\s*/\s*RADICADO\s*[:\-]?\s*(F?\d{6,10})",
        r"NRO\.?\s*FACTURA\s*[:\-]?\s*(F?\d{6,10})",
        r"N[°O]\.?\s*DE\s*FACTURA\s*[:\-]?\s*(F?\d{6,10})",
        r"FACTURA\s*:\s*(F?\d{6,10})",
        r"FACTURA\(S\)\s*[:\-]?\s*(F?\d{6,10})",
        r"ASUNTO[^:]*:\s*[^0-9]*(F?\d{6,10})",
    ]:
        m = re.search(p, text, re.DOTALL)
        if m:
            inv = normalize_ocr_token(m.group(1))
            if looks_like_invoice(inv) and clean_digits(inv) not in noise:
                return [inv]
    return []


# =========================
# SEGMENTACIÓN
# =========================
def pages_containing_invoice(pages: list, invoice: str, insurer: str) -> list:
    inv   = normalize_ocr_token(invoice)
    found = []
    for i, txt in enumerate(pages):
        t = normalize(txt)
        if insurer == "MUNDIAL DE SEGUROS S.A.":
            if mundial_page_has_reclamation(t, inv): found.append(i)
        elif insurer == "SEGUROS DEL ESTADO S.A.":
            if f"FACTURA: {inv}" in t or f"FACTURA {inv}" in t: found.append(i)
        else:
            if re.search(rf"\b{re.escape(inv)}\b", t): found.append(i)
    return found


def infer_bolivar_rgc_ranges(pages: list, invoices: list) -> dict:
    """
    REGLA validada con documento real 71-72599 (612 páginas, 18 facturas):
    - Inicio N  = página con encabezado de columnas + número de factura
    - Fin N     = misma página donde empieza factura N+1
      (contiene cierre de N e inicio de N+1 simultáneamente)
    - Última    = hasta última página
    """
    invoice_set = {normalize_ocr_token(x) for x in invoices}
    raw_starts  = _rgc_find_invoice_starts(pages)
    ordered, seen = [], set()
    for inv, pg in raw_starts:
        if inv in invoice_set and inv not in seen:
            seen.add(inv); ordered.append((inv, pg))

    # Fallback si el encabezado no matcheó
    if not ordered:
        fb, seen_fb = [], set()
        for idx, page in enumerate(pages):
            t     = normalize(page)
            noise = build_noise_numbers(t)
            for inv in invoices:
                inv_n = normalize_ocr_token(inv)
                if inv_n in seen_fb: continue
                if re.search(rf"\b{re.escape(inv_n)}\b", t) and inv_n not in noise:
                    seen_fb.add(inv_n); fb.append((inv_n, idx))
        ordered = sorted(fb, key=lambda x: x[1])

    if not ordered: return {}

    ranges = {}
    for i, (inv, start_pg) in enumerate(ordered):
        end_pg = ordered[i+1][1] if i+1 < len(ordered) else len(pages)-1
        ranges[inv] = (start_pg, end_pg)
    return ranges


def infer_envio_terceros_ranges_classic(pages: list, invoices: list) -> dict:
    """Segmentación para formato envío clásico (no RGC)."""
    starts_list = extract_envio_terceros_invoice_starts(pages)
    invoice_set = {normalize_ocr_token(x) for x in invoices}
    ordered, seen = [], set()
    for inv, pg in starts_list:
        inv_n = normalize_ocr_token(inv)
        if inv_n in invoice_set and inv_n not in seen:
            seen.add(inv_n); ordered.append((inv_n, pg))
    if not ordered: return {}
    ranges = {}
    for i, (inv, start_pg) in enumerate(ordered):
        if i+1 < len(ordered):
            next_pg = ordered[i+1][1]
            end_pg  = max(next_pg-1, start_pg)
        else:
            end_pg = len(pages)-1
        ranges[inv] = (start_pg, end_pg)
    return ranges


def infer_mundial_ranges_by_totals(pages: list, invoices: list) -> dict:
    """
    Segmenta PDFs de Mundial buscando el cierre de cada liquidación.
    Cubre tres variantes reales del campo de cierre:
      - 'Valor Pagado Actual: $X'
      - 'Valor Pagado Acumulado: $X'
      - 'Valor Pagado: $X'  (formato sin sufijo)
    """
    end_marker = re.compile(
        r"VALOR\s+PAGADO\s*(?:ACTUAL|ACUMULADO)?\s*:\s*\$[\d\.\,]+",
        re.IGNORECASE,
    )
    starts = {}
    for idx, page in enumerate(pages[1:], start=1):
        t = normalize(page)
        for inv in invoices:
            inv_n = normalize_ocr_token(inv)
            if inv_n not in starts and mundial_page_has_reclamation(t, inv_n):
                starts[inv_n] = idx
    if not starts: return {}

    ordered = sorted(
        [(inv, starts[normalize_ocr_token(inv)])
         for inv in invoices if normalize_ocr_token(inv) in starts],
        key=lambda x: x[1],
    )
    ranges = {}
    for i, (inv, start_pg) in enumerate(ordered):
        inv_n  = normalize_ocr_token(inv)
        end_pg = start_pg
        for pg_idx in range(start_pg, len(pages)):
            if i+1 < len(ordered) and pg_idx > start_pg and pg_idx >= ordered[i+1][1]:
                end_pg = ordered[i+1][1]-1
                break
            if end_marker.search(normalize(pages[pg_idx])):
                end_pg = pg_idx
                break
        if end_pg < start_pg: end_pg = start_pg
        ranges[inv_n] = (start_pg, end_pg)
    return ranges


def _sura_notification_page(pages: list) -> int:
    """
    Retorna el índice de la página que contiene la tabla de notificación
    de pago Sura (la que tiene ORDEN PAGO + VR. AUTORIZADO + TOTAL A PAGAR).
    Si no la encuentra retorna -1.
    """
    tabla_re = re.compile(
        r"ORDEN\s+PAGO.{0,60}(?:VR\.?\s*AUTORIZADO|TOTAL\s+A\s+PAGAR)",
        re.IGNORECASE | re.DOTALL,
    )
    for idx, page in enumerate(pages):
        if tabla_re.search(normalize(page)):
            return idx
    return -1


def infer_sura_ranges(pages: list, invoices: list) -> dict:
    # Detectar si hay página de notificación previa a la liquidación
    notif_pg  = _sura_notification_page(pages)
    # Las páginas de liquidación empiezan después de la tabla de notificación
    liq_start = notif_pg + 1 if notif_pg >= 0 else 0
    liq_pages = pages[liq_start:]  # subconjunto sobre el que buscamos

    end_marker = re.compile(r"VALOR\s+PAGADO\s*:\s*\$[\d\.\,]+", re.IGNORECASE)
    starts = {}
    for idx, page in enumerate(liq_pages):
        t = normalize(page)
        for inv in invoices:
            inv_n = normalize_ocr_token(inv)
            if inv_n not in starts and re.search(rf"\bFACTURA\s*:\s*{re.escape(inv_n)}\b", t):
                starts[inv_n] = idx + liq_start  # índice absoluto

    ordered = sorted(
        [(inv, starts[normalize_ocr_token(inv)])
         for inv in invoices if normalize_ocr_token(inv) in starts],
        key=lambda x: x[1],
    )
    ranges = {}
    for i, (inv, start_pg) in enumerate(ordered):
        inv_n  = normalize_ocr_token(inv)
        end_pg = start_pg
        for pg_idx in range(start_pg, len(pages)):
            if i+1 < len(ordered) and pg_idx > start_pg and pg_idx >= ordered[i+1][1]:
                end_pg = ordered[i+1][1]-1; break
            if end_marker.search(normalize(pages[pg_idx])):
                end_pg = pg_idx; break
        if end_pg < start_pg: end_pg = start_pg
        ranges[inv_n] = (start_pg, end_pg)
    return ranges


def infer_bolivar_ranges(pages: list, invoices: list) -> dict:
    starts  = extract_bolivar_invoice_starts(pages)
    ordered = [(inv, pg) for inv, pg in starts if inv in invoices]
    if len(ordered) != len(invoices):
        ordered = []
        for inv in invoices:
            hits = pages_containing_invoice(pages, inv, "SEGUROS COMERCIALES BOLIVAR S.A.")
            if hits: ordered.append((inv, min(hits)))
    if not ordered: return {}
    ordered.sort(key=lambda x: (x[1], invoices.index(x[0])))
    ranges = {}
    for i, (inv, start_pg) in enumerate(ordered):
        next_pg = ordered[i+1][1] if i+1 < len(ordered) else len(pages)-1
        end_pg  = next_pg if next_pg != start_pg else start_pg
        ranges[inv] = (start_pg, end_pg)
    return ranges


def infer_generic_segment_ranges(pages: list, invoices: list, insurer: str) -> dict:
    offset = 1 if insurer == "MUNDIAL DE SEGUROS S.A." and len(pages) > 1 else 0
    search = pages[1:] if offset else pages
    starts = []
    for inv in invoices:
        hits = pages_containing_invoice(search, inv, insurer)
        if hits: starts.append((inv, min(hits)+offset))
    starts.sort(key=lambda x: (x[1], invoices.index(x[0])))
    if not starts: return {}
    ranges = {}
    for i, (inv, start_pg) in enumerate(starts):
        if i+1 < len(starts):
            end_pg = max(starts[i+1][1]-1, start_pg)
        else:
            last   = pages_containing_invoice(search, inv, insurer)
            end_pg = (max(last)+offset) if last else start_pg
        ranges[inv] = (start_pg, end_pg)
    return ranges


def infer_page_ranges(pages: list, invoices: list, insurer: str) -> dict:
    if insurer == "SEGUROS GENERALES SURAMERICANA S.A.":
        return infer_sura_ranges(pages, invoices)

    if insurer in {"SEGUROS COMERCIALES BOLIVAR S.A.", "LA PREVISORA S.A."}:
        if is_envio_documentos_terceros_format(pages):
            if is_bolivar_rgc_format(pages):
                return infer_bolivar_rgc_ranges(pages, invoices)
            return infer_envio_terceros_ranges_classic(pages, invoices)

    if insurer == "SEGUROS COMERCIALES BOLIVAR S.A.":
        return infer_bolivar_ranges(pages, invoices)

    if insurer == "MUNDIAL DE SEGUROS S.A." and len(pages) >= 2:
        return infer_mundial_ranges_by_totals(pages, invoices)

    return infer_generic_segment_ranges(pages, invoices, insurer)


def write_pdf_range(reader: PdfReader, out_path: Path, start_idx: int, end_idx: int):
    writer = PdfWriter()
    for i in range(start_idx, end_idx+1):
        writer.add_page(reader.pages[i])
    with open(out_path, "wb") as f:
        writer.write(f)


# =========================
# RULE ENGINE
# =========================
def classify_mode(insurer: str, pages: list, pdf_path: Path | None = None) -> str:
    if insurer == "SEGUROS GENERALES SURAMERICANA S.A.":
        return "segment" if len(extract_sura_invoices(pages)) > 1 else "single"

    if insurer == "AXA COLPATRIA SEGUROS S.A.":
        if is_axa_table_format(pages):
            return "duplicate" if len(extract_axa_table_invoices(pages, pdf_path)) > 1 else "single"
        if is_axa_letter_format(pages):
            return "single"
        return "duplicate" if len(extract_axa_single_objection(pages)) > 1 else "single"

    if insurer == "SEGUROS COMERCIALES BOLIVAR S.A.":
        if is_envio_documentos_terceros_format(pages):
            return "segment" if len(extract_envio_terceros_invoices(pages)) > 1 else "single"
        return "segment" if len(extract_bolivar_invoices(pages)) > 1 else "single"

    if insurer == "LA PREVISORA S.A.":
        if is_envio_documentos_terceros_format(pages):
            return "segment" if len(extract_envio_terceros_invoices(pages)) > 1 else "single"
        return "single" if len(extract_previsora_single(pages)) <= 1 else "duplicate"

    if insurer == "HDI SEGUROS COLOMBIA S.A.":
        return "duplicate" if len(extract_hdi_multiple(pages)) > 1 else "single"

    if insurer == "SEGUROS DEL ESTADO S.A.":
        return "segment" if len(extract_estado_multiple(pages)) > 1 else "single"

    if insurer == "MUNDIAL DE SEGUROS S.A.":
        return "segment" if len(extract_mundial_summary_first_page(pages)) > 1 else "single"

    return "single"


def extract_invoices_for_doc(insurer: str, pages: list, pdf_path: Path | None = None) -> list:
    if insurer == "SEGUROS GENERALES SURAMERICANA S.A.":
        return extract_sura_invoices(pages)

    if insurer == "AXA COLPATRIA SEGUROS S.A.":
        if is_axa_table_format(pages):
            return extract_axa_table_invoices(pages, pdf_path=pdf_path)
        nums = extract_axa_letter_invoice(pages)
        if nums: return nums
        return extract_axa_single_objection(pages)

    if insurer == "SEGUROS COMERCIALES BOLIVAR S.A.":
        if is_envio_documentos_terceros_format(pages):
            return extract_envio_terceros_invoices(pages)
        starts = extract_bolivar_invoice_starts(pages)
        seen, out = set(), []
        for inv, _ in starts:
            k = inv if inv.startswith("F") else clean_digits(inv)
            if k not in seen:
                seen.add(k); out.append(k)
        return out or extract_bolivar_invoices(pages)

    if insurer == "LA PREVISORA S.A.":
        if is_envio_documentos_terceros_format(pages):
            return extract_envio_terceros_invoices(pages)
        return extract_previsora_single(pages)

    if insurer == "ASEGURADORA SOLIDARIA DE COLOMBIA S.A.":
        return extract_solidaria_single(pages)

    if insurer == "SEGUROS DEL ESTADO S.A.":
        return extract_estado_multiple(pages)

    if insurer == "MUNDIAL DE SEGUROS S.A.":
        nums = extract_mundial_summary_first_page(pages)
        if nums: return nums

    if insurer == "HDI SEGUROS COLOMBIA S.A.":
        nums = extract_hdi_multiple(pages)
        if nums: return nums

    return extract_single_generic(pages)


# =========================
# OUTPUT / REPORT
# =========================
def register(rows, invoice_name, insurer, original_name, final_path):
    rows.append({
        "factura procesada":  invoice_name,
        "aseguradora":        insurer,
        "documento original": original_name,
        "ruta destino final": str(final_path),
    })


def unique_dest(folder: Path, filename: str) -> Path:
    dest = folder / filename
    if not dest.exists(): return dest
    stem, suf = dest.stem, dest.suffix
    i = 1
    while True:
        alt = folder / f"{stem}_{i}{suf}"
        if not alt.exists(): return alt
        i += 1


def send_to_manual(pdf_path: Path, rows, reason: str = ""):
    dest = unique_dest(safe_mkdir(BASE_DIR / MANUAL_FOLDER), pdf_path.name)
    shutil.move(str(pdf_path), str(dest))
    register(rows, reason or "NO IDENTIFICADA", MANUAL_FOLDER, pdf_path.name, dest)


def remove_source(pdf_path: Path):
    """Mueve el original a PROCESADOS en lugar de borrarlo."""
    try:
        if pdf_path.exists():
            dest = unique_dest(safe_mkdir(BASE_DIR / PROCESSED_FOLDER), pdf_path.name)
            shutil.move(str(pdf_path), str(dest))
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo mover {pdf_path.name}: {e}")


# =========================
# PROCESS
# =========================
def process_pdf(pdf_path: Path, rows):
    try:
        pages  = page_texts(pdf_path)
        reader = PdfReader(str(pdf_path))
    except Exception as e:
        send_to_manual(pdf_path, rows, f"ERROR LECTURA: {e}"); return

    insurer = detect_insurer(pages)
    if not insurer:
        try:
            pages   = page_texts_force_ocr(pdf_path)
            insurer = detect_insurer(pages)
        except Exception:
            pass
    if not insurer:
        send_to_manual(pdf_path, rows, "ASEGURADORA NO IDENTIFICADA"); return

    _, _, prefixes = matrix_lookup_from_text(full_text(pages))
    mode     = classify_mode(insurer, pages, pdf_path=pdf_path)
    invoices = extract_invoices_for_doc(insurer, pages, pdf_path=pdf_path)

    # ---- AXA post-processing ----
    if insurer == "AXA COLPATRIA SEGUROS S.A.":
        invoices = restore_axa_invoice_family(pages, invoices)
        invoices = filter_invoices_by_prefixes(invoices, prefixes)
        invoices = restore_axa_invoice_family(pages, invoices)

        if is_axa_table_format(pages) and len(invoices) <= 1:
            try:
                ocr_p   = page_texts_force_ocr(pdf_path)
                ocr_inv = extract_invoices_for_doc(insurer, ocr_p, pdf_path=pdf_path)
                ocr_inv = restore_axa_invoice_family(ocr_p, ocr_inv)
                ocr_inv = filter_invoices_by_prefixes(ocr_inv, prefixes)
                ocr_inv = restore_axa_invoice_family(ocr_p, ocr_inv)
                merged, seen = [], set()
                for inv in invoices + ocr_inv:
                    inv_n = normalize_ocr_token(inv)
                    if inv_n not in seen:
                        seen.add(inv_n); merged.append(inv_n)
                if merged:
                    invoices = merged; pages = ocr_p
            except Exception:
                pass
        mode = "duplicate" if len(invoices) > 1 else "single"

    # ---- Bolívar / Previsora envío reconciliación ----
    def reconcile_envio(label):
        nonlocal invoices, pages, mode
        m = re.search(r"CANTIDAD\s+FACTURAS\s*:\s*(\d+)", normalize(full_text(pages)))
        if m:
            expected = int(m.group(1))
            if len(invoices) != expected:
                try:
                    ocr_p   = page_texts_force_ocr(pdf_path)
                    ocr_inv = extract_envio_terceros_invoices(ocr_p)
                    if len(ocr_inv) >= len(invoices):
                        invoices = ocr_inv; pages = ocr_p
                except Exception:
                    pass
        mode = "segment" if len(invoices) > 1 else "single"

    if insurer == "SEGUROS COMERCIALES BOLIVAR S.A." and is_envio_documentos_terceros_format(pages):
        reconcile_envio("BOLIVAR")
    if insurer == "LA PREVISORA S.A." and is_envio_documentos_terceros_format(pages):
        reconcile_envio("PREVISORA")

    # ---- Mundial OCR fallback ----
    if insurer == "MUNDIAL DE SEGUROS S.A." and mode == "segment":
        try:
            rt = infer_page_ranges(pages, invoices, insurer) if invoices else {}
            if not invoices or len(rt) != len(invoices):
                pages    = page_texts_force_ocr(pdf_path)
                invoices = extract_invoices_for_doc(insurer, pages, pdf_path=pdf_path)
                mode     = classify_mode(insurer, pages, pdf_path=pdf_path)
        except Exception:
            pass

    if not invoices:
        send_to_manual(pdf_path, rows, "FACTURA NO IDENTIFICADA"); return

    out_dir = safe_mkdir(BASE_DIR / insurer)

    try:
        # ---- SINGLE ----
        if mode == "single":
            final_name = arm_invoice(invoices[0], prefixes) + ".pdf"
            out_path   = unique_dest(out_dir, final_name)
            # Para Sura: si hay tabla de notificación previa, recortar desde
            # la página siguiente a esa tabla en vez de copiar el PDF completo
            if insurer == "SEGUROS GENERALES SURAMERICANA S.A.":
                notif_pg  = _sura_notification_page(pages)
                liq_start = notif_pg + 1 if notif_pg >= 0 else 0
                if liq_start > 0:
                    write_pdf_range(reader, out_path, liq_start, len(reader.pages) - 1)
                else:
                    shutil.copy2(str(pdf_path), str(out_path))
            else:
                shutil.copy2(str(pdf_path), str(out_path))
            register(rows, out_path.stem, insurer, pdf_path.name, out_path)
            remove_source(pdf_path)
            return

        # ---- DUPLICATE ----
        if mode == "duplicate":
            for raw in invoices:
                final_name = arm_invoice(raw, prefixes) + ".pdf"
                out_path   = unique_dest(out_dir, final_name)
                shutil.copy2(str(pdf_path), str(out_path))
                register(rows, out_path.stem, insurer, pdf_path.name, out_path)
            remove_source(pdf_path)
            return

        # ---- SEGMENT ----
        if mode == "segment":
            ranges = infer_page_ranges(pages, invoices, insurer)

            # Validación de conteo para formatos envío
            if insurer in {"SEGUROS COMERCIALES BOLIVAR S.A.", "LA PREVISORA S.A."} \
                    and is_envio_documentos_terceros_format(pages):
                m = re.search(r"CANTIDAD\s+FACTURAS\s*:\s*(\d+)", normalize(full_text(pages)))
                starts_list  = extract_envio_terceros_invoice_starts(pages)
                envio_invs, seen = [], set()
                for inv, _ in starts_list:
                    inv_n = normalize_ocr_token(inv)
                    if inv_n not in seen:
                        seen.add(inv_n); envio_invs.append(inv_n)
                if m:
                    expected = int(m.group(1))
                    if len(envio_invs) == expected:
                        invoices = envio_invs
                        ranges   = infer_page_ranges(pages, invoices, insurer)
                    else:
                        label = "BOLIVAR" if insurer == "SEGUROS COMERCIALES BOLIVAR S.A." else "PREVISORA"
                        send_to_manual(pdf_path, rows, f"{label} ESPERABA {expected} Y LEYO {len(envio_invs)}")
                        return
                elif envio_invs:
                    invoices = envio_invs
                    ranges   = infer_page_ranges(pages, invoices, insurer)

            elif insurer == "SEGUROS COMERCIALES BOLIVAR S.A.":
                m = re.search(r"CANTIDAD\s+FACTURAS\s*:\s*(\d+)", normalize(full_text(pages)))
                if m:
                    expected = int(m.group(1))
                    starts_list  = extract_bolivar_invoice_starts(pages)
                    bolivar_invs, seen = [], set()
                    for inv, _ in starts_list:
                        k = inv if inv.startswith("F") else clean_digits(inv)
                        if k not in seen:
                            seen.add(k); bolivar_invs.append(k)
                    if len(bolivar_invs) == expected:
                        invoices = bolivar_invs
                        ranges   = infer_bolivar_ranges(pages, invoices)
                    else:
                        send_to_manual(pdf_path, rows,
                                       f"BOLIVAR ESPERABA {expected} Y LEYO {len(bolivar_invs)}")
                        return

            if len(ranges) != len(invoices):
                send_to_manual(pdf_path, rows, f"RANGOS INCOMPLETOS {len(ranges)}/{len(invoices)}")
                return

            for raw in invoices:
                key = normalize_ocr_token(raw)
                if key not in ranges:
                    send_to_manual(pdf_path, rows, f"RANGO FALTANTE {raw}"); return
                start, end = ranges[key]
                if end < start: end = start
                final_name = arm_invoice(raw, prefixes) + ".pdf"
                out_path   = unique_dest(out_dir, final_name)
                write_pdf_range(reader, out_path, start, end)
                register(rows, out_path.stem, insurer, pdf_path.name, out_path)

            remove_source(pdf_path)
            return

        send_to_manual(pdf_path, rows, "MODO NO SOPORTADO")

    except Exception as e:
        send_to_manual(pdf_path, rows, f"ERROR PROCESO: {e}")


# =========================
# MAIN
# =========================
def main():
    mostrar_marca_agua()

    if not MATRIX:
        print("[ERROR CRÍTICO] La matriz está vacía. Verifique MATRIZ_IPS.xlsx.")
        input("Presione Enter para salir...")
        return

    safe_mkdir(BASE_DIR / MANUAL_FOLDER)
    safe_mkdir(BASE_DIR / PROCESSED_FOLDER)
    rows = []

    pdfs  = sorted([p for p in INPUT_DIR.iterdir()
                    if p.is_file() and p.suffix.lower() == ".pdf"
                    and p.parent == INPUT_DIR])
    total = len(pdfs)
    print(f"[INFO] {total} archivos PDF encontrados.\n")

    for idx, pdf in enumerate(pdfs, 1):
        print(f"[{idx}/{total}] {pdf.name}")
        process_pdf(pdf, rows)

    report = pd.DataFrame(rows, columns=[
        "factura procesada", "aseguradora",
        "documento original", "ruta destino final",
    ])
    report.to_excel(BASE_DIR / REPORT_NAME, index=False)
    print(f"\n[OK] Proceso finalizado. Reporte: {BASE_DIR / REPORT_NAME}")
    input("Presione Enter para cerrar...")


if __name__ == "__main__":
    main()