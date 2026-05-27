"""
==============================================================
  SISTEMA DE AUDITORIA - BOT GLOSAS V5
  Desarrollado por: DESARROLLO E INNOVACION SALUD NET
  Propietario     : Salud-Net
  Version         : v1.0
  Licencia        : Uso interno autorizado
==============================================================
"""

import imaplib
import email
import zipfile
import io
import email.utils
import os
import re
import unicodedata
import pandas as pd
from email.header import decode_header
from datetime import datetime
from pdfminer.high_level import extract_text

# ---------------------------------------------------------------
# CLAIMONLINE — DESCARGA AUTOMATICA DE ZIPs AXA PROTEGIDOS
# Requiere: pip install playwright && playwright install chromium
# ---------------------------------------------------------------

def _cl_extraer_links(msg) -> list:
    """
    Busca URLs de claimonline.com.co en el cuerpo del correo.
    Estrategia doble:
    1. Regex sobre texto plano (URLs escritas directamente)
    2. Parseo de href en HTML (links con texto visible distinto a la URL)
       Ej: <a href="https://axa.claimonline.com.co/...">Objecion_123.zip</a>
    """
    patron = re.compile(
        r"https?://[^\s\"'<>]*claimonline\.com\.co[^\s\"'<>]*",
        re.IGNORECASE,
    )
    # Regex para extraer href de etiquetas <a> directamente sin BeautifulSoup
    href_re = re.compile(
        r'href=["\']?(https?://[^\s"\'<>]*claimonline\.com\.co[^\s"\'<>]*)',
        re.IGNORECASE,
    )
    urls, seen = [], set()

    for part in msg.walk():
        ct = part.get_content_type()
        if ct not in ("text/plain", "text/html"):
            continue
        try:
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            texto = payload.decode(
                part.get_content_charset() or "utf-8", errors="replace"
            )
            # Capa 1: URL visible en texto plano
            for url in patron.findall(texto):
                url = url.rstrip(".,;)")
                if url not in seen:
                    seen.add(url)
                    urls.append(url)
            # Capa 2: href en HTML (link con texto visible diferente)
            if ct == "text/html":
                for url in href_re.findall(texto):
                    url = url.rstrip(".,;)\"'")
                    if url not in seen:
                        seen.add(url)
                        urls.append(url)
        except Exception:
            continue
    return urls


def es_correo_claimonline(msg) -> bool:
    """Retorna True si el correo es una notificacion de Claimonline/AXA."""
    remitente = (msg.get("From", "") or "").lower()
    if "claimonline" in remitente:
        return True
    return len(_cl_extraer_links(msg)) > 0


def _cl_descargar_zip(url: str, carpeta_destino: str) -> str | None:
    """
    Descarga el ZIP de Claimonline con dos estrategias:

    Estrategia 1 — Playwright (Chromium headless):
      Abre la URL, espera el boton 'Descargar Archivo', hace clic y
      captura el archivo descargado.

    Estrategia 2 — requests (fallback):
      Intenta descargar la URL directamente como archivo binario.
      Funciona cuando el servidor sirve el ZIP sin necesidad de clic.

    Retorna la ruta local del ZIP descargado, o None si ambas fallan.
    """
    os.makedirs(carpeta_destino, exist_ok=True)

    # ── Estrategia 1: Playwright ────────────────────────────────────────────
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
        print(f"       [Claimonline] Iniciando Playwright...")

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
            page    = context.new_page()

            # Log de errores de consola del navegador para diagnóstico
            page.on("console", lambda m: print(f"       [Browser] {m.type}: {m.text}") if m.type == "error" else None)

            print(f"       [Claimonline] Navegando a: {url[:80]}...")
            try:
                # domcontentloaded es menos estricto que networkidle — carga mas rapido
                page.goto(url, timeout=30_000, wait_until="domcontentloaded")
                # Espera adicional para que cargue el boton
                page.wait_for_timeout(3000)
            except PWTimeout:
                print("       [Claimonline] Timeout al cargar la pagina — intentando igual...")

            # Capturar HTML de la pagina para diagnóstico si no encuentra el boton
            html_pagina = page.content()
            print(f"       [Claimonline] Pagina cargada ({len(html_pagina)} chars)")

            boton = None
            selectores = [
                "button:has-text('Descargar Archivo')",
                "button:has-text('Descargar')",
                "a:has-text('Descargar Archivo')",
                "a:has-text('Descargar')",
                "input[type='button'][value*='Descargar']",
                "input[type='submit'][value*='Descargar']",
                "[onclick*='descargar']",
                "[onclick*='download']",
            ]
            for selector in selectores:
                try:
                    loc = page.locator(selector).first
                    if loc.count() > 0:
                        loc.wait_for(state="visible", timeout=5_000)
                        boton = loc
                        print(f"       [Claimonline] Boton encontrado: {selector}")
                        break
                except Exception:
                    continue

            if boton is None:
                # Mostrar texto visible de la pagina para diagnóstico
                texto_visible = page.inner_text("body")[:300].replace("\n", " ")
                print(f"       [Claimonline] Boton NO encontrado. Contenido pagina: {texto_visible}")
                browser.close()
            else:
                try:
                    print("       [Claimonline] Haciendo clic en Descargar Archivo...")
                    with page.expect_download(timeout=60_000) as dl_info:
                        boton.click()
                    download   = dl_info.value
                    nombre_zip = download.suggested_filename or "claimonline.zip"
                    zip_path   = os.path.join(carpeta_destino, nombre_zip)
                    download.save_as(zip_path)
                    browser.close()
                    if os.path.exists(zip_path) and os.path.getsize(zip_path) > 0:
                        print(f"       [Claimonline] ZIP descargado OK: {nombre_zip} ({os.path.getsize(zip_path)} bytes)")
                        return zip_path
                    else:
                        print("       [Claimonline] ZIP descargado pero vacio o no existe.")
                        browser.close()
                except Exception as e:
                    print(f"       [Claimonline] Error al hacer clic/descargar: {e}")
                    try: browser.close()
                    except Exception: pass

    except ImportError:
        print("       [Claimonline] Playwright no instalado.")
        print("       Ejecute: pip install playwright && playwright install chromium")
    except Exception as e:
        print(f"       [Claimonline] Error Playwright inesperado: {e}")
        import traceback
        traceback.print_exc()

    # ── Estrategia 2: requests directo ──────────────────────────────────────
    print("       [Claimonline] Intentando descarga directa con requests...")
    try:
        import urllib.request
        nombre_zip = "claimonline_" + re.sub(r"[^\w]", "_", url[-30:]) + ".zip"
        zip_path   = os.path.join(carpeta_destino, nombre_zip)
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "*/*",
        }
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            contenido = resp.read()
        if len(contenido) > 100:
            with open(zip_path, "wb") as f:
                f.write(contenido)
            print(f"       [Claimonline] Descarga directa OK: {len(contenido)} bytes")
            return zip_path
        else:
            print(f"       [Claimonline] Respuesta directa muy pequeña ({len(contenido)} bytes) — no es el ZIP")
    except Exception as e:
        print(f"       [Claimonline] Descarga directa fallida: {e}")

    print("       [Claimonline] Ambas estrategias fallaron — ZIP no descargado.")
    return None


def _cl_descomprimir(zip_path: str, password: str) -> list:
    """
    Descomprime un ZIP protegido y retorna lista de (filename, pdf_bytes).
    """
    resultado = []
    if not zip_path or not os.path.exists(zip_path):
        return resultado
    if not password:
        print("       [Claimonline] Sin password configurado para esta IPS.")
        return resultado
    try:
        pwd_bytes = str(password).encode("utf-8")
        with zipfile.ZipFile(zip_path, "r") as zf:
            print(f"       [Claimonline] Archivos en ZIP: {len(zf.namelist())}")
            for nombre in zf.namelist():
                if not nombre.lower().endswith(".pdf"):
                    continue
                try:
                    pdf_bytes = zf.read(nombre, pwd=pwd_bytes)
                    fn = re.sub(r'[<>:"/\\|?*\r\n\t]', "_", os.path.basename(nombre))
                    fn = re.sub(r" +", " ", fn).strip()
                    resultado.append((fn, pdf_bytes))
                    print(f"       [Claimonline] PDF extraido: {fn}")
                except RuntimeError:
                    print(f"       [Claimonline] Password incorrecto para {nombre}")
                except Exception as e:
                    print(f"       [Claimonline] Error extrayendo {nombre}: {e}")
    except zipfile.BadZipFile:
        print(f"       [Claimonline] Archivo ZIP invalido.")
    except Exception as e:
        print(f"       [Claimonline] Error al abrir ZIP: {e}")
    return resultado


def _cl_buscar_password(asunto: str, remitente: str,
                         zip_path: str, ips_dict: dict) -> str | None:
    """
    Busca la password del ZIP identificando la IPS por NIT en el asunto,
    nombre del ZIP o remitente. Si no encuentra NIT, intenta por nombre.
    """
    fuentes = [
        asunto or "",
        os.path.basename(zip_path) if zip_path else "",
        remitente or "",
    ]
    for fuente in fuentes:
        # Patrón específico para nombre de ZIP de Claimonline: Objecion_NIT-fecha.zip
        for nit in re.findall(r"[Oo]bjecion_(\d{7,11})-\d{8}", fuente):
            nit_base = nit[:9]
            if nit_base in ips_dict:
                pwd = ips_dict[nit_base].get("password", "")
                if pwd:
                    print(f"       [Claimonline] IPS por NIT {nit_base} (ZIP) — password OK")
                    return str(pwd).strip()
        # Búsqueda general de NIT de 9-11 dígitos
        for nit in re.findall(r"\b\d{9,11}\b", fuente):
            nit_base = nit[:9]
            if nit_base in ips_dict:
                pwd = ips_dict[nit_base].get("password", "")
                if pwd:
                    print(f"       [Claimonline] IPS por NIT {nit_base} — password OK")
                    return str(pwd).strip()

    # Fallback: buscar por nombre/equivalente en el asunto
    if asunto:
        asunto_upper = asunto.upper()
        mejor = (None, 0)
        for nit, datos in ips_dict.items():
            for equiv in datos.get("equivalentes_norm", []):
                if equiv and len(equiv) > 5 and equiv in asunto_upper:
                    if len(equiv) > mejor[1]:
                        mejor = (nit, len(equiv))
        if mejor[0]:
            pwd = ips_dict[mejor[0]].get("password", "")
            if pwd:
                print(f"       [Claimonline] IPS por nombre — password OK")
                return str(pwd).strip()
    return None


def _cl_extraer_nombre_ips_cuerpo(msg) -> str:
    """
    Lee el cuerpo del correo Claimonline y extrae el nombre de la IPS.
    El cuerpo siempre dice: 'Apreciados Señores NOMBRE IPS'
    Retorna el nombre en mayúsculas o string vacío si no lo encuentra.
    """
    patron = re.compile(
        r"Apreciados\s+Se[ñn]ores?\s+(.+?)(?:\s*\n|\s*<|\s*$)",
        re.IGNORECASE,
    )
    for part in msg.walk():
        ct = part.get_content_type()
        if ct not in ("text/plain", "text/html"):
            continue
        try:
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            texto = payload.decode(
                part.get_content_charset() or "utf-8", errors="replace"
            )
            if ct == "text/html":
                texto = re.sub(r"<[^>]+>", " ", texto)
                texto = re.sub(r"&nbsp;", " ", texto)
                texto = re.sub(r"&[a-z]+;", " ", texto)
            m = patron.search(texto)
            if m:
                nombre = re.sub(r"\s+", " ", m.group(1)).strip()
                if len(nombre) > 3:
                    print(f"       [Claimonline] IPS en cuerpo: {nombre}")
                    return nombre.upper()
        except Exception:
            continue
    return ""


def _cl_buscar_ips_por_nombre(nombre_cuerpo: str, ips_dict: dict):
    """
    Busca la IPS comparando el nombre del cuerpo del correo contra el Excel.
    Prioridad:
    1. Nombre completo normalizado (IPS NOMBRE COMPLETO)
    2. Equivalentes y nombres de sede
    En cada capa busca subcadena exacta — el más largo gana.
    Retorna (nit, datos) del mejor match o (None, None).
    """
    if not nombre_cuerpo:
        return None, None

    import unicodedata
    def _norm(t):
        t = str(t).upper().strip()
        t = unicodedata.normalize("NFD", t)
        t = "".join(c for c in t if unicodedata.category(c) != "Mn")
        # Quitar sufijos legales para comparación
        for sfx in [" SAS", " S.A.S", " S.A", " LTDA", " EU", " IPS"]:
            if t.endswith(sfx):
                t = t[:-len(sfx)].strip()
        return t

    nombre_n = _norm(nombre_cuerpo)
    mejor    = (None, None, 0)

    # ── Capa 1: nombre completo ──────────────────────────────────────────────
    for nit, datos in ips_dict.items():
        cand_n = _norm(datos.get("nombre", ""))
        if not cand_n or len(cand_n) < 4:
            continue
        if cand_n in nombre_n or nombre_n in cand_n:
            if len(cand_n) > mejor[2]:
                mejor = (nit, datos, len(cand_n))

    if mejor[0]:
        print(f"       [Claimonline] IPS por nombre completo: {mejor[1]['nombre']} (NIT {mejor[0]})")
        return mejor[0], mejor[1]

    # ── Capa 2: equivalentes y sedes ─────────────────────────────────────────
    for nit, datos in ips_dict.items():
        for equiv in datos.get("equivalentes_norm", []):
            if not equiv or len(equiv) < 4:
                continue
            equiv_n = _norm(equiv)
            if equiv_n in nombre_n or nombre_n in equiv_n:
                if len(equiv_n) > mejor[2]:
                    mejor = (nit, datos, len(equiv_n))

    if mejor[0]:
        print(f"       [Claimonline] IPS por equivalente: {mejor[1]['nombre']} (NIT {mejor[0]})")
        return mejor[0], mejor[1]

    print(f"       [Claimonline] IPS no identificada para: {nombre_cuerpo}")
    return None, None


def procesar_correo_claimonline(msg, ips_dict: dict, carpeta_temp: str,
                                 asunto: str = "", fecha=None,
                                 remitente: str = "") -> list:
    """
    Proceso completo para correos Claimonline. TODO lo que venga del ZIP
    se guarda en la ruta de la IPS bajo ZIP_COLPATRIA — el desapilador
    nunca toca esta carpeta.

    Flujo:
    1. Lee cuerpo del correo → identifica IPS → ciudad, NIT, nombre
    2. Toma fecha del correo → año, mes, día (reglas de salida estándar)
    3. Descarga el ZIP con Playwright
    4. Intenta descomprimir con password del Excel
       - Si abre: guarda cada PDF en Ciudad/NIT-IPS/Año/Mes/Día/AXA COLPATRIA.../ZIP_COLPATRIA/
       - Si no abre: guarda el ZIP en la misma ruta ZIP_COLPATRIA/
    5. Marca correo como leido en ambos casos
       Solo NO marca como leido si no pudo descargar el ZIP.
    """
    import shutil
    resultado = []
    links = _cl_extraer_links(msg)
    if not links:
        print("       [Claimonline] No se encontraron links en el correo.")
        return resultado

    print(f"       [Claimonline] Links encontrados: {len(links)}")
    fecha_uso = fecha or datetime.now()

    # Paso 1: identificar IPS desde el cuerpo del correo
    nombre_cuerpo      = _cl_extraer_nombre_ips_cuerpo(msg)
    nit_ips, datos_ips = _cl_buscar_ips_por_nombre(nombre_cuerpo, ips_dict)

    # Determinar carpeta de destino ZIP_COLPATRIA usando reglas estándar de salida
    if nit_ips and datos_ips:
        carpeta_destino = crear_ruta(
            datos_ips.get("ciudad", "SIN CIUDAD"),
            nit_ips,
            datos_ips.get("nombre", "IPS"),
            "AXA COLPATRIA SEGUROS S.A.",
            fecha_uso,
        )
        carpeta_destino = os.path.join(carpeta_destino, "ZIP_COLPATRIA")
        print(f"       [Claimonline] Ruta destino: .../{datos_ips.get('ciudad','')}/.../{fecha_uso.strftime('%d')}/AXA COLPATRIA.../ZIP_COLPATRIA/")
    else:
        carpeta_destino = os.path.join(RUTA_BASE, "ZIP_COLPATRIA_SIN_IPS")
        print(f"       [Claimonline] IPS no identificada — usando ZIP_COLPATRIA_SIN_IPS")

    os.makedirs(carpeta_destino, exist_ok=True)

    for url in links:
        print(f"       [Claimonline] Procesando: {url[:80]}...")

        # Paso 2: descargar ZIP
        zip_path = _cl_descargar_zip(url, carpeta_temp)
        if not zip_path:
            print("       [Claimonline] No se pudo descargar el ZIP — correo NO marcado como leido.")
            continue

        nombre_zip = os.path.basename(zip_path)

        # Paso 3: buscar password
        password = None
        if nit_ips and datos_ips:
            pwd = datos_ips.get("password", "")
            if pwd:
                print(f"       [Claimonline] Password por IPS del cuerpo — OK")
                password = str(pwd).strip()
        if not password:
            password = _cl_buscar_password(asunto, remitente, zip_path, ips_dict)

        # Paso 4: intentar descomprimir
        pdfs = []
        if password:
            pdfs = _cl_descomprimir(zip_path, password)
        else:
            print("       [Claimonline] Sin password — ZIP se guardará sin abrir.")

        def _destino_unico(carpeta, nombre):
            destino = os.path.join(carpeta, nombre)
            contador = 1
            while os.path.exists(destino):
                base, ext = os.path.splitext(nombre)
                destino = os.path.join(carpeta, f"{base}_{contador}{ext}")
                contador += 1
            return destino

        if pdfs:
            # Abrió correctamente: guardar cada PDF en ZIP_COLPATRIA
            try: os.remove(zip_path)
            except Exception: pass
            for fn, pdf_bytes in pdfs:
                destino_pdf = _destino_unico(carpeta_destino, fn)
                try:
                    with open(destino_pdf, "wb") as f:
                        f.write(pdf_bytes)
                    print(f"       [Claimonline] PDF guardado: {os.path.basename(destino_pdf)}")
                except Exception as e:
                    print(f"       [Claimonline] Error guardando PDF {fn}: {e}")
        else:
            # No abrió: guardar el ZIP en ZIP_COLPATRIA
            destino_zip = _destino_unico(carpeta_destino, nombre_zip)
            try:
                shutil.move(zip_path, destino_zip)
                print(f"       [Claimonline] ZIP guardado: {os.path.basename(destino_zip)}")
            except Exception as e:
                print(f"       [Claimonline] Error guardando ZIP: {e}")
                try: os.remove(zip_path)
                except Exception: pass

        # Marcador especial: correo se marca leido (ZIP descargado y ubicado)
        resultado.append((
            f"__ZIP_IPS__{nombre_zip}",
            b"",
            asunto,
            fecha_uso,
            remitente,
        ))

    return resultado

# ---------------------------------------------------------------
# IDENTIDAD
# ---------------------------------------------------------------
def datos_identidad():
    return {
        "sistema"         : "BOT GLOSAS - DESCARGA Y ORGANIZACION",
        "propietario"     : "Salud-Net",
        "desarrollado_por": "DESARROLLO E INNOVACION SALUD NET",
        "version"         : "v1.0",
        "licencia"        : "Uso interno autorizado",
    }

# ---------------------------------------------------------------
# RUTAS
# ---------------------------------------------------------------
DIRECTORIO_BASE = os.path.dirname(os.path.abspath(__file__))
EXCEL_CONFIG    = os.path.join(DIRECTORIO_BASE, "NOMBRES Y NIT EQUIVALENTES.xlsx")
EXCEL_ASUNTOS   = os.path.join(DIRECTORIO_BASE, "asuntos_correos_glosas.xlsx")
RUTA_BASE       = os.path.join(DIRECTORIO_BASE, "DESCARGA DESDE CORREOS")
REPORTE_EXCEL   = os.path.join(DIRECTORIO_BASE, "REPORTE_GLOSAS.xlsx")
CARPETA_NO_ID   = os.path.join(RUTA_BASE, "_NO_IDENTIFICADOS")
MAX_NOMBRE_PDF  = 80

MESES = {
    "01": "01 - ENERO",    "02": "02 - FEBRERO",  "03": "03 - MARZO",
    "04": "04 - ABRIL",    "05": "05 - MAYO",     "06": "06 - JUNIO",
    "07": "07 - JULIO",    "08": "08 - AGOSTO",   "09": "09 - SEPTIEMBRE",
    "10": "10 - OCTUBRE",  "11": "11 - NOVIEMBRE","12": "12 - DICIEMBRE",
}

SUFIJOS = [" S.A.S", " SAS", " S.A", " SA", " LTDA", " E.U", " EU", " IPS"]

# ---------------------------------------------------------------
# NORMALIZAR
# ---------------------------------------------------------------
def normalizar(texto):
    if not texto:
        return ""
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto

def normalizar_nombre(texto):
    norm = normalizar(texto)
    for sufijo in SUFIJOS:
        if norm.endswith(normalizar(sufijo)):
            norm = norm[:-len(normalizar(sufijo))].strip()
    return norm

def similitud(texto_a, texto_b):
    palabras_a = {p for p in texto_a.split() if len(p) > 2}
    palabras_b = {p for p in texto_b.split() if len(p) > 2}
    if not palabras_a or not palabras_b:
        return 0.0
    return len(palabras_a & palabras_b) / len(palabras_a | palabras_b)

UMBRAL_SIMILITUD        = 0.70  # Matching de nombres IPS
UMBRAL_SIMILITUD_ASUNTO = 0.45  # Matching flexible de asuntos de correo

# ---------------------------------------------------------------
# EXTRAER PATRON ESTABLE DE UN ASUNTO
# Elimina numeros, fechas, horas, emails y codigos variables
# para quedarse con palabras clave que no cambian entre correos
# ---------------------------------------------------------------
def extraer_patron(asunto_norm):
    """
    Extrae palabras clave estables de un asunto normalizado.
    Elimina numeros, fechas, emails, simbolos y codigos variables.
    El texto de entrada ya viene sin tildes (normalizar lo quita).
    """
    txt = re.sub(r"[()\[\]{}<>]", " ", asunto_norm)
    txt = re.sub(r"\S+@\S+", " ", txt)
    txt = re.sub(r"[^\w\s]", " ", txt)
    txt = re.sub(r"([A-Z])([0-9])", r"\1 \2", txt)
    txt = re.sub(r"([0-9])([A-Z])", r"\1 \2", txt)
    txt = re.sub(r"\b\d+\b", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    # Excluir palabras cortas y palabras de relleno
    excluir = {"PARA", "CASO", "DESDE", "COMO", "ESTE", "ESTA", "MAIL",
               "EMAIL", "FROM", "REGISTERED", "CERTIFICADO"}
    palabras = [p for p in txt.split() if len(p) > 3 and p not in excluir]
    return " ".join(palabras[:5])

# ---------------------------------------------------------------
# CARGAR CONFIGURACION
# ---------------------------------------------------------------
def cargar_configuracion():
    print("  Cargando matriz de IPS...")
    wb_ips  = pd.read_excel(EXCEL_CONFIG, sheet_name="ESTRUCTURA")
    wb_cor  = pd.read_excel(EXCEL_CONFIG, sheet_name="CORREOS")
    wb_aseg = pd.read_excel(EXCEL_CONFIG, sheet_name="ASEGURADORAS")

    cuentas = []
    for _, row in wb_cor.iterrows():
        correo = str(row.iloc[0]).strip()
        clave  = str(row.iloc[1]).strip()
        if correo and clave and correo.lower() != "correo" and "@" in correo:
            cuentas.append({"correo": correo, "password": clave})

    wb_ips.columns = wb_ips.columns.str.strip().str.upper()
    ips_dict = {}
    col_equiv        = [c for c in wb_ips.columns if c.startswith("EQUIVALENTE")]
    col_dirs         = [c for c in wb_ips.columns if c.startswith("DIR")]
    col_sedes_nombre = [c for c in wb_ips.columns if c.startswith("NOMBRE_SEDE")]

    for _, row in wb_ips.iterrows():
        nit    = str(row["NIT"]).strip().split(".")[0][:9]
        nombre = str(row["IPS NOMBRE COMPLETO"]).strip()
        ciudad = ""
        if "CIUDAD" in wb_ips.columns:
            val = row.get("CIUDAD")
            if pd.notna(val) and str(val).strip():
                ciudad = str(val).strip().upper()
        if not ciudad:
            ciudad = "SIN CIUDAD"

        equivalentes = [nombre]
        for c in col_equiv + col_sedes_nombre:
            val = row.get(c)
            if pd.notna(val) and str(val).strip():
                equivalentes.append(str(val).strip())
        dirs = []
        for c in col_dirs:
            val = row.get(c)
            if pd.notna(val) and str(val).strip():
                dirs.append(normalizar(str(val).strip()))

        # Leer password del ZIP de Claimonline si existe la columna
        password_zip = ""
        if "PASSWORD" in wb_ips.columns:
            val_pwd = row.get("PASSWORD")
            if pd.notna(val_pwd) and str(val_pwd).strip() and str(val_pwd).strip().lower() != "nan":
                password_zip = str(val_pwd).strip()

        ips_dict[nit] = {
            "nombre"           : nombre,
            "ciudad"           : ciudad,
            "equivalentes_norm": [normalizar_nombre(e) for e in equivalentes],
            "nombre_norm"      : normalizar_nombre(nombre),
            "dirs"             : dirs,
            "password"         : password_zip,
        }

    print("  Cargando aseguradoras...")
    wb_aseg.columns = wb_aseg.columns.str.strip().str.upper()
    aseg_dict = {}
    col_aseg_equiv = [c for c in wb_aseg.columns if c.startswith("EQUIVALENTE")]
    for _, row in wb_aseg.iterrows():
        nombre_carpeta = str(row["ASEGURADORA_CARPETA"]).strip()
        for c in ["ASEGURADORA_CARPETA"] + col_aseg_equiv:
            val = row.get(c)
            if pd.notna(val) and str(val).strip():
                aseg_dict[normalizar(str(val).strip())] = nombre_carpeta

    # ── Patrones del Excel de asuntos ──────────────────────────
    # Se generan DOS tipos de patron por cada fila:
    # 1. Patron completo normalizado (para coincidencia exacta)
    # 2. Patron de palabras clave (para coincidencia flexible sin numeros)
    print("  Cargando asuntos validos desde Excel...")
    patrones_asuntos = []   # lista de strings normalizados
    try:
        wb_asuntos = pd.read_excel(EXCEL_ASUNTOS)
        for v in wb_asuntos.iloc[:, 0].dropna():
            val = str(v).strip()
            if not val or val.upper() == "ASUNTO":
                continue
            val_norm = normalizar(val)
            # Patron 1: normalizado completo
            if val_norm and val_norm not in patrones_asuntos:
                patrones_asuntos.append(val_norm)
            # Patron 2: palabras clave sin numeros/fechas
            patron_clave = extraer_patron(val_norm)
            if patron_clave and len(patron_clave) > 5 and patron_clave not in patrones_asuntos:
                patrones_asuntos.append(patron_clave)
        print(f"  Patrones generados: {len(patrones_asuntos)}")
    except Exception as e:
        print(f"  [AVISO] Excel de asuntos no cargado: {e}")

    # aseg_passwords: requerido por app.py como quinto valor de retorno.
    # Las passwords por IPS ya estan en ips_dict["password"].
    # Este diccionario queda disponible para usos futuros por aseguradora.
    aseg_passwords = {}
    return cuentas, ips_dict, aseg_dict, patrones_asuntos, aseg_passwords

# ---------------------------------------------------------------
# VALIDAR ASUNTO
# El Excel de asuntos es la UNICA fuente de verdad.
# Un correo sin asunto + adjuntos .eml tambien es valido.
#
# Estrategia de matching por capas:
# 1. Subcadena exacta: patron completo dentro del asunto
# 2. Palabras clave: todas las palabras del patron (>3 chars)
#    presentes en el asunto, sin importar el orden ni palabras
#    intermedias como "DE", "Y", "EN"
# ---------------------------------------------------------------
def limpiar_para_matching(texto_norm):
    """
    Limpia el texto para matching:
    - Reemplaza guiones bajos por espacios: Objecion_123 -> Objecion 123
    - Separa palabras pegadas a simbolos: m.)SEGUROS -> m  SEGUROS
    - Separa transiciones letra->numero y numero->letra: CMVIQ034000 -> CMVIQ 034000
    """
    txt = re.sub(r"_", " ", texto_norm)          # guion bajo -> espacio
    txt = re.sub(r"[()\[\]{}<>]", " ", txt)
    txt = re.sub(r"[^\w\s]", " ", txt)
    # Separar donde letras se pegan a numeros y viceversa
    txt = re.sub(r"([A-Z])([0-9])", r"\1 \2", txt)
    txt = re.sub(r"([0-9])([A-Z])", r"\1 \2", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

def _quitar_prefijos_reenvio(asunto_norm):
    """
    Elimina prefijos de reenvío/respuesta al inicio del asunto:
    'FWD:', 'RE:', 'RV:', 'RES:', 'FW:' con variaciones de espacios.
    Retorna el asunto sin esos prefijos para comparar contra patrones.
    """
    return re.sub(r"^\s*(FWD?|RE|RV|RES)\s*:\s*", "", asunto_norm).strip()


def asunto_es_valido(asunto_norm, patrones_asuntos):
    if not asunto_norm.strip():
        return False  # sin asunto se maneja por es_contenedor_eml

    # Evaluar tanto el asunto completo como sin prefijo Fwd:/Re:
    asunto_sin_prefijo = _quitar_prefijos_reenvio(asunto_norm)
    candidatos_asunto  = list({asunto_norm, asunto_sin_prefijo})  # únicos

    for asunto_a_evaluar in candidatos_asunto:
        asunto_limpio   = limpiar_para_matching(asunto_a_evaluar)
        palabras_asunto = set(asunto_limpio.split())
        mejor_score     = 0.0

        for patron in patrones_asuntos:
            if not patron:
                continue
            patron_limpio = limpiar_para_matching(patron)

            # Capa 1: subcadena exacta
            if patron_limpio in asunto_limpio:
                return True

            palabras_patron = [p for p in patron_limpio.split() if len(p) > 3]

            # Capa 1b: patron sin palabras largas (solo palabras <=3 chars o vacias)
            # Ej: asunto "OBJECION ." -> patron_limpio="OBJECION", palabras_patron=["OBJECION"]
            # Si palabras_patron vacio, comparar cada palabra del patron contra el asunto
            if not palabras_patron:
                patron_todas = [p for p in patron_limpio.split() if p]
                if patron_todas and all(p in palabras_asunto for p in patron_todas):
                    return True
                continue

            # Capa 2: todas las palabras clave presentes
            if all(p in palabras_asunto for p in palabras_patron):
                return True

            # Capa 3: fraccion de palabras del patron presentes en el asunto
            score = sum(1 for p in palabras_patron if p in palabras_asunto) / len(palabras_patron)
            if score > mejor_score:
                mejor_score = score

            # Capa 4: similitud de secuencia de caracteres (typos, variaciones leves)
            import difflib
            ratio = difflib.SequenceMatcher(None, patron_limpio, asunto_limpio).ratio()
            if ratio > mejor_score:
                mejor_score = ratio

        # Aceptar si supera el umbral de similitud
        if mejor_score >= UMBRAL_SIMILITUD_ASUNTO:
            return True

    return False

# ---------------------------------------------------------------
# DETECTAR IPS
# ---------------------------------------------------------------
def _normalizar_dir_desc(texto: str) -> str:
    """
    Normaliza una dirección para comparación robusta:
    quita tildes, colapsa espacios alrededor de guiones y puntos.
    """
    t = normalizar(texto)
    t = re.sub(r"\s*-\s*", "-", t)
    t = re.sub(r"\s*\.\s*", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def detectar_ips(texto_pdf_norm, ips_dict, texto_contexto_norm=""):
    """
    Detecta la IPS con el siguiente orden de prioridad:

    1. NIT delimitado en el texto del PDF
    2. Nombre completo normalizado en el texto del PDF
    3. Equivalentes y nombres de sede (el más largo gana) en texto PDF
    4. Similitud de nombre (frases del texto vs nombre_norm)
    5. Direcciones normalizadas en texto PDF
    6. Repetir pasos 2-4 sobre texto_contexto_norm (asunto + remitente)
       solo si no hubo match en el PDF.
    """
    # ── Paso 1: NIT delimitado en texto del PDF ──────────────────────────────
    for nit_encontrado in re.findall(r"\b\d{9,11}\b", texto_pdf_norm):
        nit_base = nit_encontrado[:9]
        if nit_base in ips_dict:
            return nit_base, ips_dict[nit_base]["nombre"], ips_dict[nit_base]["ciudad"]

    # ── Paso 2: nombre completo en texto del PDF ─────────────────────────────
    best_nombre = (None, None, None, 0)
    for nit, datos in ips_dict.items():
        nombre_n = datos["nombre_norm"]
        if nombre_n and nombre_n in texto_pdf_norm and len(nombre_n) > best_nombre[3]:
            best_nombre = (nit, datos["nombre"], datos["ciudad"], len(nombre_n))
    if best_nombre[0]:
        return best_nombre[0], best_nombre[1], best_nombre[2]

    # ── Paso 3: equivalentes en texto del PDF (el más largo gana) ────────────
    candidatos = []
    for nit, datos in ips_dict.items():
        for equiv in datos["equivalentes_norm"]:
            if equiv and len(equiv) > 4 and equiv in texto_pdf_norm:
                candidatos.append((len(equiv), nit, datos["nombre"], datos["ciudad"]))
    if candidatos:
        candidatos.sort(reverse=True)
        _, nit, nombre, ciudad = candidatos[0]
        return nit, nombre, ciudad

    # ── Paso 4: similitud de nombre en texto del PDF ─────────────────────────
    mejores = []
    palabras_texto = texto_pdf_norm.split()
    for inicio in range(len(palabras_texto)):
        for largo in range(3, 9):
            frase = " ".join(palabras_texto[inicio:inicio + largo])
            for nit, datos in ips_dict.items():
                sim = similitud(frase, datos["nombre_norm"])
                if sim >= UMBRAL_SIMILITUD:
                    mejores.append((sim, nit, datos["nombre"], datos["ciudad"]))
    if mejores:
        mejores.sort(reverse=True)
        _, nit, nombre, ciudad = mejores[0]
        return nit, nombre, ciudad

    # ── Paso 5: dirección normalizada en texto del PDF ───────────────────────
    texto_dir = _normalizar_dir_desc(texto_pdf_norm)
    for nit, datos in ips_dict.items():
        for dir_raw in datos.get("dirs", []):
            dir_n = _normalizar_dir_desc(dir_raw)
            if dir_n and len(dir_n) > 8 and dir_n in texto_dir:
                return nit, datos["nombre"], datos["ciudad"]

    # ── Paso 6: fallback en contexto (asunto + remitente) ────────────────────
    if texto_contexto_norm:
        best_ctx = (None, None, None, 0)
        for nit, datos in ips_dict.items():
            nombre_n = datos["nombre_norm"]
            if nombre_n and nombre_n in texto_contexto_norm and len(nombre_n) > best_ctx[3]:
                best_ctx = (nit, datos["nombre"], datos["ciudad"], len(nombre_n))
        if best_ctx[0]:
            return best_ctx[0], best_ctx[1], best_ctx[2]

        candidatos_ctx = []
        for nit, datos in ips_dict.items():
            for equiv in datos["equivalentes_norm"]:
                if equiv and len(equiv) > 4 and equiv in texto_contexto_norm:
                    candidatos_ctx.append((len(equiv), nit, datos["nombre"], datos["ciudad"]))
        if candidatos_ctx:
            candidatos_ctx.sort(reverse=True)
            _, nit, nombre, ciudad = candidatos_ctx[0]
            return nit, nombre, ciudad

    return None, None, None

# ---------------------------------------------------------------
# DETECTAR ASEGURADORA
# ---------------------------------------------------------------
def _palabras_clave(texto_norm):
    """Retorna el conjunto de palabras significativas (mas de 3 letras)."""
    return {p for p in texto_norm.split() if len(p) > 3}

def _equiv_coincide(equiv_norm, palabras_fuente):
    """
    Verifica si un equivalente coincide con el texto de una fuente.
    Reglas:
    1. El equivalente debe tener al menos 2 palabras significativas
       para evitar falsos positivos con palabras genericas como
       BOLIVAR (departamento) o ESTADO (sustantivo comun).
    2. Todas las palabras significativas del equivalente deben estar
       presentes como palabras individuales en el texto.
    """
    palabras_equiv = [p for p in equiv_norm.split() if len(p) > 3]
    if len(palabras_equiv) < 2:
        return False
    return all(p in palabras_fuente for p in palabras_equiv)

def detectar_aseguradora(texto_norm, asunto_norm, remitente_norm, aseg_dict,
                          filename_norm=""):
    """
    Busca la aseguradora en cuatro fuentes en orden de prioridad:
    1. Texto extraido del PDF
    2. Asunto del correo
    3. Remitente del correo
    4. Nombre del archivo PDF

    Matching por palabras clave individuales — no subcadena exacta.
    Requiere minimo 2 palabras significativas por equivalente para
    evitar falsos positivos con palabras genericas (BOLIVAR, ESTADO).
    El equivalente mas largo que coincida gana.
    """
    for fuente in [texto_norm, asunto_norm, remitente_norm, filename_norm]:
        if not fuente:
            continue
        palabras_fuente = _palabras_clave(fuente)
        candidatos = []
        for equiv_norm, nombre_carpeta in aseg_dict.items():
            if equiv_norm and _equiv_coincide(equiv_norm, palabras_fuente):
                candidatos.append((len(equiv_norm.split()), nombre_carpeta))
        if candidatos:
            candidatos.sort(reverse=True)
            return candidatos[0][1]
    return "_ASEGURADORA_NO_IDENTIFICADA"

# ---------------------------------------------------------------
# CREAR RUTA
# ---------------------------------------------------------------
def crear_ruta(ciudad, nit, ips_nombre, aseguradora, fecha_correo):
    def limpiar(txt):
        return re.sub(r'[<>:"/\\|?*]', "", str(txt)).strip()
    mes_texto = MESES.get(fecha_correo.strftime("%m"), fecha_correo.strftime("%m"))
    ruta = os.path.join(
        RUTA_BASE,
        limpiar(ciudad),
        f"{nit} - {limpiar(ips_nombre)}",
        str(fecha_correo.year),
        mes_texto,
        fecha_correo.strftime("%d"),
        limpiar(aseguradora),
    )
    os.makedirs(ruta, exist_ok=True)
    return ruta

# ---------------------------------------------------------------
# DECODIFICAR ASUNTO
# ---------------------------------------------------------------
def decodificar_asunto(msg):
    partes = decode_header(msg.get("Subject", ""))
    resultado = []
    for parte, charset in partes:
        if isinstance(parte, bytes):
            try:
                resultado.append(parte.decode(charset or "utf-8", errors="replace"))
            except Exception:
                resultado.append(parte.decode("latin-1", errors="replace"))
        else:
            resultado.append(str(parte))
    return "".join(resultado)

# ---------------------------------------------------------------
# DECODIFICAR NOMBRE DE ARCHIVO
# ---------------------------------------------------------------
def decodificar_filename(filename):
    partes = decode_header(filename)
    resultado = ""
    for fp, fcharset in partes:
        if isinstance(fp, bytes):
            resultado += fp.decode(fcharset or "utf-8", errors="replace")
        else:
            resultado += str(fp)
    # Eliminar saltos de linea, tabulaciones y caracteres de control
    resultado = re.sub(r'[\r\n\t]', " ", resultado)
    # Eliminar caracteres invalidos en rutas Windows
    resultado = re.sub(r'[<>:"/\\|?*]', "_", resultado)
    # Colapsar espacios multiples
    resultado = re.sub(r' +', " ", resultado).strip()
    return resultado

# ---------------------------------------------------------------
# OBTENER FECHA DEL CORREO
# ---------------------------------------------------------------
def obtener_fecha_correo(msg):
    try:
        tupla = email.utils.parsedate_tz(msg.get("Date", ""))
        if tupla:
            return datetime.fromtimestamp(email.utils.mktime_tz(tupla))
    except Exception:
        pass
    return datetime.now()

# ---------------------------------------------------------------
# NOMBRE SEGURO (evita rutas largas en Windows)
# ---------------------------------------------------------------
def nombre_seguro(carpeta, filename):
    base, ext = os.path.splitext(filename)
    if len(base) > MAX_NOMBRE_PDF:
        base = base[:MAX_NOMBRE_PDF].strip()
    nombre_final = base + ext
    ruta_total   = os.path.join(carpeta, nombre_final)
    if len(ruta_total) > 255:
        espacio = 255 - len(carpeta) - len(ext) - 2
        nombre_final = (base[:espacio].strip() if espacio > 10
                        else f"doc_{datetime.now().strftime('%H%M%S')}") + ext
    return nombre_final

# ---------------------------------------------------------------
# RESOLVER DESTINO (duplicados legibles)
# ---------------------------------------------------------------
def resolver_destino(carpeta, filename):
    filename = nombre_seguro(carpeta, filename)
    destino  = os.path.join(carpeta, filename)
    if not os.path.exists(destino):
        return destino, "OK"
    base, ext = os.path.splitext(filename)
    if len(base) > MAX_NOMBRE_PDF - 12:
        base = base[:MAX_NOMBRE_PDF - 12].strip()
    destino_dup = os.path.join(carpeta, f"{base} (duplicado){ext}")
    if not os.path.exists(destino_dup):
        return destino_dup, "DUPLICADO"
    contador = 2
    while True:
        destino_dup = os.path.join(carpeta, f"{base} (duplicado {contador}){ext}")
        if not os.path.exists(destino_dup):
            return destino_dup, "DUPLICADO"
        contador += 1

# ---------------------------------------------------------------
# DETECTAR .EML ADJUNTO (incluyendo octet-stream de Gmail)
# ---------------------------------------------------------------
def es_zip_eml(fn):
    """Detecta adjuntos .zip.eml por su nombre de archivo."""
    fn_lower = fn.lower()
    return fn_lower.endswith(".zip.eml")

def extraer_pdfs_de_zip_eml(part):
    """
    Procesa un adjunto .zip.eml:
    1. Extrae el .eml del adjunto
    2. Dentro del .eml busca un .zip
    3. Descomprime el .zip
    4. Retorna todos los PDFs que encuentre dentro
    Retorna lista de (filename, pdf_bytes)
    """
    resultado = []
    try:
        # Paso 1: Obtener el contenido del .eml
        raw_eml = part.get_payload(decode=True)
        if not raw_eml:
            return resultado

        # Paso 2: Parsear el .eml
        sub_msg = email.message_from_bytes(raw_eml)

        # Paso 3: Buscar el .zip dentro del .eml
        for sub_part in sub_msg.walk():
            sub_fn  = sub_part.get_filename() or ""
            sub_ct  = sub_part.get_content_type()
            sub_fn_lower = sub_fn.lower()

            if (sub_fn_lower.endswith(".zip") or
                sub_ct in ("application/zip", "application/x-zip-compressed",
                           "application/octet-stream")):
                zip_bytes = sub_part.get_payload(decode=True)
                if not zip_bytes:
                    continue
                try:
                    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                        for nombre_zip in zf.namelist():
                            if nombre_zip.lower().endswith(".pdf"):
                                pdf_bytes = zf.read(nombre_zip)
                                fn_limpio = decodificar_filename(
                                    os.path.basename(nombre_zip)
                                )
                                resultado.append((fn_limpio, pdf_bytes))
                except Exception:
                    continue

    except Exception:
        pass
    return resultado

EXTENSIONES_ARCHIVO = {
    ".pdf", ".xlsx", ".xls", ".doc", ".docx", ".zip", ".rar",
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff",
    ".txt", ".csv", ".xml", ".html", ".htm", ".ppt", ".pptx",
}

def _tiene_extension_conocida(fn):
    """Retorna True si el filename tiene una extension de archivo conocida (no-email)."""
    fn = fn.lower()
    for ext in EXTENSIONES_ARCHIVO:
        if fn.endswith(ext):
            return True
    return False

def es_adjunto_eml(part):
    ct = part.get_content_type()
    fn = (part.get_filename() or "").lower()
    if ct == "message/rfc822":
        return True
    if fn.endswith(".eml"):
        return True
    # octet-stream sin extension conocida — puede ser un .eml sin extension
    # Se verifica intentando parsear el contenido como email
    if ct == "application/octet-stream" and not _tiene_extension_conocida(fn):
        try:
            raw = part.get_payload(decode=True)
            if raw and len(raw) > 50:
                # Verificar que tenga headers tipicos de email
                cabecera = raw[:2000].decode("utf-8", errors="replace")
                if any(h in cabecera for h in ["From:", "To:", "Subject:", "Date:", "MIME-Version:"]):
                    return True
        except Exception:
            pass
    return False

def extraer_sub_msg(part):
    try:
        ct = part.get_content_type()
        if ct == "message/rfc822":
            payload = part.get_payload()
            if isinstance(payload, list):
                return payload[0]
            elif isinstance(payload, email.message.Message):
                return payload
        raw = part.get_payload(decode=True)
        if raw:
            return email.message_from_bytes(raw)
    except Exception:
        pass
    return None

# ---------------------------------------------------------------
# EXTRAER PDFs DE UN MENSAJE (directo o dentro de .eml)
# ---------------------------------------------------------------
def extraer_pdfs_de_mensaje(msg):
    resultado = []
    asunto_principal    = decodificar_asunto(msg)
    fecha_principal     = obtener_fecha_correo(msg)
    remitente_principal = msg.get("From", "")

    for part in msg.walk():
        content_type = part.get_content_type()
        raw_filename = part.get_filename()

        # PDF directo
        if content_type == "application/pdf" or (
            raw_filename and raw_filename.lower().endswith(".pdf")
        ):
            filename  = decodificar_filename(raw_filename) if raw_filename else "adjunto.pdf"
            pdf_bytes = part.get_payload(decode=True)
            if pdf_bytes:
                resultado.append((filename, pdf_bytes,
                                   asunto_principal, fecha_principal, remitente_principal))

        # Correo incrustado .eml
        elif es_adjunto_eml(part):
            sub_msg = extraer_sub_msg(part)
            if not sub_msg:
                continue
            asunto_sub    = decodificar_asunto(sub_msg)
            fecha_sub     = obtener_fecha_correo(sub_msg)
            remitente_sub = sub_msg.get("From", "")
            for sub_part in sub_msg.walk():
                sub_fn = sub_part.get_filename()
                if sub_part.get_content_type() == "application/pdf" or (
                    sub_fn and sub_fn.lower().endswith(".pdf")
                ):
                    fn        = decodificar_filename(sub_fn) if sub_fn else "adjunto_eml.pdf"
                    pdf_bytes = sub_part.get_payload(decode=True)
                    if pdf_bytes:
                        resultado.append((fn, pdf_bytes,
                                           asunto_sub, fecha_sub, remitente_sub))
    return resultado

# ---------------------------------------------------------------
# DETECTAR SI UN CORREO ES CONTENEDOR DE .eml SIN ASUNTO
# ---------------------------------------------------------------
def es_contenedor_eml_sin_asunto(msg):
    asunto = decodificar_asunto(msg).strip()
    if asunto:
        return False
    for part in msg.walk():
        if es_adjunto_eml(part):
            return True
    return False

# ---------------------------------------------------------------
# GUARDAR UN PDF Y REGISTRAR
# ---------------------------------------------------------------
def guardar_pdf(pdf_bytes, filename, asunto, fecha_correo,
                remitente, correo_cuenta, ips_dict, aseg_dict,
                sesion_id=""):
    temp_pdf = os.path.join(
        os.environ.get("TEMP", os.path.expanduser("~")),
        f"glosa_tmp_{filename[:40]}"
    )
    try:
        with open(temp_pdf, "wb") as f:
            f.write(pdf_bytes)

        texto_pdf           = extract_text(temp_pdf)
        texto_norm          = normalizar(texto_pdf)
        asunto_norm         = normalizar(asunto)
        remitente_norm      = normalizar(remitente)
        # Separar texto del PDF del contexto del correo para que
        # detectar_ips busque primero en el PDF (más confiable)
        # y use asunto+remitente solo como fallback
        texto_contexto_norm = asunto_norm + " " + remitente_norm

        nit, ips_nombre, ciudad = detectar_ips(texto_norm, ips_dict,
                                               texto_contexto_norm=texto_contexto_norm)
        aseguradora = detectar_aseguradora(
            texto_norm, asunto_norm, remitente_norm, aseg_dict
        )

        if nit is None:
            os.makedirs(CARPETA_NO_ID, exist_ok=True)
            destino, estado = resolver_destino(CARPETA_NO_ID, filename)
            with open(destino, "wb") as f:
                f.write(pdf_bytes)
            print(f"       [NO IDENTIFICADO] {os.path.basename(destino)}")
            ips_nombre = "NO IDENTIFICADA"
            nit        = "NO IDENTIFICADO"
            ciudad     = "NO IDENTIFICADA"
        else:
            carpeta = crear_ruta(ciudad, nit, ips_nombre, aseguradora, fecha_correo)
            destino, estado = resolver_destino(carpeta, filename)
            with open(destino, "wb") as f:
                f.write(pdf_bytes)
            tag = "[DUPLICADO]" if estado == "DUPLICADO" else "[OK]"
            print(f"       {tag} {os.path.basename(destino)}")
            print(f"             Ciudad: {ciudad} | IPS: {ips_nombre} | Aseg: {aseguradora}")

        return {
            "sesion"       : sesion_id,
            "fecha_correo" : fecha_correo.strftime("%Y-%m-%d %H:%M"),
            "correo_origen": correo_cuenta,
            "ciudad"       : ciudad,
            "ips"          : ips_nombre,
            "nit"          : nit,
            "aseguradora"  : aseguradora,
            "archivo"      : os.path.basename(destino),
            "ruta"         : destino,
            "estado"       : estado,
        }, estado

    except Exception as e:
        print(f"       [ERROR] {filename[:60]}: {e}")
        return None, "ERROR"
    finally:
        if os.path.exists(temp_pdf):
            try:
                os.remove(temp_pdf)
            except Exception:
                pass

# ---------------------------------------------------------------
# ACTUALIZAR REPORTE EXCEL
# ---------------------------------------------------------------
def actualizar_reporte(registros):
    if not registros:
        return
    COLUMNAS = ["sesion", "fecha_correo", "correo_origen", "ciudad", "ips", "nit",
                "aseguradora", "archivo", "ruta", "estado"]
    df_nuevo = pd.DataFrame(registros, columns=COLUMNAS)
    if os.path.exists(REPORTE_EXCEL):
        try:
            df_viejo = pd.read_excel(REPORTE_EXCEL)
            df_viejo = df_viejo.rename(columns={"fecha": "fecha_correo"})
            for col in COLUMNAS:
                if col not in df_viejo.columns:
                    df_viejo[col] = ""
            df_viejo = df_viejo[COLUMNAS]
            df_viejo = df_viejo[
                df_viejo["estado"].notna() &
                df_viejo["estado"].astype(str).str.strip().ne("") &
                df_viejo["estado"].astype(str).str.strip().ne("None") &
                df_viejo["archivo"].notna() &
                df_viejo["archivo"].astype(str).str.strip().ne("")
            ]
            df_final = pd.concat([df_viejo, df_nuevo], ignore_index=True)
        except Exception:
            df_final = df_nuevo
    else:
        df_final = df_nuevo
    os.makedirs(os.path.dirname(REPORTE_EXCEL), exist_ok=True)
    df_final.to_excel(REPORTE_EXCEL, index=False)
    print(f"\n  Reporte guardado: {len(registros)} registro(s) -> {REPORTE_EXCEL}")

# ---------------------------------------------------------------
# CONECTAR
# ---------------------------------------------------------------
def conectar(email_cuenta, password):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(email_cuenta, password)
    mail.select("inbox")
    return mail

# ---------------------------------------------------------------
# OBTENER ASUNTO DE UN CORREO SIN ABRIRLO (PEEK)
# No marca el correo como leido
# ---------------------------------------------------------------
def peek_asunto(mail, num):
    try:
        st, hdr = mail.fetch(num, "(BODY.PEEK[HEADER.FIELDS (SUBJECT)])")
        if st != "OK" or not hdr or not hdr[0]:
            return ""
        raw = hdr[0][1] if isinstance(hdr[0], tuple) else b""
        msg_hdr = email.message_from_bytes(raw)
        return decodificar_asunto(msg_hdr).strip()
    except Exception:
        return ""

# ---------------------------------------------------------------
# PROCESAR UNA CUENTA
# Flujo:
# 1. Traer IDs de todos los UNSEEN
# 2. Leer asunto con PEEK (sin marcar leido)
# 3. Validar asunto contra patrones del Excel
# 4. Si valido o es contenedor .eml sin asunto: abrir y procesar
# 5. Si no valido: ignorar completamente
# ---------------------------------------------------------------
def procesar_cuenta(cuenta, ips_dict, aseg_dict, patrones_asuntos, sesion_id=""):
    correo_cuenta = cuenta["correo"]
    password      = cuenta["password"]
    print(f"\n  Conectando: {correo_cuenta}")

    registros        = []
    total_ok         = 0
    total_duplicados = 0
    total_no_id      = 0

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(correo_cuenta, password)
    except Exception as e:
        print(f"  [ERROR DE CONEXION] {correo_cuenta}: {e}")
        return registros, 0, 0, 0

    # Listar bandejas disponibles para detectar nombres reales
    try:
        # Usar UTF-7 modificado que es el encoding oficial de IMAP para nombres de carpetas
        # Esto evita el error 'ascii' codec con caracteres como — en nombres de etiquetas
        mail.literal = None
        status_list, bandejas_raw = mail.list()
        bandejas_disponibles = set()
        if status_list == "OK":
            for b in bandejas_raw:
                if not b:
                    continue
                try:
                    # Decodificar con utf-8, ignorando errores para no romper con caracteres raros
                    decoded = b.decode("utf-8", errors="ignore")
                    # Filtrar caracteres no-ASCII del nombre para que mail.select() funcione
                    partes = decoded.strip().split('"')
                    if len(partes) >= 3:
                        nombre = partes[-2] if partes[-1].strip() == "" else partes[-1].strip().strip('"')
                    else:
                        nombre = decoded.split()[-1].strip('"')
                    # Solo agregar si el nombre es ASCII puro — evita error en mail.select()
                    if nombre and all(ord(c) < 128 for c in nombre):
                        bandejas_disponibles.add(nombre)
                except Exception:
                    continue
    except Exception:
        bandejas_disponibles = set()

    # Bandejas a revisar — se detectan dinamicamente segun lo disponible en la cuenta
    BANDEJAS = [
        "INBOX",
        "[Gmail]/Promotions",
        "[Gmail]/Updates",
        "[Gmail]/Social",
        "[Gmail]/Promociones",
        "[Gmail]/Notificaciones",
        "[Gmail]/Actualizaciones",
        "INBOX/Promotions",
        "INBOX/Updates",
        "INBOX/Notifications",
    ]

    todos_unseen_por_bandeja = {}
    bandejas_a_revisar = []

    # Primero agregar INBOX siempre
    bandejas_a_revisar.append("INBOX")

    # Luego agregar bandejas disponibles que coincidan con las que queremos
    keywords = ["promot", "updat", "social", "notif", "actuali"]
    for b in bandejas_disponibles:
        b_lower = b.lower()
        if any(kw in b_lower for kw in keywords):
            if b not in bandejas_a_revisar:
                bandejas_a_revisar.append(b)

    # Si no encontro ninguna por nombre, intentar las de la lista fija
    if len(bandejas_a_revisar) == 1:
        for bandeja in BANDEJAS:
            if bandeja != "INBOX" and bandeja not in bandejas_a_revisar:
                bandejas_a_revisar.append(bandeja)

    for bandeja in bandejas_a_revisar:
        try:
            # Encodear nombre de bandeja para evitar error con caracteres
            # especiales como em-dash (—) que no son ASCII
            bandeja_enc = bandeja.encode("utf-8").decode("ascii", errors="replace") \
                if not all(ord(c) < 128 for c in bandeja) else bandeja
            status, _ = mail.select(bandeja_enc, readonly=False)
            if status != "OK":
                continue
            status, data = mail.search(None, "UNSEEN")
            if status == "OK" and data[0]:
                ids = data[0].split()
                if ids:
                    todos_unseen_por_bandeja[bandeja] = ids
                    print(f"  {bandeja}: {len(ids)} sin leer")
        except Exception:
            continue

    total_unseen = sum(len(v) for v in todos_unseen_por_bandeja.values())
    if total_unseen == 0:
        print("  Sin correos sin leer en ninguna bandeja.")
        try:
            mail.logout()
        except Exception:
            pass
        return registros, 0, 0, 0

    # Aplanar todos los IDs con su bandeja de origen
    todos_unseen = []
    for bandeja, ids in todos_unseen_por_bandeja.items():
        for id_msg in ids:
            todos_unseen.append((bandeja, id_msg))

    print(f"  Total sin leer: {total_unseen}")

    # Filtrar por asunto usando PEEK (sin abrir ni marcar como leido)
    ids_a_procesar = []
    ids_contenedor = []
    asuntos_omitidos = {}  # para diagnostico agrupado

    for bandeja, num in todos_unseen:
        try:
            mail.select(bandeja, readonly=False)
        except Exception:
            continue

        asunto = peek_asunto(mail, num)
        asunto_norm = normalizar(asunto)

        if asunto_norm == "":
            try:
                st, data_full = mail.fetch(num, "(BODY.PEEK[])")
                if st == "OK" and data_full and data_full[0]:
                    raw = data_full[0][1]
                    msg_tmp = email.message_from_bytes(raw)
                    if es_contenedor_eml_sin_asunto(msg_tmp):
                        ids_contenedor.append((bandeja, num, msg_tmp))
            except Exception:
                pass
        elif asunto_es_valido(asunto_norm, patrones_asuntos):
            ids_a_procesar.append((bandeja, num))
        else:
            # Agrupar omitidos por patron de asunto (primeras 60 chars)
            clave = asunto[:60] if asunto else "(sin asunto)"
            asuntos_omitidos[clave] = asuntos_omitidos.get(clave, 0) + 1

    total_validos = len(ids_a_procesar) + len(ids_contenedor)
    print(f"  Validos por asunto  : {len(ids_a_procesar)}")
    print(f"  Contenedores .eml   : {len(ids_contenedor)}")
    print(f"  Omitidos            : {len(todos_unseen) - total_validos}")
    if asuntos_omitidos:
        print("  Asuntos omitidos (muestra):")
        for asunto_txt, qty in sorted(asuntos_omitidos.items(), key=lambda x: -x[1])[:10]:
            print(f"    x{qty:3}  {asunto_txt}")

    if total_validos == 0:
        print("  Nada que procesar.")
        mail.logout()
        return registros, 0, 0, 0

    # ── Procesar correos con asunto valido ──
    for i, (bandeja, num) in enumerate(ids_a_procesar, 1):
        try:
            mail.select(bandeja, readonly=False)
        except Exception:
            continue

        st, data_full = mail.fetch(num, "(RFC822)")
        if st != "OK" or not data_full or not data_full[0]:
            continue

        raw_email    = data_full[0][1]
        msg          = email.message_from_bytes(raw_email)
        asunto       = decodificar_asunto(msg)
        fecha_correo = obtener_fecha_correo(msg)
        remitente    = msg.get("From", "")

        print(f"\n  [{i}/{len(ids_a_procesar)}] {asunto[:75]}")
        print(f"       Bandeja: {bandeja}")
        print(f"       Fecha: {fecha_correo.strftime('%Y-%m-%d %H:%M')}")

        pdfs = extraer_pdfs_de_mensaje(msg)

        # ── Correos Claimonline: link en cuerpo, sin PDF adjunto directo ──
        if not pdfs and es_correo_claimonline(msg):
            print("       [Claimonline] Correo con link detectado — iniciando descarga...")
            carpeta_temp = os.path.join(
                os.environ.get("TEMP", os.path.expanduser("~")),
                "glosas_claimonline_tmp"
            )
            pdfs = procesar_correo_claimonline(
                msg,
                ips_dict,
                carpeta_temp,
                asunto=asunto,
                fecha=fecha_correo,
                remitente=remitente,
            )
            if not pdfs:
                print("       [Claimonline] No se pudieron obtener PDFs del link.")

        if not pdfs:
            print("       Sin PDFs adjuntos.")
            continue

        print(f"       PDFs encontrados: {len(pdfs)}")
        correo_procesado = False

        for filename, pdf_bytes, asunto_src, fecha_src, remitente_src in pdfs:
            # Marcador especial: ZIP ubicado en ruta IPS por Claimonline
            if filename.startswith("__ZIP_IPS__") or filename.startswith("__MANUAL__"):
                print(f"       [Claimonline] ZIP guardado en ruta IPS — correo se marca leido.")
                correo_procesado = True
                continue

            registro, estado = guardar_pdf(
                pdf_bytes, filename,
                asunto_src or asunto,
                fecha_src or fecha_correo,
                remitente_src or remitente,
                correo_cuenta, ips_dict, aseg_dict,
                sesion_id=sesion_id
            )
            if registro:
                registros.append(registro)
                correo_procesado = True
                if estado == "OK":           total_ok += 1
                elif estado == "DUPLICADO":  total_duplicados += 1
                else:                        total_no_id += 1

        if correo_procesado:
            mail.store(num, "+FLAGS", "\\Seen")
            print("       -> Marcado como leido.")
        else:
            print("       -> Sin PDFs procesables. NO marcado como leido.")

    # ── Procesar contenedores .eml sin asunto ──
    for j, (bandeja, num, msg) in enumerate(ids_contenedor, 1):
        try:
            mail.select(bandeja, readonly=False)
        except Exception:
            continue

        remitente    = msg.get("From", "")
        fecha_correo = obtener_fecha_correo(msg)

        print(f"\n  [EML {j}/{len(ids_contenedor)}] (sin asunto) — {remitente[:60]}")
        print(f"       Bandeja: {bandeja}")
        print(f"       Fecha: {fecha_correo.strftime('%Y-%m-%d %H:%M')}")

        pdfs = extraer_pdfs_de_mensaje(msg)

        # ── Verificar si algún sub-.eml es un correo de Claimonline ──────────
        if not pdfs:
            for part in msg.walk():
                if not es_adjunto_eml(part):
                    continue
                sub_msg = extraer_sub_msg(part)
                if not sub_msg:
                    continue
                if es_correo_claimonline(sub_msg):
                    print("       [Claimonline] Sub-.eml con link detectado — iniciando descarga...")
                    asunto_sub   = decodificar_asunto(sub_msg)
                    fecha_sub    = obtener_fecha_correo(sub_msg)
                    remitente_sub = sub_msg.get("From", "")
                    carpeta_temp = os.path.join(
                        os.environ.get("TEMP", os.path.expanduser("~")),
                        "glosas_claimonline_tmp"
                    )
                    pdfs_cl = procesar_correo_claimonline(
                        sub_msg,
                        ips_dict,
                        carpeta_temp,
                        asunto=asunto_sub,
                        fecha=fecha_sub or fecha_correo,
                        remitente=remitente_sub or remitente,
                    )
                    if pdfs_cl:
                        pdfs = pdfs_cl
                        break
                    else:
                        print("       [Claimonline] No se pudieron obtener archivos del sub-.eml.")

        if not pdfs:
            print("       Sin PDFs dentro de los .eml.")
            continue

        print(f"       PDFs encontrados: {len(pdfs)}")
        correo_procesado = False

        for filename, pdf_bytes, asunto_src, fecha_src, remitente_src in pdfs:
            # Marcador especial: ZIP ubicado en ruta IPS por Claimonline
            if filename.startswith("__ZIP_IPS__") or filename.startswith("__MANUAL__"):
                print(f"       [Claimonline] ZIP guardado en ruta IPS — correo se marca leido.")
                correo_procesado = True
                continue

            registro, estado = guardar_pdf(
                pdf_bytes, filename,
                asunto_src, fecha_src or fecha_correo,
                remitente_src or remitente,
                correo_cuenta, ips_dict, aseg_dict,
                sesion_id=sesion_id
            )
            if registro:
                registros.append(registro)
                correo_procesado = True
                if estado == "OK":           total_ok += 1
                elif estado == "DUPLICADO":  total_duplicados += 1
                else:                        total_no_id += 1

        if correo_procesado:
            mail.store(num, "+FLAGS", "\\Seen")
            print("       -> Marcado como leido.")
        else:
            print("       -> Sin PDFs procesables. NO marcado como leido.")

    try:
        mail.logout()
    except Exception:
        pass
    return registros, total_ok, total_duplicados, total_no_id

# ---------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------
def procesar_correos():
    cuentas, ips_dict, aseg_dict, patrones_asuntos, _ = cargar_configuracion()
    print(f"\n  Cuentas de correo   : {len(cuentas)}")
    print(f"  IPS cargadas        : {len(ips_dict)}")
    print(f"  Equiv. aseguradoras : {len(aseg_dict)}")
    print(f"  Patrones de asunto  : {len(patrones_asuntos)}")

    os.makedirs(CARPETA_NO_ID, exist_ok=True)
    os.makedirs(RUTA_BASE, exist_ok=True)

    # Timestamp de inicio de sesion — identifica todos los registros de esta ejecucion
    sesion_id = datetime.now().strftime("%Y-%m-%d %H:%M")

    todos_registros = []
    gran_ok = gran_duplicados = gran_no_id = 0

    print("\n" + "=" * 55)

    for idx, cuenta in enumerate(cuentas, 1):
        print(f"\n[CUENTA {idx}/{len(cuentas)}] {cuenta['correo']}")
        print("-" * 55)
        registros, ok, dup, no_id = procesar_cuenta(
            cuenta, ips_dict, aseg_dict, patrones_asuntos, sesion_id
        )
        todos_registros.extend(registros)
        gran_ok += ok; gran_duplicados += dup; gran_no_id += no_id
        print(f"\n  Cuenta finalizada -> OK: {ok} | Duplicados: {dup} | No ID: {no_id}")

    actualizar_reporte(todos_registros)

    print("\n" + "=" * 55)
    print("  RESUMEN FINAL - TODAS LAS CUENTAS")
    print("=" * 55)
    for cuenta in cuentas:
        print(f"  {cuenta['correo']}")
    print("-" * 55)
    print(f"  PDFs guardados OK   : {gran_ok}")
    print(f"  PDFs duplicados     : {gran_duplicados}")
    print(f"  No identificados    : {gran_no_id}")
    print(f"  Total en reporte    : {gran_ok + gran_duplicados + gran_no_id}")
    print("=" * 55)

# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
def main():
    info = datos_identidad()
    sep  = "=" * 55
    print(sep)
    print(f"  {info['sistema']}")
    print(f"  {info['propietario']} | {info['version']}")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(sep)
    try:
        procesar_correos()
    except Exception as e:
        print(f"\n[ERROR CRITICO] {e}")
        import traceback
        traceback.print_exc()
    print(f"\n  {info['desarrollado_por']}")
    print(sep)

# ---------------------------------------------------------------
# MODO PRUEBA — Solo verifica asuntos, NO descarga ni marca leidos
# Ejecutar con: python bot_glosas_v5.py --prueba
# ---------------------------------------------------------------
def prueba_asuntos():
    """
    Conecta a cada cuenta, lee los asuntos de los correos sin leer
    usando PEEK (sin marcarlos como leidos) y reporta cuales
    procesaria y cuales omitira. No descarga ningun PDF.
    """
    print("=" * 55)
    print("  MODO PRUEBA — SIN DESCARGAS NI CAMBIOS")
    print("=" * 55)

    cuentas, ips_dict, aseg_dict, patrones_asuntos, _ = cargar_configuracion()
    print(f"\n  Patrones cargados: {len(patrones_asuntos)}\n")

    for idx, cuenta in enumerate(cuentas, 1):
        correo_cuenta = cuenta["correo"]
        password      = cuenta["password"]
        print(f"\n[CUENTA {idx}] {correo_cuenta}")
        print("-" * 55)

        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(correo_cuenta, password)
        except Exception as e:
            print(f"  [ERROR CONEXION] {e}")
            continue

        # Detectar bandejas disponibles
        try:
            _, bandejas_raw = mail.list()
            bandejas_disponibles = set()
            for b in (bandejas_raw or []):
                if b:
                    decoded = b.decode("utf-8", errors="replace")
                    partes  = decoded.strip().split('"')
                    nombre  = partes[-2] if len(partes) >= 3 and partes[-1].strip() == "" else partes[-1].strip().strip('"')
                    bandejas_disponibles.add(nombre)
        except Exception:
            bandejas_disponibles = set()

        keywords_bandejas = ["promot", "updat", "social", "notif", "actuali"]
        bandejas_dinamicas = [b for b in bandejas_disponibles
                              if any(kw in b.lower() for kw in keywords_bandejas)]
        bandejas_fijas = [
            "[Gmail]/Promotions", "[Gmail]/Updates", "[Gmail]/Social",
            "[Gmail]/Promociones", "[Gmail]/Notificaciones", "[Gmail]/Actualizaciones",
        ]
        bandejas_candidatas = ["INBOX"]
        for b in bandejas_dinamicas + bandejas_fijas:
            if b not in bandejas_candidatas:
                bandejas_candidatas.append(b)

        procesaria    = []
        omitira       = {}
        total_unseen  = 0

        for bandeja in bandejas_candidatas:
            try:
                status, _ = mail.select(bandeja, readonly=True)  # readonly=True — no modifica nada
                if status != "OK":
                    continue
                status, data = mail.search(None, "UNSEEN")
                if status != "OK" or not data[0]:
                    continue
                ids = data[0].split()
                total_unseen += len(ids)

                for num in ids:
                    asunto = peek_asunto(mail, num)
                    asunto_norm = normalizar(asunto)

                    if asunto_norm == "":
                        # Verificar si tiene .eml adjunto sin abrir completamente
                        try:
                            st, hdr = mail.fetch(num, "(BODY.PEEK[MIME])")
                            asunto = "(sin asunto — posible contenedor .eml)"
                        except Exception:
                            pass
                        procesaria.append({
                            "bandeja": bandeja,
                            "asunto" : asunto or "(sin asunto)",
                            "motivo" : "sin asunto — se abre para verificar .eml"
                        })
                    elif asunto_es_valido(asunto_norm, patrones_asuntos):
                        procesaria.append({
                            "bandeja": bandeja,
                            "asunto" : asunto[:80],
                            "motivo" : "patron coincide"
                        })
                    else:
                        clave = asunto[:60] if asunto else "(sin asunto)"
                        omitira[clave] = omitira.get(clave, 0) + 1

            except Exception as e:
                print(f"  {bandeja}: error — {e}")
                continue

        print(f"  Total sin leer    : {total_unseen}")
        print(f"  PROCESARIA        : {len(procesaria)}")
        print(f"  OMITIRA           : {total_unseen - len(procesaria)}")

        if procesaria:
            print("\n  --- CORREOS QUE PROCESARIA ---")
            for c in procesaria:
                print(f"  [SI] [{c['bandeja']}] {c['asunto']}")
                print(f"        Motivo: {c['motivo']}")

        if omitira:
            print("\n  --- ASUNTOS QUE OMITIRA (top 15) ---")
            for txt, qty in sorted(omitira.items(), key=lambda x: -x[1])[:15]:
                print(f"  [NO] x{qty:3}  {txt}")

        try:
            mail.logout()
        except Exception:
            pass

    print("\n" + "=" * 55)
    print("  FIN PRUEBA — Ningun correo fue modificado")
    print("=" * 55)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--prueba":
        prueba_asuntos()
    else:
        main()