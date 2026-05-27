import os
import re
import shutil
from pathlib import Path
from openpyxl import Workbook
import fitz  # PyMuPDF


def datos_identidad():
    return {
        "sistema": "RENOMBRADOR DE AUDITORIA",
        "propietario": "Salud-Net",
        "desarrollado_por": "DESARROLLO E INNOVACIÓN SALUD NET",
        "version": "v1.0",
        "licencia": "Uso interno autorizado",
    }


def mostrar_marca_agua():
    datos = datos_identidad()

    try:
        os.system(f'title {datos["sistema"]} - {datos["propietario"]}')
    except:
        pass

    print("\n" + "=" * 68)
    print(f'{datos["sistema"]:^68}')
    print("=" * 68)
    print(f'PROPIETARIO : {datos["propietario"]}')
    print(f'DESARROLLADO: {datos["desarrollado_por"]}')
    print(f'VERSION     : {datos["version"]}')
    print(f'LICENCIA    : {datos["licencia"]}')
    print("=" * 68 + "\n")


def extract_text_pdf(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text()
    except:
        pass
    return text.upper()


def extract_factura_filename(nombre):
    match = re.search(r'([A-Z]?\d+)-(\d+)', nombre.upper())
    if match:
        return match.group(1) + match.group(2)
    return None


def extract_factura_text(text):
    text = text.upper()
    text = text.replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)

    matches = re.findall(r'\b([A-Z]?\d{1,3})\s*-\s*(\d{3,})\b', text)

    if matches:
        prefijo, numero = matches[-1]
        return f"{prefijo}{numero}"

    return None


# 🔥 FIX NIT (ya probado)
def extract_nit(text):
    text = text.replace(".", "").replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)

    match = re.search(r'\b(\d{9})\s+\d\b', text)
    if match:
        return match.group(1)

    match = re.search(r'\b(\d{10})\b', text)
    if match:
        return match.group(1)[:9]

    match = re.search(r'\b(\d{9})\b', text)
    if match:
        return match.group(1)

    return None


# 🔥 DETECCIÓN ASEGURADORA CON REGEX Y PRIORIDADES
def detectar_aseguradora(text):
    text = text.upper().replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)

    candidatos = []

    reglas = [
        # ESTADO
        (r'\bSEGUROS\s+DEL\s+ESTADO\s+S\.?\s*A\.?\b', "SEGUROS DEL ESTADO S.A.", 100),
        (r'\bSEGUROS\s+DEL\s+ESTADO\b', "SEGUROS DEL ESTADO S.A.", 95),
        (r'\bSEGUROS\s+ESTADO\b', "SEGUROS DEL ESTADO S.A.", 85),

        # PREVISORA
        (r'\bLA\s+PREVISORA\s+S\.?\s*A\.?\b', "LA PREVISORA S.A.", 100),
        (r'\bLA\s+PREVISORA\b', "LA PREVISORA S.A.", 95),
        (r'\bPREVISORA\b', "LA PREVISORA S.A.", 70),

        # MUNDIAL
        (r'\bMUNDIAL\s+DE\s+SEGUROS\s+S\.?\s*A\.?\b', "MUNDIAL DE SEGUROS S.A.", 100),
        (r'\bMUNDIAL\s+DE\s+SEGUROS\b', "MUNDIAL DE SEGUROS S.A.", 95),
        (r'\bSEGUROS\s+MUNDIAL\b', "MUNDIAL DE SEGUROS S.A.", 90),
        (r'\bCOMPAÑIA\s+MUNDIAL\b', "MUNDIAL DE SEGUROS S.A.", 85),
        (r'\bCOMPANIA\s+MUNDIAL\b', "MUNDIAL DE SEGUROS S.A.", 85),
        (r'\bMUNDIAL\b', "MUNDIAL DE SEGUROS S.A.", 60),

        # SURA
        (r'\bSEGUROS\s+GENERALES\s+SURAMERICANA\s+S\.?\s*A\.?\b', "SEGUROS GENERALES SURAMERICANA S.A.", 100),
        (r'\bSEGUROS\s+GENERALES\s+SURAMERICANA\b', "SEGUROS GENERALES SURAMERICANA S.A.", 95),
        (r'\bSURAMERICANA\b', "SEGUROS GENERALES SURAMERICANA S.A.", 80),
        (r'\bSURA\b', "SEGUROS GENERALES SURAMERICANA S.A.", 60),

        # BOLIVAR
        # OJO: se elimina "BOLIVAR" suelto para no confundir con departamento/ciudad
        (r'\bSEGUROS\s+COMERCIALES\s+BOLIVAR\s+S\.?\s*A\.?\b', "SEGUROS COMERCIALES BOLIVAR S.A.", 100),
        (r'\bSEGUROS\s+COMERCIALES\s+BOLIVAR\b', "SEGUROS COMERCIALES BOLIVAR S.A.", 95),
        (r'\bSEGUROS\s+BOLIVAR\b', "SEGUROS COMERCIALES BOLIVAR S.A.", 85),
    ]

    for patron, nombre, puntaje in reglas:
        if re.search(patron, text, re.IGNORECASE):
            candidatos.append((puntaje, nombre, patron))

    if not candidatos:
        return None

    candidatos.sort(key=lambda x: x[0], reverse=True)
    return candidatos[0][1]


def procesar_pdf(pdf_path, base, revision):
    nombre = pdf_path.name

    text = extract_text_pdf(pdf_path)
    aseguradora = detectar_aseguradora(text)

    factura_texto = extract_factura_text(text)
    factura_nombre = extract_factura_filename(nombre)
    nit = extract_nit(text)

    factura = factura_texto if factura_texto else factura_nombre

    # DEBUG (puedes quitar después)
    print(f"Aseguradora: {aseguradora} | NIT: {nit} | Factura: {factura}")

    # 🔥 LÓGICA DE RENOMBRE
    if aseguradora == "MUNDIAL DE SEGUROS S.A." and factura and nit:
        nuevo = f"R-{nit}-{factura}.pdf"

    elif aseguradora == "LA PREVISORA S.A." and factura and nit:
        nuevo = f"CRC_{nit}_{factura}.pdf"

    elif aseguradora == "SEGUROS COMERCIALES BOLIVAR S.A." and factura:
        nuevo = f"{factura}_PRG_1.pdf"

    elif aseguradora in [
        "SEGUROS GENERALES SURAMERICANA S.A.",
        "SEGUROS DEL ESTADO S.A."
    ] and factura:
        nuevo = f"{factura}.pdf"

    else:
        shutil.move(str(pdf_path), revision / nombre)
        return ("REVISION", nombre, "")

    destino = base / nuevo

    if destino.exists():
        destino = base / f"DUPLICADO_{nuevo}"

    pdf_path.rename(destino)

    return ("OK", nombre, destino.name)


def main():
    mostrar_marca_agua()

    base = Path.cwd()

    revision = base / "revision_manual"
    revision.mkdir(exist_ok=True)

    archivos = list(base.glob("*.pdf"))

    if not archivos:
        print("No se encontraron PDFs para procesar en:", base)
        input("Presione una tecla para salir...")
        return

    log_file = base / "REPORTE_FINAL.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    ws.append(["estado", "archivo_original", "archivo_nuevo"])

    ok, rev, err = 0, 0, 0

    for pdf in archivos:
        try:
            estado, original, nuevo = procesar_pdf(pdf, base, revision)
            ws.append([estado, original, nuevo])

            if estado == "OK":
                ok += 1
            elif estado == "REVISION":
                rev += 1
            else:
                err += 1

        except Exception as e:
            ws.append(["ERROR", pdf.name, str(e)])
            err += 1

    wb.save(log_file)

    print("\nResumen:")
    print(f"OK: {ok}")
    print(f"Revision: {rev}")
    print(f"Error: {err}")
    print(f"Reporte: {log_file}")

    input("\nPresione una tecla para continuar...")


if __name__ == "__main__":
    main()