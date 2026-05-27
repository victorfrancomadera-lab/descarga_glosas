import os
import re
import sys
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger

# =========================
# IDENTIDAD
# =========================
def datos_identidad():
    return {
        "sistema": "UNIFICAR CORTE DIGITAL",
        "propietario": "Salud-Net",
        "desarrollado_por": "DESARROLLO E INNOVACIÓN SALUD NET",
        "version": "v1.4",
        "licencia": "Uso interno autorizado",
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


def extraer_base_y_orden(nombre: str):
    """
    Extrae la clave de agrupación (base numérica) y el valor de orden de un nombre de archivo.

    Patrones soportados:
      109664          → base="109664", orden=(0, "")
      109664 (2)      → base="109664", orden=(2, "")
      109664OK        → base="109664", orden=(0, "OK")
      109664SERV      → base="109664", orden=(0, "SERV")
      109664 (2)OK    → base="109664", orden=(2, "OK")

    Retorna (base, orden_tuple) o (None, None) si no hay número base.
    """
    # Regex flexible:
    #   grupo 1 → número base obligatorio
    #   grupo 2 → sufijo alfabético ANTES del paréntesis (opcional)
    #   grupo 3 → número entre paréntesis (opcional)
    #   grupo 4 → sufijo alfabético DESPUÉS del paréntesis (opcional)
    match = re.match(
        r"^(\d+)"           # número base (obligatorio)
        r"([A-Za-z]*)"      # sufijo alfa antes del paréntesis (ej: SERV, OK)
        r"(?:\s*\((\d+)\))?" # número entre paréntesis opcional  (ej: (2))
        r"([A-Za-z]*)$",    # sufijo alfa después del paréntesis (opcional)
        nombre
    )
    if not match:
        return None, None

    base = match.group(1)

    # Prioridad de orden:
    # 1. número entre paréntesis  → orden numérico principal
    # 2. sufijo alfabético        → orden secundario alfabético
    num_paren = int(match.group(3)) if match.group(3) else 0
    sufijo    = (match.group(2) or "") + (match.group(4) or "")

    return base, (num_paren, sufijo.upper())


def main():
    mostrar_marca_agua()

    BASE_DIR = script_dir()
    ruta = BASE_DIR / "UNIFICAR - CORTE DIGITAL"

    print("Carpeta detectada:")
    print(f'"{ruta}"\n')

    if not ruta.exists():
        print("La carpeta no existe.")
        return

    archivos = [
        p for p in ruta.iterdir()
        if p.is_file() and p.suffix.lower() == ".pdf"
    ]

    grupos = {}  # { base_numerica: [(orden_tuple, Path), ...] }

    for archivo in archivos:
        nombre = archivo.stem
        base, orden = extraer_base_y_orden(nombre)

        if base is None:
            print(f"  [OMITIDO] '{archivo.name}' — nombre no reconocido.")
            continue

        grupos.setdefault(base, []).append((orden, archivo))

    if not grupos:
        print("No se encontraron archivos PDF válidos para unificar.")
        return

    procesados = 0

    for base, lista in grupos.items():

        # Si hay un único archivo, renombrarlo limpio si tiene sufijo
        if len(lista) == 1:
            archivo_unico = lista[0][1]
            salida_final  = ruta / f"{base}.pdf"

            # Comparar nombres en minúsculas para no fallar en Windows
            if archivo_unico.name.lower() == salida_final.name.lower():
                print(f"  [OK]         '{archivo_unico.name}' — ya tiene nombre limpio.")
                continue

            try:
                # Si ya existe el destino limpio, eliminarlo primero
                if salida_final.exists():
                    os.remove(salida_final)
                # shutil.move es más robusto que Path.rename en Windows
                shutil.move(str(archivo_unico), str(salida_final))
                print(f"  [RENOMBRADO] '{archivo_unico.name}' → '{salida_final.name}'")
                procesados += 1
            except Exception as e:
                print(f"  [ERROR]      No se pudo renombrar '{archivo_unico.name}': {e}")
            continue

        # Ordenar por (num_paréntesis, sufijo_alfa)
        # Criterio:
        #   • Sin sufijo y sin paréntesis  → primero  (0, "")
        #   • Con paréntesis               → por número ascendente
        #   • Con sufijo alfabético        → orden alfabético como desempate
        lista_ordenada = sorted(lista, key=lambda x: x[0])
        archivos_ordenados = [item[1] for item in lista_ordenada]

        print(f"\nUnificando base '{base}':")
        for idx, pdf in enumerate(archivos_ordenados, 1):
            print(f"   {idx}. {pdf.name}")

        salida_temp  = ruta / f"{base}_TEMP.pdf"
        salida_final = ruta / f"{base}.pdf"

        try:
            with PdfMerger() as merger:
                for pdf in archivos_ordenados:
                    merger.append(str(pdf))
                merger.write(str(salida_temp))

            # Eliminar archivos fuente
            for pdf in archivos_ordenados:
                if pdf.exists():
                    os.remove(pdf)

            # Reemplazar destino si ya existía
            if salida_final.exists():
                os.remove(salida_final)

            salida_temp.rename(salida_final)
            print(f"   ✔ Unificado → '{salida_final.name}'")
            procesados += 1

        except Exception as e:
            if salida_temp.exists():
                try:
                    os.remove(salida_temp)
                except Exception:
                    pass
            print(f"   ✘ Error al procesar '{base}': {e}")

    print(f"\n{'=' * 68}")
    print(f"Proceso terminado. Archivos procesados: {procesados}")
    print("=" * 68)


if __name__ == "__main__":
    main()
