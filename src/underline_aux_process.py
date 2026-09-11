import re
from os.path import join
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd


def normalizar(texto):
    """
    Elimina cualquier carácter que no sea alfanumérico y convierte a mayúsculas.
    Ejemplo: 'F-52344' -> 'F52344', 'F/52344' -> 'F52344'
    """
    return re.sub(r"[^a-zA-Z0-9]", "", str(texto)).upper()


def procesar_referencia(valor):
    valor_str = str(valor).strip().replace(" ", "")
    if "," in valor_str:
        primer_parte = valor_str.split(",")[0].strip()
    else:
        primer_parte = valor_str
    return primer_parte


def procesar_xlsx(archivo_excel):
    df = pd.read_excel(archivo_excel)
    df["REFERENCIA"] = df["REFERENCIA"].apply(procesar_referencia)

    # Referencias originales para el reporte final
    referencias_originales = df["REFERENCIA"].dropna().unique().tolist()

    # Mapeo de normalizada -> original para poder decir qué original se encontró
    # Ej: {'F52344': 'F-52344'}
    dict_mapeo = {normalizar(r): r for r in referencias_originales}
    referencias_limpias = list(dict_mapeo.keys())

    print(
        f"[+] Referencias procesadas: {referencias_originales[:10]}... (total: {len(referencias_originales)})"
    )

    return referencias_originales, referencias_limpias, dict_mapeo


def process_pdf(pdf_path, pdf_salida, referencias_limpias, dict_mapeo):
    doc = fitz.open(pdf_path)
    found_references = set()  # Guardaremos las originales aquí

    for pagina_num in range(len(doc)):
        pagina = doc[pagina_num]
        print(f"[+] Procesando página {pagina_num + 1} de {len(doc)}...")

        dict_texto = pagina.get_text("dict")
        lineas_agrupadas = {}

        for block in dict_texto["blocks"]:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                y_coord = round(line["bbox"][1], 0)
                if y_coord not in lineas_agrupadas:
                    lineas_agrupadas[y_coord] = []
                lineas_agrupadas[y_coord].append(line)

        for y in sorted(lineas_agrupadas.keys()):
            lineas_de_esta_fila = lineas_agrupadas[y]
            texto_fila_completa = ""
            y0_min, y1_max = 9999, 0

            for l in lineas_de_esta_fila:
                y0_min = min(y0_min, l["bbox"][1])
                y1_max = max(y1_max, l["bbox"][3])
                for span in l["spans"]:
                    texto_fila_completa += span["text"] + " "

            texto_fila_completa = texto_fila_completa.strip()

            # --- LÓGICA DE NORMALIZACIÓN ---
            texto_fila_norm = normalizar(texto_fila_completa)

            # Buscamos coincidencias usando las versiones limpias
            coincidencias_encontradas_norm = [
                ref_norm
                for ref_norm in referencias_limpias
                if ref_norm in texto_fila_norm
            ]

            if coincidencias_encontradas_norm:
                texto_lower = texto_fila_completa.lower()
                es_diario = "diario" in texto_lower
                es_ingreso_egreso = any(k in texto_lower for k in ["ingreso", "egreso"])

                if es_diario and not es_ingreso_egreso:
                    full_row_rect = fitz.Rect(
                        0, y0_min - 1, pagina.rect.width, y1_max + 1
                    )
                    highlight = pagina.add_highlight_annot(full_row_rect)
                    highlight.set_colors(stroke=(1, 1, 0))
                    highlight.update()

                    # Guardamos las referencias ORIGINALES encontradas para el reporte
                    for c_norm in coincidencias_encontradas_norm:
                        ref_original = dict_mapeo[c_norm]
                        found_references.add(ref_original)

                    print(
                        f"    [✓] Renglón DIARIO detectado y remarcado: '{texto_fila_completa[:60]}...'"
                    )

    # Guardar PDF
    doc.save(pdf_salida)
    doc.close()
    print(f"[+] PDF guardado en: {pdf_salida}")

    return found_references


def proceso_estadistico(referencias_originales, found_references):
    # === 3. Estadístico (Igual al original pero con lógica de normalización) ===
    total = len(referencias_originales)
    encontradas = len(found_references)
    no_encontradas = total - encontradas

    print("\n" + "=" * 55)
    print("ESTADÍSTICO DE REFERENCIAS")
    print("=" * 55)
    print(f"Total de referencias procesadas       : {total}")
    print(f"Referencias encontradas y remarcadas  : {encontradas}")
    print(f"Referencias NO encontradas            : {no_encontradas}")
    print("=" * 55)

    if no_encontradas > 0:
        # Identificar cuáles de las originales no están en el set de encontradas
        no_encontradas_lista = [
            r for r in referencias_originales if r not in found_references
        ]
        print(f"Muestra de no encontradas: {no_encontradas_lista[:50]}")


def main_underline_aux_process(self, path, suff):
    path = Path(path).resolve()
    # === 1. Procesar el Excel ===
    archivo_excel = path / f"{suff}_referencias.xlsx"

    if not archivo_excel.exists():
        msg = f"No se encontró el archivo necesario: {archivo_excel}"
        self.text_console_log(msg, "ERROR")
        # Este RAISE detiene la función AQUÍ MISMO y va directo al except de la GUI
        raise FileNotFoundError(msg)

    # Paso 1, obtener la información del auxiliar procesando el xlsx
    referencias_originales, referencias_limpias, dict_mapeo = procesar_xlsx(
        archivo_excel
    )

    # === 2. Buscar en PDF agrupando por filas virtuales ===
    pdf_name = "aux"
    pdf_path = path / f"{pdf_name}.pdf"
    pdf_salida = path / f"{pdf_name}_con_subrayados.pdf"

    if not pdf_path.exists():
        msg = f"No se encontró el archivo necesario: {pdf_path}"
        self.text_console_log(msg, "ERROR")
        # Este RAISE detiene la función AQUÍ MISMO y va directo al except de la GUI
        raise FileNotFoundError(msg)

    # Paso 2,
    found_references = process_pdf(
        pdf_path, pdf_salida, referencias_limpias, dict_mapeo
    )

    # PAso 3,
    proceso_estadistico(referencias_originales, found_references)
