import glob
import io
import os
import re
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
from pdf2image import convert_from_path
from PIL import Image, ImageFilter

nombres_hojas = [
    "Trazabilidad Materia Prima",
    "Factura",
    "Orden de compra",
    "Espectómetro",
    "Análisis Químico",
    "Inspección de Componentes",
    "Almacenario",
    "Proceso",
    "Venta Final",
]

mapeo_documentos_simple = {
    "Almacenario": "Almacenario",
    "F": "Factura",
    "OC": "Orden de compra",
    "Analisis Quimico": "Análisis Químico",
    "Analisis Quimicos": "Análisis Químico",
    "IC": "Inspección de Componentes",
}

ABREV_MAP = {
    "AQ": "Análisis Químico",
    "IC": "Inspección de Componentes",
    "E": "Espectómetro",
    "A": "Almacenario",
    "OC": "Orden de compra",
    "F": "Factura",
}

# ==============================================
# FUNCIONES DE PROCESAMIENTO DE IMAGEN
# ==============================================


def trim_white_space(img):
    """Mantiene la lógica original para documentos PDF (que suelen ser fondo blanco)"""
    if img.mode != "RGB":
        img = img.convert("RGB")
    threshold = 245
    mask = Image.new("L", img.size, 0)
    pixels = img.load()
    mask_pixels = mask.load()
    for y in range(img.height):
        for x in range(img.width):
            r, g, b = pixels[x, y]
            if r < threshold or g < threshold or b < threshold:
                mask_pixels[x, y] = 255
    mask = mask.filter(ImageFilter.MaxFilter(3))
    bbox = mask.getbbox()
    return img.crop(bbox) if bbox else img


def comprimir_imagen(ruta_imagen, calidad=75, max_ancho=1000):
    """Abre una imagen y decide si mantener PNG (transparencia) o convertir a JPEG."""
    img = Image.open(ruta_imagen)
    if img.width > max_ancho:
        ratio = max_ancho / img.width
        img = img.resize((max_ancho, int(img.height * ratio)), Image.Resampling.LANCZOS)

    img_comprimida = io.BytesIO()

    # --- MEJORA AQUÍ: DETECCIÓN DE TRANSPARENCIA ---
    # Si la imagen tiene canal alfa (RGBA) o es paleta con transparencia (P)
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        # Guardamos como PNG para preservar la transparencia y evitar el fondo negro
        img.save(img_comprimida, format="PNG", optimize=True)
    else:
        # Si es una imagen opaca (como las capturas de PDF), usamos JPEG para comprimir peso
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.save(img_comprimida, format="JPEG", quality=calidad, optimize=True)

    img_comprimida.seek(0)
    return img_comprimida


# ==============================================
# LOGICA DE DETECCIÓN (Se mantiene idéntica)
# ==============================================


def detectar_tipo_documento_simple(nombre_pdf):
    nombre_lower = nombre_pdf.lower()
    if "analisis quimico" in nombre_lower:
        return "Analisis Quimico"
    if "analisis quimicos" in nombre_lower:
        return "Analisis Quimicos"
    if "inspeccion de componentes" in nombre_lower:
        return "Inspeccion de Componentes"
    if "solicitud de inspeccion" in nombre_lower:
        return "Solicitud de Inspeccion"
    if nombre_lower.startswith("almacenario"):
        return "Almacenario"
    if any(nombre_lower.startswith(prefix) for prefix in ["f-", "f.", "f "]):
        return "F"
    if nombre_lower.startswith("oc"):
        return "OC"
    return None


def analizar_instrucciones_complejas(nombre_pdf):
    instrucciones = []
    match_poliza = re.search(r"(?:F|OC|f|oc)[- ]?(\d+)", nombre_pdf)
    poliza = match_poliza.group(1) if match_poliza else None

    if nombre_pdf.upper().startswith("F"):
        instrucciones.append((1, "Factura"))

    matches = re.finditer(r"(\d+)([A-Z]+)|([A-Z]+)[- ]?(\d+)", nombre_pdf.upper())
    for m in matches:
        num_pag = int(m.group(1)) if m.group(1) else int(m.group(4))
        abrev = m.group(2) if m.group(2) else m.group(3)

        if abrev == "NC":
            instrucciones.append((num_pag, "DISCARD"))
        elif abrev in ABREV_MAP:
            tipo_doc = ABREV_MAP[abrev]
            if not (num_pag == 1 and tipo_doc == "Factura"):
                instrucciones.append((num_pag, tipo_doc))
    return poliza, instrucciones


def crear_unificar_pdf_xlsx(
    pdf_dirs, dpi, assets_dir, img_process_path, excel_output_dir
):
    # ==============================================
    # PROCESAMIENTO PRINCIPAL DE PDFs (Se mantiene idéntica)
    # ==============================================
    print("🔄 Iniciando procesamiento unificado de PDFs...\n")

    imagenes_por_poliza = {}

    for pdf_dir in pdf_dirs:
        pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))

        for pdf_path in pdf_files:
            nombre_pdf = os.path.basename(pdf_path).replace(".pdf", "")
            print(f"📄 Procesando: {nombre_pdf}.pdf")

            es_complejo = (
                "_" in nombre_pdf
                or "NC" in nombre_pdf.upper()
                or re.search(r"\d+[A-Z]{1,2}", nombre_pdf.upper())
            )

            try:
                pages = convert_from_path(pdf_path, dpi=dpi)

                if es_complejo:
                    poliza, instrucciones = analizar_instrucciones_complejas(nombre_pdf)
                    if poliza:
                        if poliza not in imagenes_por_poliza:
                            imagenes_por_poliza[poliza] = {}
                        for num_pag, tipo_doc in instrucciones:
                            if tipo_doc == "DISCARD":
                                continue
                            idx = num_pag - 1
                            if 0 <= idx < len(pages):
                                output_path = os.path.join(
                                    assets_dir, f"tmp_{poliza}_{tipo_doc}.png"
                                )
                                trim_white_space(pages[idx]).save(output_path, "PNG")
                                imagenes_por_poliza[poliza][tipo_doc] = output_path
                else:
                    tipo_doc_key = detectar_tipo_documento_simple(nombre_pdf)
                    match_poliza = re.search(r"(\d+)", nombre_pdf)
                    if tipo_doc_key and match_poliza:
                        poliza = match_poliza.group(1)
                        nombre_hoja_final = mapeo_documentos_simple[tipo_doc_key]
                        if poliza not in imagenes_por_poliza:
                            imagenes_por_poliza[poliza] = {}
                        output_path = os.path.join(
                            assets_dir, f"tmp_{poliza}_{tipo_doc_key}.png"
                        )
                        trim_white_space(pages[0]).save(output_path, "PNG")
                        imagenes_por_poliza[poliza][nombre_hoja_final] = output_path

            except Exception as e:
                print(f"   ❌ Error procesando PDF: {e}")

    # ==============================================
    # CREACIÓN DE EXCEL (Ajustada para respetar PNG)
    # ==============================================
    print("\n📋 Generando archivos Excel...")

    for poliza, docs in imagenes_por_poliza.items():
        print(f"🔧 Creando Excel póliza {poliza}...")
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        for nombre_hoja in nombres_hojas:
            ws = wb.create_sheet(title=nombre_hoja)
            ws.sheet_view.showGridLines = False
            ws.column_dimensions["A"].width = 25

            ruta_a_insertar = None

            if nombre_hoja == "Proceso":
                if os.path.exists(img_process_path):
                    ruta_a_insertar = img_process_path
                    print(f"   ⭐ Insertando imagen de Proceso (PNG)")
                else:
                    print(f"   ⚠️ No se encontró la imagen en: {img_process_path}")
            elif nombre_hoja in docs:
                ruta_a_insertar = docs[nombre_hoja]
                print(f"   ✅ Insertando documento: {nombre_hoja}")

            if ruta_a_insertar:
                try:
                    img_data = comprimir_imagen(ruta_a_insertar)

                    # Guardamos temporalmente para openpyxl
                    # Usamos extensión .png para asegurar compatibilidad con transparencia
                    temp_file_path = os.path.join(
                        assets_dir, f"final_tmp_{poliza}_{nombre_hoja}.png"
                    )

                    with open(temp_file_path, "wb") as f:
                        f.write(img_data.getvalue())

                    img_xl = XLImage(temp_file_path)
                    if nombre_hoja != "Proceso":
                        img_xl.width, img_xl.height = 550, 450
                    ws.add_image(img_xl, "A1")
                except Exception as e:
                    print(f"   ❌ Error insertando imagen en {nombre_hoja}: {e}")

        wb.save(
            os.path.join(excel_output_dir, f"{poliza}_Plantilla_Control_Calidad.xlsx")
        )

    # ==============================================
    # LIMPIEZA
    # ==============================================
    print("\n🧹 Limpiando archivos temporales...")
    archivos_tmp = glob.glob(os.path.join(assets_dir, "tmp_*.*")) + glob.glob(
        os.path.join(assets_dir, "final_tmp_*.*")
    )
    for f in archivos_tmp:
        try:
            os.remove(f)
        except:
            pass

    print("🎉 Proceso completado exitosamente.")


def conciliar_y_actualizar(
    directorio_facturas, facturas_folios, indice_liberacion, sheet_name
):

    cabecera_trazabilidad = [
        "Folio de Solicitud",
        "Lote Proveedor",
        "#lotes",
        "Aleación",
        "Proveedor",
        "Peso (Kg)",
        "Hora de Recepción",
        "Fecha de Liberación",
        "Mes de Liberación",
        "Inspector",
        "Estatus de Liberación",
        "Escaneado",
        "Observaciones",
    ]

    try:
        # 1. Cargamos el objeto Excel para inspeccionar las hojas
        xl = pd.ExcelFile(facturas_folios)

        # 2. Validamos si el nombre existe en la lista de hojas
        # xl.sheet_names devuelve una lista con los nombres de todas las hojas
        if sheet_name in xl.sheet_names:
            hoja_a_cargar = sheet_name
        else:
            # Si no existe, usamos 0 para que cargue la primera hoja (la activa por defecto)
            hoja_a_cargar = 0
            print(f"La hoja '{sheet_name}' no existe. Cargando hoja por defecto.")

        # 3. Leemos el DataFrame usando la variable definida
        df_a = pd.read_excel(xl, usecols=[0, 1], sheet_name=hoja_a_cargar)
        df_b = pd.read_excel(indice_liberacion)

        def limpiar_folio(serie):
            return (
                serie.astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
                .replace("nan", "")
            )

        df_a["folio_clean"] = limpiar_folio(df_a.iloc[:, 0])
        df_b["folio_clean"] = limpiar_folio(df_b.iloc[:, 0])

        print(f"--- Iniciando Búsqueda y Actualización en {directorio_facturas} ---")

        for _, fila_a in df_a.iterrows():
            folio_buscado = fila_a["folio_clean"]
            factura_nombre = str(fila_a.iloc[1]).strip()

            if folio_buscado == "" or folio_buscado == "None":
                continue

            coincidencia_b = df_b[df_b["folio_clean"] == folio_buscado]

            if not coincidencia_b.empty:
                fila_datos_b = coincidencia_b.iloc[0]
                datos_para_insertar = [
                    fila_datos_b.get(col, "") for col in cabecera_trazabilidad
                ]

                if not os.path.exists(directorio_facturas):
                    print(f"❌ El directorio no existe: {directorio_facturas}")
                    return

                archivo_encontrado = None
                for f in os.listdir(directorio_facturas):
                    if f.upper().startswith(factura_nombre.upper()) and f.endswith(
                        ".xlsx"
                    ):
                        archivo_encontrado = os.path.join(directorio_facturas, f)
                        break

                if archivo_encontrado:
                    actualizar_excel_seguro(
                        archivo_encontrado, cabecera_trazabilidad, datos_para_insertar
                    )
                else:
                    print(f"   ⚠️ No se halló archivo para factura: {factura_nombre}")

        print(f"\n--- Proceso Finalizado ---")

    except Exception as e:
        print(f"❌ Error crítico en el flujo principal: {e}")


def actualizar_excel_seguro(ruta_archivo, cabecera, datos):
    """
    Actualiza el Excel iniciando en A2, sin dejar filas vacías,
    aplicando color azul a la cabecera y formato de fecha en col H.
    """
    nombre_hoja = "Trazabilidad Materia Prima"
    folio_nuevo = str(datos[0]).strip()

    relleno_azul = PatternFill(
        start_color="0070C0", end_color="0070C0", fill_type="solid"
    )
    fuente_blanca = Font(color="FFFFFF", bold=True)

    try:
        wb = load_workbook(ruta_archivo)

        if nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
        else:
            ws = wb.create_sheet(nombre_hoja)

        # --- Lógica de Cabecera (Fila 1) ---
        if ws.cell(row=1, column=1).value is None:
            for col_idx, column_title in enumerate(cabecera, 1):
                celda = ws.cell(row=1, column=col_idx, value=column_title)
                celda.fill = relleno_azul
                celda.font = fuente_blanca
        else:
            # Asegurar estilo aunque ya exista
            for col_idx in range(1, len(cabecera) + 1):
                ws.cell(row=1, column=col_idx).fill = relleno_azul
                ws.cell(row=1, column=col_idx).font = fuente_blanca

        # --- Prevención de Duplicados ---
        existe = False
        for row in range(1, ws.max_row + 1):
            valor_celda = (
                str(ws.cell(row=row, column=1).value).strip().replace(".0", "")
            )
            if valor_celda == folio_nuevo:
                existe = True
                break

        if not existe:
            # --- Buscar la primera fila vacía real (iniciando en fila 2) ---
            proxima_fila = 2
            while ws.cell(row=proxima_fila, column=1).value is not None:
                proxima_fila += 1

            # Insertar los datos en la fila encontrada
            for col_idx, valor in enumerate(datos, 1):
                ws.cell(row=proxima_fila, column=col_idx, value=valor)

            # --- Formato de Fecha Corta en Columna H (8) ---
            ws.cell(row=proxima_fila, column=8).number_format = "DD/MM/YYYY"

            wb.save(ruta_archivo)
            print(
                f"   ✅ Folio {folio_nuevo} añadido en fila {proxima_fila} de {os.path.basename(ruta_archivo)}"
            )
        else:
            print(f"   ℹ️ El folio {folio_nuevo} ya existe. Omitiendo.")

        wb.close()

    except PermissionError:
        print(f"   ❌ Error: Archivo abierto {os.path.basename(ruta_archivo)}")
    except Exception as e:
        print(f"   ❌ Error al actualizar {os.path.basename(ruta_archivo)}: {e}")


def main_flujograma_process(self, path):
    name = path.split("\\") if "/" not in path else path.split("/")
    name = name[-1]
    path = Path(path).resolve()

    assets_dir = f"{path}/assets"
    excel_output_dir = path / name
    img_process_path = (
        path / "proceso.png"  # Imagen fija para la hoja 'Proceso'
    )

    img_process_path = (
        img_process_path
        if img_process_path.exists()
        else "icons/flujograma_process.png"
    )

    facturas_folios = path / "facturas_folios.xlsx"
    indice_liberacion = path / "Indice_Liberacion_Lingote_2025.xlsx"

    # ============================================================
    # VALIDACIÓN PREVIA
    # ============================================================
    if not path.exists() or not path.is_dir():
        msg = f"La ruta no existe o no es un directorio: {path}"
        if hasattr(self, "text_console_log"):
            self.text_console_log(msg, "ERROR")
        raise FileNotFoundError(msg)

    # Buscar al menos un PDF (cualquier nombre)
    pdf_files = list(path.glob("*.pdf"))

    if not pdf_files:
        msg = "No se ha encontrado archivo para procesar"
        if hasattr(self, "text_console_log"):
            self.text_console_log(msg, "ERROR")
        raise FileNotFoundError(msg)

    # ============================================================
    # VALIDACIÓN PREVIA (Aquí detona y sale antes de intentar nada)
    # Si no existen los insumos, lanzamos el error de inmediato
    # ============================================================
    if not facturas_folios.exists() or not indice_liberacion.exists():
        faltante = (
            "facturas_folios.xlsx"
            if not facturas_folios.exists()
            else "Indice_Liberacion_Lingote_2025.xlsx"
        )
        msg = f"No se encontró el archivo necesario: {faltante}"
        self.text_console_log(msg, "ERROR")
        # Este RAISE detiene la función AQUÍ MISMO y va directo al except de la GUI
        raise FileNotFoundError(msg)

    dpi = 150

    os.makedirs(assets_dir, exist_ok=True)
    # os.makedirs(excel_output_dir, exist_ok=True)

    crear_unificar_pdf_xlsx([path], dpi, assets_dir, img_process_path, path)

    conciliar_y_actualizar(path, facturas_folios, indice_liberacion, "")
