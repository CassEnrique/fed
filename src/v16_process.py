import argparse
import copy
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DIR_ENCABEZADO = {
    "A": "Poliza",
    "B": "Número",
    "C": "Fecha factura",
    "D": "Cliente",
    "E": "RFC",
    "F": "Descripcion",
    "G": "Referencia",
    "H": "Base Gravable MXN",
    "I": "IVA 16% MXN",
    "J": "Retencion de iva 16% MXN",
    "K": "TOTAL MXN",
    "L": "TIPO CAMBIO",
    "M": "Base Gravable USD",
    "N": "IVA 16% USD",
    "O": "Retencion IVA 16% USD",
    "P": "Total USD",
    "Q": "Fecha de cobro",
    "R": "Forma de pago",
    "S": "Cobranza referenciada estado de cuenta",
}


def obtener_indices(headers_buscados, diccionario, incremento=0):
    """
    Devuelve:
        - Un entero (el índice) si solo se busca un encabezado.
        - Una lista de índices si se buscan varios encabezados.
    """
    # Diccionario inverso: valor → índice (más eficiente)
    valor_a_indice = {
        valor: indice + incremento for indice, valor in enumerate(diccionario.values())
    }

    # Obtener resultados
    resultado = [valor_a_indice[h] for h in headers_buscados if h in valor_a_indice]
    # Si solo hay un resultado → devolver el número directamente
    if len(resultado) == 1:
        return resultado[0]

    return resultado


def argumentos_consola():
    # ====================== ARGUMENTOS DESDE CONSOLA ======================
    parser = argparse.ArgumentParser(description="Filtro de egresos vs proveedores")
    parser.add_argument("anio", type=int, help="Año (ejemplo: 2025)")
    parser.add_argument("mes", help="Mes en formato abreviado (ej: sep, oct, nov)")
    args = parser.parse_args()

    ANIO = args.anio  # ← Variable 1
    MES = args.mes.lower()  # ← Variable 2

    print(f"🔹 Año recibido: {ANIO}")
    print(f"🔹 Mes recibido: {MES}")

    return ANIO, MES


def limpiar_concepto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto)
    for palabra in ["DEPOSITO", "AJUSTE", "PAGO"]:
        texto = re.sub(rf"\b{palabra}\b", "", texto, flags=re.IGNORECASE)
    return texto.strip()


def formatear_excel_sin_diarios(archivo_entrada, archivo_salida):
    """
    Formatea Excel omitiendo registros con tipo 'Diario'
    e inserta el nombre del cliente en la columna D
    Omite la fila ['Cuenta', 'Nombre', 'Saldo Inicial']
    Elimina columna Saldo y añade nuevas columnas
    """

    # Cargar workbook original
    wb_original = load_workbook(archivo_entrada)
    ws_original = wb_original.active

    # Crear nuevo workbook
    wb_nuevo = Workbook()
    ws_nuevo = wb_nuevo.active
    ws_nuevo.title = "Datos"

    # Definir encabezados con los nuevos nombres
    encabezados = list(DIR_ENCABEZADO.values())

    # Escribir encabezados en el nuevo archivo
    for col, encabezado in enumerate(encabezados, 1):
        celda = ws_nuevo.cell(row=1, column=col)
        celda.value = encabezado
        celda.font = Font(bold=True, color="FFFFFF", size=11)
        celda.fill = PatternFill(
            start_color="366092", end_color="366092", fill_type="solid"
        )
        celda.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Estilos
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Variables para rastrear el cliente actual
    cliente_actual = ""
    fila_nueva = 2

    # Leer datos del archivo original
    for row_idx, row in enumerate(
        ws_original.iter_rows(
            min_row=1, max_row=ws_original.max_row, values_only=False
        ),
        1,
    ):
        # Obtener valores de las celdas
        valores = [celda.value for celda in row]

        # **OMITIR la fila ['Cuenta', 'Nombre'] - eliminando espacios en blanco**
        if valores[0] and valores[1]:
            val0_clean = str(valores[0]).replace(" ", "").strip().lower()
            val1_clean = str(valores[1]).replace(" ", "").strip().lower()
            if val0_clean == "cuenta" and val1_clean == "nombre":
                print(f"Fila {row_idx} omitida (Cuenta/Nombre)")
                continue

        # Identificar encabezados de cliente (tienen formato específico)
        # Detectar si esta fila contiene el nombre del cliente (patrón: Número-Cliente | Nombre)
        if (
            valores[0]
            and valores[1]
            and "-" in str(valores[0])
            and isinstance(valores[1], str)
        ):
            # Verificar si es un encabezado de cliente
            if not valores[2] and not valores[3]:  # Si las columnas 3 y 4 están vacías
                cliente_actual = valores[1].strip()
                print(f"Cliente detectado: {cliente_actual}")
                continue

        # Saltar la fila de encabezados de columnas ("Fecha", "Tipo", etc.)
        if valores[0] == "Fecha" or not valores[0]:
            continue

        # Validar que sea una fila de datos válida
        if not valores[0] or not valores[1]:
            continue

        # **FILTRO: Omitir registros con tipo 'Diario'**
        if valores[1] and str(valores[1]).strip().lower() == "diario":
            print(f"Fila {row_idx} omitida (Diario)")
            continue

        # Insertar datos en el nuevo archivo
        nuevos_valores = [
            valores[1],  # A - Poliza
            valores[2],  # B - Número
            "",  # C - Fecha factura
            limpiar_concepto(valores[3]),  # D - Cliente
            "",  # E - RFC
            "",  # F - Descripcion
            valores[4],  # G - Referencia
            f"=I{fila_nueva}/0.16",  # H - Base Gravable MXN
            valores[6],  # I - IVA 16% MXN
            "",  # J - Retencion de iva 16% MXN
            f"=IF(H{fila_nueva}=0, 0, SUM(H{fila_nueva} + I{fila_nueva} - J{fila_nueva}))",  # K - TOTAL MXN
            "",  # L - TIPO CAMBIO
            f'=IF(P{fila_nueva}="", "", IF(J{fila_nueva}="", P{fila_nueva} / 1.16, P{fila_nueva}))',  # M - Base Gravable USD
            f'=IF(M{fila_nueva}="", "", M{fila_nueva}  * 0.16)',  # N - IVA 16% USD
            f'=IF(N{fila_nueva}="", "", N{fila_nueva})',  # O - Retencion IVA 16% USD
            "",  # P - Total USD
            valores[0],  # Q - Fecha de cobro
            "TRANSFERENCIA BANCARIA",  # R - Forma de pago
            "",  # S - Cobranza referenciada estado de cuenta
        ]

        # Escribir fila en el nuevo archivo
        for col, valor in enumerate(nuevos_valores, 1):
            celda = ws_nuevo.cell(row=fila_nueva, column=col)
            celda.value = valor
            # celda.border = thin_border
            celda.alignment = Alignment(horizontal="left", vertical="center")

            # Alineación para columnas numéricas
            print("#%#" * 30, col)
            if col in [
                9,
                10,
                11,
                12,
                13,
                14,
                15,
                16,
                17,
            ]:  # Número, Cargos, Abonos, etc.
                celda.alignment = Alignment(horizontal="right", vertical="center")
                # Formato de números
                if isinstance(valor, (int, float)):
                    celda.number_format = "$#,##0.00"

        fila_nueva += 1

    # Congelar encabezado
    ws_nuevo.freeze_panes = "A2"

    # Ajustar anchos de columnas
    anchos = {
        "A": 10,  # Poliza
        "B": 20,  # Número
        "C": 20,  # Fecha factura
        "D": 20,  # Cliente
        "E": 20,  # RFC
        "F": 20,  # Descripcion
        "G": 20,  # Referencia
        "H": 20,  # Base Gravable MXN
        "I": 20,  # IVA 16% MXN
        "J": 30,  # Retencion de iva 16% MXN
        "K": 20,  # TOTAL MXN
        "L": 20,  # TIPO CAMBIO
        "M": 20,  # Base Gravable USD
        "N": 20,  # IVA 16% USD
        "O": 30,  # Retencion IVA 16% USD
        "P": 20,  # Total USD
        "Q": 20,  # Fecha de cobro
        "R": 30,  # Forma de pago
        "S": 40,  # Cobranza referenciada estado de cuenta
    }
    for col, ancho in anchos.items():
        ws_nuevo.column_dimensions[col].width = ancho

    # Guardar archivo
    ws_nuevo.sheet_view.showGridLines = False
    wb_nuevo.save(archivo_salida)
    print(f"\n✓ Archivo formateado guardado en: {archivo_salida}")
    print(f"✓ Total de registros procesados: {fila_nueva - 2}")
    print(f"✓ Registros de tipo 'Diario' omitidos")
    print(f"✓ Fila 'Cuenta/Nombre/Saldo Inicial' omitida")
    print(f"✓ Columna 'Saldo' eliminada")
    print(f"✓ Columna 'Descripcion' completada: 'Venta Productos de Aluminio'")


def actualizar_archivo_con_retenido(archivo_rete, archivo_formateado):
    """
    Procesa agosto_usd.xlsx y actualiza archivo_formateado.xlsx
    sin crear un nuevo archivo.
    Busca coincidencias por Referencia y actualiza columnas M y N.
    """

    # Cargar agosto_usd.xlsx
    wb_usd = load_workbook(archivo_rete)
    ws_usd = wb_usd.active

    # Diccionario para almacenar datos de USD (Referencia -> {Cargos, Abonos})
    datos_usd = {}

    cliente_actual = ""

    # Leer datos del archivo USD
    print(f"\n--- Procesando {archivo_rete} ---")
    for row_idx, row in enumerate(
        ws_usd.iter_rows(min_row=1, max_row=ws_usd.max_row, values_only=False),
        1,
    ):
        valores = [celda.value for celda in row]

        # **OMITIR la fila ['Cuenta', 'Nombre']**
        if valores[0] and valores[1]:
            val0_clean = str(valores[0]).replace(" ", "").strip().lower()
            val1_clean = str(valores[1]).replace(" ", "").strip().lower()
            if val0_clean == "cuenta" and val1_clean == "nombre":
                continue

        # Detectar cliente
        if (
            valores[0]
            and valores[1]
            and "-" in str(valores[0])
            and isinstance(valores[1], str)
        ):
            if not valores[2] and not valores[3]:
                cliente_actual = valores[1].strip()
                print(f"Cliente USD detectado: {cliente_actual}")
                continue

        # Saltar encabezados
        if valores[0] == "Fecha" or not valores[0]:
            continue

        # Validar fila de datos
        if not valores[0] or not valores[1]:
            continue

        # **FILTRO: Omitir Diarios**
        if valores[1] and str(valores[1]).strip().lower() == "diario":
            print(f"Fila {row_idx} USD omitida (Diario)")
            continue

        # Almacenar datos: Referencia -> {Cargos (col F=5), Abonos (col G=6)}
        referencia = valores[4] if len(valores) > 4 else ""  # Columna F (Referencia)
        cargos = valores[5] if len(valores) > 5 else ""  # Columna G (Cargos)

        print("#" * 30, referencia)

        if referencia:
            referencia_clean = re.sub(r"\s+", "", str(referencia)).upper()
            # referencia_clean = str(referencia).strip().replace(" ", "").upper()
            datos_usd[referencia_clean] = {
                "cargos": cargos,
            }
            print(f"Referencia USD: {referencia_clean} - Cargos: {cargos}")

    # Cargar archivo_formateado.xlsx para actualizar
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    registros_actualizados = 0

    # Iterar sobre archivo_formateado.xlsx
    print(f"\n--- Actualizando {archivo_formateado} ---")
    for row_idx in range(2, ws_formateado.max_row + 1):  # Comenzar desde fila 2
        # Columna H = Referencia/ Factura/ PL
        referencia_col = ws_formateado.cell(row=row_idx, column=7).value

        if referencia_col:
            referencia_clean = re.sub(r"\s+", "", str(referencia_col)).upper()
            # referencia_clean = str(referencia_col).strip().replace(" ", "").upper()

            # Buscar en datos_usd
            if referencia_clean in datos_usd:
                # Actualizar columna J (Retencion de iva 16% MXN) con Cargos
                ws_formateado.cell(row=row_idx, column=10).value = datos_usd[
                    referencia_clean
                ]["cargos"]

                # Aplicar formato a los valores numéricos
                for col in [
                    9,
                    10,
                    11,
                    12,
                    13,
                    14,
                    15,
                    16,
                    17,
                ]:
                    celda = ws_formateado.cell(row=row_idx, column=col)
                    if isinstance(celda.value, (int, float)):
                        celda.number_format = "#,##0.00"

                registros_actualizados += 1
                print(f"Fila {row_idx}: Referencia {referencia_clean} actualizada")

    # Guardar archivo_formateado.xlsx SIN crear uno nuevo
    wb_formateado.save(archivo_formateado)
    print(f"\n✓ Archivo {archivo_formateado} actualizado sin crear nuevo archivo")
    print(f"✓ Total de registros actualizados: {registros_actualizados}")
    print(f"✓ Registros de tipo 'Diario' omitidos")
    print(f"✓ Fila 'Cuenta/Nombre/Saldo Inicial' omitida")
    print(f"✓ Columna 'Saldo' eliminada")


def agrupar_por_poliza(archivo_formateado):
    """
    Agrupa registros por póliza (Número en columna C)
    Añade border-bottom 1.75pt negro al final de cada grupo
    Inserta fórmula de suma en columna N (Importe Cobrado USD) para cada grupo
    Imprime resumen de sumas por consola
    """

    # Cargar archivo_formateado.xlsx
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    # Diccionario para agrupar filas por póliza
    grupos_poliza = {}
    resumen_sumas = []
    numero_grupo = 1

    # Recorrer todas las filas de datos (comenzando desde fila 2)
    for row_idx in range(2, ws_formateado.max_row + 1):
        numero_poliza = ws_formateado.cell(row=row_idx, column=2).value  # Columna C

        if numero_poliza:
            numero_poliza_str = str(numero_poliza).strip()
            if numero_poliza_str not in grupos_poliza:
                grupos_poliza[numero_poliza_str] = []
            grupos_poliza[numero_poliza_str].append(row_idx)

    print(f"\n--- Agrupando por Póliza ---")
    print(f"Total de grupos de pólizas encontrados: {len(grupos_poliza)}")

    # Border grueso para separación
    thick_border = Border(
        # left=Side(style="thin"),
        # right=Side(style="thin"),
        # top=Side(style="thin"),
        bottom=Side(style="thick", color="000000"),  # 1.75 pt negro
    )

    # Procesar cada grupo de póliza
    for poliza, filas in grupos_poliza.items():
        primera_fila = filas[0]
        ultima_fila = filas[-1]

        # Rango de suma: M[primera_fila]:M[ultima_fila]
        rango_suma = f"M{primera_fila}:M{ultima_fila}"
        formula = f"=SUM({rango_suma})"

        # Colocar la fórmula en la última fila, columna N (col 14)
        # celda_formula = ws_formateado.cell(row=ultima_fila, column=17)
        # celda_formula.value = formula
        # celda_formula.number_format = "#,##0.00"
        # celda_formula.alignment = Alignment(horizontal="right", vertical="center")

        # Aplicar border-bottom grueso a toda la última fila del grupo
        for col in range(1, 20):  # 24 columnas
            celda = ws_formateado.cell(row=ultima_fila, column=col)
            celda.border = thick_border

        # Calcular suma para consola
        suma_grupo = 0
        for fila in filas:
            valor = ws_formateado.cell(row=fila, column=11).value  # Columna M
            if valor and isinstance(valor, (int, float)):
                suma_grupo += valor

        resumen_sumas.append(
            {
                "grupo": numero_grupo,
                "poliza": poliza,
                "suma": suma_grupo,
                "filas": f"{primera_fila}:{ultima_fila}",
            }
        )

        print(f"\nGrupo D{numero_grupo}:")
        print(f"  Póliza: {poliza}")
        print(f"  Filas: {primera_fila} a {ultima_fila}")
        print(f"  Suma (Cobranza Recibida USD): ${suma_grupo:,.2f}")
        print(f"  Fórmula en N{ultima_fila}: {formula}")

        columna_referencia_cobranza_edo = obtener_indices(
            ["Cobranza referenciada estado de cuenta"], DIR_ENCABEZADO, 1
        )
        celda_ref_edo = ws_formateado.cell(
            row=ultima_fila, column=columna_referencia_cobranza_edo
        )
        celda_ref_edo.value = f"I.{numero_grupo}"
        celda_ref_edo.font = Font(
            bold=True, color="DC0042", size=12
        )  # Negritas y color rojo
        celda_ref_edo.alignment = Alignment(horizontal="center", vertical="center")

        numero_grupo += 1

    # Imprimir resumen consolidado
    print(f"\n{'=' * 70}")
    print(f"RESUMEN DE PÓLIZAS Y SUMAS")
    print(f"{'=' * 70}")
    total_general = 0
    for resumen in resumen_sumas:
        print(
            f"D{resumen['grupo']}: Póliza {resumen['poliza']:>15} | Suma: ${resumen['suma']:>12,.2f} | Filas {resumen['filas']}"
        )
        total_general += resumen["suma"]

    print(f"{'-' * 70}")
    print(f"TOTAL GENERAL: ${total_general:,.2f}")
    print(f"{'=' * 70}\n")

    # Guardar archivo actualizado
    wb_formateado.save(archivo_formateado)
    print(f"✓ Archivo {archivo_formateado} actualizado con agrupación de pólizas")


def actualizar_con_ventas(archivo_ventas, archivo_formateado):
    """
    Actualiza archivo_formateado.xlsx con datos de ventas_2025.xlsx
    sin crear un nuevo archivo.
    Relaciona: Folio (col B) de ventas_2025  <->  Referencia (col G) de archivo_formateado
    Actualiza las columnas:
        C  -> Fecha factura          (Emisión)
        E  -> RFC                    (Receptor RFC)
        F  -> Descripcion            (Conceptos Descripción)
        P  -> Total USD              (Total Original XML)
    Respeta fórmulas y estilos existentes.
    """

    print(f"\n--- Actualizando con {archivo_ventas} ---")

    # Cargar archivo fuente (ventas_2025.xlsx)
    wb_ventas = load_workbook(archivo_ventas, data_only=True)
    ws_ventas = wb_ventas.active

    # Diccionario: Folio → {Emisión, Receptor RFC, Conceptos Descripción, Total Original XML}
    datos_ventas = {}

    # 1. Crear un diccionario para mapear el nombre de la cabecera con su índice
    col_map = {}

    # Asumimos que las cabeceras están en la fila 1
    for cell in ws_ventas[1]:
        if cell.value:
            # Guardamos el nombre (limpio de espacios) y su número de columna
            col_map[str(cell.value).strip()] = cell.column

    # Definir una función auxiliar para obtener el valor de forma segura
    def get_val(row_idx, header_name):
        col_idx = col_map.get(header_name)
        if col_idx:
            return ws_ventas.cell(row=row_idx, column=col_idx).value
        return None  # O un valor por defecto si la columna no existe

    # 2. Tu bucle principal actualizado
    for row_idx in range(2, ws_ventas.max_row + 1):
        # Buscamos el folio usando el nombre de la cabecera "Folio"
        folio = get_val(row_idx, "Folio")

        if folio:
            folio_clean = str(folio).strip()
            datos_ventas[folio_clean] = {
                "emision": get_val(row_idx, "Emisión"),
                "receptor_rfc": get_val(row_idx, "Receptor RFC"),
                "descripcion": get_val(row_idx, "Conceptos Descripción"),
                "total_xml": get_val(row_idx, "Total Original XML"),
                "moneda": get_val(row_idx, "Moneda"),
            }

    print(f"Total de folios cargados desde ventas: {len(datos_ventas)}")

    # Cargar archivo_formateado.xlsx para actualizar (sin crear nuevo)
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    registros_actualizados = 0

    for row_idx in range(2, ws_formateado.max_row + 1):
        referencia = ws_formateado.cell(
            row=row_idx, column=7
        ).value  # Columna G = Referencia

        if referencia:
            # === LIMPIEZA DE REFERENCIA ===
            match = re.search(r"F-?\s*(\d+)", str(referencia).strip())
            ref_clean = (
                match.group(1)
                if match
                else re.sub(r"[^0-9]", "", str(referencia).strip())
            )

            if ref_clean in datos_ventas:
                datos = datos_ventas[ref_clean]

                # C - Fecha factura
                ws_formateado.cell(row=row_idx, column=3).value = datos["emision"]

                # E - RFC
                ws_formateado.cell(row=row_idx, column=5).value = datos["receptor_rfc"]

                # F - Descripcion
                ws_formateado.cell(row=row_idx, column=6).value = datos["descripcion"]

                tipo_moneda = datos["moneda"]
                if tipo_moneda == "USD":
                    # Q - Total USD
                    ws_formateado.cell(row=row_idx, column=16).value = datos[
                        "total_xml"
                    ]

                # Formato de número en Total USD
                if isinstance(datos["total_xml"], (int, float)):
                    ws_formateado.cell(
                        row=row_idx, column=17
                    ).number_format = "#,##0.00"

                registros_actualizados += 1
                print(f"Fila {row_idx}: Referencia {ref_clean} actualizada")

    # Guardar sobre el mismo archivo
    wb_formateado.save(archivo_formateado)

    print(f"\n✓ Actualización con ventas_2025 completada")
    print(f"✓ Total de registros actualizados: {registros_actualizados}")


def parse_fecha_excel(valor):
    """Convierte cualquier formato de fecha del Excel a 'YYYY-MM-DD'"""
    if valor is None:
        return None

    # Si ya es datetime
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d")

    # Convertir a string y limpiar
    fecha_str = str(valor).strip()

    # Intentar conversión normal primero
    try:
        return pd.to_datetime(fecha_str).strftime("%Y-%m-%d")
    except:
        pass

    # Manejar formato "01/Ago/2025"
    meses_es = {
        "Ene": "01",
        "Feb": "02",
        "Mar": "03",
        "Abr": "04",
        "May": "05",
        "Jun": "06",
        "Jul": "07",
        "Ago": "08",
        "Sep": "09",
        "Oct": "10",
        "Nov": "11",
        "Dic": "12",
    }

    try:
        partes = fecha_str.split("/")
        if len(partes) == 3:
            dia = partes[0].zfill(2)
            mes_abbr = partes[1].capitalize()
            anio = partes[2]
            mes = meses_es.get(mes_abbr, "01")
            return f"{anio}-{mes}-{dia}"
    except:
        pass

    return None


def actualizar_tipo_cambio(archivo_cambio, archivo_salida):
    # ============================================
    CSV_FILE = archivo_cambio
    EXCEL_FILE = archivo_salida
    HOJA = "Datos"

    COL_FECHA_CSV = "Fecha_Determinacion"
    COL_FECHA_EXCEL = "Fecha de cobro"
    MAPEO_COLUMNAS = {"Tipo_de_Cambio": "TIPO CAMBIO"}
    FORMATO_FECHA = "%Y-%m-%d"
    # ============================================

    print("Cargando archivos...")

    # 1. Leer CSV
    df_csv = pd.read_csv(CSV_FILE)
    df_csv[COL_FECHA_CSV] = pd.to_datetime(
        df_csv[COL_FECHA_CSV], format=FORMATO_FECHA, errors="coerce"
    )

    # 2. Crear diccionario de actualizaciones
    df_csv_unicos = df_csv.drop_duplicates(subset=[COL_FECHA_CSV], keep="last")
    df_csv_unicos["fecha_str"] = df_csv_unicos[COL_FECHA_CSV].dt.strftime("%Y-%m-%d")
    actualizaciones = df_csv_unicos.set_index("fecha_str")[
        list(MAPEO_COLUMNAS.keys())
    ].to_dict("index")

    print(f"CSV: {len(df_csv)} registros")
    print(f"Fechas disponibles para actualizar: {len(actualizaciones)}")

    # 3. Cargar Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb[HOJA]

    # 4. Buscar columnas
    header_row = 1
    col_indices = {}
    for csv_col, excel_col in MAPEO_COLUMNAS.items():
        for cell in ws[header_row]:
            if cell.value == excel_col:
                col_indices[excel_col] = cell.column
                break

    fecha_col_idx = None
    for cell in ws[header_row]:
        if cell.value == COL_FECHA_EXCEL:
            fecha_col_idx = cell.column
            break

    if fecha_col_idx is None:
        print(f"ERROR: No se encontró la columna '{COL_FECHA_EXCEL}'")
        return

    # 5. Actualizar filas
    actualizados = 0
    for row in range(2, ws.max_row + 1):
        fecha_val = ws.cell(row=row, column=fecha_col_idx).value
        fecha_str = parse_fecha_excel(fecha_val)

        if fecha_str is None:
            continue

        if fecha_str in actualizaciones:
            for csv_col, excel_col in MAPEO_COLUMNAS.items():
                if excel_col in col_indices:
                    ws.cell(
                        row=row, column=col_indices[excel_col]
                    ).value = actualizaciones[fecha_str][csv_col]
            actualizados += 1

    print(f"Registros actualizados: {actualizados}")

    wb.save(EXCEL_FILE)
    print(f"Archivo guardado: {EXCEL_FILE}")


def agregar_sumatoria_global(archivo_salida):
    HOJA = "Datos"
    wb = load_workbook(archivo_salida)
    ws = wb[HOJA]

    # Lista de letras de columnas que quieres sumar
    columnas_a_sumar = ["H", "I", "J", "K", "M", "N", "O", "P"]

    last_row = ws.max_row
    proxima_fila = last_row + 2

    # Definimos el formato: #,##0.00 incluye separador de miles y 2 decimales
    formato_contable = "$#,##0.00"

    # Iteramos directamente sobre las letras de las columnas
    for col in columnas_a_sumar:
        # Usamos la sintaxis de coordenadas ws["A1"] que es más legible
        formula = f"=SUM({col}2:{col}{last_row})"
        celda = ws[f"{col}{proxima_fila}"]
        celda.value = formula

        celda.number_format = formato_contable

        # Opcional: poner el total en negrita
        celda.font = celda.font.copy(bold=True)

    # GUARDAR FUERA DEL BUCLE: Esto es vital para el rendimiento
    wb.save(archivo_salida)


def main_v16_process(path):
    # ANIO, MES = argumentos_consola()
    # dir = f"v16/{ANIO}/{MES}"
    # assets = "v16/assets"

    path = Path(path).resolve()
    file_name = "v16.xlsx"

    archivo_tras = path / "tras.xlsx"
    archivo_rete = path / "rete.xlsx"
    archivo_salida = path / file_name
    archivo_fuente = path / "ventas.xlsx"
    archivo_cambio = path / "cambio_obligaciones.csv"

    # Paso 1: Crear archivo_formateado.xlsx desde agosto_mxn.xlsx

    # ============================================================
    # 1. Procesar DIOT (el que genera el archivo problemático)
    # ============================================================
    try:
        formatear_excel_sin_diarios(archivo_tras, archivo_salida)
        print(f"✅ DIOT procesado correctamente → {archivo_salida}")
    except Exception as e:
        print(f"❌ Error en formatear_excel_sin_diarios: {e}")
        return  # o raise, según prefieras

    # Paso 2: Actualizar archivo_formateado.xlsx con datos de agosto_usd.xlsx
    actualizar_archivo_con_retenido(archivo_rete, archivo_salida)

    # Paso 3: Agrupar por póliza y añadir sumas
    agrupar_por_poliza(archivo_salida)

    # Paso 4: Actualizar archivo_formateado.xlsx con datos de ventas.xlsx
    actualizar_con_ventas(archivo_fuente, archivo_salida)

    # Paso 5: Actualizar archivo_formateado.xlsx con datos de cambio_obligaciones.xlsx
    actualizar_tipo_cambio(archivo_cambio, archivo_salida)

    # Paso 6: Agregar sumatoria total para columnas (H, I, J, K, M, N, O, P)
    agregar_sumatoria_global(archivo_salida)
