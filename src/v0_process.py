import argparse
import re
from datetime import datetime
from difflib import SequenceMatcher

import openpyxl
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from rapidfuzz import fuzz, process

# from openpyxl.utils import get_column_letter

DIR_ENCABEZADO = {
    "A": "Fecha Registro",
    "B": "Tipo",
    "C": "Número",
    "D": "Cliente",
    "E": "Concepto",
    "F": "Descripcion",
    "G": "Gravable IVA",
    "H": "Referencia Factura PL",
    "I": "Folio Fiscal",
    "J": "NC Aplicadas a Cobranza MXN",
    "K": "Cobranza Recibida MXN",
    "L": "Tipo de Cambio",
    "M": "NC Aplicadas a Cobranza USD",
    "N": "Cobranza Recibida USD",
    "O": "Importe Cobrado USD",
    "P": "Referencia de Cobranza Edo Cuenta",
    "Q": "BANCO",
    "R": "CUENTA BANCARIA",
    "S": "FORMA DE PAGO",
    "T": "Fecha Cobro",
    "U": "Pedimento",
    "V": "INCOTERM",
    "W": "Tipo de Pedimento",
    "X": "Valor Comercial",
    "Y": "Fecha de Pedimento",
}

LIST_LENGTHS = [
    20,  # Fecha Registro
    12,  # Tipo
    12,  # Número
    25,  # Cliente
    30,  # Concepto
    15,  # Descripcion
    20,  # Gravable IVA
    30,  # Referencia Factura PL
    40,  # Folio Fiscal
    30,  # NC Aplicadas a Cobranza MXN
    30,  # Cobranza Recibida MXN
    20,  # Tipo de Cambio
    40,  # NC Aplicadas a Cobranza USD
    30,  # Cobranza Recibida USD
    30,  # Importe Cobrado USD
    40,  # Referencia de Cobranza Edo Cuenta
    10,  # BANCO
    20,  # CUENTA BANCARIA
    20,  # FORMA DE PAGO
    20,  # Fecha Cobro
    20,  # Pedimento
    20,  # INCOTERM
    20,  # Tipo de Pedimento
    20,  # Valor Comercial
    20,  # Fecha de Pedimento
]


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


def obtener_grupo_polizas(ws):
    grupos_poliza = {}

    # Recorrer todas las filas de datos (comenzando desde fila 2)
    columna_numero = obtener_indices(["Número"], DIR_ENCABEZADO, 1)
    for row_idx in range(2, ws.max_row + 1):
        numero_poliza = ws.cell(row=row_idx, column=columna_numero).value  # Columna C

        if numero_poliza:
            numero_poliza_str = str(numero_poliza).strip()
            if numero_poliza_str not in grupos_poliza:
                grupos_poliza[numero_poliza_str] = []
            grupos_poliza[numero_poliza_str].append(row_idx)

    return grupos_poliza


def obtener_columnas(headers_buscados, diccionario_encabezados):
    """
    Devuelve:
        - La columna (string) si solo se busca **un** encabezado.
        - Una lista de columnas si se buscan **varios**.
    """
    # Creamos un diccionario inverso: encabezado → columna (una sola vez)
    header_a_columna = {
        header: columna for columna, header in diccionario_encabezados.items()
    }

    # Obtenemos los resultados
    resultado = [header_a_columna[h] for h in headers_buscados if h in header_a_columna]

    # Si solo hay un resultado → devolver el valor directamente (no una lista)
    if len(resultado) == 1:
        return resultado[0]

    return resultado


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
            valores[0],  # A - Fecha Registro
            valores[1],  # B - Tipo
            valores[2],  # C - Número
            cliente_actual,  # D - Cliente (NUEVO)
            valores[3],  # E - Concepto
            "Venta Productos de Aluminio",  # F - Descripcion (predeterminado)
            "0%",  # G - Gravable IVA (vacío)
            valores[4],  # H - Referencia Factura PL
            "",  # I - Folio Fiscal
            valores[5],  # J - NC Aplicadas a Cobranza MXN
            valores[6],  # K - Cobranza Recibida MXN
            "",  # L - Tipo de Cambio (vacío)
            "",  # M - NC Aplicadas a Cobranza USD (vacío)
            "",  # N - Cobranza Recibida USD (vacío)
            "",  # O - Importe Cobrado USD (vacío)
            "",  # P - Referencia de Cobranza Edo Cuenta (vacío)
            "HSBC",  # Q - BANCO (vacío)
            "021225070037154302",  # R - CUENTA BANCARIA (vacío)
            "Trasferencia electronica",  # S - FORMA DE PAGO (vacío)
            valores[0],  # T - Fecha Cobro (vacío)
            "",  # U - Pedimento (vacío)
            "",  # V - INCOTERM (vacío)
            "",  # W - Tipo de Pedimento (vacío)
            "",  # X - Valor Comercial (vacío)
            "",  # Y - Fecha de Pedimento (vacío)
        ]
        columnas_formato_numerico = obtener_indices(
            [
                "NC Aplicadas a Cobranza MXN",
                "Cobranza Recibida MXN",
                "Tipo de Cambio",
                "NC Aplicadas a Cobranza USD",
                "Cobranza Recibida USD",
                "Importe Cobrado USD",
                "Referencia de Cobranza Edo Cuenta",
                "Tipo de Pedimento",
            ],
            DIR_ENCABEZADO,
            1,
        )

        # Escribir fila en el nuevo archivo
        for col, valor in enumerate(nuevos_valores, 1):
            es_negativo = isinstance(valor, (int, float)) and valor < 0

            condicion = (
                columnas_formato_numerico[0]
                if col == columnas_formato_numerico[1] and es_negativo
                else col
            )

            celda = ws_nuevo.cell(row=fila_nueva, column=condicion)
            celda.value = valor
            # celda.border = thin_border
            celda.alignment = Alignment(horizontal="left", vertical="center")

            # Alineación para columnas numéricas
            if col in columnas_formato_numerico:  # Número, Cargos, Abonos, etc.
                celda.alignment = Alignment(horizontal="right", vertical="center")
                # Formato de números
                if isinstance(valor, (int, float)):
                    celda.number_format = "#,##0.00"

        fila_nueva += 1

    # Congelar encabezado
    ws_nuevo.freeze_panes = "A2"

    # Ajustar anchos de columnas
    keys_columns = list(DIR_ENCABEZADO.keys())
    anchos = dict(zip(keys_columns, LIST_LENGTHS))
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


def actualizar_archivo_con_usd(archivo_usd, archivo_formateado):
    """
    Procesa agosto_usd.xlsx y actualiza archivo_formateado.xlsx
    sin crear un nuevo archivo.
    Busca coincidencias por Referencia y actualiza columnas M y N.
    """

    # Cargar agosto_usd.xlsx
    wb_usd = load_workbook(archivo_usd)
    ws_usd = wb_usd.active

    # Diccionario para almacenar datos de USD (Referencia -> {Cargos, Abonos})
    datos_usd = {}

    cliente_actual = ""

    # Leer datos del archivo USD
    print(f"\n--- Procesando {archivo_usd} ---")
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
        referencia = valores[4] if len(valores) > 4 else ""  # Columna E (Referencia)
        cargos = (
            valores[5] if (len(valores) > 5 and valores[5] is not None) else ""
        )  # Columna F (Cargos)
        abonos = (
            valores[6] if (len(valores) > 6 and valores[6] is not None) else ""
        )  # Columna G (Abonos)

        if referencia:
            if isinstance(referencia, float) and referencia.is_integer():
                ref_key = str(int(referencia))
            else:
                ref_key = str(referencia).strip()
                if ref_key.endswith(".0"):
                    ref_key = ref_key[:-2]
            datos_usd[ref_key] = {"cargos": cargos, "abonos": abonos}
            print(f"Referencia USD: {ref_key} - Cargos: {cargos}, Abonos: {abonos}")

    # Cargar archivo_formateado.xlsx para actualizar
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    registros_actualizados = 0

    # Iterar sobre archivo_formateado.xlsx
    print(f"\n--- Actualizando {archivo_formateado} ---")
    columna_referencia = obtener_indices(["Referencia Factura PL"], DIR_ENCABEZADO, 1)
    columna_cobranza = obtener_indices(
        ["NC Aplicadas a Cobranza USD"], DIR_ENCABEZADO, 1
    )
    columna_importe = obtener_indices(["Cobranza Recibida USD"], DIR_ENCABEZADO, 1)
    for row_idx in range(2, ws_formateado.max_row + 1):  # Comenzar desde fila 2
        # Columna I = Referencia Factura PL
        referencia_col = ws_formateado.cell(
            row=row_idx, column=columna_referencia
        ).value

        if referencia_col:
            referencia_clean = str(referencia_col).strip()

            # Buscar en datos_usd
            if referencia_clean in datos_usd:
                v_abonos = datos_usd[referencia_clean]["abonos"]
                v_cargos = datos_usd[referencia_clean]["cargos"]

                if isinstance(v_abonos, (int, float)) and v_abonos < 0:
                    ws_formateado.cell(
                        row=row_idx, column=columna_cobranza
                    ).value = v_abonos
                    ws_formateado.cell(row=row_idx, column=columna_importe).value = ""
                else:
                    ws_formateado.cell(
                        row=row_idx, column=columna_cobranza
                    ).value = v_cargos
                    ws_formateado.cell(
                        row=row_idx, column=columna_importe
                    ).value = v_abonos

                # # Actualizar columna N (Cobranza Recibida USD) con Cargos
                # ws_formateado.cell(row=row_idx, column=columna_cobranza).value = datos_usd[referencia_clean]["abonos"] if datos_usd[referencia_clean]["abonos"] < 0 else datos_usd[referencia_clean]["cargos"]

                # # Actualizar columna O (Importe Cobrado USD) con Abonos
                # ws_formateado.cell(row=row_idx, column=columna_importe).value = '' if datos_usd[referencia_clean]["abonos"] < 0 else datos_usd[referencia_clean]["abonos"]

                # Aplicar formato a los valores numéricos
                for col in [12, 13, 14, 15]:
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


def actualizar_archivo_con_hsbc(archivo_hsbc, archivo_formateado):
    """
    Procesa hsbc.xlsx y actualiza archivo_formateado.xlsx
    sin crear un nuevo archivo.
    """

    # ============================================
    # Obtener meses y años válidos (usando parse_fecha_excel)
    # ============================================
    df_formateado = pd.read_excel(archivo_formateado, sheet_name="Datos")
    fechas_formateado = df_formateado.iloc[:, 0].apply(parse_fecha_excel)

    meses_validos = set()
    for f_str in fechas_formateado:
        if f_str:
            try:
                f = datetime.strptime(f_str, "%Y-%m-%d")
                meses_validos.add((f.month, f.year))
            except:
                pass

    print(f"Meses/años válidos encontrados: {meses_validos}")

    # Cargar hsbc.xlsx
    wb_usd = load_workbook(archivo_hsbc)
    ws_usd = wb_usd.active

    datos_usd = {}
    cliente_actual = ""

    print(f"\n--- Procesando {archivo_hsbc} ---")
    for row_idx, row in enumerate(
        ws_usd.iter_rows(min_row=1, max_row=ws_usd.max_row, values_only=False), 1
    ):
        valores = [celda.value for celda in row]
        fecha_val = valores[0] if len(valores) > 0 else None

        # Omitir fila de encabezado "Cuenta | Nombre"
        if valores[0] and valores[1]:
            if (
                str(valores[0]).replace(" ", "").lower() == "cuenta"
                and str(valores[1]).replace(" ", "").lower() == "nombre"
            ):
                continue

        # Detectar cliente
        if (
            valores[0]
            and valores[1]
            and "-" in str(valores[0])
            and isinstance(valores[1], str)
            and not valores[2]
            and not valores[3]
        ):
            cliente_actual = valores[1].strip()
            print(f"Cliente USD detectado: {cliente_actual}")
            continue

        # Saltar encabezados y filas vacías
        if valores[0] == "Fecha" or not valores[0]:
            continue
        if not valores[1]:
            continue

        # === VALIDACIÓN DE FECHA (corregida) ===
        fecha_parseada = parse_fecha_excel(fecha_val)
        if not fecha_parseada:
            print(f"Fila {row_idx} omitida (fecha inválida)")
            continue

        try:
            f = datetime.strptime(fecha_parseada, "%Y-%m-%d")
            if (f.month, f.year) not in meses_validos:
                print(f"Fila {row_idx} omitida por fecha ({fecha_parseada})")
                continue
        except:
            print(f"Fila {row_idx} omitida (fecha inválida)")
            continue

        # Filtro Diario / Egresos
        if (
            str(valores[1]).strip().lower() in ["diario", "egresos"]
            and valores[5] is not None
        ):
            print(f"Fila {row_idx} HSBC omitida (Diario y Egresos)")
            continue

        # Almacenar datos
        numero = valores[2] if len(valores) > 2 else ""
        cargos = valores[5] if len(valores) > 5 else ""

        if numero and cargos is not None:
            try:
                numero_int = int(numero)
                datos_usd[numero_int] = {"cargos": cargos, "numero": numero_int}
                print(
                    f"Referencia USD: {numero} - Cargos: {cargos}, Número: {numero_int}"
                )
            except:
                pass

    # ============================================
    # Actualizar archivo_formateado.xlsx
    # ============================================
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    registros_actualizados = 0
    columna_numero = obtener_indices(["Número"], DIR_ENCABEZADO, 1)
    columna_importe = obtener_indices(["Importe Cobrado USD"], DIR_ENCABEZADO, 1)

    grupos_polizas = obtener_grupo_polizas(ws_formateado)
    grupos_polizas = {k: v[-1] for k, v in grupos_polizas.items()}

    print(f"\n--- Actualizando {archivo_formateado} ---")
    for poliza, fila in grupos_polizas.items():
        # Colocar la fórmula en la última fila, columna O (col 15)
        celda_importe = ws_formateado.cell(row=fila, column=columna_importe)

        if poliza is None:
            continue

        poliza_int = int(poliza) if isinstance(poliza, (int, float, str)) else None

        if poliza_int is not None and poliza_int in datos_usd:
            celda_importe.value = datos_usd[poliza_int]["cargos"]
            celda_importe.number_format = "#,##0.00"
            celda_importe.alignment = Alignment(horizontal="right", vertical="center")

            registros_actualizados += 1
            print(f"Fila {fila}: Número {poliza_int} actualizada")

    wb_formateado.save(archivo_formateado)

    print(f"\n✓ Archivo actualizado correctamente")
    print(f"✓ Total de registros actualizados: {registros_actualizados}")


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
    grupos_poliza = obtener_grupo_polizas(ws_formateado)
    resumen_sumas = []
    numero_grupo = 1

    print(f"\n--- Agrupando por Póliza ---")
    print(f"Total de grupos de pólizas encontrados: {len(grupos_poliza)}")

    # Border grueso para separación
    thick_border = Border(
        bottom=Side(style="thick", color="000000"),  # 1.75 pt negro
    )

    # Procesar cada grupo de póliza
    columna_importe_cobrado = obtener_indices(
        ["Importe Cobrado USD"], DIR_ENCABEZADO, 1
    )
    columna_cobranza_recibida = obtener_indices(
        ["Cobranza Recibida USD"], DIR_ENCABEZADO, 1
    )
    columna_referencia_cobranza_edo = obtener_indices(
        ["Referencia de Cobranza Edo Cuenta"], DIR_ENCABEZADO, 1
    )
    longitud_cabecera = len(DIR_ENCABEZADO.keys()) + 1

    col_rango_suma = obtener_columnas(["Cobranza Recibida USD"], DIR_ENCABEZADO)
    col_nc_aplicadas_cobranza = obtener_columnas(
        ["NC Aplicadas a Cobranza USD"], DIR_ENCABEZADO
    )

    for poliza, filas in grupos_poliza.items():
        primera_fila = filas[0]
        ultima_fila = filas[-1]

        # Rango de suma: M[primera_fila]:M[ultima_fila]
        rango_resta = f"{col_nc_aplicadas_cobranza}{primera_fila}:{col_nc_aplicadas_cobranza}{ultima_fila}"
        rango_suma = f"{col_rango_suma}{primera_fila}:{col_rango_suma}{ultima_fila}"
        formula = f"=SUM({rango_suma}) - SUM({rango_resta})"

        # Colocar la fórmula en la última fila, columna O (col 15)
        celda_formula = ws_formateado.cell(
            row=ultima_fila, column=columna_importe_cobrado
        )
        celda_formula.value = formula
        celda_formula.number_format = "#,##0.00"
        celda_formula.alignment = Alignment(horizontal="right", vertical="center")

        # Aplicar border-bottom grueso a toda la última fila del grupo
        for col in range(1, longitud_cabecera):  # 26 columnas
            celda = ws_formateado.cell(row=ultima_fila, column=col)
            celda.border = thick_border

        # Calcular suma para consola
        suma_grupo = 0
        for fila in filas:
            valor = ws_formateado.cell(
                row=fila, column=columna_cobranza_recibida
            ).value  # Columna N
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

        celda_ref_edo = ws_formateado.cell(
            row=ultima_fila, column=columna_referencia_cobranza_edo
        )
        celda_ref_edo.value = f"D{numero_grupo}"
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


def conciliar_pedimentos_polizas(archivo_salida, billing_order, bosch, zfnvh, expo):
    def clean_ref(val):
        if pd.isna(val) or val == "":
            return ""
        val = str(val).strip()
        if val.endswith(".0"):
            val = val[:-2]
        match = re.search(r"F[-\s]*0*(\d+)", val, re.IGNORECASE)
        if match:
            return match.group(1)
        digits = re.sub(r"\D", "", val)
        return digits.lstrip("0") or digits

    def get_fuzzy_match(df, ref_clean, threshold=0.90):
        unique_vals = df["ref_clean"].unique()
        matches = [
            v
            for v in unique_vals
            if v and SequenceMatcher(None, v, ref_clean).ratio() >= threshold
        ]
        if matches:
            return df[df["ref_clean"].isin(matches)]
        return pd.DataFrame()

    print("Cargando archivos...")

    # Nombres de columnas origen (ADX)
    cols_adx_interes = [
        "Packing list",
        "Factura fiscal",
        "PEDIMENTO",
        "CLAVE",
        "INCOTERM",
        "PRECIO PAGADO/VALOR COMERCIAL",
        "PAGO DE PEDIMENTO",
        "REFERENCIA AA",
    ]

    aax = pd.read_excel(archivo_salida)
    abx = pd.read_excel(billing_order, usecols=["Document number", "Packing List"])
    acx = pd.read_excel(bosch, usecols=["FI Document", "External Delivery Note"])
    adx = pd.read_excel(expo, sheet_name="AllMonths", usecols=cols_adx_interes)
    aex = pd.read_excel(zfnvh, usecols=["DocumentNo", "Delivery note"])

    # Mapeo de columnas
    (
        col_ref_aax,
        col_pedimento_aax,
        col_tipo_aax,
        col_incoterm_aax,
        col_val_com_aax,
        col_fecha_aax,
    ) = (
        "Referencia Factura PL",
        "Pedimento",
        "Tipo de Pedimento",
        "INCOTERM",
        "Valor Comercial",
        "Fecha de Pedimento",
    )
    col_fi_doc_acx, col_ext_del_note_acx = "FI Document", "External Delivery Note"
    col_doc_num_abx, col_packing_list_abx = "Document number", "Packing List"
    col_doc_num_aex, col_del_not_aex = "DocumentNo", "Delivery note"

    print("Pre-procesando datos...")
    acx["ref_clean"] = acx[col_fi_doc_acx].astype(str).apply(clean_ref)
    abx["ref_clean"] = abx[col_doc_num_abx].astype(str).apply(clean_ref)
    aex["ref_clean"] = aex[col_doc_num_aex].astype(str).apply(clean_ref)

    # --- NUEVA LÓGICA DE LOOKUP MULTI-COLUMNA ---
    # Agrupamos por packing list y creamos un diccionario que contiene listas de valores únicos para cada columna
    adx_grouped = adx.groupby(adx["Packing list"].astype(str).str.strip())

    adx_lookup_completo = {}
    for note, group in adx_grouped:
        adx_lookup_completo[note] = {
            "facturas": sorted(group["Factura fiscal"].dropna().astype(str).unique()),
            "pedimentos": sorted(group["PEDIMENTO"].dropna().astype(str).unique()),
            "claves": sorted(group["CLAVE"].dropna().astype(str).unique()),
            "incoterms": sorted(group["INCOTERM"].dropna().astype(str).unique()),
            "precios": sorted(
                group["PRECIO PAGADO/VALOR COMERCIAL"].dropna().astype(str).unique()
            ),
            "pagos": sorted(group["PAGO DE PEDIMENTO"].dropna().astype(str).unique()),
            "referencias_aa": sorted(
                group["REFERENCIA AA"].dropna().astype(str).unique()
            ),
        }

    memo_resultados = {}
    unique_refs = aax[col_ref_aax].dropna().unique()

    print(f"Procesando {len(unique_refs)} referencias únicas...")

    for ref_original in unique_refs:
        ref_original_str = str(ref_original).strip()
        ref_clean = clean_ref(ref_original_str)

        if not ref_clean:
            continue

        external_notes = set()

        # Búsqueda en Cascada (ACX -> ABX -> AEX)
        matches = acx[acx["ref_clean"] == ref_clean]
        if matches.empty:
            matches = get_fuzzy_match(acx, ref_clean)

        if not matches.empty:
            external_notes = set(
                matches[col_ext_del_note_acx].dropna().astype(str).str.strip()
            )
        else:
            matches = abx[abx["ref_clean"] == ref_clean]
            if matches.empty:
                matches = get_fuzzy_match(abx, ref_clean)
            if not matches.empty:
                for val in matches[col_packing_list_abx].dropna().astype(str):
                    parts = val.split("|")
                    external_notes.add(
                        parts[1].strip() if len(parts) > 1 else parts[0].strip()
                    )

        if not external_notes:
            matches = aex[aex["ref_clean"] == ref_clean]
            if matches.empty:
                matches = get_fuzzy_match(aex, ref_clean)
            if not matches.empty:
                external_notes = set(
                    matches[col_del_not_aex].dropna().astype(str).str.strip()
                )

        # --- RECOPILACIÓN DE RESULTADOS DE ADX ---
        res_consolidado = {
            "facturas": [],
            "pedimentos": [],
            "claves": [],
            "incoterms": [],
            "precios": [],
            "pagos": [],
            "referencias_aa": [],
        }

        for note in external_notes:
            if note in adx_lookup_completo:
                data = adx_lookup_completo[note]
                for key in res_consolidado:
                    res_consolidado[key].extend(data[key])

        # Limpiar duplicados y convertir a string
        final_vals = {}
        for key, vals in res_consolidado.items():
            unique_vals = sorted(list(set(vals)))
            val_str = ", ".join(unique_vals)
            # Formato especial para factura
            if key == "facturas" and val_str and not val_str.startswith("F-"):
                val_str = f"F- {val_str}"
            final_vals[key] = val_str

        if any(final_vals.values()):
            memo_resultados[ref_original_str] = final_vals

    # --- ACTUALIZACIÓN DE EXCEL ---
    print("\nActualizando archivo Excel...")
    wb = load_workbook(archivo_salida)
    ws = wb.active

    # Mapeo de encabezados del Excel de salida vs nuestras claves de resultado
    # Ajusta los nombres de la izquierda según cómo se llamen las columnas en 'archivo_salida'
    target_cols = {
        col_ref_aax: "facturas",
        col_pedimento_aax: "pedimentos",
        col_tipo_aax: "claves",
        col_incoterm_aax: "incoterms",
        col_val_com_aax: "precios",
        col_fecha_aax: "pagos",
    }

    col_indices = {}
    for idx, cell in enumerate(ws[1], 1):
        header = str(cell.value).strip() if cell.value else ""
        if header in target_cols:
            col_indices[target_cols[header]] = idx

    actualizados = 0
    col_idx_ref = col_indices.get(
        "facturas"
    )  # Usamos la columna de referencia original para buscar en el memo

    if col_idx_ref:
        for row in range(2, ws.max_row + 1):
            cell_ref = ws.cell(row=row, column=col_idx_ref)
            val_orig = str(cell_ref.value).strip() if cell_ref.value else ""

            if val_orig in memo_resultados:
                datos_nuevos = memo_resultados[val_orig]

                # Actualizar cada columna encontrada
                for key, col_idx in col_indices.items():
                    valor_str = datos_nuevos[key]

                    if valor_str:  # Solo escribir si hay datos
                        celda = ws.cell(row=row, column=col_idx)

                        if key == "precios":
                            try:
                                celda.value = float(valor_str)
                                celda.number_format = "#,##0.00"
                            except:
                                celda.value = valor_str
                        elif key == "pagos":
                            try:
                                fecha_dt = pd.to_datetime(valor_str).to_pydatetime()

                                celda.value = fecha_dt
                                # celda.number_format = '[$-080A]dd/mmm/yyyy'
                                celda.number_format = "dd/mm/yyyy"
                            except:
                                celda.value = valor_str
                        else:
                            celda.value = valor_str

                actualizados += 1

    wb.save(archivo_salida)
    print(f"✅ Proceso completado. Filas modificadas: {actualizados}")


def actualizar_con_ventas(archivo_ventas, archivo_formateado):
    """
    Actualiza archivo_formateado.xlsx con datos de ventas.xlsx
    sin crear un nuevo archivo.
    Relaciona: Folio (col B) de ventas_2025  <->  Referencia (col G) de archivo_formateado
    Actualiza las columnas:
        I  -> Folio Fiscal          (Emisión)
    Respeta fórmulas y estilos existentes.
    """

    print(f"\n--- Actualizando con {archivo_ventas} ---")

    # Cargar archivo fuente (ventas_2025.xlsx)
    wb_ventas = load_workbook(archivo_ventas, data_only=True)
    ws_ventas = wb_ventas.active

    # Diccionario: Folio → {Emisión, Receptor RFC, Conceptos Descripción, Total Original XML}
    datos_ventas = {}

    # Crear diccionario para mapear el nombre de la cabecera con su indice
    col_map = {}

    # Asumimos que las cabeceras estan en la fila 1
    for cell in ws_ventas[1]:
        if cell.value:
            # Guardar el nombre (limpio de espacios) y su número de columna
            col_map[str(cell.value).strip()] = cell.column

    # Definir una función auxiliar para obtener el valor de forma segura
    def get_val(row_idx, header_name):
        col_idx = col_map.get(header_name)
        if col_idx:
            return ws_ventas.cell(row=row_idx, column=col_idx).value
        return None

    for row_idx in range(2, ws_ventas.max_row + 1):
        folio = get_val(row_idx, "Folio")  # Columna B = Folio
        if folio:
            folio_clean = str(folio).strip()
            datos_ventas[folio_clean] = {
                "uuid": get_val(row_idx, "UUID"),  # C - Emisión
            }

    print(f"Total de folios cargados desde ventas: {len(datos_ventas)}")

    # Cargar archivo_formateado.xlsx para actualizar (sin crear nuevo)
    wb_formateado = load_workbook(archivo_formateado)
    ws_formateado = wb_formateado.active

    registros_actualizados = 0

    for row_idx in range(2, ws_formateado.max_row + 1):
        referencia = ws_formateado.cell(
            row=row_idx, column=8
        ).value  # Columna H = Referencia Factura PL

        if referencia:
            # === LIMPIEZA DE REFERENCIA ===
            match = re.search(r"F-?\s*(\d+)", str(referencia).strip())
            ref_clean = (
                match.group(1).lstrip("0")
                if match
                else re.sub(r"[^0-9]", "", str(referencia).strip())
            )

            if ref_clean in datos_ventas:
                datos = datos_ventas[ref_clean]

                # I - Folio Fiscal
                ws_formateado.cell(row=row_idx, column=9).value = datos["uuid"]

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
    COL_FECHA_EXCEL = "Fecha Registro"
    MAPEO_COLUMNAS = {"Tipo_de_Cambio": "Tipo de Cambio"}
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

    # Lista de columnas a las cuales se agregara la sumatoria
    columnas_a_sumar = ["J", "K", "M", "N", "O"]

    last_row = ws.max_row
    proxima_fila = last_row + 2

    # Definimos el formato: #,##0.00 incluye separador de miles y 2 decimales
    formato_contable = "$#,##0.00"

    # Iteramos directamente sobre las letras de las columnas
    for col in columnas_a_sumar:
        # Usamos la sintaxis de coordenadas ws["A1"] que es más legible
        ridx = columnas_a_sumar.index(col) - 1
        columna_resta = columnas_a_sumar[ridx]

        formula_resta = f"SUM({columna_resta}2:{columna_resta}{last_row})"
        formula_suma = f"SUM({col}2:{col}{last_row})"

        # formula = (
        #     f"={formula_suma} - {formula_resta}" if col == "K" else f"={formula_suma}"
        # )
        formula = f"={formula_suma}"
        celda = ws[f"{col}{proxima_fila}"]
        celda.value = formula

        celda.number_format = formato_contable

        # Opcional: poner el total en negritas
        celda.font = celda.font.copy(bold=True)

    # Definir el relleno amarillo
    amarillo = PatternFill(
        start_color="FFFF00",  # Amarillo puro
        end_color="FFFF00",
        fill_type="solid",
    )

    formula_total_ventas = f"=K{proxima_fila} - J{proxima_fila}"
    celda_ventas_exportacion = ws[f"J{proxima_fila + 3}"]
    celda_total_ventas_expo = ws[f"K{proxima_fila + 3}"]

    celda_ventas_exportacion.value = "Ventas 0% Exportación"
    celda_total_ventas_expo.value = formula_total_ventas
    celda_total_ventas_expo.fill = amarillo

    wb.save(archivo_salida)


def actualizar_archivo_b(ruta_a, ruta_b):
    print("Cargando archivos...")

    # 1. Cargar archivo_a (solo lectura de datos para mayor velocidad)
    # Usamos data_only=True por si la referencia en A es resultado de una fórmula
    wb_a = openpyxl.load_workbook(ruta_a, data_only=True)
    ws_a = wb_a.active  # O usa wb_a["NombreDeLaHoja"]

    # Extraer todas las referencias de la Columna G (columna 7) del archivo_a
    # Guardamos en un set() para que la búsqueda sea ultra rápida
    referencias_a = set()
    for row in range(2, ws_a.max_row + 1):
        valor = ws_a.cell(row=row, column=7).value  # Columna G
        if valor is not None:
            referencias_a.add(str(valor).strip())

    wb_a.close()  # Ya no necesitamos el archivo A

    # 2. Cargar archivo_b (Modo completo para preservar estilos y fórmulas)
    wb_b = openpyxl.load_workbook(ruta_b)
    ws_b = wb_b.active

    reporte_borrados = []
    total_procesados = 0

    # 3. Iterar archivo_b de ABAJO hacia ARRIBA
    # Columna H es la columna 8
    max_fila_b = ws_b.max_row

    print(f"Analizando {max_fila_b} filas en el Archivo B...")

    for row_idx in range(max_fila_b, 1, -1):
        ref_pl_b = ws_b.cell(row=row_idx, column=8).value  # Columna H

        if ref_pl_b is not None:
            ref_pl_str = str(ref_pl_b).strip()

            if ref_pl_str in referencias_a:
                # Guardar info para el reporte antes de borrar
                reporte_borrados.append(ref_pl_str)

                # Borrar la fila completa
                ws_b.delete_rows(row_idx)
                total_procesados += 1

    # 4. Guardar los cambios en el MISMO archivo
    if total_procesados > 0:
        wb_b.save(ruta_b)
        print(f"Éxito: Se actualizaron los cambios en {ruta_b}")
    else:
        print("No se encontraron coincidencias para borrar.")

    # 5. Mostrar reporte
    print("\n--- REPORTE DE FACTURAS ELIMINADAS ---")
    if reporte_borrados:
        for ref in reporte_borrados:
            print(f"Eliminada: {ref}")
        print(f"Total eliminadas: {len(reporte_borrados)}")
    else:
        print("Cero filas eliminadas.")


def main_v0_process(path):
    # ANIO, MES = argumentos_consola()
    # dir = f"v0/{ANIO}/{MES}"
    # assets = "v0/assets"
    # public = "public"

    archivo_mxn = f"{path}/mxn.xlsx"
    archivo_usd = f"{path}/usd.xlsx"
    archivo_salida = f"{path}/v0.xlsx"

    archivo_ventas = f"{path}/ventas.xlsx"
    archivo_cambio = "cambio_obligaciones.csv"
    archivo_hsbc = f"{path}/hsbc.xlsx"

    billing_order = f"{path}/billing_order.xlsx"
    bosch = f"{path}/bosch.xlsx"
    zfnvh = f"{path}/zfnvh.xlsx"
    expo = f"{path}/EXPO.xlsx"

    file_v16 = f"{path}/v16.xlsx"

    # Paso 1: Crear archivo_formateado.xlsx desde agosto_mxn.xlsx
    formatear_excel_sin_diarios(archivo_mxn, archivo_salida)

    # Paso 2: Actualizar archivo_formateado.xlsx con datos de agosto_usd.xlsx
    actualizar_archivo_con_usd(archivo_usd, archivo_salida)

    # Paso 4: Cruce de datos
    conciliar_pedimentos_polizas(archivo_salida, billing_order, bosch, zfnvh, expo)

    # Paso 5: Actualizar archivo_formateado.xlsx respecto a ventas.xlsx
    actualizar_con_ventas(archivo_ventas, archivo_salida)

    # Paso 6: Actualizar archivo_formateado.xlsx con cambio_obligaciones.csv
    actualizar_tipo_cambio(archivo_cambio, archivo_salida)

    # Paso 7: Actualizar archivo_formateado.xlsx con hsbc.xlsx
    actualizar_archivo_con_hsbc(archivo_hsbc, archivo_salida)

    # Paso 8: Conciliar facturas existentes en V16 y borrar en V0
    actualizar_archivo_b(file_v16, archivo_salida)

    # Paso 3: Agrupar por póliza y añadir sumas
    agrupar_por_poliza(archivo_salida)

    # Paso 8: Agregar sumatoria total para columnas (J, K, M, N, O)
    agregar_sumatoria_global(archivo_salida)
