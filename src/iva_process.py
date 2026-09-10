import os
import re
from pathlib import Path

import openpyxl  # Librería para manejar estilos y fórmulas
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from rapidfuzz import process, utils


def procesar_diot_template(archivo_origen, archivo_salida):
    # 1. Leer el archivo DIOT
    try:
        df_origen = pd.read_excel(archivo_origen)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{archivo_origen}'")
        return

    # 2. Definir la cabecera final
    cabecera_final = [
        "Fecha",
        "Tipo",
        "Número",
        "Proveedor",
        "RFC",
        "Factura",
        "FOLIO FISCAL",
        "CONCEPTO",
        "Subtotal Base 16%",
        "Base exenta IVA",
        "IVA 16%",
        "Ret IVA",
        "ISR Retenido",
        "Impto. Loc. Ret.",
        "Total",
        "Moneda",
        "TC",
        "Subtotal USD",
        "Base exenta IVA",
        "IVA 16% USD",
        "Ret IVA USD",
        "Total USD",
        "Fecha Pago",
        "Banco",
        "Monto",
        "Referencia Edo Cta",
    ]

    # 3. Preparar los bloques de datos (Filtrar Tipo, EXCLUIR RFC genérico y Ordenar)
    tipos = ["Egresos", "Ingresos", "Diario", "Genericos_Buscar", "Genericos"]
    bloques = {}
    rfc_generico = "XAXX010101000"

    for t in tipos:
        # Filtrar por tipo
        df_temp = df_origen[df_origen["Tipo"] == t].copy()

        # FILTRO CRÍTICO: Excluir RFC genérico (Columna 'Concepto' en el origen)
        df_temp = df_temp[df_temp["Concepto"].str.strip() != rfc_generico]

        # Ordenar por RFC
        df_temp = df_temp.sort_values(by="Concepto")
        bloques[t] = df_temp

    # 1. Definimos el patrón regex
    # patron_regex = r"^\d{2} \d{2} \d{4} \d{6,}$"
    # patron_regex = (r"^\s*\d{1,2}\s+(?:\d{2,3}(?:\s+\d{3,4})+\s+\d{6,}|\d{6}\s+\d{8,})\s*$")
    patron_regex = r"^\s*\d{1,3}(?:\s+\d{1,8})+\s+\d{6,}\s*$"

    # 2. Creamos la máscara booleana (la condición lógica)
    # Esto genera una serie de True/False para cada fila
    filtro_genericos = df_origen["Concepto"].str.strip() == rfc_generico
    filtro_coincide = df_origen["Referencia"].str.contains(
        f"COMISION|VENTAS|DEPOSITO|{patron_regex}", case=False, na=True, regex=True
    )

    # 3. Separamos el DataFrame usando la máscara
    # Los que SÍ cumplen la condición
    df_temp_si = df_origen[filtro_genericos & filtro_coincide].copy()

    # Los que NO cumplen la condición (usando ~ para negar el filtro)
    df_temp_no = df_origen[filtro_genericos & ~filtro_coincide].copy()
    bloques["Genericos_Buscar"] = df_temp_no
    bloques["Genericos"] = df_temp_si

    # 4. Configurar el libro de Excel con Openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Hoja1"

    # --- ESTILOS ---
    relleno_salmon = PatternFill(
        start_color="f6c6ad", end_color="f6c6ad", fill_type="solid"
    )
    border_negro = Side(border_style="medium", color="000000")
    estilo_borde = Border(
        left=border_negro, right=border_negro, top=border_negro, bottom=border_negro
    )
    fuente_cabecera = Font(name="Arial", size=8, bold=True)
    fuente_datos = Font(name="Arial", size=8)

    # 5. Escribir Cabecera
    for col_num, titulo in enumerate(cabecera_final, 1):
        celda = ws.cell(row=1, column=col_num, value=titulo)
        celda.font = fuente_cabecera
        celda.fill = relleno_salmon
        celda.border = estilo_borde
        celda.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[celda.column_letter].width = max(len(titulo) + 5, 12)

    # 6. Escribir bloques de datos con fila de separación
    fila_actual = 2

    for t in tipos:
        df_bloque = bloques[t]

        if df_bloque.empty:
            continue

        for _, row in df_bloque.iterrows():
            # Mapeo de columnas según tu especificación
            ws.cell(row=fila_actual, column=1, value=row["Fecha"])  # Col A
            ws.cell(row=fila_actual, column=2, value=row["Tipo"])  # Col B
            ws.cell(row=fila_actual, column=3, value=row["Número"])  # Col C
            ws.cell(row=fila_actual, column=5, value=row["Concepto"])  # Col E (RFC)
            ws.cell(
                row=fila_actual, column=6, value=row["Referencia"]
            )  # Col F (Factura)
            ws.cell(
                row=fila_actual,
                column=11,
                value=row["Cargos"]
                if pd.notna(row["Cargos"]) and row["Cargos"] != ""
                else -row["Abonos"],
            )  # Col K (IVA 16%)

            # --- INSERTAR FÓRMULAS DINÁMICAS ---
            # Nota: Los índices del diccionario se suman +1 para coincidir con las columnas de Excel
            formulas = {
                9: lambda row: f"=$K{row}/0.16",  # Subtotal base 16%
                15: lambda row: (
                    f"=$I{row} + $K{row} - SUM(L{row}:N{row})"
                ),  # Total (según tu índice 12)
                18: lambda row: f'=IF(V{row}="","",V{row}/1.16)',  # Subtotal USD
                20: lambda row: f'=IF($R{row}="","",$R{row}*0.16)',  # IVA USD
                22: lambda row: (
                    f'=IF(OR(P{row}="MXN",P{row}="",Q{row}=""),"",O{row}/Q{row})'
                ),  # Total USD
                23: lambda row: f"=$A{row}",  # Fecha Pago
                24: lambda row: (
                    f'=IF($P{row}="MXN","Banco HSBC 8881",IF($P{row}="USD","Banco Base 2018",""))'
                ),  # Banco
            }

            for col_idx, formula in formulas.items():
                celda_f = ws.cell(row=fila_actual, column=col_idx)
                celda_f.value = formula(fila_actual)

            # Aplicar fuente 8pt a toda la fila
            for c in range(1, 27):
                ws.cell(row=fila_actual, column=c).font = fuente_datos

            fila_actual += 1

        # SALTO DE FILA: Al terminar un bloque, dejamos una fila vacía
        fila_actual += 1

    # --- CONFIGURACIÓN FINAL ---
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = (
        f"A1:{ws.cell(row=1, column=len(cabecera_final)).column_letter}1"
    )

    wb.save(archivo_salida)
    print(f"Archivo '{archivo_salida}' generado exitosamente (RFC genérico excluido).")


def procesar_iva_acreditable(file_principal, file_b):
    # 1. Configuración de rutas
    # dir_path = "iva/2025/oct"
    # assets_path = "iva/assets"

    # file_principal = f"{dir_path}/cedula_principales_proveedores.xlsx"
    # file_b = f"{assets_path}/Archivo_B.xlsx"

    print("Cargando archivos...")
    # Cargamos el principal con pandas para procesar la lógica
    df_principal = pd.read_excel(file_principal)
    df_b = pd.read_excel(file_b)

    # --- MAPEO DE COLUMNAS POR POSICIÓN (Base 0 para Pandas) ---
    # Usamos estas variables para saber qué columnas actualizar en el Excel final
    idx_prov = 3  # D
    idx_rfc = 4  # E
    idx_fact = 5  # F
    idx_uuid = 6  # G
    idx_conc = 7  # H
    idx_iva_l = 10  # K
    idx_r_iva = 11  # L
    idx_isr = 12  # M
    idx_loc = 13  # N
    idx_mon = 15  # P
    idx_tot = 21  # V

    # Columnas del Archivo B
    uuid_col_b = df_b.columns[12]
    serie_col_b = df_b.columns[14]
    folio_col_b = df_b.columns[15]
    rfc_col_b = df_b.columns[34]
    nombre_col_b = df_b.columns[35]
    concepto_col_b = df_b.columns[69]
    iva_col_b = df_b.columns[82]
    ret_iva_col_b = df_b.columns[86]
    isr_ret_col_b = df_b.columns[87]
    loc_ret_col_b = df_b.columns[88]
    total_xml_col_b = df_b.columns[91]
    moneda_col_b = df_b.columns[94]

    def limpiar_factura(texto):
        if pd.isna(texto):
            return ""
        texto = str(texto).upper().strip()
        if texto.endswith(".0"):
            texto = texto[:-2]
        texto = re.sub(r"^F\s*-\s*|^F\s+|^F(?=\d)", "", texto)
        match = re.search(r"([A-Z]*\s*\d+)", texto)
        return match.group(0).strip() if match else texto

    def limpiar_monto(valor):
        try:
            return abs(float(valor))
        except:
            return 0.0

    print("Limpiando y preparando datos...")

    rfc_a_excluir = "HMI950125KG8"
    df_principal["rfc_clean"] = (
        df_principal.iloc[:, idx_rfc].astype(str).str.strip().str.upper()
    )
    df_principal["factura_clean"] = df_principal.iloc[:, idx_fact].apply(
        limpiar_factura
    )

    df_b[rfc_col_b] = df_b[rfc_col_b].astype(str).str.strip().str.upper()
    df_b_filtrado = df_b[df_b[rfc_col_b] != rfc_a_excluir].copy()

    # 2. Estructura de búsqueda Robusta
    mapeo_complejo = {}
    for _, row in df_b_filtrado.iterrows():
        rfc = str(row[rfc_col_b])
        uuid = str(row[uuid_col_b]).strip().upper()
        info_row = {
            "uuid": uuid,
            "nombre": str(row[nombre_col_b]).strip().upper(),
            "concepto": str(row[concepto_col_b]).strip(),
            "iva_log": limpiar_monto(row[iva_col_b]),
            "ret_iva": row[ret_iva_col_b],
            "isr_ret": row[isr_ret_col_b],
            "loc_ret": row[loc_ret_col_b],
            "moneda": row[moneda_col_b],
            "total_xml": row[total_xml_col_b],
        }
        serie = (
            str(row[serie_col_b]).strip().upper() if pd.notna(row[serie_col_b]) else ""
        )
        folio = str(row[folio_col_b]).strip().upper()
        if folio.endswith(".0"):
            folio = folio[:-2]

        if rfc not in mapeo_complejo:
            mapeo_complejo[rfc] = {
                "llaves": {},
                "detalles": {},
                "nombre_defecto": info_row["nombre"],
            }

        mapeo_complejo[rfc]["detalles"][uuid] = info_row
        if folio != "NAN" and folio != "":
            mapeo_complejo[rfc]["llaves"][folio] = uuid
            if serie and serie != "NAN":
                mapeo_complejo[rfc]["llaves"][f"{serie} {folio}"] = uuid
                mapeo_complejo[rfc]["llaves"][f"{serie}{folio}"] = uuid

    lista_rfcs_b = list(mapeo_complejo.keys())
    memo_fuzzy_rfc = {}

    # 3. Función de búsqueda
    def buscar_y_actualizar(row):
        rfc_actual = row["rfc_clean"]
        out = {
            "uuid": row.iloc[idx_uuid],
            "prov": row.iloc[idx_prov],
            "conc": row.iloc[idx_conc],
            "r_iva": row.iloc[idx_r_iva],
            "isr": row.iloc[idx_isr],
            "loc": row.iloc[idx_loc],
            "mon": row.iloc[idx_mon],
            "tot": row.iloc[idx_tot],
        }
        if rfc_actual == rfc_a_excluir or rfc_actual == "NAN" or not rfc_actual:
            return pd.Series(list(out.values()))

        if rfc_actual in memo_fuzzy_rfc:
            rfc_match = memo_fuzzy_rfc[rfc_actual]
        else:
            match = process.extractOne(
                rfc_actual,
                lista_rfcs_b,
                processor=utils.default_process,
                score_cutoff=90,
            )
            rfc_match = match[0] if match else None
            memo_fuzzy_rfc[rfc_actual] = rfc_match

        if rfc_match:
            datos_prov = mapeo_complejo[rfc_match]
            factura_in = str(row["factura_clean"])
            iva_in = limpiar_monto(row.iloc[idx_iva_l])
            uuid_found = None

            if factura_in != "":
                if factura_in in datos_prov["llaves"]:
                    uuid_found = datos_prov["llaves"][factura_in]
                else:
                    for ll, uv in datos_prov["llaves"].items():
                        if ll in factura_in or factura_in in ll:
                            uuid_found = uv
                            break
            if not uuid_found and factura_in != "":
                for u_c in datos_prov["detalles"].keys():
                    u_s = u_c.replace("-", "")
                    f_s = factura_in.replace(" ", "")
                    if (
                        u_c.startswith(factura_in)
                        or u_c.endswith(factura_in)
                        or u_s.startswith(f_s)
                        or u_s.endswith(f_s)
                    ):
                        uuid_found = u_c
                        break
            if (
                not uuid_found
                and factura_in != ""
                and (
                    rfc_match.endswith(factura_in)
                    or rfc_match.endswith(factura_in.replace(" ", ""))
                )
            ):
                uuid_found = next(iter(datos_prov["detalles"].keys()))
            if not uuid_found and iva_in > 0:
                for u_c, info in datos_prov["detalles"].items():
                    if iva_in == info["iva_log"] or int(iva_in) == int(info["iva_log"]):
                        uuid_found = u_c
                        break
                    diff = abs(iva_in - info["iva_log"])
                    if (1 - (diff / iva_in) if iva_in != 0 else 0) >= 0.90:
                        uuid_found = u_c
                        break

            if uuid_found:
                res = datos_prov["detalles"][uuid_found]
                return pd.Series(
                    [
                        res["uuid"],
                        res["nombre"],
                        res["concepto"],
                        res["ret_iva"],
                        res["isr_ret"],
                        res["loc_ret"],
                        res["moneda"],
                        res["total_xml"],
                    ]
                )
            else:
                out["prov"] = datos_prov["nombre_defecto"]
                return pd.Series(list(out.values()))
        return pd.Series(list(out.values()))

    print("Procesando lógica de búsqueda...")
    columnas_destino = [
        "tmp_uuid",
        "tmp_prov",
        "tmp_conc",
        "tmp_riva",
        "tmp_isr",
        "tmp_loc",
        "tmp_mon",
        "tmp_tot",
    ]
    df_principal[columnas_destino] = df_principal.apply(buscar_y_actualizar, axis=1)

    # 4. APLICAR CAMBIOS AL EXCEL ORIGINAL PRESERVANDO ESTILOS Y FÓRMULAS
    print("Abriendo archivo original para inyectar datos...")
    wb = openpyxl.load_workbook(
        file_principal, data_only=False
    )  # data_only=False preserva fórmulas
    ws = wb.active  # O usa wb["NombreDeTuHoja"]

    # Mapeo de columnas de Pandas (Base 0) a Excel (Base 1)
    # Ejemplo: idx_uuid es 6 (G), en Excel es columna 7.
    mapeo_excel = {
        idx_uuid: "tmp_uuid",
        idx_prov: "tmp_prov",
        idx_conc: "tmp_conc",
        idx_r_iva: "tmp_riva",
        idx_isr: "tmp_isr",
        idx_loc: "tmp_loc",
        idx_mon: "tmp_mon",
        idx_tot: "tmp_tot",
    }

    print("Escribiendo datos en celdas específicas...")
    # Empezamos en la fila 2 (asumiendo que la 1 son encabezados)
    for i, row_data in df_principal.iterrows():
        excel_row = i + 2
        for col_idx, tmp_col_name in mapeo_excel.items():
            # Solo escribimos si el valor es diferente al original para ser eficientes
            nuevo_valor = row_data[tmp_col_name]
            ws.cell(row=excel_row, column=col_idx + 1).value = nuevo_valor

    # 5. Guardado final sobre el mismo archivo
    print(f"Guardando cambios en {file_principal}...")
    wb.save(file_principal)

    print(f"\n¡Proceso terminado con éxito!")
    print(
        "Se han actualizado los datos directamente en el archivo original sin afectar estilos ni fórmulas."
    )


def procesar_genericos(archivo):
    # Ruta del archivo
    # dir = "iva/2025/oct"
    # archivo = f"{dir}/cedula_principales_proveedores.xlsx"
    hoja_nombre = "Hoja1"

    # Regex para la factura numérica compleja (se mantiene)
    patron_factura = re.compile(r"^\s*\d{1,3}(?:\s+\d{1,8})+\s+\d{6,}\s*$")

    # Patrón para buscar palabras clave en Factura (COMISION, VENTAS, DEPOSITO)
    patron_palabras = re.compile(r"\b(?:COMISION|VENTAS|DEPOSITO)\b", re.IGNORECASE)

    # Cargamos el libro en modo de preservación de formato
    libro = openpyxl.load_workbook(archivo, keep_vba=False, data_only=False)

    # Verificamos si existe la hoja
    if hoja_nombre not in libro.sheetnames:
        raise ValueError(
            f"La hoja '{hoja_nombre}' no existe en el archivo. Hojas disponibles: {libro.sheetnames}"
        )

    hoja = libro[hoja_nombre]

    # Iterar sobre las filas desde la 2 (asumiendo encabezados en fila 1)
    for fila in range(2, hoja.max_row + 1):
        rfc_celda = hoja.cell(row=fila, column=5)  # Columna E (RFC)
        rfc = rfc_celda.value
        if not isinstance(rfc, str):
            continue  # Saltar si RFC no es cadena

        # === NORMALIZACIÓN DEL RFC ===
        rfc_limpio = str(rfc).strip().upper()

        proveedor_celda = hoja.cell(row=fila, column=4)  # Columna D (Proveedor)
        factura_celda = hoja.cell(row=fila, column=6)  # Columna F (Factura)
        folio_celda = hoja.cell(row=fila, column=7)  # Columna G (FOLIO FISCAL)
        moneda_celda = hoja.cell(row=fila, column=16)  # Columna P (Moneda)

        if rfc_limpio == "HMI950125KG8":
            # RFC específico: HSBC
            proveedor_celda.value = (
                "HSBC MEXICO SA INSTITUCION DE BANCA MULTIPLE GRUPO FINANCIERO HSBC"
            )
            folio_celda.value = "-"
            moneda_celda.value = "MXN"

        elif rfc_limpio == "XAXX010101000":
            factura = factura_celda.value

            # Condición 1: Cumple con el patrón numérico original
            cond1 = isinstance(factura, str) and patron_factura.match(factura.strip())

            # Condición 2: Contiene COMISION, VENTAS o DEPOSITO (insensible a mayúsculas)
            cond2 = isinstance(factura, str) and patron_palabras.search(factura.strip())

            # Condición 3: Factura es vacía (None o solo espacios)
            cond3 = factura is None or (
                isinstance(factura, str) and factura.strip() == ""
            )

            if cond1 or cond2 or cond3:
                proveedor_celda.value = "Publico en General"
                folio_celda.value = "-"
                moneda_celda.value = "MXN"

    # Guardar sobre el mismo archivo
    libro.save(archivo)

    print(
        f"Archivo '{archivo}' actualizado correctamente sin perder formato ni fórmulas."
    )


def main_iva_process(self, path):
    path = Path(path).resolve()  # convierte a Path y resuelve rutas absolutas

    file_name = "cedula_iva_acreditable_100.xlsx"
    file_path = path / file_name  # usa Path para juntar rutas correctamente

    diot_path = path / "diot.xlsx"
    egresos_path = path / "egresos.xlsx"

    # ============================================================
    # VALIDACIÓN PREVIA (Aquí detona y sale antes de intentar nada)
    # Si no existen los insumos, lanzamos el error de inmediato
    # ============================================================
    if not diot_path.exists() or not egresos_path.exists():
        faltante = "diot.xlsx" if not diot_path.exists() else "egresos.xlsx"
        msg = f"No se encontró el archivo necesario: {faltante}"
        self.text_console_log(msg, "ERROR")
        # Este RAISE detiene la función AQUÍ MISMO y va directo al except de la GUI
        raise FileNotFoundError(msg)

    # ============================================================
    # 1. Procesar DIOT (el que genera el archivo problemático)
    # ============================================================
    try:
        procesar_diot_template(diot_path, file_path)
        print(f"✅ DIOT procesado correctamente → {file_path}")
        self.text_console_log(
            f"✅ DIOT procesado correctamente → {file_path}", "PROCESS"
        )
    except Exception as e:
        print(f"❌ Error en procesar_diot_template: {e}")
        self.text_console_log(f"❌ Error en procesar_diot_template: {e}", "ERROR")
        raise FileNotFoundError("¡Es imposible seguir sin el archivo DIOT!")
        # return  # o raise, según prefieras

    # ============================================================
    # 2. Verificar que el archivo exista ANTES de continuar
    # ============================================================
    if not file_path.exists():
        print(f"❌ ERROR: El archivo no fue creado: {file_path}")
        print(
            "Revisa la función 'procesar_diot_template' (posible fallo silencioso o ruta incorrecta)"
        )
        self.text_console_log(
            f"❌ ERROR: El archivo no fue creado: {file_path}", "ERROR"
        )
        self.text_console_log(
            "Revisa la función 'procesar_diot_template' (posible fallo silencioso o ruta incorrecta)",
            "ERROR",
        )
        raise FileNotFoundError(f"El archivo no fue creado: {file_path.name}")

    # ============================================================
    # 3. Llamar a las siguientes funciones (ahora es seguro)
    # ============================================================
    try:
        procesar_iva_acreditable(file_path, egresos_path)
        procesar_genericos(file_path)
        print("✅ Proceso IVA completado exitosamente")
        self.text_console_log("✅ Proceso IVA completado exitosamente", "PROCESS")
    except FileNotFoundError as e:
        print(f"❌ Archivo no encontrado: {e}")
        self.text_console_log(f"❌ Archivo no encontrado: {e}", "INFO")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        self.text_console_log(f"❌ Error inesperado: {e}", "ERROR")

    # procesar_iva_acreditable(file_path, egresos_path)
    # procesar_genericos(file_path)
