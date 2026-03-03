import re
import os
import openpyxl
import unicodedata
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


MESES_ES = {
    "01": "ENERO",
    "02": "FEBRERO",
    "03": "MARZO",
    "04": "ABRIL",
    "05": "MAYO",
    "06": "JUNIO",
    "07": "JULIO",
    "08": "AGOSTO",
    "09": "SEPTIEMBRE",
    "10": "OCTUBRE",
    "11": "NOVIEMBRE",
    "12": "DICIEMBRE",
}

def _convertir_num(valor):
    if valor is None:
        return 0
    if isinstance(valor, (int, float)):
        return valor
    if isinstance(valor, str):
        try:
            return float(valor.replace(",", "."))
        except ValueError:
            return 0
    return 0

def normalizar(texto):
    if not texto:
        return ""
    # 1. Convertir a mayúsculas y quitar espacios
    texto = str(texto).upper().replace(" ", "")
    # 2. Quitar conectores comunes que causan discordancia (como "DE")
    texto = texto.replace("DE", "")
    # 3. Quitar tildes
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                    if unicodedata.category(c) != 'Mn')
    # 4. Limpiar cualquier caracter que no sea letra o número
    texto = re.sub(r'[^A-Z0-9]', '', texto)
    return texto

def obtener_diccionario_metodos(ventas_ws, day):
    """
    Obtiene los métodos de pago para el día indicado de manera independiente,
    calculando el bloque de filas usando la columna B y las filas que contienen TOTAL.
    Genera dinámicamente los nombres de los métodos a partir de las celdas de la columna B.
    """

    # Detectar todas las filas donde aparece 'TOTAL' en columna 1 o 2 (merge)
    total_indices = []
    for fila in range(3, ventas_ws.max_row + 1):
        val_col1 = ventas_ws.cell(row=fila, column=1).value
        val_col2 = ventas_ws.cell(row=fila, column=2).value
        val = val_col1 or val_col2
        if val and "TOTAL" in str(val).strip().upper():
            total_indices.append(fila)

    if not total_indices:
        print("⚠️ No se encontraron filas con 'TOTAL' en columna 1 o 2.")
        return {}

    # Calcular inicio y fin del bloque para el día indicado
    inicio = 3 if day == 1 else total_indices[day - 2] + 1
    fin = total_indices[day - 1] # - 1 if day - 1 < len(total_indices) else ventas_ws.max_row

    # Construir lista de filas del bloque
    filas_dia = list(range(inicio, fin + 1))

    # Extraer dinámicamente los nombres de métodos de la columna B del bloque
    nombres_metodos = []
    for fila in filas_dia:
        val_col2 = ventas_ws.cell(row=fila, column=2).value
        val_col1 = ventas_ws.cell(row=fila, column=1).value
        val = val_col2 or val_col1
        if val and str(val).strip():
            nombres_metodos.append(str(val).strip())

    # Mapear método -> (nombre_sin_espacios, fila correspondiente en el bloque)
    metodo_pago_map = {}
    for idx, metodo in enumerate(nombres_metodos):
        metodo_pago_map[normalizar(metodo)] = (metodo.replace(" ", ""), filas_dia[idx])

    print(f"\nMétodos de pago para el día {day} (filas {inicio}-{fin}):")
    for k, (nombre, fila) in metodo_pago_map.items():
        print(f" - {nombre} (fila {fila})")

    return metodo_pago_map

def detectar_cajas(cierre_ws):
    """
    Detecta las cajas en el archivo de cierre.
    Soporta:
      - 'RYV 905-021'   -> toma 021 -> '21'
      - '071 - 1906'    -> toma 071 -> '71'
      - '071' (solo)   -> toma 071 -> '71'
    Intenta primero la fila 3 (tal como tienes), y si no encuentra nada, prueba filas 1,2,4.
    Devuelve lista de tuplas (col_index, caja_str) en el mismo orden que aparecen en el archivo.
    """
    def try_row(row):
        found = []
        for col in range(3, cierre_ws.max_column + 1):
            cell = cierre_ws.cell(row=row, column=col).value
            if cell is None:
                continue
            raw = str(cell).strip()
            raw_up = raw.upper()

            # 1) Formato RYV ...-NNN  -> tomar último número
            if "RYV" in raw_up:
                m = re.search(r'(\d+)\s*$', raw)
                if m:
                    caja_num = str(int(m.group(1)))
                    found.append((col, caja_num))
                continue

            # 2) Formato '071 - 1906'  -> tomar parte izquierda antes del guion
            if "-" in raw:
                left = raw.split("-", 1)[0].strip()
                m = re.search(r'(\d+)', left)
                if m:
                    caja_num = str(int(m.group(1)))
                    found.append((col, caja_num))
                else:
                    # si no hay dígitos a la izquierda, intentar derecha
                    m2 = re.search(r'(\d+)', raw)
                    if m2:
                        caja_num = str(int(m2.group(1)))
                        found.append((col, caja_num))
                continue

            # 3) Formato solo número '071' o '71' o ' 071 '
            m = re.search(r'(\d+)', raw)
            if m:
                caja_num = str(int(m.group(1)))
                found.append((col, caja_num))
                continue

            # si no cumple ninguno, se salta
        return found

    # Intentar fila 3 primero (tu caso esperado)
    cajas = try_row(3)

    # Si no detectó nada, intentar filas alternativas (por robustez)
    if not cajas:
        for r in (1, 2, 4):
            cajas = try_row(r)
            if cajas:
                print(f"Info: detección de cajas falló en fila 3, se usó fila {r} como respaldo.")
                break

    # Si aún no hay cajas, imprimir debug útil y devolver vacío
    if not cajas:
        sample = []
        # mostrar primeras 12 columnas para inspección
        for col in range(1, min(13, cierre_ws.max_column + 1)):
            sample.append((col, cierre_ws.cell(row=3, column=col).value))
        print("⚠️ No se detectaron cajas en la fila 3 ni filas alternativas.")
        print("Muestra fila 3 (col, valor):", sample)
        return []

    # Debug detallado: col -> raw -> caja
    print("Debug detección cajas (col -> raw -> caja):")
    for col, caja in cajas:
        raw = cierre_ws.cell(row=3, column=col).value
        print(f" - col {col}: {raw!r} -> caja {caja}")

    return cajas

def transfer_cierre_to_ventas():
    # Buscar archivos CIERRE TOTAL
    cierre_files = []
    pattern = re.compile(r"CIERRE\s*TOTAL\s+(\d+)\s+(\d{2})-(\d{2})", re.IGNORECASE)

    for file in os.listdir():
        if file.lower().endswith((".xls", ".xlsx")):
            match = pattern.search(file)
            if match:
                local, dia, mes = match.group(1), match.group(2), match.group(3)
                mes_nombre = MESES_ES.get(mes)
                if not mes_nombre:
                    print(f"⚠️ No se reconoce el mes '{mes}' en archivo {file}")
                    continue
                cierre_files.append((local, int(dia), mes_nombre, file))

    if not cierre_files:
        print("⚠️ No se encontraron archivos CIERRE TOTAL")
        return

    print("\nArchivos encontrados:")
    for local, dia, mes_nombre, fname in cierre_files:
        print(f" - Local {local}, Día {dia}, Mes {mes_nombre}, archivo: {fname}")

    # Abrir archivo destino (usando el mes del último archivo encontrado)
    ano_actual = datetime.now().year
    archivo_destino = f"{mes_nombre} {ano_actual}.xlsx"
    try:
        ventas_wb = openpyxl.load_workbook(archivo_destino)
        ventas_ws = ventas_wb["LOCALES DETALLE"]
        print(f"✓ Archivo de destino cargado: {archivo_destino}")
    except Exception as e:
        print(f"❌ Error al cargar el archivo de destino {archivo_destino}: {e}")
        return

    # Construir mapa de columnas destino: número de caja -> [col1, col2, ...]
    dest_cols_map = defaultdict(list)
    for col in range(3, ventas_ws.max_column + 1):
        header = ventas_ws.cell(row=2, column=col).value
        if header is None:
            continue
        header_str = str(header).strip()
        m = re.search(r'(\d+)', header_str)
        if not m:
            continue
        key = str(int(m.group(1)))  # normalizar (ej "021" -> "21")
        dest_cols_map[key].append(col)

    print("\nMapa de columnas destino por caja (header fila 2):")
    for k, cols in dest_cols_map.items():
        print(f" - Caja {k}: columnas {cols}")

    # Procesar cada archivo de cierre
    for local, day, mes_nombre, cierre_file in cierre_files:
        try:
            cierre_wb = openpyxl.load_workbook(cierre_file, data_only=True)
            cierre_ws = cierre_wb.active

            # Obtener métodos de pago del día en destino (fila correcta según bloque)
            metodo_pago_map = obtener_diccionario_metodos(ventas_ws, day)
            if not metodo_pago_map:
                print(f"⚠️ No hay métodos detectados para el día {day} en destino. Saltando archivo {cierre_file}")
                continue

            # Identificar columnas de cada caja en el cierre (origen)
            cajas = detectar_cajas(cierre_ws)
            print(f"\nCajas detectadas en {cierre_file}: {[c for _, c in cajas]}")

            # Extraer valores por método de pago desde el cierre
            cierre_data = {}
            for metodo_norm, (metodo_destino, target_row) in metodo_pago_map.items():
                valores = []
                encontrado = False
                for row in range(1, cierre_ws.max_row + 1):
                    metodo_cell = cierre_ws.cell(row=row, column=1).value
                    if not metodo_cell:
                        continue
                    if normalizar(str(metodo_cell).strip()) == metodo_norm:
                        encontrado = True
                        for col, caja_num in cajas:
                            cell_value = cierre_ws.cell(row=row, column=col).value
                            valores.append((caja_num, cell_value))
                        break
                if not encontrado:
                    for col, caja_num in cajas:
                        valores.append((caja_num, None))
                cierre_data[metodo_destino] = (target_row, valores)

            print("\nResumen de match de métodos de pago:")
            print(f" - Métodos esperados: {len(metodo_pago_map)}")
            print(f" - Métodos encontrados en CIERRE: {len(cierre_data)}")

            # Transferir datos al archivo de destino
            for metodo, (target_row, valores) in cierre_data.items():
                if target_row is None:
                    print(f"⚠️ No hay fila destino para el método {metodo}. Se omite.")
                    continue
                print(f"\n--- Escribiendo datos para {metodo} ---")
                counters = defaultdict(int)
                for caja_num, valor in valores:
                    dest_list = dest_cols_map.get(str(caja_num))
                    if not dest_list:
                        print(f"⚠️ No se encontró columna destino para la caja {caja_num}")
                        continue
                    idx = counters[str(caja_num)]
                    if idx >= len(dest_list):
                        print(f"⚠️ No hay columnas suficientes en destino para la caja {caja_num} (ocurrencia {idx})")
                        continue
                    col_to_write = dest_list[idx]
                    ventas_ws.cell(row=target_row, column=col_to_write).value = _convertir_num(valor)
                    print(f"Escribiendo {valor if valor is not None else 0} en Fila {target_row}, Columna {get_column_letter(col_to_write)} (Caja {caja_num}, occ {idx + 1})")
                    counters[str(caja_num)] += 1

            print(f"Datos transferidos correctamente para el local {local}, día {day}")

        except Exception as e:
            print(f"Error al procesar {cierre_file}: {e}")

    transfer_control_efectivo(cierre_files, archivo_destino)

    # Guardar cambios en el archivo destino
    try:
        ventas_wb.save(archivo_destino)
        print(f"\n✓ Archivo guardado exitosamente: {archivo_destino}")
    except Exception as e:
        print(f"❌ Error al guardar archivo: {e}")

def transfer_control_efectivo(cierre_files, archivo_destino):
    """
    Procesa los archivos de cierre y transfiere la info de vendedores, cajas y efectivo
    a la hoja CONTROL DE EFECTIVO del archivo destino.
    Cada día ocupa un bloque de 3 columnas consecutivas: A-C, E-G, I-K, ...
    Cada cajero se coloca uno debajo del otro: fila 4, 5, 6, ...
    """
    try:
        ventas_wb = openpyxl.load_workbook(archivo_destino)
        control_ws = ventas_wb["CONTROL DE EFECTIVO"]
    except Exception as e:
        print(f"❌ Error al cargar el archivo de destino {archivo_destino}: {e}")
        return

    # Colores de relleno según local
    verde_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    amarillo_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for local, day, mes_nombre, cierre_file in cierre_files:
        try:
            cierre_wb = openpyxl.load_workbook(cierre_file, data_only=True)
            cierre_ws = cierre_wb.active

            # Buscar filas clave
            fila_vendedor = fila_caja = fila_efectivo = None
            for fila in range(1, cierre_ws.max_row + 1):
                val = cierre_ws.cell(row=fila, column=1).value
                val_up = str(val).upper() if val else ""
                if val_up == "VENDEDOR":
                    fila_vendedor = fila
                elif val_up.startswith("ARQUEO Z"):
                    fila_caja = fila
                elif val_up == "EFECTIVO":
                    fila_efectivo = fila

            if not fila_vendedor or not fila_caja or not fila_efectivo:
                print(f"⚠️ No se detectaron filas clave en {cierre_file}. Saltando.")
                continue

            # Calcular columna inicial del día según la fórmula estándar
            start_col = int((day - 0.75) * 4)
            print(f"\nProcesando {cierre_file} (Día {day}) → columna inicial: {start_col}")

            # Recorrer las columnas del cierre
            cajero_idx = 0
            for col in range(3, cierre_ws.max_column + 1):
                vendedor = cierre_ws.cell(row=fila_vendedor, column=col).value
                caja_raw = cierre_ws.cell(row=fila_caja, column=col).value
                efectivo = cierre_ws.cell(row=fila_efectivo, column=col).value

                if not vendedor and not caja_raw and not efectivo:
                    continue  # columna vacía

                # Extraer número de caja (parte antes del guion)
                caja_num = str(caja_raw).split("-")[0].strip() if caja_raw else ""

                # Determinar fila de destino según el cajero
                row_dest = 4 + cajero_idx  # fila 4,5,6,...

                # Escribir datos en las 3 columnas del bloque del día
                control_ws.cell(row=row_dest, column=start_col).value = vendedor
                control_ws.cell(row=row_dest, column=start_col + 1).value = caja_num
                control_ws.cell(row=row_dest, column=start_col + 2).value = _convertir_num(efectivo)

                # Aplicar color según local
                fill = verde_fill if str(local) == "905" else amarillo_fill
                for c in range(start_col, start_col + 3):
                    control_ws.cell(row=row_dest, column=c).fill = fill

                print(f"  → Fila {row_dest}: Vendedor={vendedor}, Caja={caja_num}, Efectivo={efectivo}")
                cajero_idx += 1

        except Exception as e:
            print(f"❌ Error procesando {cierre_file}: {e}")

    # Guardar cambios
    try:
        ventas_wb.save(archivo_destino)
        print(f"\n✓ Archivo guardado exitosamente: {archivo_destino}")
    except Exception as e:
        print(f"❌ Error al guardar archivo: {e}")

if __name__ == "__main__":
    transfer_cierre_to_ventas()
    print("Debug")
    # Después de cargar los archivos y definir las variables necesarias


