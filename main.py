import re
import os
import openpyxl
import unicodedata
from datetime import datetime
from collections import defaultdict
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN Y DICCIONARIOS ---
MESES_ES = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL",
    "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO",
    "09": "SEPTIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE",
}

def _convertir_num(valor):
    if valor is None: return 0
    if isinstance(valor, (int, float)): return valor
    if isinstance(valor, str):
        try:
            return float(valor.replace(",", ".").strip())
        except ValueError:
            return 0
    return 0

def normalizar(texto):
    if not texto: return ""
    texto = str(texto).upper().replace(" ", "")
    # Quitar tildes
    return "".join(c for c in unicodedata.normalize('NFD', texto)
                  if unicodedata.category(c) != 'Mn')

# --- FUNCIONES DE APOYO ---

def obtener_diccionario_metodos(ws, dia):
    """Busca la fila donde empieza el bloque del día en LOCALES DETALLE."""
    metodos = {}
    bloque_encontrado = False
    for row in range(1, ws.max_row + 1):
        cell_val = str(ws.cell(row=row, column=1).value).strip()
        if cell_val == str(dia):
            bloque_encontrado = True
        if bloque_encontrado:
            metodo_nombre = ws.cell(row=row, column=2).value
            if metodo_nombre:
                norm = normalizar(metodo_nombre)
                metodos[norm] = (metodo_nombre, row)
            if cell_val == "TOTAL": break
    return metodos

def detectar_cajas(ws_cierre):
    """Detecta en qué columnas del CIERRE están las cajas."""
    cajas = []
    for col in range(3, ws_cierre.max_column + 1):
        val = ws_cierre.cell(row=3, column=col).value
        if val and "RYV" in str(val):
            m = re.search(r'(\d+)', str(val))
            if m:
                cajas.append((col, str(int(m.group(1)))))
    return cajas

# --- FUNCIÓN DE CONTROL DE EFECTIVO ---

def transfer_control_efectivo(ventas_wb, cierre_files):
    """
    Procesa los vendedores y efectivo de forma dinámica hacia CONTROL DE EFECTIVO.
    Agrupa los vendedores del mismo día hacia abajo.
    """
    try:
        control_ws = ventas_wb["CONTROL DE EFECTIVO"]
    except:
        print("⚠️ No se encontró la hoja 'CONTROL DE EFECTIVO'")
        return

    verde_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    amarillo_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # NUEVO: Diccionario para llevar el registro de la siguiente fila libre por cada día.
    # Arranca en la fila 4 para todos los días.
    filas_por_dia = defaultdict(lambda: 4)

    for local, day, mes_nombre, cierre_file in cierre_files:
        try:
            cierre_wb = openpyxl.load_workbook(cierre_file, data_only=True)
            cierre_ws = cierre_wb.active

            # Búsqueda dinámica de filas en el archivo de entrada
            f_vend = f_caja = f_efec = None
            for r in range(1, 15):
                v = str(cierre_ws.cell(row=r, column=1).value).strip().upper()
                if v == "VENDEDOR": f_vend = r
                elif "ARQUEO Z" in v or "CAJA" in v: f_caja = r
                elif v == "EFECTIVO": f_efec = r

            if not f_vend or not f_efec: continue

            # Columna de inicio para este día
            start_col = ((day - 1) * 4) + 1
            
            for col in range(3, cierre_ws.max_column + 1):
                vendedor = cierre_ws.cell(row=f_vend, column=col).value
                monto = _convertir_num(cierre_ws.cell(row=f_efec, column=col).value)
                caja_raw = cierre_ws.cell(row=f_caja, column=col).value if f_caja else ""

                if vendedor and monto > 0:
                    # Obtenemos la fila destino actual para este día
                    row_dest = filas_por_dia[day]
                    caja_clean = str(caja_raw).split("-")[0].strip() if caja_raw else ""

                    # Asignamos: start_col = Vendedor, start_col+1 = Caja, start_col+2 = Monto
                    control_ws.cell(row=row_dest, column=start_col).value = vendedor
                    control_ws.cell(row=row_dest, column=start_col + 1).value = caja_clean
                    control_ws.cell(row=row_dest, column=start_col + 2).value = monto

                    # Aplicamos color de relleno
                    fill = verde_fill if str(local) == "905" else amarillo_fill
                    for c_style in range(start_col, start_col + 3):
                        control_ws.cell(row=row_dest, column=c_style).fill = fill
                    
                    # INCREMENTAMOS el contador de filas PARA ESTE DÍA, 
                    # así el próximo vendedor (o el próximo archivo del mismo día) irá abajo
                    filas_por_dia[day] += 1
                    
            print(f"✓ Efectivo actualizado para el local {local}, día {day} en columna {get_column_letter(start_col)}")
            
        except Exception as e:
            print(f"❌ Error en efectivo {cierre_file}: {e}")

# --- FUNCIÓN PRINCIPAL ---

def transfer_cierre_to_ventas():
    # 1. Identificar archivos
    cierre_files = []
    pattern = re.compile(r"CIERRE\s*TOTAL\s+(\d+)\s+(\d{2})-(\d{2})", re.IGNORECASE)

    for file in os.listdir():
        if file.lower().endswith((".xlsx", ".xls")):
            match = pattern.search(file)
            if match:
                local, dia, mes = match.group(1), match.group(2), match.group(3)
                mes_nombre = MESES_ES.get(mes)
                if mes_nombre:
                    cierre_files.append((local, int(dia), mes_nombre, file))

    if not cierre_files:
        print("⚠️ No se encontraron archivos de cierre.")
        return

    # Usar el mes del primer archivo para determinar el destino
    archivo_destino = f"{cierre_files[0][2]} {datetime.now().year}.xlsx"
    
    try:
        ventas_wb = openpyxl.load_workbook(archivo_destino)
        ventas_ws = ventas_wb["LOCALES DETALLE"]
    except Exception as e:
        print(f"❌ Error al cargar destino: {e}")
        return

    # Mapeo de columnas en LOCALES DETALLE (Fila 2)
    dest_cols_map = defaultdict(list)
    for col in range(3, ventas_ws.max_column + 1):
        header = ventas_ws.cell(row=2, column=col).value
        if header:
            m = re.search(r'(\d+)', str(header))
            if m:
                dest_cols_map[str(int(m.group(1)))].append(col)

    # Procesar LOCALES DETALLE
    for local, day, mes_nombre, cierre_file in cierre_files:
        try:
            cierre_wb = openpyxl.load_workbook(cierre_file, data_only=True)
            cierre_ws = cierre_wb.active
            
            metodo_pago_map = obtener_diccionario_metodos(ventas_ws, day)
            cajas_origen = detectar_cajas(cierre_ws)

            for row in range(1, cierre_ws.max_row + 1):
                metodo_orig = cierre_ws.cell(row=row, column=1).value
                if not metodo_orig: continue
                
                norm_orig = normalizar(metodo_orig)
                if norm_orig in metodo_pago_map:
                    target_row = metodo_pago_map[norm_orig][1]
                    
                    counters = defaultdict(int)
                    for col_orig, caja_num in cajas_origen:
                        valor = cierre_ws.cell(row=row, column=col_orig).value
                        dest_list = dest_cols_map.get(caja_num)
                        
                        if dest_list:
                            idx = counters[caja_num]
                            if idx < len(dest_list):
                                ventas_ws.cell(row=target_row, column=dest_list[idx]).value = _convertir_num(valor)
                                counters[caja_num] += 1
            print(f"✓ Detalle procesado: {cierre_file}")
        except Exception as e:
            print(f"❌ Error procesando detalle {cierre_file}: {e}")

    # --- LLAMADA A LA FUNCIÓN DE CONTROL DE EFECTIVO ---
    print("\nActualizando Control de Efectivo...")
    transfer_control_efectivo(ventas_wb, cierre_files)

    # GUARDAR TODO
    try:
        ventas_wb.save(archivo_destino)
        print(f"\n🚀 PROCESO FINALIZADO EXITOSAMENTE: {archivo_destino}")
    except Exception as e:
        print(f"❌ Error al guardar: {e}")

if __name__ == "__main__":
    transfer_cierre_to_ventas()