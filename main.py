import os
import re
import openpyxl
from datetime import datetime

MESES_ES = {
    "01": "ENERO", "02": "FEBRERO", "03": "MARZO", "04": "ABRIL",
    "05": "MAYO", "06": "JUNIO", "07": "JULIO", "08": "AGOSTO",
    "09": "SEPTIEMBRE", "10": "OCTUBRE", "11": "NOVIEMBRE", "12": "DICIEMBRE",
}

def _convertir_num(valor):
    if valor is None or str(valor).strip() == "":
        return 0
    if isinstance(valor, (int, float)):
        return valor
    if isinstance(valor, str):
        try:
            clean_val = valor.replace("$", "").replace(".", "").replace(",", ".")
            return float(clean_val)
        except ValueError:
            return 0
    return 0

def clean_header(text):
    return str(text).strip().upper().replace(" ", "")

def transferir_facturas():
    print("Iniciando proceso de traspaso de facturas...")
    
    facturas_files = []
    pattern = re.compile(r"FACTURAS\s+(\d+)\s+(\d{2})-(\d{2})", re.IGNORECASE)

    for file in os.listdir():
        if file.lower().endswith(".xlsx") and "FACTURAS" in file.upper():
            match = pattern.search(file)
            if match:
                local, dia, mes = match.group(1), match.group(2), match.group(3)
                mes_nombre = MESES_ES.get(mes)
                if mes_nombre:
                    facturas_files.append((local, int(dia), mes_nombre, file))

    if not facturas_files:
        print("⚠️ No se encontraron archivos de FACTURAS (.xlsx) en esta carpeta.")
        input("Presiona Enter para salir...")
        return

    datos_por_dia = {} 
    
    for local, day, mes_nombre, fname in facturas_files:
        print(f"\n--- Leyendo datos de: {fname} ---")
        try:
            wb_origen = openpyxl.load_workbook(fname, data_only=True)
            ws_origen = wb_origen.active

            header_map = {}
            for col in range(1, ws_origen.max_column + 1):
                val = ws_origen.cell(row=1, column=col).value
                if val:
                    header_map[clean_header(val)] = col
            
            col_serie = header_map.get("SERIE") or header_map.get("FECHA")
            col_nombre = header_map.get("NOMBRE") or header_map.get("CLIENTE")
            col_numero = header_map.get("NÚMERO") or header_map.get("NUMERO") or header_map.get("FACTURA")
            col_total = header_map.get("TOTAL") or header_map.get("MONTO")

            if not all([col_serie, col_nombre, col_numero, col_total]):
                print(f"⚠️ Faltan columnas clave en {fname}. Saltando archivo.")
                continue

            if day not in datos_por_dia:
                datos_por_dia[day] = []

            for row in range(2, ws_origen.max_row + 1):
                serie = str(ws_origen.cell(row=row, column=col_serie).value or "").strip()
                nombre = str(ws_origen.cell(row=row, column=col_nombre).value or "").strip()
                numero_val = ws_origen.cell(row=row, column=col_numero).value
                numero = str(numero_val).strip() if numero_val is not None else ""
                total = ws_origen.cell(row=row, column=col_total).value

                if not serie or serie.upper() == "NONE" or "TOTAL" in serie.upper():
                    continue

                datos_por_dia[day].append({
                    'serie': serie,
                    'nombre': nombre,
                    'numero': numero,
                    'total': _convertir_num(total)
                })
        except Exception as e:
            print(f"❌ Error al leer {fname}: {e}")

    if not datos_por_dia:
        print("\n⚠️ No se encontraron datos válidos en los archivos para procesar.")
        input("Presiona Enter para salir...")
        return

    _, _, mes_nombre_dest, _ = facturas_files[0]
    ano_actual = datetime.now().year
    archivo_destino = f"{mes_nombre_dest} {ano_actual}.xlsx"

    try:
        ventas_wb = openpyxl.load_workbook(archivo_destino)
        hoja_nombre = "FACTURA A COBRAR"
        if hoja_nombre in ventas_wb.sheetnames:
            ventas_ws = ventas_wb[hoja_nombre]
            print(f"\n✓ Archivo Maestro cargado: {archivo_destino} | Hoja: {hoja_nombre}")
        else:
            print(f"❌ La hoja '{hoja_nombre}' no existe en {archivo_destino}")
            input("Presiona Enter para salir...")
            return
    except Exception as e:
        print(f"❌ Error al cargar el archivo destino: {e}. ¿Está abierto en Excel? Ciérralo e intenta de nuevo.")
        input("Presiona Enter para salir...")
        return

    dest_header_map = {}
    for row_idx in range(1, 4):
        for col in range(1, ventas_ws.max_column + 1):
            val = ventas_ws.cell(row=row_idx, column=col).value
            if val:
                key = clean_header(val)
                if key not in dest_header_map:
                    dest_header_map[key] = col

    col_dest_fecha = dest_header_map.get("FECHA") or dest_header_map.get("SERIE") or 1
    col_dest_nombre = dest_header_map.get("NOMBRE") or dest_header_map.get("CLIENTE") or 3
    col_dest_nfact = dest_header_map.get("N°FACT") or dest_header_map.get("FACTURA") or dest_header_map.get("NFACT") or 4
    col_dest_monto = dest_header_map.get("MONTO") or dest_header_map.get("TOTAL") or 5

    facturas_agregadas = 0

    for dia in sorted(datos_por_dia.keys(), reverse=True):
        facturas_del_dia = datos_por_dia[dia]
        fila_marcador = None
        col_marcador = None
        
        # 1. Encontrar el marcador del día en la parte de abajo
        for r in range(1, ventas_ws.max_row + 1):
            encontrado = False
            for c in range(1, 4): 
                celda_val = ventas_ws.cell(row=r, column=c).value
                if celda_val is not None:
                    if isinstance(celda_val, (int, float)) and int(celda_val) == int(dia):
                        encontrado = True
                    elif isinstance(celda_val, str):
                        texto_limpio = celda_val.strip()
                        if texto_limpio == str(dia) or texto_limpio == f"{dia}.0":
                            encontrado = True
                
                if encontrado:
                    fila_marcador = r
                    col_marcador = c
                    break
            if encontrado:
                break
        
        if fila_marcador:
            fila_insercion = fila_marcador # Valor por defecto si no encuentra el encabezado
            
            # 2. SUBIR COMO ASCENSOR buscando la fecha o el encabezado
            for fila_arriba in range(fila_marcador - 1, 0, -1):
                val_arriba = ventas_ws.cell(row=fila_arriba, column=col_marcador).value
                es_fecha = False
                
                if val_arriba is not None:
                    # Detectamos si es objeto datetime o texto de fecha/encabezado
                    if isinstance(val_arriba, datetime):
                        es_fecha = True
                    elif isinstance(val_arriba, str):
                        txt = val_arriba.strip().upper()
                        if "FECHA" in txt or "SERIE" in txt or re.search(r'\d{1,2}[-/]\d{1,2}', txt):
                            es_fecha = True
                
                if es_fecha:
                    # Comprobamos la celda aún más arriba para asegurarnos de que estamos en la primera (el tope)
                    val_mas_arriba = ventas_ws.cell(row=fila_arriba - 1, column=col_marcador).value if fila_arriba > 1 else None
                    es_mas_arriba_fecha = False
                    
                    if val_mas_arriba is not None:
                        if isinstance(val_mas_arriba, datetime):
                            es_mas_arriba_fecha = True
                        elif isinstance(val_mas_arriba, str):
                            txt_mas = val_mas_arriba.strip().upper()
                            if "FECHA" in txt_mas or "SERIE" in txt_mas or re.search(r'\d{1,2}[-/]\d{1,2}', txt_mas):
                                es_mas_arriba_fecha = True
                    
                    # Si la de más arriba ya no es fecha, ¡hemos llegado al techo de la lista!
                    if not es_mas_arriba_fecha:
                        fila_insercion = fila_arriba + 1 # Insertamos justo debajo del encabezado/fecha
                        break

            cantidad_a_insertar = len(facturas_del_dia)
            print(f"\n✓ Día {dia}: Marcador en fila {fila_marcador}. Inicio de lista en fila {fila_insercion - 1}. Insertando en fila {fila_insercion}...")
            
            # 3. Insertar e ingresar datos
            ventas_ws.insert_rows(fila_insercion, amount=cantidad_a_insertar)
            
            for i, dato in enumerate(facturas_del_dia):
                fila_escritura = fila_insercion + i
                
                ventas_ws.cell(row=fila_escritura, column=col_dest_fecha).value = dato['serie']
                ventas_ws.cell(row=fila_escritura, column=col_dest_nombre).value = dato['nombre']
                
                try:
                    ventas_ws.cell(row=fila_escritura, column=col_dest_nfact).value = int(float(dato['numero'])) if dato['numero'] else ""
                except ValueError:
                    ventas_ws.cell(row=fila_escritura, column=col_dest_nfact).value = dato['numero']
                    
                ventas_ws.cell(row=fila_escritura, column=col_dest_monto).value = dato['total']

                print(f"  → Insertada: {dato['serie']} | {dato['nombre']} | Fact: {dato['numero']} | Monto: {dato['total']}")
                facturas_agregadas += 1
        else:
            print(f"\n⚠️ NO SE ENCONTRÓ el marcador final para el día {dia}. No se pudieron insertar.")

    if facturas_agregadas > 0:
        try:
            ventas_wb.save(archivo_destino)
            print(f"\n✅ ¡ÉXITO! Se traspasaron {facturas_agregadas} facturas al archivo '{archivo_destino}'.")
        except Exception as e:
            print(f"\n❌ Error al guardar: {e}. Asegúrate de que '{archivo_destino}' NO ESTÉ ABIERTO en Excel.")
    else:
        print("\n⚠️ No se modificó el archivo maestro.")

    input("\nPresiona Enter para cerrar el programa...")

if __name__ == "__main__":
    transferir_facturas()