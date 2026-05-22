import os
import re
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract
import streamlit as st

# === CONFIGURACIÓN GLOBAL DEL OCR ===
# NOTA: Si lo corres en tu PC local, mantén esta línea. Si lo subes a internet, 
# el servidor debe tener Tesseract instalado y la ruta cambiará.
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# =====================================================================
# FUNCIONES AUXILIARES Y PROCESAMIENTO (Tu lógica original intacta)
# =====================================================================
def convertir_bloque_a_24h(bloque_12h):
    try:
        bloque_12h = bloque_12h.upper().replace(" ", "")
        match = re.search(r"(\d{1,2}:\d{2})(AM|PM)-(\d{1,2}:\d{2})(AM|PM)", bloque_12h)
        if not match: return bloque_12h
        
        hora1, m1, hora2, m2 = match.groups()
        h1, min1 = map(int, hora1.split(":"))
        if m1 == "PM" and h1 != 12: h1 += 12
        if m1 == "AM" and h1 == 12: h1 = 0
        
        h2, min2 = map(int, hora2.split(":"))
        if m2 == "PM" and h2 != 12: h2 += 12
        if m2 == "AM" and h2 == 12: h2 = 0
        
        return f"{h1:02d}:{min1:02d}-{h2:02d}:{min2:02d}"
    except:
        return bloque_12h

def interpretar_dias(texto_dias):
    columnas = []
    if "L" in texto_dias: columnas.append(2)
    if "Ma" in texto_dias or "MA" in texto_dias: columnas.append(3)
    if "Mi" in texto_dias or "MI" in texto_dias: columnas.append(4)
    if "J" in texto_dias: columnas.append(5)
    if "V" in texto_dias: columnas.append(6)
    return columnas

def procesar_texto_ocr(texto_raw, preferencias):
    matriz_clases = {}
    bloques_horarios = set()
    lineas = texto_raw.split("\n")
    materia_actual = None
    maestro_excluido = preferencias["maestro_excluido"]

    for linea in lineas:
        linea = linea.strip()
        if not linea: continue
        
        if "créd" in linea.lower() or "cred" in linea.lower():
            materia_actual = re.sub(r'[^A-ZÁÉÍÓÚÑ\s]', '', linea.split("(")[0]).strip()
            continue
            
        if "lugares" in linea.lower() or "dispon" in linea.lower() or "pm" in linea.lower() or "am" in linea.lower() or "06:00" in linea:
            linea_limpia = linea.replace("=", "-").replace("~", "-")
            match_hora = re.search(r"(\d{1,2}:\d{2}\s*[A-Z]{2}\s*-\s*\d{1,2}:\d{2}\s*[A-Z]{2})", linea_limpia, re.IGNORECASE)
            
            if match_hora and materia_actual:
                bloque_12h = match_hora.group(1)
                bloque_24h = convertir_bloque_a_24h(bloque_12h)
                partes_linea = linea_limpia.split(bloque_12h)
                
                maestro = "POR SELECCIONAR"
                if len(partes_linea) > 1:
                    maestro = partes_linea[1].split("Lugares")[0].strip()
                
                if maestro_excluido and maestro_excluido in maestro.upper(): continue
                if preferencias["horarios"] and bloque_24h not in preferencias["horarios"]: continue
                
                texto_antes_hora = partes_linea[0]
                lista_dias = interpretar_dias(texto_antes_hora)
                
                bloques_horarios.add(bloque_24h)
                if bloque_24h not in matriz_clases:
                    matriz_clases[bloque_24h] = {}
                    
                for dia_num in lista_dias:
                    mapa_inverso = {2: "L", 3: "MA", 4: "MI", 5: "J", 6: "V"}
                    if not preferencias["dias"] or mapa_inverso[dia_num] in preferencias["dias"]:
                        matriz_clases[bloque_24h][dia_num] = (materia_actual, maestro)

    return sorted(list(bloques_horarios)), matriz_clases

# Modificado para guardar el Excel en memoria y permitir la descarga web
def crear_excel_horario_en_memoria(nombre_alumno, bloques_horarios, matriz_clases):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Horario Semanal"
    ws.views.sheetView[0].showGridLines = True

    NAVY_DARK = "2B4C7E"       
    BLUE_LIGHT = "F0F4F8"      
    WHITE = "FFFFFF"
    GRAY_BORDER = "D3D3D3"

    font_title = Font(name="Segoe UI", size=15, bold=True, color=WHITE)
    font_header = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_sub = Font(name="Segoe UI", size=9, italic=True, color="DDDDDD")
    font_body = Font(name="Segoe UI", size=11, bold=True, color="2C3E50")
    font_time = Font(name="Segoe UI", size=10, bold=True, color="34495E")

    fill_title = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_header = PatternFill(start_color="3B629B", end_color="3B629B", fill_type="solid")
    fill_zebra = PatternFill(start_color=BLUE_LIGHT, end_color=BLUE_LIGHT, fill_type="solid")
    fill_white = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
    fill_clase = PatternFill(start_color="E6EEF8", end_color="E6EEF8", fill_type="solid")

    thin_side = Side(border_style="thin", color=GRAY_BORDER)
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = "CONTROL DE CARGA ACADÉMICA AUTOMATIZADA"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = align_center
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:F2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Horario Personalizado de Carga Académica - Alumno: {nombre_alumno}"
    sub_cell.font = font_sub
    sub_cell.fill = fill_title
    sub_cell.alignment = align_center
    ws.row_dimensions[2].height = 18

    columnas = ["Horario", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    for col_num, header_text in enumerate(columnas, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.value = header_text
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_data
    ws.row_dimensions[4].height = 25

    row_idx = 5
    for i, bloque in enumerate(bloques_horarios):
        cell_time = ws.cell(row=row_idx, column=1)
        cell_time.value = bloque 
        cell_time.font = font_time
        cell_time.alignment = align_center
        cell_time.border = border_data
        cell_time.fill = fill_zebra if i % 2 == 0 else fill_white
        
        for col_idx in range(2, 7):
            cell_data = ws.cell(row=row_idx, column=col_idx)
            cell_data.border = border_data
            cell_data.fill = fill_zebra if i % 2 == 0 else fill_white
            
            if bloque in matriz_clases and col_idx in matriz_clases[bloque]:
                materia, prof = matriz_clases[bloque][col_idx]
                cell_data.value = f"{materia}\n\n{prof}"
                cell_data.font = font_body
                cell_data.alignment = align_center
                cell_data.fill = fill_clase
            else:
                cell_data.value = "-"
                cell_data.font = Font(name="Segoe UI", size=10, color="CCCCCC")
                cell_data.alignment = align_center

        ws.row_dimensions[row_idx].height = 60  
        row_idx += 1

    ws.column_dimensions['A'].width = 22
    for col_idx in range(2, 7):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 30

    # Guardar en un buffer de memoria en vez de en el disco duro
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =====================================================================
# INTERFAZ WEB CON STREAMLIT
# =====================================================================
st.set_page_config(page_title="Generador de Horarios", page_icon="📅", layout="centered")

st.title("📅 Generador Automatizado de Horarios")
st.markdown("Sube la captura de tu carga académica y descarga tu horario optimizado en Excel.")

# Formulario de configuración (Preferencias del Alumno)
st.sidebar.header("⚙️ Preferencias del Alumno")
dias_input = st.sidebar.text_input("Días disponibles (ej: L, Ma, Mi)", placeholder="Dejar vacío para todos")
horas_input = st.sidebar.text_input("Horarios preferidos en 24h (ej: 14:00-16:00)", placeholder="Dejar vacío para todos")
maestro_input = st.sidebar.text_input("Maestro a excluir de la lista")

# Variables de control
dias_preferidos = [dia.strip().upper() for dia in dias_input.split(",")] if dias_input.strip() else []
horas_preferidas = [hora.strip() for hora in horas_input.split(",")] if horas_input.strip() else []
maestro_excluido = maestro_input.strip().upper() if maestro_input.strip() else None

preferencias = {
    "dias": dias_preferidos,
    "horarios": horas_preferidas,
    "maestro_excluido": maestro_excluido
}

# Inputs principales en la página
nombre_alumno = st.text_input("👤 Nombre completo del alumno:")
archivo_subido = st.file_uploader("📸 Sube la captura de pantalla de tu horario (JPG, JPEG, PNG)", type=["jpg", "jpeg", "png"])

if archivo_subido and nombre_alumno:
    st.success("¡Imagen cargada con éxito!")
    
    # Botón para activar el proceso
    if st.button("🚀 Procesar Horario y Generar Excel"):
        with st.spinner("Procesando la imagen con OCR..."):
            try:
                # Abrir imagen directamente desde la web
                imagen_real = Image.open(archivo_subido)
                texto_real_ocr = pytesseract.image_to_string(imagen_real)
                
                bloques, matriz_final = procesar_texto_ocr(texto_real_ocr, preferencias)
                
                # Salvavidas de respaldo (Tu lógica original intacta)
                if len(matriz_final) == 0:
                    st.warning("Se detectó ruido visual en la captura. Aplicando datos de respaldo...")
                    texto_respaldo_perfecto = """
                    ☑ INTRODUCCION A LA ECONOMIA (6 créditos)
                    C2A - L,Mi 04:00 PM-06:00 PM - DR. FELIPE HERNANDEZ Lugares disponibles: 38
                    ☑ BASES TEORICAS DE LA INVESTIGACION CIENTIFICA (6 créditos)
                    C2A - L,Mi 02:00 PM-04:00 PM - MC. JOSE RAMON OLIVO Lugares disponibles: 38
                    ☑ LOGICA MATEMATICA (6 créditos)
                    C2A - J-V 04:00 PM-06:00 PM - LIC. BEATRIZ ANGELICA TOSCANO Lugares disponibles: 38
                    ☑ SISTEMAS OPERATIVOS (6 créditos)
                    C2A - Ma,J 02:00 PM-04:00 PM - DR. GABRIEL ZEPEDA Lugares disponibles: 37
                    ☑ CALCULO DIFERENCIAL (6 créditos)
                    C2A - Ma,J 06:00 PM-08:00 PM - LIC. SERGIO OMAR RODRIGUEZ Lugares disponibles: 38
                    """
                    bloques, matriz_final = procesar_texto_ocr(texto_respaldo_perfecto, preferencias)
                
                # Generar el Excel en memoria
                excel_data = crear_excel_horario_en_memoria(nombre_alumno, bloques, matriz_final)
                
                # Botón de descarga nativo de la web
                st.balloons()
                st.download_button(
                    label="📥 Descargar Horario en Excel",
                    data=excel_data,
                    file_name=f"Horario_{nombre_alumno.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Hubo un problema al procesar el archivo: {e}")
elif not nombre_alumno and archivo_subido:
    st.info("Por favor introduce el nombre del alumno para continuar.")
