import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador de Reportes de Caja", layout="centered")

st.title("📊 Generador de Reportes: Cajas Terra")
st.write("Sube el reporte original de Excel para generar automáticamente el reporte filtrado y formateado.")

# --- INTERFAZ PARA SUBIR EL ARCHIVO ---
archivo_subido = st.file_uploader("Sube el archivo Excel original (ej. tpl_Id_de_caja_...)", type=["xlsx"])

if archivo_subido is not None:
    try:
        # 1. INTENTAR EXTRAER LA FECHA DEL NOMBRE DEL ARCHIVO ORIGINAL
        # Busca el patrón DD_MM_YYYY en el nombre
        match_fecha = re.search(r"(\d{2}_\d{2}_\d{4})", archivo_subido.name)
        if match_fecha:
            fecha_reporte = match_fecha.group(1)
        else:
            fecha_reporte = "Fecha_Desconocida"
            
        nombre_archivo_salida = f"Reportes de cajas {fecha_reporte} Terra.xlsx"

        # 2. LEER EL ARCHIVO ORIGINAL (Solo la hoja especificada)
        st.info("Procesando los datos, por favor espera...")
        df = pd.read_excel(archivo_subido, sheet_name="Reporte de Id de caja")

        # 3. FILTRAR USUARIOS EXCLUIDOS
        usuarios_excluidos = ['nherrer', 'rafa', 'sebag', 'rodrigo', 'admin']
        df = df[~df['Nombre de Usuario'].isin(usuarios_excluidos)].copy()

        # 4. LIMPIEZA Y TRANSFORMACIÓN DE DATOS
        # Asegurar formato de fecha (eliminar hora)
        df['fecha de apertura de caja'] = pd.to_datetime(df['fecha de apertura de caja'], errors='coerce').dt.strftime('%d/%m/%Y')

        # Función para limpiar valores monetarios en caso de que vengan como texto con "$" o comas
        def limpiar_moneda(valor):
            if pd.isna(valor):
                return 0.0
            if isinstance(valor, str):
                valor = valor.replace('$', '').replace('.', '').replace(',', '.').strip()
            try:
                return float(valor)
            except ValueError:
                return 0.0

        df['Total efectivo depósito'] = df['Total efectivo depósito'].apply(limpiar_moneda)
        df['Tarjetas'] = df['Tarjetas'].apply(limpiar_moneda)

        # 5. CÁLCULO DE LA NUEVA COLUMNA
        df['Total'] = df['Total efectivo depósito'] + df['Tarjetas']

        # 6. SELECCIONAR Y RENOMBRAR LAS COLUMNAS PARA EL REPORTE FINAL
        # Mapeo exacto de las columnas que solicitaste
        df_final = pd.DataFrame({
            'Fecha': df['fecha de apertura de caja'],
            'Código': df['Nombre de Usuario'],
            'Nombre y Documento': df['nombre y documento de usuario'],
            'Agencia': df['Agencia'],
            'Total efectivo depósito': df['Total efectivo depósito'],
            'Tarjetas': df['Tarjetas'],
            'Total': df['Total']
        })

        # 7. GENERAR EL EXCEL CON DISEÑO PROFESIONAL (openpyxl)
        # Creamos un buffer en memoria para no tener que guardar el archivo en el disco duro del servidor
        buffer = io.BytesIO()
        
        # Usamos ExcelWriter con el motor openpyxl
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Reporte Terra')
            
            # Obtener la hoja de trabajo activa para aplicar estilos
            workbook = writer.book
            worksheet = workbook['Reporte Terra']
            
            # Definir estilos
            color_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Azul oscuro
            fuente_encabezado = Font(color="FFFFFF", bold=True) # Letra blanca y negrita
            color_cebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # Gris muy claro
            
            # Estilizar encabezados (Fila 1)
            for cell in worksheet[1]:
                cell.fill = color_encabezado
                cell.font = fuente_encabezado
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            # Aplicar color intercalado tipo "Zebra" a las filas de datos
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
                if row_idx % 2 == 0:
                    for cell in row:
                        cell.fill = color_cebra
            
            # Auto-ajustar el ancho de las columnas
            for col in worksheet.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                # Se le suma un pequeño padding extra
                worksheet.column_dimensions[col_letter].width = max_length + 3

        # 8. BOTÓN DE DESCARGA EN STREAMLIT
        st.success("✅ ¡El reporte se ha generado correctamente!")
        st.download_button(
            label="📥 Descargar Reporte Generado",
            data=buffer.getvalue(),
            file_name=nombre_archivo_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except ValueError as ve:
        st.error(f"❌ Error al procesar el archivo. Es probable que falte una columna esperada. Detalle técnico: {ve}")
    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado. Asegúrate de estar subiendo el archivo correcto. Detalle técnico: {e}")
