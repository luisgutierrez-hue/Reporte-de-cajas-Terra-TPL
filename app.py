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

archivo_subido = st.file_uploader("Sube el archivo Excel original (ej. tpl_Id_de_caja_...)", type=["xlsx"])

if archivo_subido is not None:
    try:
        # 1. EXTRAER LA FECHA DEL NOMBRE DEL ARCHIVO ORIGINAL
        match_fecha = re.search(r"(\d{2}_\d{2}_\d{4})", archivo_subido.name)
        if match_fecha:
            fecha_reporte = match_fecha.group(1)
        else:
            fecha_reporte = "Fecha_Desconocida"
            
        nombre_archivo_salida = f"Reportes de cajas {fecha_reporte} Terra.xlsx"

        st.info("Procesando los datos, por favor espera...")
        
        # 2. ENCONTRAR AUTOMÁTICAMENTE LA FILA DE LOS ENCABEZADOS
        # Leemos el archivo sin encabezados primero para buscar en qué fila están
        df_raw = pd.read_excel(archivo_subido, sheet_name="Reporte de Id de caja", header=None)
        
        # Buscamos la fila que contiene el texto 'Nombre de Usuario'
        try:
            fila_encabezado = df_raw[df_raw.eq('Nombre de Usuario').any(axis=1)].index[0]
        except IndexError:
            st.error("❌ No se encontró la columna 'Nombre de Usuario' en el archivo. Verifica que sea el formato correcto.")
            st.stop()
            
        # Volvemos a leer el archivo, pero ahora diciéndole exactamente dónde empiezan los datos
        df = pd.read_excel(archivo_subido, sheet_name="Reporte de Id de caja", header=fila_encabezado)

        # 3. FILTRAR USUARIOS EXCLUIDOS
        usuarios_excluidos = ['nherrer', 'rafa', 'sebag', 'rodrigo', 'richard', 'admin']
        # Convertimos a minúsculas por si acaso alguien lo escribe diferente
        df = df[~df['Nombre de Usuario'].str.lower().isin([u.lower() for u in usuarios_excluidos])].copy()

        # 4. LIMPIEZA Y TRANSFORMACIÓN DE DATOS
        # Asegurar formato de fecha (eliminar hora)
        df['Fecha y hora apertura caja'] = pd.to_datetime(df['Fecha y hora apertura caja'], errors='coerce').dt.strftime('%d/%m/%Y')

        # Función para limpiar valores monetarios 
        def limpiar_moneda(valor):
            if pd.isna(valor):
                return 0.0
            if isinstance(valor, str):
                valor = valor.replace('$', '').replace('.', '').replace(',', '.').strip()
            try:
                return float(valor)
            except ValueError:
                return 0.0

        # Limpiar columnas de dinero
        df['Total efectivo depósito'] = df['Total efectivo depósito'].apply(limpiar_moneda)
        df['Tarjetas'] = df['Tarjetas'].apply(limpiar_moneda)

        # 5. CÁLCULO DE LA NUEVA COLUMNA
        df['Total'] = df['Total efectivo depósito'] + df['Tarjetas']

        # 6. SELECCIONAR Y RENOMBRAR LAS COLUMNAS PARA EL REPORTE FINAL
        df_final = pd.DataFrame({
            'Fecha': df['Fecha y hora apertura caja'],
            'Código': df['Nombre de Usuario'],
            'Nombre y Documento': df['Nombre y Documento'],
            'Agencia': df['Agencia'],
            'Total efectivo depósito': df['Total efectivo depósito'],
            'Tarjetas': df['Tarjetas'],
            'Total': df['Total']
        })

        # 7. GENERAR EL EXCEL CON DISEÑO PROFESIONAL
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Reporte Terra')
            
            workbook = writer.book
            worksheet = workbook['Reporte Terra']
            
            # Estilos
            color_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            fuente_encabezado = Font(color="FFFFFF", bold=True)
            color_cebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            
            # Formato de moneda para columnas de dinero
            from openpyxl.styles import numbers
            formato_moneda = '#,##0' 
            
            # Estilizar encabezados
            for cell in worksheet[1]:
                cell.fill = color_encabezado
                cell.font = fuente_encabezado
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            # Aplicar color intercalado y formato de números
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
                if row_idx % 2 == 0:
                    for cell in row:
                        cell.fill = color_cebra
                
                # Dar formato de número a las columnas de Total efectivo, Tarjetas y Total (columnas E, F, G)
                row[4].number_format = formato_moneda # Total efectivo depósito
                row[5].number_format = formato_moneda # Tarjetas
                row[6].number_format = formato_moneda # Total
            
            # Auto-ajustar ancho de columnas
            for col in worksheet.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[col_letter].width = max_length + 3

        # 8. BOTÓN DE DESCARGA
        st.success("✅ ¡El reporte se ha generado correctamente sin errores!")
        st.download_button(
            label="📥 Descargar Reporte Generado",
            data=buffer.getvalue(),
            file_name=nombre_archivo_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado. Detalle técnico: {e}")
