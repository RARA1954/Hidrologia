import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os
import numpy as np
from io import BytesIO

# ==============================================================================
# CONFIGURACIÓN - EDITAR VALORES SEGÚN NECESIDADES
# ==============================================================================

# Archivo de entrada (CSV o Excel)
ARCHIVO_ENTRADA = "C:/1.PYTHON/Descarga_Python/Historicos-Estacion 2004.xlsx"  # Ruta del archivo
# Si es archivo CSV, especificar HOJA_EXCEL = None
# Si es archivo Excel, especificar la hoja # Ejemplo: "Daily_Summary" o 0 para la primera hoja
HOJA_EXCEL = "Hoja1"

# Estación o fuente de datos (se usa para Data_Source)
estacion = 'Paso de Indios'

# Archivo de salida
ARCHIVO_SALIDA = f"C:/1.PYTHON/Descarga_Python/{estacion}_reporte_precipitaciones.xlsx"

# ==============================================================================
# Variable de control: columna a procesar
# Cambia el nombre exactamente a la columna que contenga los datos de interés
colum_mane = 'precipitacion'
# ==============================================================================

# Función para convertir el formato de fecha
def convertir_fecha(fecha):  
    if pd.isna(fecha):  # Verificar si la fecha es NaN  
        return fecha  
    try:  
        # Convertir la fecha de día/mes/año a un objeto datetime  
        fecha_convertida = pd.to_datetime(fecha, format='%d/%m/%Y %H:%M')  
        # Devolver la fecha en el nuevo formato  
        return fecha_convertida.strftime('%m/%d/%Y %H:%M')  
    except ValueError:
        # Intentar convertir la fecha sin especificar el formato
        try:
            fecha_convertida = pd.to_datetime(fecha)
            return fecha_convertida.strftime('%m/%d/%Y %H:%M')
        except ValueError:
            return fecha  # Devolver la fecha original si hay un error

# ==============================================================================
# Función principal de procesamiento de precipitaciones
# ==============================================================================
def procesar_precipitaciones(archivo_entrada, archivo_salida='reporte_precipitaciones.xlsx', hoja=None, fuente_data=None, columna_procesar=None):
    """
    Procesa datos de precipitaciones diarias y genera reportes mensuales/anuales
    
    Parameters:
    archivo_entrada (str): Ruta del archivo CSV o XLSX con datos diarios
    archivo_salida (str): Nombre del archivo Excel de salida
    hoja (str/int): Nombre o índice de la hoja (solo para Excel)
    fuente_data (str): Nombre de la fuente de datos para Data_Source
    columna_procesar (str): Nombre lógico de la columna de datos a procesar (like se ve en el conjunto)
    
    Returns:
    str: Ruta del archivo generado
    """
    
    print(f"Procesando archivo: {archivo_entrada}")
    if hoja is not None:
        print(f"Hoja seleccionada: {hoja}")
    
    try:
        # Leer archivo de entrada
        if archivo_entrada.lower().endswith('.csv'):
            df = pd.read_csv(archivo_entrada)
        elif archivo_entrada.lower().endswith(('.xlsx', '.xls')):
            if hoja is not None:
                df = pd.read_excel(archivo_entrada, sheet_name=hoja)
            else:
                df = pd.read_excel(archivo_entrada)
        else:
            raise ValueError("Formato no soportado. Use archivos CSV o Excel (.xlsx/.xls)")
        
        print(f"Datos cargados: {len(df)} registros")
        print("Columnas disponibles:", list(df.columns))
        
    except FileNotFoundError:
        raise FileNotFoundError(f"No se encontró el archivo: {archivo_entrada}")
    except Exception as e:
        raise Exception(f"Error al leer el archivo: {str(e)}")
    
    # Detectar nombres de columnas (flexibilidad en nombres)
    fecha_col = None
    precip_col = None
    
    for col in df.columns:
        col_lower = col.lower().strip()
        if any(word in col_lower for word in ['fecha', 'date', 'time', 'dia','fecha y hora']):
            fecha_col = col
        elif any(word in col_lower for word in ['precipitacion', 'precipitation', 'lluvia', 'rain', 'pp', 'prec','pd_pt','pd']):
            precip_col = col
    
    if fecha_col is None or precip_col is None:
        print("\n❌ Error: No se encontraron las columnas requeridas")
        print("Columnas disponibles:", list(df.columns))
        print("Se buscan columnas que contengan:")
        print("- Para fecha: 'fecha', 'date', 'time', 'dia','Fecha y Hora'")
        print("- Para precipitación: 'precipitacion', 'precipitation', 'lluvia', 'rain', 'pp', 'prec','pd_pt','pd'")
        raise ValueError("No se encontraron las columnas de fecha y precipitacion requeridas")
    
    print(f"✅ Columnas detectadas: Fecha='{fecha_col}', Precipitación='{precip_col}'")
    
    # Preparar datos
    df_clean = df[[fecha_col, precip_col]].copy()
    df_clean.columns = ['Fecha', 'Precipitacion']
    
    # Nueva columna de fuente de datos
    df_clean['Data_Source'] = fuente_data if fuente_data is not None else estacion
    # Nueva columna de control de procesamiento
    df_clean['Columna_A_Procesar'] = columna_procesar if columna_procesar is not None else colum_mane
    
    # Aplicar la función de conversión de formato de fecha
    df_clean['Fecha'] = df_clean['Fecha'].apply(convertir_fecha)
    df_clean = df_clean.dropna(subset=['Fecha'])
    
    # Convertir precipitación a numérico
    df_clean['Precipitacion'] = pd.to_numeric(df_clean['Precipitacion'], errors='coerce')
    df_clean = df_clean.dropna(subset=['Precipitacion'])
    
    # Extraer año y mes (convertir primero a datetime)
    df_clean['Año'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.year
    df_clean['Mes'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.month
    df_clean['Nombre_Mes'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.strftime('%B')
    
    print(f"Datos procesados: {len(df_clean)} registros válidos")
    fecha_min = pd.to_datetime(df_clean['Fecha'].min(), format='%m/%d/%Y %H:%M')
    fecha_max = pd.to_datetime(df_clean['Fecha'].max(), format='%m/%d/%Y %H:%M')
    print(f"Período: {fecha_min.strftime('%Y-%m-%d')} a {fecha_max.strftime('%Y-%m-%d')}")
    
    # 1. PRECIPITACIONES MENSUALES
    precipitacion_mensual = df_clean.groupby(['Año', 'Mes']).agg({
        'Precipitacion': 'sum'
    }).reset_index()
    
    # Crear tabla pivote para reporte mensual (años en filas, meses en columnas)
    tabla_mensual = precipitacion_mensual.pivot_table(
        index='Año',
        columns='Mes',
        values='Precipitacion',
        fill_value=0
    )
    
    # Renombrar columnas con nombres de meses
    nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    
    # Asegurar que solo se usen los meses que existen en los datos
    meses_disponibles = sorted(tabla_mensual.columns)
    tabla_mensual.columns = [nombres_meses[i-1] for i in meses_disponibles]
    
    # Agregar columna de total anual
    tabla_mensual['Total Anual'] = tabla_mensual.sum(axis=1)
    
    # 2. PRECIPITACIONES ANUALES
    precipitacion_anual = df_clean.groupby('Año').agg({
        'Precipitacion': ['sum', 'mean', 'std', 'min', 'max', 'count']
    }).round(2)
    
    precipitacion_anual.columns = ['Total', 'Promedio Diario', 'Desv. Estándar', 
                                   'PDMínA', 'PDMáxA', 'Días con Datos']
    precipitacion_anual = precipitacion_anual.reset_index()
    
    # 3. ESTADÍSTICAS MENSUALES
    estadisticas_mensuales = df_clean.groupby('Mes').agg({
        'Precipitacion': ['count', 'sum', 'mean', 'std', 'min', 'max', 'median']
    }).round(2)
    
    estadisticas_mensuales.columns = ['N° Registros', 'Total', 'Promedio', 'Desv. Estándar', 
                                      'Mínimo', 'Máximo', 'Mediana']
    
    # Renombrar índice con nombres de meses
    meses_estadisticas = sorted(estadisticas_mensuales.index)
    estadisticas_mensuales.index = [nombres_meses[i-1] for i in meses_estadisticas]
    
    # 4. ESTADÍSTICAS ANUALES
    estadisticas_anuales = precipitacion_anual[['Total', 'Promedio Diario']].describe().round(2)
    
    # 5. HISTOGRAMA DE PROMEDIOS MENSUALES (opcional)
    print("\n📊 Generando histograma de promedios mensuales...")
    tabla_para_promedios = tabla_mensual.drop(columns=['Total Anual'], errors='ignore')
    promedios_mensuales = tabla_para_promedios.mean().reset_index()
    promedios_mensuales.columns = ['Mes', 'Promedio']
    
    plt.figure(figsize=(12, 6))
    # Sustituir estilo problemático por uno disponible
    plt.style.use('ggplot')
    ax = plt.bar(promedios_mensuales['Mes'], promedios_mensuales['Promedio'], color='blue')
    plt.title(f'Est. {estacion} - Precipitación Promedio Mensual', fontsize=16, fontweight='bold')
    plt.xlabel('Mes', fontsize=12, fontweight='bold')
    plt.ylabel('Precipitación (mm)', fontsize=12, fontweight='bold')
    plt.xticks(rotation=45)
    plt.tight_layout()
    # Guardar el histograma como imagen
    ruta_base = os.path.splitext(archivo_salida)[0]
    ruta_histograma = f"{ruta_base}_histograma.png"
    plt.savefig(ruta_histograma, dpi=300, bbox_inches='tight')
    print(f"✅ Histograma guardado como: {ruta_histograma}")
    img_buf = BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    img_buf.seek(0)
    plt.close()

    # EXPORTAR A EXCEL
    print(f"\n📊 Generando archivo Excel: {archivo_salida}")
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'fg_color': '#B8CCE4',
            'border': 1
        })
        number_format = workbook.add_format({'num_format': '#,##0.0'})
        year_format = workbook.add_format({'num_format': '0'})
        
        # HOJA 1: Precipitaciones Mensuales
        tabla_mensual.to_excel(writer, sheet_name='Precipitaciones Mensuales', startrow=2)
        worksheet1 = writer.sheets['Precipitaciones Mensuales']
        worksheet1.merge_range('A1:N1', 'PRECIPITACIONES MENSUALES (mm)', title_format)
        for col_num, value in enumerate(['Año'] + list(tabla_mensual.columns)):
            worksheet1.write(2, col_num, value, header_format)
        worksheet1.set_column('A:A', 8, year_format)
        worksheet1.set_column('B:N', 12, number_format)
        
        # HOJA 2: Precipitaciones Anuales
        precipitacion_anual.to_excel(writer, sheet_name='Precipitaciones Anuales', index=False, startrow=2)
        worksheet2 = writer.sheets['Precipitaciones Anuales']
        worksheet2.merge_range('A1:G1', 'PRECIPITACIONES ANUALES (mm)', title_format)
        for col_num, value in enumerate(precipitacion_anual.columns):
            worksheet2.write(2, col_num, value, header_format)
        worksheet2.set_column('A:A', 10, year_format)
        worksheet2.set_column('B:G', 15, number_format)
        
        # HOJA 3: Histograma de Promedios Mensuales
        worksheet3 = workbook.add_worksheet('Histograma Mensual')
        worksheet3.merge_range('A1:G1', 'HISTOGRAMA DE PRECIPITACIÓN PROMEDIO MENSUAL', title_format)
        worksheet3.insert_image('B3', 'histograma', {'image_data': img_buf, 'x_scale': 0.8, 'y_scale': 0.8})
        promedios_mensuales.to_excel(writer, sheet_name='Histograma Mensual', startrow=25, index=False)
        for i, col in enumerate(promedios_mensuales.columns):
            worksheet3.write(25, i, col, header_format)
        worksheet3.set_column('A:A', 8)
        worksheet3.set_column('B:C', 15, number_format)
        
        # HOJA 4: Full_Raw_Temps (con Data_Source y columna de procesamiento)
        # df_clean.to_excel(writer, sheet_name='Full_Raw_Temps', index=False)
    
    print(f"✅ Archivo generado exitosamente: {os.path.abspath(archivo_salida)}")
    print(f"📈 Resumen del procesamiento:")
    print(f"   - Años procesados: {df_clean['Año'].nunique()}")
    print(f"   - Rango: {df_clean['Año'].min()} - {df_clean['Año'].max()}")
    print(f"   - Total registros: {len(df_clean)}")
    
    return archivo_salida

def main():
    """Función principal"""
    print("="*100)
    print("🌧️  PROCESADOR DE DATOS DE PRECIPITACIÓN")
    print("="*100)
    print(f"📁 Archivo de entrada: {ARCHIVO_ENTRADA}")
    print(f"📄 Hoja Excel: {HOJA_EXCEL if HOJA_EXCEL else 'Primera hoja'}")
    print(f"💾 Archivo de salida: {ARCHIVO_SALIDA}")
    print("="*100)
    
    try:
        procesar_precipitaciones(
            ARCHIVO_ENTRADA,
            ARCHIVO_SALIDA,
            HOJA_EXCEL,
            estacion,
            columna_procesar=colum_mane
        )
        print("\n🎉 ¡Procesamiento completado con éxito!")
        print("📊 Se generaron varias hojas en el archivo Excel:")
        print("   - Precipitaciones Mensuales")
        print("   - Precipitaciones Anuales, PDMínA, PDMáxA")
        print("   - Histograma de Promedios Mensuales")
        print("   - Full_Raw_Temps (con Data_Source y Columna_A_Procesar)")
        
    except Exception as e:
        print(f"\n❌ Error durante el procesamiento:")
        print(f"   {str(e)}")
        print("\n💡 Verifique:")
        print("   - Que el archivo de entrada exista")
        print("   - Que tenga columnas de fecha y precipitacion")
        print("   - Que la hoja especificada exista (si es Excel)")
        raise

if __name__ == "__main__":
    main()

# ==============================================================================