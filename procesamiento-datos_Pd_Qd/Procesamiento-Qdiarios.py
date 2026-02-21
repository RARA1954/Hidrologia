# ===================================================================================================
# CONFIGURACIÓN - EDITAR ESTOS VALORES SEGÚN SUS NECESIDADES
# ===================================================================================================

# Archivo de entrada (CSV o Excel)
ARCHIVO_ENTRADA = "C:/1.PYTHON/Descarga_Python/CCC Caudales Vertidos.xlsx"  # Cambiar por la ruta de su archivo

# Si es archivo CSV, especificar HOJA_EXCEL = None
# Si es archivo Excel, especificar la hoja # Ejemplo: "Hoja1" o 0 para la primera hoja
HOJA_EXCEL = "PG Vertido"

rio = 'Río Neuquén'
estacion = 'Embalse Los Barreales'

# Archivo de salida
ARCHIVO_SALIDA = f"C:/1.PYTHON/Descarga_Python/{estacion}_reporte_caudales.xlsx"

# ===================================================================================================
import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter
import os
import sys
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# Función para convertir el formato de fecha  
def convertir_fecha(fecha):  
    if pd.isna(fecha):  # Verificar si la fecha es NaN  
        return fecha  
    try:  
        # Intentar convertir con varios formatos comunes
        formatos = [
            '%d/%m/%Y %H:%M',
            '%Y-%m-%d %H:%M',
            '%Y-%m-%d %H:%M:%S',
            '%m/%d/%Y %H:%M',
            '%d/%m/%Y',
            '%Y-%m-%d',
        ]
        for fmt in formatos:
            try:
                fecha_convertida = pd.to_datetime(fecha, format=fmt)
                return fecha_convertida.strftime('%m/%d/%Y %H:%M')
            except (ValueError, TypeError):
                pass
        # Si ninguno coincide, intentar parseo general
        fecha_convertida = pd.to_datetime(fecha)
        return fecha_convertida.strftime('%m/%d/%Y %H:%M')
    except Exception:
        return fecha  # Devolver la fecha original si hay un error

def procesar_caudales(archivo_entrada, archivo_salida='reporte_caudales.xlsx', hoja=None):
    """
    Procesa datos de caudales diarios y genera reportes mensuales/anuales
    
    Parameters:
    archivo_entrada (str): Ruta del archivo CSV o XLSX con datos diarios
    archivo_salida (str): Nombre del archivo Excel de salida
    hoja (str/int): Nombre o índice de la hoja (solo para Excel)
    
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
    caudal_col = None
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(word in col_lower for word in ['fecha', 'date', 'time', 'dia','fecha y hora']):
            fecha_col = col
        elif any(word in col_lower for word in ['caudal', 'flow', 'discharge', 'q', 'descarga', 'flujo','aforo','qd', 'qmd']):
            caudal_col = col
    
    if fecha_col is None or caudal_col is None:
        print("\n❌ Error: No se encontraron las columnas requeridas")
        print("Columnas disponibles:", list(df.columns))
        print("Se buscan columnas que contengan:")
        print("- Para fecha: 'fecha', 'date', 'time', 'dia','Fecha y Hora'")
        print("- Para caudal: 'caudal', 'flow', 'discharge', 'q', 'descarga', 'flujo'")
        raise ValueError("No se encontraron las columnas de fecha y caudal requeridas")
    
    print(f"✅ Columnas detectadas: Fecha='{fecha_col}', Caudal='{caudal_col}'")
    
    # Preparar datos
    df_clean = df[[fecha_col, caudal_col]].copy()
    df_clean.columns = ['Fecha', 'Caudal']
    
    # Normalizar decimales en Caudal (reemplazar coma por punto)
    # Convierte cualquier valor no numérico después de reemplazar comas
    df_clean['Caudal'] = df_clean['Caudal'].astype(str).str.replace(',', '.', regex=False)
    
    # Aplicar la función de conversión de formato de fecha
    df_clean['Fecha'] = df_clean['Fecha'].apply(convertir_fecha)
    df_clean = df_clean.dropna(subset=['Fecha'])
    
    # Convertir caudal a numérico
    df_clean['Caudal'] = pd.to_numeric(df_clean['Caudal'], errors='coerce')
    df_clean = df_clean.dropna(subset=['Caudal'])
    
    # Extraer año y mes (convertir primero a datetime)
    df_clean['Año'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.year
    df_clean['Mes'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.month
    df_clean['Nombre_Mes'] = pd.to_datetime(df_clean['Fecha'], format='%m/%d/%Y %H:%M').dt.strftime('%B')
    
    print(f"Datos procesados: {len(df_clean)} registros válidos")
    fecha_min = pd.to_datetime(df_clean['Fecha'].min(), format='%m/%d/%Y %H:%M')
    fecha_max = pd.to_datetime(df_clean['Fecha'].max(), format='%m/%d/%Y %H:%M')
    print(f"Período: {fecha_min.strftime('%Y-%m-%d')} a {fecha_max.strftime('%Y-%m-%d')}")
    
    # 1. CAUDALES MENSUALES
    caudal_mensual = df_clean.groupby(['Año', 'Mes']).agg({
        'Caudal': 'mean'  # usar media para caudales
    }).reset_index()
    
    # Crear tabla pivote para reporte mensual (años en filas, meses en columnas)
    tabla_mensual = caudal_mensual.pivot_table(
        index='Año',
        columns='Mes',
        values='Caudal',
        fill_value=0
    )
    
    # Renombrar columnas con nombres de meses
    nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Augosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    # Nota: corregir 'Augosto' a 'Agosto' si se usa en tu entorno
    nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    
    # Asegurar que solo se usen los meses que existen en los datos
    meses_disponibles = sorted(tabla_mensual.columns)
    tabla_mensual.columns = [nombres_meses[i-1] for i in meses_disponibles]
    
    # Agregar columna de promedio anual
    tabla_mensual['Promedio Anual'] = tabla_mensual.mean(axis=1)
    
    # 2. CAUDALES ANUALES
    caudal_anual = df_clean.groupby('Año').agg({
        'Caudal': ['mean', 'std', 'min', 'max', 'count']
    }).round(2)
    
    caudal_anual.columns = ['Promedio', 'Desv. Estándar', 
                           'QDMínA', 'QDMáxA', 'Días con Datos']
    caudal_anual = caudal_anual.reset_index()
    
    # 3. ESTADÍSTICAS MENSUALES
    estadisticas_mensuales = df_clean.groupby('Mes').agg({
        'Caudal': ['count', 'mean', 'std', 'min', 'max', 'median']
    }).round(2)
    
    estadisticas_mensuales.columns = ['N° Registros', 'Promedio', 'Desv. Estándar', 
                                     'Mínimo', 'Máximo', 'Mediana']
    
    # Renombrar índice con nombres de meses
    meses_estadisticas = sorted(estadisticas_mensuales.index)
    estadisticas_mensuales.index = [nombres_meses[i-1] for i in meses_estadisticas]
    
    # 4. ESTADÍSTICAS ANUALES
    estadisticas_anuales = caudal_anual[['Promedio']].describe().round(2)
    
    # 5. CREAR HISTOGRAMA DE PROMEDIOS MENSUALES
    print("\n📊 Generando histograma de caudales promedios mensuales...")
    
    # Calcular el promedio para cada mes a partir de tabla_mensual, excluyendo Promedio Anual
    tabla_para_promedios = tabla_mensual.drop(columns=['Promedio Anual'], errors='ignore')
    promedios_mensuales = tabla_para_promedios.mean().reset_index()
    promedios_mensuales.columns = ['Mes', 'Promedio']
    
    # Configurar estilo del gráfico
    plt.figure(figsize=(12, 6))
    sns.set_style("whitegrid")
    
    # Crear el histograma/gráfico de barras
    ax = sns.barplot(x='Mes', y='Promedio', data=promedios_mensuales, color='Blue') # Color simple
    
    # Añadir etiquetas y título
    plt.title(f'{rio} ({estacion}) - Caudal Promedio Mensual', fontsize=16, fontweight='bold')
    plt.xlabel('Mes', fontsize=12, fontweight='bold')
    plt.ylabel('Caudal (m³/s)', fontsize=12, fontweight='bold')
    plt.xticks(rotation=45)
    
    # Añadir valores sobre las barras
    for i, bar in enumerate(ax.patches):
        ax.text(i, bar.get_height() + 0.3, 
                f'{bar.get_height():.1f}', 
                ha='center', va='bottom', 
                fontsize=10)
    
    plt.tight_layout()
    
    # Guardar el histograma como imagen
    ruta_base = os.path.splitext(archivo_salida)[0]
    ruta_histograma = f"{ruta_base}_histograma.png"
    plt.savefig(ruta_histograma, dpi=300, bbox_inches='tight')
    print(f"✅ Histograma guardado como: {ruta_histograma}")
    
    # Guardar histograma para incluir en Excel
    img_buf = BytesIO()
    plt.savefig(img_buf, format='png', dpi=300, bbox_inches='tight')
    img_buf.seek(0)
    
    # EXPORTAR A EXCEL
    print(f"\n📊 Generando archivo Excel: {archivo_salida}")
    
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
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
        
        # HOJA 1: Caudales Mensuales
        tabla_mensual.to_excel(writer, sheet_name='Caudales Mensuales', startrow=2)
        worksheet1 = writer.sheets['Caudales Mensuales']
        
        # Título
        worksheet1.merge_range('A1:N1', 'CAUDALES MENSUALES (m³/s)', title_format)
        
        # Formatear encabezados
        for col_num, value in enumerate(['Año'] + list(tabla_mensual.columns)):
            worksheet1.write(2, col_num, value, header_format)
        
        # Formatear números
        worksheet1.set_column('A:A', 8, year_format)
        worksheet1.set_column('B:N', 12, number_format)
        
        # HOJA 2: Caudales Anuales
        caudal_anual.to_excel(writer, sheet_name='Caudales Anuales', index=False, startrow=2)
        worksheet2 = writer.sheets['Caudales Anuales']
        
        # Título
        worksheet2.merge_range('A1:G1', 'CAUDALES ANUALES (m³/s)', title_format)
        
        for col_num, value in enumerate(caudal_anual.columns):
            worksheet2.write(2, col_num, value, header_format)
        
        worksheet2.set_column('A:A', 8, year_format)
        worksheet2.set_column('B:G', 15, number_format)
        
        # HOJA 3: Histograma de Promedios Mensuales
        worksheet3 = workbook.add_worksheet('Histograma Mensual')
        
        # Título
        worksheet3.merge_range('A1:G1', 'HISTOGRAMA DE CAUDAL PROMEDIO MENSUAL', title_format)
        
        # Insertar el histograma en Excel
        worksheet3.insert_image('B3', 'histograma', {'image_data': img_buf, 'x_scale': 0.8, 'y_scale': 0.8})
        
        # Añadir tabla de datos de promedios mensuales
        promedios_mensuales.to_excel(writer, sheet_name='Histograma Mensual', startrow=25, index=False)
        for i, col in enumerate(promedios_mensuales.columns):
            worksheet3.write(25, i, col, header_format)
        
        worksheet3.set_column('A:A', 8)
        worksheet3.set_column('B:C', 15, number_format)
    
    print(f"✅ Archivo generado exitosamente: {os.path.abspath(archivo_salida)}")
    print(f"📈 Resumen del procesamiento:")
    print(f"   - Años procesados: {df_clean['Año'].nunique()}")
    print(f"   - Rango: {df_clean['Año'].min()} - {df_clean['Año'].max()}")
    print(f"   - Total registros: {len(df_clean)}")
    
    return archivo_salida

def main():
    """Función principal"""
    print("="*100)
    print("🌊 PROCESADOR DE DATOS DE CAUDALES")
    print("="*100)
    print(f"📁 Archivo de entrada: {ARCHIVO_ENTRADA}")
    print(f"📄 Hoja Excel: {HOJA_EXCEL if HOJA_EXCEL else 'Primera hoja'}")
    print(f"💾 Archivo de salida: {ARCHIVO_SALIDA}")
    print("="*100)
    
    try:
        procesar_caudales(ARCHIVO_ENTRADA, ARCHIVO_SALIDA, HOJA_EXCEL)
        print("\n🎉 ¡Procesamiento completado con éxito!")
        print("📊 Se generaron 3 hojas en el archivo Excel:")
        print("   1. Caudales Mensuales")
        print("   2. Caudales Anuales, QDMínA, QDMáxA") 
        print("   3. Histograma de Caudales Promedios Mensuales")
        
    except Exception as e:
        print(f"\n❌ Error durante el procesamiento:")
        print(f"   {str(e)}")
        print("\n💡 Verifique:")
        print("   - Que el archivo de entrada exista")
        print("   - Que tenga columnas de fecha y caudal")
        print("   - Que la hoja especificada exista (si es Excel)")
        sys.exit(1)

if __name__ == "__main__":
    main()

# ===================================================================================================