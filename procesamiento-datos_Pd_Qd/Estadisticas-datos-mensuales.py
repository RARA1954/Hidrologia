import pandas as pd
import numpy as np
from scipy import stats  # Importar el módulo stats de scipy
import os

# Ruta del archivo de entrada (ajusta según sea necesario; usa el archivo adjunto como guía)
input_file = r'C:\1.PYTHON\Descarga_Python\Cinco_Saltos_PM_1993_2025.xlsx'  # Cambia esto a la ruta real de tu archivo

estación = 'Cinco Saltos'

# Ruta de salida
output_dir = r'C:\1.PYTHON\Descarga_Python'
output_file = os.path.join(output_dir, f'{estación} - estadisticos_mensuales.xlsx')

# Asegúrate de que el directorio de salida exista
os.makedirs(output_dir, exist_ok=True)

# Leer el archivo Excel (asumiendo que los datos están en la primera hoja)
df = pd.read_excel(input_file)

# Asumir estructura: columna 'Año' y columnas para cada mes (e.g., 'Enero', 'Febrero', ..., 'Diciembre')
# Ajusta los nombres de las columnas de meses si es necesario
month_columns = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# Verificar si las columnas existen (basado en el archivo de guía)
available_months = [col for col in month_columns if col in df.columns]

if not available_months:
    raise ValueError("No se encontraron columnas de meses en el archivo. Ajusta los nombres en el script.")

# Lista de estadísticos a calcular
stats_list = [
    'Cantidad de datos',
    'Media',
    'Mediana',
    'Moda',
    'Varianza',
    'Desvío Estándar',
    'Coef. Variación (%)',
    'Mínimo',
    'Máximo',
    'Rango',
    'Sesgo',
    'Sesgo Estandarizado',
    'Curtosis',
    'Curtosis Estandarizada',
    'Suma'
]

# DataFrame para almacenar los resultados
results = pd.DataFrame(index=stats_list, columns=available_months)

# Función auxiliar para redondear solo si no es NaN
def safe_round(val, nd=2):
    if pd.isna(val):
        return val
    return round(val, nd)

# Función para calcular estadísticos de una serie de datos (maneja NaNs ignorándolos)
# En calculate_stats, aplicar rounding a 2 decimales
def calculate_stats(data):
    n = data.notna().sum()
    if n == 0:
        return {stat: np.nan for stat in stats_list}
    
    data = pd.to_numeric(data, errors='coerce')
    
    mean = safe_round(np.nanmean(data), 2)
    median = safe_round(np.nanmedian(data), 2)
    variance = safe_round(np.nanvar(data, ddof=1) if n > 1 else np.nan, 2)
    std = safe_round(np.nanstd(data, ddof=1) if n > 1 else np.nan, 2)
    cv = (std/mean)*100
    min_val = safe_round(np.nanmin(data), 2)
    max_val = safe_round(np.nanmax(data), 2)
    range_val = safe_round((max_val - min_val) if n > 0 else np.nan, 2)
    sum_val = safe_round(np.nansum(data), 2)
    
    try:
        mode_result = stats.mode(data, nan_policy='omit')
        mode = mode_result.mode[0] if len(mode_result.mode) > 0 else np.nan
        mode = safe_round(mode, 2)
    except Exception:
        mode = np.nan
    
    skew = stats.skew(data, nan_policy='omit') if n > 2 else np.nan
    skew = safe_round(skew, 4) if not pd.isna(skew) else np.nan
    
    kurt = stats.kurtosis(data, nan_policy='omit') if n > 3 else np.nan
    kurt = safe_round(kurt, 4) if not pd.isna(kurt) else np.nan
    
    se_skew = np.sqrt(6 / n) if n > 0 else np.nan
    standardized_skew = (skew / se_skew) if (not pd.isna(skew) and not pd.isna(se_skew) and se_skew != 0) else np.nan
    standardized_skew = safe_round(standardized_skew, 4) if not pd.isna(standardized_skew) else np.nan
    
    se_kurt = np.sqrt(24 / n) if n > 0 else np.nan
    standardized_kurt = (kurt / se_kurt) if (not pd.isna(kurt) and not pd.isna(se_kurt) and se_kurt != 0) else np.nan
    standardized_kurt = safe_round(standardized_kurt, 4) if not pd.isna(standardized_kurt) else np.nan
    
    return {
        'Cantidad de datos': n,
        'Media': mean,
        'Mediana': median,
        'Moda': mode,
        'Varianza': variance,
        'Desvío Estándar': std,
        'Coef. Variación (%)': round(cv,2),
        'Mínimo': min_val,
        'Máximo': max_val,
        'Rango': range_val,
        'Sesgo': skew,
        'Sesgo Estandarizado': standardized_skew,
        'Curtosis': kurt,
        'Curtosis Estandarizada': standardized_kurt,
        'Suma': sum_val
    }

# Calcular para cada mes
for month in available_months:
    data = df[month]
    month_stats = calculate_stats(data)  # Renombrado para evitar conflicto con el módulo 'stats'
    for stat, value in month_stats.items():
        results.at[stat, month] = value

# Exportar a Excel
results.to_excel(output_file, index=True)

print(f"Archivo exportado exitosamente a: {output_file}")