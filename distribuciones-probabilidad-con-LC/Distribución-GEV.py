'''
 Script para ajustar la distribución General de Valores Extremos (GEV)
 a precipitaciones diarias máximas anuales (PDMA) y
 caudales diarios máximos anuales (QDMA),
 con límites de confianza del 90% y 95%

 by Rapa 2024

'''

##############################################################################################################

import pandas as pd  
import numpy as np  
import matplotlib.pyplot as plt  
import scipy.stats as stats  
from scipy.interpolate import interp1d  
import matplotlib.ticker as mtick 


##########################################################################################################

# Indicar la ruta del archivo Excel de entrada  
input_file_path = 'C:/1.PYTHON/Descarga_Python/PDMA_Lindero_Atravesado.xlsx'  

# Establecer la hoja donde se encuentran los datos
nombre_hoja ='Hoja1'

# Definir la variable para la columna de precipitación o caudales 
nombre_columna = 'PDMA'  # Cambia este valor si es necesario  

# Establecer la estación de medición
estación = 'Lindero Atravesado'

##########################################################################################################

# Leer el archivo Excel  
data = pd.read_excel(input_file_path,sheet_name= nombre_hoja)

# Imprimir los nombres de las columnas del DataFrame  
print(" " * 100)
print("-" * 100)
print(" " * 100)
print("Columnas en el archivo de entrada:", data.columns.tolist())  
print(" " * 100)

# Asegurarse de que la columna de caudales esté en formato numérico  
data[nombre_columna] = pd.to_numeric(data[nombre_columna], errors='coerce')  

# Eliminar filas con valores NaN  
data = data.dropna(subset=[nombre_columna])  

# Ajustar la distribución GEV  
params = stats.genextreme.fit(data[nombre_columna])  
c, loc, scale = params  

# Crear un rango de valores para la gráfica  
x = np.linspace(np.min(data[nombre_columna]), np.max(data[nombre_columna]), len(data[nombre_columna]))  

# Calcular la CDF de la GEV ajustada  
cdf = stats.genextreme.cdf(x, c, loc=loc, scale=scale)  

# Calcular la CDF empírica  
data_sorted = np.sort(data[nombre_columna])  
empirical_cdf = np.arange(1, len(data_sorted) + 1) / len(data_sorted)  

# Interpolar la CDF empírica para que coincida con el rango de x  
interp_func = interp1d(data_sorted, empirical_cdf, bounds_error=False, fill_value=(0, 1))  
empirical_cdf_interp = interp_func(x)  

# Calcular R²  
sst = np.sum((empirical_cdf_interp - np.mean(empirical_cdf_interp))**2)  # Suma total de cuadrados  
ssr = np.sum((empirical_cdf_interp - cdf)**2)  # Suma de cuadrados de los residuos  
r_squared = 1 - (ssr / sst)  # Coeficiente de determinación  
r_squared_percentage = r_squared * 100  # Convertir a porcentaje  

# Calcular límites de confianza del 90% y 95%  
n = len(data[nombre_columna])  # Número de observaciones  
cdf_lower_90 = np.maximum(0, cdf - 1.645 * np.sqrt(cdf * (1 - cdf) / n))  # 90% confianza  
cdf_upper_90 = np.minimum(1, cdf + 1.645 * np.sqrt(cdf * (1 - cdf) / n))  
cdf_lower_95 = np.maximum(0, cdf - 1.96 * np.sqrt(cdf * (1 - cdf) / n))  # 95% confianza  
cdf_upper_95 = np.minimum(1, cdf + 1.96 * np.sqrt(cdf * (1 - cdf) / n))  

# Contar la cantidad de datos que caen dentro de cada límite de confianza  
count_within_90 = np.sum((data[nombre_columna] >= np.min(x[cdf_lower_90 > 0])) & (data[nombre_columna] <= np.max(x[cdf_upper_90 < 1])))  
count_within_95 = np.sum((data[nombre_columna] >= np.min(x[cdf_lower_95 > 0])) & (data[nombre_columna] <= np.max(x[cdf_upper_95 < 1])))  

# Imprimir los resultados  
print("-" * 100)
print(" " * 100)
print(f'cantidad de datos: {n}')
print(f'Cantidad de datos dentro de los límites de confianza del 90%: {count_within_90}')  
print(f'Cantidad de datos dentro de los límites de confianza del 95%: {count_within_95}') 
print(f'Porcentaje de datos que caen en los límites de confianza del 90%: {((count_within_90 / n) * 100):.2f}%')  
print(f'Porcentaje de datos que caen en los límites de confianza del 95%: {((count_within_95 / n) * 100):.2f}%')  
print(" " * 100)
print("-" * 100)

# Calcular los valores asociados a las recurrencias  
recurrencias = [2, 5, 10, 25, 50, 100, 200, 500, 1000, 2000, 5000, 10000]  
valores_recurrencia = []  

for T in recurrencias:  
    P_a = 1 - (1 / T)  # Probabilidad acumulada asociada  
    valor = stats.genextreme.ppf(P_a, c, loc=loc, scale=scale)  # Cuantil inverso de la GEV  
    valores_recurrencia.append(valor)  

# Crear un DataFrame con los resultados  
resultados = pd.DataFrame({  
    'Recurrencia (años)': recurrencias,  
    'Valor asociado (mm)': valores_recurrencia  
})  

# Crear un DataFrame con las CDF y límites de confianza  
cdf_data = pd.DataFrame({  
    'x': x,  
    'CDF GEV': cdf,  
    'CDF Empírica': empirical_cdf_interp,  
    'Límite Inferior 90%': cdf_lower_90,  
    'Límite Superior 90%': cdf_upper_90,  
    'Límite Inferior 95%': cdf_lower_95,  
    'Límite Superior 95%': cdf_upper_95  
})  

# Crear un DataFrame con los parámetros de la GEV  
parametros_gev = pd.DataFrame({  
    'Parámetro': ['Forma (c)', 'Ubicación (loc)', 'Escala (scale)', 'R² (%)'],  
    'Valor': [c, loc, scale, r_squared_percentage]  
})  

# Exportar los resultados a un archivo Excel  
output_file_path = 'C:/1.PYTHON/Descarga_Python/Resultados_GEV.xlsx'  
with pd.ExcelWriter(output_file_path) as writer:  
    cdf_data.to_excel(writer, sheet_name='CDF y Límites', index=False)
    parametros_gev.to_excel(writer, sheet_name='Parámetros GEV', index=False)  
    resultados.to_excel(writer, sheet_name='Valores Recurrencia', index=False)  

print(" " * 100)
print(f'Resultados exportados a {output_file_path}')  
print(" " * 100)
print("-" * 100)  

# Graficar la CDF ajustada y la CDF empírica  
plt.figure(figsize=(10, 6))  
plt.plot(x, cdf, label='GEV Ajustada', color='magenta', linewidth=2.5)  
plt.scatter(data_sorted, empirical_cdf, label='Fex', color='blue', marker='o')  # CDF empírica como puntos  

# Dibujar los límites de confianza como líneas  
plt.plot(x, cdf_lower_90, color='red', linestyle='-.', label='Límite Inferior 90%')  
plt.plot(x, cdf_upper_90, color='red', linestyle='--', label='Límite Superior 90%')  
plt.plot(x, cdf_lower_95, color='green', linestyle='-.', label='Límite Inferior 95%')  
plt.plot(x, cdf_upper_95, color='green', linestyle='--', label='Límite Superior 95%')  

# Formatear el eje y como porcentaje  
plt.gca().yaxis.set_major_formatter(mtick.PercentFormatter(1.0))


# Título y etiquetas  
plt.title(f'Ajuste GEV con Límites de Confianza - Est. {estación}', fontweight='bold')  
plt.xlabel('Precipitación Diaria Máxima Anual (mm)', fontweight='bold')  
plt.ylabel('Probabilidad de No Excedencia', fontweight='bold')    

# Añadir R² a la leyenda  
plt.legend(loc='lower right', frameon=True, shadow=True, facecolor='white', framealpha=0.95, edgecolor="black")  
plt.gca().get_legend().get_texts()[0].set_text(f'GEV Ajustada (R² = {r_squared_percentage:.2f}%)')


# plt.legend(loc='lower right')  
# plt.gca().get_legend().get_texts()[0].set_text(f'GEV Ajustada (R² = {r_squared_percentage:.2f}%)')  # Actualizar la leyenda de GEV Ajustada  

# Añadir grillas mayor y menor  
plt.grid(which='both', color='grey', linestyle='-', linewidth=0.5)  # Grilla mayor  
plt.minorticks_on()  # Activar ticks menores  
plt.grid(which='minor', color='lightgrey', linestyle=':', linewidth=0.5)  # Grilla menor  

# Guardar la figura  
plt.savefig(f'C:/1.PYTHON/Descarga_Python/Ajuste_GEV_Límites_Confianza.png', dpi=1800)  

####################################################################################################
print('                                                                                        ')
print('PARAMETROS GEV')
print('shape:', c, '\nloc:', loc, '\nscale:', scale, '\nR2:', f"{r_squared_percentage:.2f}%")
print('                                                                                        ')
print('')
# Imprimir PTR -TR
# Formatear los valores de 'PTR (mm)' a dos decimales  
resultados['Valor asociado (mm)'] = resultados['Valor asociado (mm)'].map('{:.2f}'.format)  
print(resultados) 


