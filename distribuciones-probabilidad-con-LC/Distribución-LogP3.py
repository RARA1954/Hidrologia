'''
 Script para ajustar la distribución LogPearson III (LP3)
 a caudales diarios máximos anuales (QDMA),
 con límites de confianza del 90% y 95%

 se utiliza 'stats.pearson3'

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
input_file_path = "C:/1.PYTHON/Descarga_Python/QDMA.xlsx"  # Cambia esto a la ruta de tu archivo  

# Establecer la hoja donde se encuentran los datos
nombre_hoja ='Sheet1'

# Definir la variable para la columna de precipitación o caudales 
nombre_columna = 'QDMA'  # Cambia este valor si es necesario  

# Establecer la estación de medición
estación = 'Paso de Indios'

##########################################################################################################

# Leer el archivo Excel  
data = pd.read_excel(input_file_path, sheet_name= nombre_hoja)

# Imprimir los nombres de las columnas del DataFrame  
print("Columnas en el archivo de entrada:", data.columns.tolist())

# Asegurarse de que la columna de caudales esté en formato numérico  
data[nombre_columna] = pd.to_numeric(data[nombre_columna], errors='coerce')    

# Eliminar filas con valores NaN  
data = data.dropna(subset=[nombre_columna])  

# Aplicar logaritmo a los datos para ajustar la distribución Log Pearson III  
log_data = np.log(data[nombre_columna])  

# Ajustar la distribución Pearson III a los datos logarítmicos  
skew, loc, scale = stats.pearson3.fit(log_data)  

# Crear un rango de valores para la gráfica  
x = np.linspace(np.min(data[nombre_columna]), np.max(data[nombre_columna]), len(data[nombre_columna]))  

# Calcular la CDF de la Log Pearson III ajustada  
log_x = np.log(x)  # Transformar x al espacio logarítmico  
cdf = stats.pearson3.cdf(log_x, skew, loc=loc, scale=scale)  

# Calcular la CDF empírica  
data_sorted = np.sort(data[nombre_columna])  
empirical_cdf = np.arange(1, len(data_sorted) + 1) / len(data_sorted)  

# Interpolar la CDF empírica para que coincida con el rango de x  
interp_func = interp1d(data_sorted, empirical_cdf, bounds_error=False, fill_value=(0, 1))  
empirical_cdf_interp = interp_func(x)  

# Calcular R²  
sst = np.sum((empirical_cdf_interp - np.mean(empirical_cdf_interp))**2)  
ssr = np.sum((empirical_cdf_interp - cdf)**2)  
r_squared = 1 - (ssr / sst)  
r_squared_percentage = r_squared * 100  

# Calcular límites de confianza del 90% y 95%  
alpha_90 = 0.10  
alpha_95 = 0.05  
n = len(data[nombre_columna])  

# Calcular los límites de confianza para la CDF  
cdf_lower_90 = np.maximum(0, cdf - 1.645 * np.sqrt(cdf * (1 - cdf) / n))  
cdf_upper_90 = np.minimum(1, cdf + 1.645 * np.sqrt(cdf * (1 - cdf) / n))  

cdf_lower_95 = np.maximum(0, cdf - 1.96 * np.sqrt(cdf * (1 - cdf) / n))  
cdf_upper_95 = np.minimum(1, cdf + 1.96 * np.sqrt(cdf * (1 - cdf) / n))  

# Contar la cantidad de datos que caen dentro de cada límite de confianza  
count_within_90 = np.sum((data[nombre_columna] >= np.min(x[cdf_lower_90 > 0])) & (data[nombre_columna] <= np.max(x[cdf_upper_90 < 1])))  
count_within_95 = np.sum((data[nombre_columna] >= np.min(x[cdf_lower_95 > 0])) & (data[nombre_columna] <= np.max(x[cdf_upper_95 < 1])))  

# Imprimir los resultados  
print('')
print('')
print('Cantidad de datos:', n)  
print(f'Cantidad de datos dentro de los límites de confianza del 90%: {count_within_90}')  
print(f'Cantidad de datos dentro de los límites de confianza del 95%: {count_within_95}')   
print(f'Porcentaje de datos que caen en los límites de confianza del 90%: {((count_within_90 / n) * 100):.2f}%')  
print(f'Porcentaje de datos que caen en los límites de confianza del 95%: {((count_within_95 / n) * 100):.2f}%')   

# Calcular los valores asociados a las recurrencias  
recurrencias = [2, 5, 10, 25, 50, 100, 200, 500, 1000, 5000, 10000]  
valores_recurrencia = []  

for T in recurrencias:  
    P_a = 1 - (1 / T)  # Probabilidad acumulada asociada  
    valor = stats.pearson3.ppf(P_a, skew, loc=loc, scale=scale)  # Cuantil inverso de la Pearson III  
    valores_recurrencia.append(np.exp(valor))  # Transformar de vuelta al espacio original  

# Crear un DataFrame con los resultados  
resultados = pd.DataFrame({  
    'Recurrencia (años)': recurrencias,  
    'Valor asociado (m³/s)': valores_recurrencia  
})  

# Crear un DataFrame con las CDF y límites de confianza  
cdf_data = pd.DataFrame({  
    'x': x,  
    'CDF LP3': cdf,  
    'CDF Empírica': empirical_cdf_interp,  
    'Límite Inferior 90%': cdf_lower_90,  
    'Límite Superior 90%': cdf_upper_90,  
    'Límite Inferior 95%': cdf_lower_95,  
    'Límite Superior 95%': cdf_upper_95  
})  

# Crear un DataFrame con los parámetros de la distribución  
parametros_lp3 = pd.DataFrame({  
    'Parámetro': ['Asimetría (skew)', 'Ubicación (loc)', 'Escala (scale)', 'R² (%)'],  
    'Valor': [skew, loc, scale, r_squared_percentage]  
})  

# Exportar los resultados a un archivo Excel  
output_file_path = 'C:/1.PYTHON/Descarga_Python/Resultados_LP3.xlsx'  
with pd.ExcelWriter(output_file_path) as writer:  
    cdf_data.to_excel(writer, sheet_name='CDF y Límites', index=False)  
    parametros_lp3.to_excel(writer, sheet_name='Parámetros LP3', index=False)  
    resultados.to_excel(writer, sheet_name='Valores Recurrencia', index=False)  

print('')
print('')
print(f'Resultados exportados a {output_file_path}')  

# Graficar la CDF ajustada y la CDF empírica  
plt.figure(figsize=(10, 6))  
plt.plot(x, cdf, label='LP3 Ajustada', color='magenta', linewidth=2.5)  
plt.scatter(data_sorted, empirical_cdf, label='Fex', color='blue', marker='o')  

# Dibujar los límites de confianza como líneas  
plt.plot(x, cdf_lower_90, color='red', linestyle='--', label='Límite Inferior 90%')  
plt.plot(x, cdf_upper_90, color='red', linestyle='--', label='Límite Superior 90%')  

plt.plot(x, cdf_lower_95, color='green', linestyle='-.', label='Límite Inferior 95%')  
plt.plot(x, cdf_upper_95, color='green', linestyle='-.', label='Límite Superior 95%')  

# Formatear el eje y como porcentaje  
plt.gca().yaxis.set_major_formatter(mtick.PercentFormatter(1.0))



# Título y etiquetas  
plt.title(f'Ajuste Log Pearson III (LP3) con Límites de Confianza - Estación {estación}', fontweight='bold') 
plt.xlabel('Caudal Máximo Anual Instantaneo (m³/s)', fontweight='bold')  
plt.ylabel('Probabilidad de No Excedencia', fontweight='bold')  

# Añadir R² a la leyenda  
plt.legend(loc='lower right', frameon=True, shadow=True, facecolor='white', framealpha=0.95, edgecolor="black")  
plt.gca().get_legend().get_texts()[0].set_text(f'LP3 Ajustada (R² = {r_squared_percentage:.2f}%)')

# plt.legend(loc='lower right')  
# plt.gca().get_legend().get_texts()[0].set_text(f'LP3 Ajustada (R² = {r_squared_percentage:.2f}%)')  

# Añadir grillas mayor y menor  
plt.grid(which='both', color='grey', linestyle='-', linewidth=0.5)  
plt.minorticks_on()  
plt.grid(which='minor', color='lightgrey', linestyle=':', linewidth=0.5)  

# Guardar la figura  
plt.savefig(f'C:/1.PYTHON/Descarga_Python/Ajuste Log Pearson III con Límites de Confianza', dpi=1800)  

# Mostrar el gráfico  
# plt.show()

################################################################################################################

# Asegúrate de que log_data sea un array de números  
log_data = np.array(log_data, dtype=float)  

# Calcular los momentos usando stats.describe  
desc = stats.describe(log_data)  
mean = desc.mean  
var = desc.variance  
skewness = stats.skew(log_data)  
kurtosis = stats.kurtosis(log_data)  

print('')
print('')
print('Estadística de los logaritmos de los datos')
print(f"Media: {mean:.4f}")  
print(f"Varianza: {var:.4f}")  
print(f"Sesgo: {skewness:.4f}")  
print(f"Curtosis: {kurtosis:.4f}") 

####################################################################################################
print('                                                                                        ')
print('PARAMETROS LP3')
print('shape:', skew, '\nloc:', loc, '\nscale:', scale, '\nR2:', f"{r_squared_percentage:.2f}%")
print('                                                                                        ')
print('')
# Imprimir Valor asociado -TR
# Formatear los valores de 'Valor asociado (mm)' a dos decimales  
resultados['Valor asociado (m³/s)'] = resultados['Valor asociado (m³/s)'].map('{:.2f}'.format)  
print(resultados) 
