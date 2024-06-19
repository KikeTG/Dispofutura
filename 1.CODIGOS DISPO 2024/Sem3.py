#  COPIAR ARCHIVO BASE SP, MAESTRA AFM, DEFINIR CARPETA DE DESTINO PRINCIPAL, CREAR CARPETA SEMANA ACTUAL, TUBO SEMANAL, STOCK TIENDA, STOCK, TRANSITO CONSOLIDADO
# ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
import os
import shutil
from datetime import datetime
import calendar
import pandas as pd

user_home_dir = os.path.expanduser('~')

# Obtener el mes actual en formato de nombre (por ejemplo, 'octubre')
# Obtener el mes actual y el mes siguiente
mes_actual_numero = datetime.now().strftime('%m')
mes_siguiente_numero = int(mes_actual_numero) % 12 + 1  # Calcula el mes siguiente
mes_actual = datetime.now().strftime('%B').lower()
mes_siguiente = calendar.month_name[mes_siguiente_numero].lower()

# Mapeo de nombres de meses en inglés a español
meses = {
    'january': 'enero',
    'february': 'febrero',
    'march': 'marzo',
    'april': 'abril',
    'may': 'mayo',
    'june': 'junio',
    'july': 'julio',
    'august': 'agosto',
    'september': 'septiembre',
    'october': 'octubre',
    'november': 'noviembre',
    'december': 'diciembre'
}
# Obtener el nombre del mes en español
mes_siguiente = meses.get(mes_siguiente, mes_siguiente)


# Ruta de la carpeta donde se encuentran los archivos
# carpeta_origen = r'C:\Users\Etorres\Inchcape\Open to Buy_OTB\2023 OK'

carpeta_origen = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos', 'Planificación y Compras AFM', 'Open to Buy_OTB', '2024')

# RUTA DEL "C:\Users\Etorres\Inchcape\Maestros Actualizables\Maestra Aftermarket Actualizable.xlsx"
# ruta_AFMACTUALIZABLE = r'C:\Users\Etorres\Inchcape\Maestros Actualizables\Maestra Aftermarket Actualizable.xlsx'

ruta_AFMACTUALIZABLE = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos', 'Planificación y Compras Maestros', 'Vigencias', 'Maestra Retail', 'Maestra Aftermarket Actualizable.xlsx')


# Identificar la subcarpeta que contiene el mes actual en su nombre
subcarpetas = [d for d in os.listdir(carpeta_origen) if os.path.isdir(os.path.join(carpeta_origen, d))]
subcarpeta_mes = next((d for d in subcarpetas if mes_actual in d.lower()), None)
subcarpeta_mes_siguiente = next((d for d in subcarpetas if mes_siguiente in d.lower()), None)

if subcarpeta_mes_siguiente:
    carpeta_origen_mes_siguiente = os.path.join(carpeta_origen, subcarpeta_mes_siguiente)
else:
    raise ValueError(f"No se encontró ninguna subcarpeta que contenga el mes '{mes_siguiente}' en su nombre.")

archivos_otb = [f for f in os.listdir(carpeta_origen_mes_siguiente) 
                if os.path.isfile(os.path.join(carpeta_origen_mes_siguiente, f)) 
                and 'SP' in f]

if not archivos_otb:
    raise ValueError(f"No se encontraron archivos del OTB Aprobado que contengan 'SP' y '{mes_siguiente}' en su nombre.")

# Seleccionar el archivo más reciente
archivo_otb_reciente = max(archivos_otb, key=lambda x: os.path.getmtime(os.path.join(carpeta_origen_mes_siguiente, x)))

# Ruta completa del archivo OTB Aprobado más reciente
ruta_otb_reciente = os.path.join(carpeta_origen_mes_siguiente, archivo_otb_reciente) 


# carpeta_destino_principal = r'C:\Users\Etorres\OneDrive - Inchcape\Archivo Base Dispo'

carpeta_destino_principal = os.path.join(user_home_dir, 'OneDrive - Inchcape', 'Archivo Base Dispo')


# Obtener el número de la semana actual
semana_actual = datetime.now().isocalendar()[1]

# Ruta de la carpeta de destino con la semana actual
carpeta_destino_semana = os.path.join(carpeta_destino_principal, f'Sem {semana_actual}')

# Verificar si la carpeta de destino con la semana actual existe o no
if not os.path.exists(carpeta_destino_semana):
    # Si no existe, crear la carpeta
    os.makedirs(carpeta_destino_semana)


# Verificar si el archivo "Maestra Aftermarket Actualizable.xlsx" existe
if os.path.exists(ruta_AFMACTUALIZABLE):
    try:
        # Copiar el archivo "Maestra Aftermarket Actualizable.xlsx" a la carpeta de la semana actual
        shutil.copy2(ruta_AFMACTUALIZABLE, carpeta_destino_semana)
        print(f'Archivo "Maestra Aftermarket Actualizable.xlsx" copiado a la carpeta de la semana actual.')
        
        # Verificar si la copia fue exitosa
        archivo_copiado = os.path.join(carpeta_destino_semana, os.path.basename(ruta_AFMACTUALIZABLE))
        if os.path.exists(archivo_copiado):
            print("La copia del archivo se realizó correctamente.")
        else:
            print("No se encontró el archivo copiado en la carpeta de destino.")
    except Exception as e:
        print(f"Ocurrió un error al copiar el archivo: {e}")
else:
    print(f'El archivo "{ruta_AFMACTUALIZABLE}" no existe o no se encuentra en la ruta especificada.')


ruta_destino_otb = os.path.join(carpeta_destino_semana, archivo_otb_reciente)
if os.path.exists(ruta_destino_otb):
    print(f'El archivo del OTB Aprobado ya existe en la carpeta de destino: "{ruta_destino_otb}"')
else:
    # Copiar el archivo del OTB Aprobado a la carpeta de destino
    shutil.copy2(ruta_otb_reciente, carpeta_destino_semana)
    print(f'Archivo del OTB Aprobado copiado a la carpeta de destino.')


from datetime import datetime, timedelta

# Obtener el primer día del mes actual
primer_dia_mes_actual = datetime.now().replace(day=1)

# Restar un día para obtener el último día del mes anterior
ultimo_dia_mes_anterior = primer_dia_mes_actual - timedelta(days=1)

# Extraer el mes anterior
mes_anterior_numero = ultimo_dia_mes_anterior.strftime('%m')

# cod actual s4 mes anterior
# ruta_base = "C:\\Users\\Etorres\\OneDrive - Inchcape\\Planificación y Compras Maestros\\2023\\"
ruta_base = os.path.join(user_home_dir, "Inchcape", "Planificación y abastecimiento AFM - Documentos", "Planificación y Compras Maestros", "2024")

nombre_carpeta = f"2024-{mes_anterior_numero}"
ruta_carpeta_mes = os.path.join(ruta_base, nombre_carpeta)
# Obtener el año actual en formato de cadena
anioo_actual = datetime.now().strftime('%Y')

# Crear la ruta de la carpeta del mes actual
nombre_carpeta_actual = f"{anioo_actual}-{mes_actual_numero}"
ruta_carpeta_mes_actual = os.path.join(ruta_base, nombre_carpeta_actual)

archivo_encontrado = None
for archivo in os.listdir(ruta_carpeta_mes_actual):
    if archivo.startswith("COD_ACTUAL_S4"):
        archivo_encontrado = archivo
        break

if archivo_encontrado:
    ruta_origen = os.path.join(ruta_carpeta_mes_actual, archivo_encontrado)
    shutil.copy(ruta_origen, carpeta_destino_semana)
    print(f"Archivo {archivo_encontrado} copiado exitosamente a {carpeta_destino_semana}")
else:
    print("No se encontró el archivo que comienza con 'COD_ACTUAL_S4'")
def es_formato_fecha_valido(nombre_carpeta):
    try:
        datetime.strptime(nombre_carpeta, "%Y-%m-%d")
        return True
    except ValueError:
        return False


# Ruta donde están las carpetas con fechas
# ruta_kpi_reportes = r'C:\Users\Etorres\OneDrive - Inchcape\KPI\Tubo Semanal'

ruta_kpi_reportes = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos','KPI', 'Tubo Semanal')

# Obtener una lista de todas las carpetas en la ruta_kpi_reportes que cumplan con el formato de fecha
carpetas = [d for d in os.listdir(ruta_kpi_reportes) if os.path.isdir(os.path.join(ruta_kpi_reportes, d)) and es_formato_fecha_valido(d)]

# Ordenar las carpetas por su fecha
carpetas_ordenadas = sorted(carpetas, key=lambda x: datetime.strptime(x, "%Y-%m-%d"), reverse=True)

# Tomar la carpeta más cercana a la fecha actual (la primera en la lista ordenada)
carpeta_mas_cercana = carpetas_ordenadas[0]

# Ruta completa de la carpeta seleccionada
ruta_carpeta_seleccionada = os.path.join(ruta_kpi_reportes, carpeta_mas_cercana)

# Buscar el archivo que contiene "Stock Tiendas" en su nombre
archivo_stock_tiendas = next((f for f in os.listdir(ruta_carpeta_seleccionada) if "Stock Tiendas" in f), None)

# Si encontramos el archivo, copiamos a la carpeta_destino_semana
if archivo_stock_tiendas:
    ruta_origen_stock = os.path.join(ruta_carpeta_seleccionada, archivo_stock_tiendas)
    shutil.copy2(ruta_origen_stock, carpeta_destino_semana)
    print(f'Archivo "{archivo_stock_tiendas}" copiado a la carpeta destino "{carpeta_destino_semana}".')
else:
    print(f"No se encontró ningún archivo con 'Stock Tiendas' en su nombre en la carpeta {carpeta_mas_cercana}.")

archivo_stock_tiendas2 = next((f for f in os.listdir(ruta_carpeta_seleccionada) if "Stock S4" in f and not ('Tiendas' in f or 'Pa' in f or 'Centro' in f or 'R3' in f)), None)

#     # Si encontramos el archivo, copiamos a la carpeta_destino_semana
if archivo_stock_tiendas2:
    ruta_origen_stock2 = os.path.join(ruta_carpeta_seleccionada, archivo_stock_tiendas2)
    shutil.copy2(ruta_origen_stock2, carpeta_destino_semana)
    print(f'Archivo "{archivo_stock_tiendas2}" copiado a la carpeta destino "{carpeta_destino_semana}".')
else:
    print(f"No se encontró ningún archivo con 'Stock S4' en su nombre en la carpeta {carpeta_mas_cercana}.")
archivo_stock_tiendas3 = next((f for f in os.listdir(ruta_carpeta_seleccionada) if "onsolidado" in f and 'S4' in f), None)

if archivo_stock_tiendas3:
    ruta_origen_stock3 = os.path.join(ruta_carpeta_seleccionada, archivo_stock_tiendas3)
    ruta_destino_stock3 = os.path.join(carpeta_destino_semana, archivo_stock_tiendas3)

    # Verificar si el archivo ya existe en la carpeta de destino
    if not os.path.exists(ruta_destino_stock3):
        shutil.copy2(ruta_origen_stock3, carpeta_destino_semana)
        print(f'Archivo "{archivo_stock_tiendas3}" (CONSOLIDADO) copiado a la carpeta destino "{carpeta_destino_semana}".')
    else:
        print(f'El archivo "{archivo_stock_tiendas3}" ya existe en la carpeta destino "{carpeta_destino_semana}".')
else:
    print(f"No se encontró ningún archivo con 'Stock CONSOLIDADO' en su nombre en la carpeta {ruta_carpeta_seleccionada}.")

# ruta_origen_stock3 = r'C:\Users\Etorres\OneDrive - Inchcape\Archivo Base Dispo\Sem 47\2023-11-13 TR semana 46 Consolidado.xlsx'
# ruta_origen_stock2 = r'C:\Users\Etorres\OneDrive - Inchcape\Archivo Base Dispo\Sem 47\2023-11-13 - Stock.XLSX'
# # ruta_origen_stock = r'C:\Users\Etorres\OneDrive - Inchcape\Archivo Base Dispo\Sem 47\2023-11-13 Stock Tiendas.XLSX'

# ruta_origen_stock3 = os.path.join(carpeta_destino_principal, 'Sem 47', '2023-11-13 TR semana 46 Consolidado.xlsx')
# ruta_origen_stock2 = os.path.join(carpeta_destino_principal, 'Sem 47', '2023-11-13 - Stock.XLSX')
# ruta_origen_stock = os.path.join(carpeta_destino_principal, 'Sem 47', '2023-11-13 Stock Tiendas.XLSX')

    
#BUSCAR FORECAST ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

import os
import shutil
from datetime import datetime, timedelta
import locale

# Ruta donde se encuentran las carpetas por año y mes
# directorio_forecast = r"C:\Users\Etorres\OneDrive - Inchcape\Forecast Inbound"

directorio_forecast = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos', 'Planificación y Compras AFM', 'Forecast Inbound')


# Mapeo de nombres de meses en inglés a español
meses_en_espanol = {
    "January": "Enero",
    "February": "Febrero",
    "March": "Marzo",
    "April": "Abril",
    "May": "Mayo",
    "June": "Junio",
    "July": "Julio",
    "August": "Agosto",
    "September": "Septiembre",
    "October": "Octubre",
    "November": "Noviembre",
    "December": "Diciembre"
}

# Establece el entorno local a inglés para obtener el mes en inglés
locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')

def obtener_nombre_carpeta(fecha):
    año = fecha.strftime("%Y")
    mes_inglés = fecha.strftime("%B")
    mes_español = meses_en_espanol[mes_inglés]
    mes_número = fecha.strftime("%m")
    return f"{año}-{mes_número} {mes_español}"

fecha_actual = datetime.now()
# No necesitas restarle días, ya que deseas el mes actual
nombre_carpeta_mes_actual = obtener_nombre_carpeta(fecha_actual)

# Buscar la carpeta del mes actual
ruta_carpeta_encontrada = ""
for nombre_carpeta in os.listdir(directorio_forecast):
    if nombre_carpeta_mes_actual in nombre_carpeta:
        ruta_carpeta_encontrada = os.path.join(directorio_forecast, nombre_carpeta)
        break

# Si se encontró la carpeta del mes anterior
if ruta_carpeta_encontrada:
    # Buscar el archivo con la palabra "Inbound" más reciente
    lista_archivos = [archivo for archivo in os.listdir(ruta_carpeta_encontrada) if "Inbound" in archivo]
    lista_archivos.sort(key=lambda archivo: os.path.getmtime(os.path.join(ruta_carpeta_encontrada, archivo)), reverse=True)
    
    if lista_archivos:
        archivo_reciente = lista_archivos[0]
        ruta_completa_archivo = os.path.join(ruta_carpeta_encontrada, archivo_reciente)
        
        # Copiar el archivo más reciente a la carpeta destino
        shutil.copy(ruta_completa_archivo, carpeta_destino_semana)
        print(f"El archivo {archivo_reciente} ha sido copiado a {carpeta_destino_semana} exitosamente.")
    else:
        print(f"No se encontraron archivos con la palabra 'Inbound' en {ruta_carpeta_encontrada}.")
else:
    print(f"No se encontró la carpeta para el mes anterior.")
    

#COPIAR Y PEGAR LO DEL SP Y CALCULAR LEADTIME SEMANAL  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

df = pd.read_excel(ruta_destino_otb, sheet_name="Base", usecols="A:B, D:I, K:L, N, CX, CY, ER, EU, M, ET, DM:DN", skiprows=1)

df.insert(11, 'Leadtime Semanal', (df.iloc[:, 10] / 7).round(2))
#CALCULO VIGENCIA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

nombres_columnas = ["Material SAP", "Condicion de Compra", "CES 01", "Mayorista", "Sodimac", "Easy", "Walmart", "SMU", "Tottus", "Retail (AP-AG)"]
ruta_archivo_2 = os.path.join(carpeta_destino_semana, "Maestra Aftermarket Actualizable.xlsx")
dataframe_maestra_2 = pd.read_excel(ruta_archivo_2, usecols=nombres_columnas)#, skiprows=1)

agrupado_2 = dataframe_maestra_2.groupby(by=["Material SAP", "Condicion de Compra"]).sum().reset_index()

dataframe_maestra_2['vig may'] = ((dataframe_maestra_2['CES 01'] + dataframe_maestra_2['Mayorista']) != 0).astype(int)

dataframe_maestra_2['vig gt'] = ((dataframe_maestra_2[['Sodimac', 'Easy', 'Walmart', 'SMU', 'Tottus']].sum(axis=1)) != 0).astype(int)

cond1 = dataframe_maestra_2['Retail (AP-AG)'] != 0
cond2 = dataframe_maestra_2['Condicion de Compra'] != 2
dataframe_maestra_2['vig retail'] = (cond1 & cond2).astype(int)

dataframe_maestra_2['vig total'] = ((dataframe_maestra_2[['vig may', 'vig gt', 'vig retail']].sum(axis=1)) != 0).astype(int)

#LEER CODIGO S4 MES ANTERIOR ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# 1. Buscar el archivo que comienza con 'COD_ACTUAL_S4' en la carpeta destino
archivo_cod_actual = None
for archivo in os.listdir(carpeta_destino_semana):
    if archivo.startswith("COD_ACTUAL_S4"):
        archivo_cod_actual = archivo
        break

# 2. Leer las columnas requeridas del archivo encontrado
ruta_archivo_cod_actual = os.path.join(carpeta_destino_semana, archivo_cod_actual)
df_cod_actual = pd.read_excel(ruta_archivo_cod_actual, usecols=["Nro_pieza_fabricante_1", "Cod_Actual_1"])

# Convertir las columnas a strings para asegurarse de que tengan el mismo tipo de datos
dataframe_maestra_2['Material SAP'] = dataframe_maestra_2['Material SAP'].astype(str)
df_cod_actual['Nro_pieza_fabricante_1'] = df_cod_actual['Nro_pieza_fabricante_1'].astype(str)

# 3. Realizar merge
dataframe_maestra_2 = dataframe_maestra_2.merge(df_cod_actual, left_on='Material SAP', right_on='Nro_pieza_fabricante_1', how='left')

# 4. Si no hay coincidencia en el cruce, llenar "Cod_Actual_1" con el valor de "Material SAP"
dataframe_maestra_2["Cod_Actual_1"].fillna(dataframe_maestra_2["Material SAP"], inplace=True)

# 5. Renombrar la columna y eliminar la columna adicional
dataframe_maestra_2.rename(columns={"Cod_Actual_1": "COD ACTUAL"}, inplace=True)
dataframe_maestra_2.drop(columns='Nro_pieza_fabricante_1', inplace=True)

# groupby
df_agrupado = dataframe_maestra_2.groupby('COD ACTUAL').agg({
    'vig may': 'sum',
    'vig gt': 'sum',
    'vig retail': 'sum',
    'vig total': 'sum'
}).reset_index()

# Reemplazar valores mayores a 1 por 1
cols = ['vig may', 'vig gt', 'vig retail', 'vig total']
for col in cols:
    df_agrupado[col] = df_agrupado[col].apply(lambda x: 1 if x > 1 else x)

df_agrupado.head()  
# Convertir las columnas 'Material' y 'COD ACTUAL' a tipo object
df['Material'] = df['Material'].astype(str)
df_agrupado['COD ACTUAL'] = df_agrupado['COD ACTUAL'].astype(str)

# Hacer el merge, solo con las columnas especificadas
df_resultado = df.merge(df_agrupado[['COD ACTUAL', 'vig may', 'vig gt', 'vig retail']],
                        left_on='Material', 
                        right_on='COD ACTUAL', 
                        how='left')

# Rellenar con 0 en caso de que no haya coincidencias
for columna in ['vig may', 'vig gt', 'vig retail']:
    df_resultado[columna].fillna(0, inplace=True)

# Eliminar la columna 'COD ACTUAL'
df_resultado = df_resultado.drop('COD ACTUAL', axis=1)

# Crear la columna 'Vigencia Total'
df_resultado['Vigencia Total'] = df_resultado[['vig may', 'vig gt', 'vig retail']].sum(axis=1).apply(lambda x: 1 if x != 0 else 0)

# Reemplazar df con el dataframe resultante
df = df_resultado

# OBSOLECENCIA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# Obtener la lista de columnas del dataframe
columnas = df_resultado.columns.tolist()

# Remover 'OBS Retail' y 'OBS DERCO' de la lista
columnas.remove('OBS Retail')
columnas.remove('OBS DERCO')

# Ajustar los valores negativos a 0 en 'OBS Retail' y 'OBS DERCO'
df_resultado['OBS Retail'] = df_resultado['OBS Retail'].apply(lambda x: 0 if x < 0 else x)
df_resultado['OBS DERCO'] = df_resultado['OBS DERCO'].apply(lambda x: 0 if x < 0 else x)

# Agregar 'OBS Retail' y 'OBS DERCO' al final de la lista
columnas.extend(['OBS Retail', 'OBS DERCO'])

# Reorganizar las columnas 
df_resultado = df_resultado[columnas]

df_resultado = df_resultado.copy()

# Crear la columna 'Obs Total' utilizando .loc
df_resultado.loc[:, 'Obs Total'] = (df_resultado['OBS Retail'] + df_resultado['OBS DERCO'] > 0).astype(int)

# Agregar 'Obs Total' al final de la lista de columnas
columnas.append('Obs Total')

# Reorganizar las columnas del dataframe para asegurarse de que 'Obs Total' esté al final
df_resultado = df_resultado[columnas]
#FALTANTES Y SOBRANTES ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

import re
from datetime import datetime

# Ruta de origen
# ruta_origen = r"C:\Users\Etorres\Inchcape\KPI\Faltantes y Sobrantes Retail\Consolidado Sobrantes y Faltantes AP - AGP"

ruta_origen = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos', 'KPI', 'Faltantes y Sobrantes Retail', 'Sobrantes y Faltantes AP - AGP')

# Lista todos los archivos en la ruta de origen
archivos = [f for f in os.listdir(ruta_origen) if os.path.isfile(os.path.join(ruta_origen, f))]

# Función para extraer la fecha del nombre del archivo
def extraer_fecha(archivo):
    # Intenta el primer patrón
    match = re.search(r"(\d{2}.\d{2}.\d{4})", archivo)
    if match:
        fecha_str = match.group(1)
        return datetime.strptime(fecha_str, "%d.%m.%Y")
    # Si el primer patrón no funciona, intenta el segundo
    match = re.search(r"(\d{2}.\d{2}.\d{2})", archivo)
    if match:
        fecha_str = match.group(1)
        return datetime.strptime(fecha_str, "%d.%m.%y")
    return None

# Obtener el mes y año actual
mes_actual = datetime.now().month
año_actual = datetime.now().year

# Si estamos en enero, el mes anterior sería diciembre del año pasado
if mes_actual == 1:
    mes_anterior = 12
    año_anterior = año_actual - 1
else:
    mes_anterior = mes_actual - 1
    año_anterior = año_actual

# Filtra solo los archivos que contienen una fecha en su nombre son excel y del mes pasado
# archivos_filtrados = [
#     f for f in archivos if extraer_fecha(f) 
#     and extraer_fecha(f).month == mes_anterior 
#     and extraer_fecha(f).year == año_anterior 
#     and (f.endswith('.xlsx') or f.endswith('.xls')) 
#     and "xD".lower() not in f.lower()
# ]
# Filtra solo los archivos que contienen una fecha en su nombre son excel y mes actual
archivos_filtrados = [
    f for f in archivos if extraer_fecha(f) 
    and extraer_fecha(f).month == mes_actual   # Cambiado a mes_actual
    and extraer_fecha(f).year == año_actual    # Cambiado a año_actual
    and (f.endswith('.xlsx') or f.endswith('.xls')) 
    and "xD".lower() not in f.lower()
]

# Si no hay archivos que cumplan el criterio
if not archivos_filtrados:
    print("No se encontró un archivo que cumpla con los criterios.")
    exit()

# Ordena los archivos por fecha
archivos_filtrados.sort(key=extraer_fecha, reverse=True)

# Toma el archivo con la fecha más cercana al mes actual (o el único archivo si solo hay uno)
archivo_a_mover = archivos_filtrados[0]

try:
    shutil.copy(os.path.join(ruta_origen, archivo_a_mover), carpeta_destino_semana)
    print(f"Archivo {archivo_a_mover} copiado exitosamente a {carpeta_destino_semana}")
except shutil.Error as e:
    print(f"Error al copiar el archivo: {e}")
except Exception as e:
    print(f"Error inesperado: {e}")


# Define la ruta completa de destino
ruta_destino_completa = os.path.join(carpeta_destino_semana, archivo_a_mover)
# ruta_destino_completa = r'C:\Users\Etorres\Inchcape\Archivo Base Dispo\Sem 47\Falt & Sobr AP Y AGP_06.11.23 (Mat-Cen).XLSX'
# ruta_destino_completa = os.path.join(user_home_dir, 'Inchcape', 'Archivo Base Dispo', 'Sem 47', 'Falt & Sobr AP Y AGP_06.11.23 (Mat-Cen).XLSX')

import pandas as pd

# Función auxiliar para buscar una hoja que contenga 'TB' en su nombre
def encontrar_hoja_tb(ruta_destino_completa):
    xl = pd.ExcelFile(ruta_destino_completa)
    for sheet_name in xl.sheet_names:
        if 'TB' in sheet_name:
            return sheet_name
    return None  # Retornar None si no se encuentra ninguna hoja

# Cargar la hoja correcta
hoja_tb = encontrar_hoja_tb(ruta_destino_completa)
if hoja_tb is None:
    print("No se encontró una hoja que contenga 'TB' en su nombre.")
    exit()

# Cargar solo las primeras 5 columnas
df = pd.read_excel(ruta_destino_completa, sheet_name=hoja_tb, usecols=range(5))

# Preprocesamiento para manejar variaciones en los nombres de las columnas
# Crear un mapeo de nombres esperados a posibles variaciones
column_mapping = {
    'POG': ['POG'],
    'Ultimo Slabon': ['Ultimo Slabon', 'Ultmo Slabon', 'Slabon', 'Último Eslabón'],
    'Faltante': ['Faltante', 'Faltantes', 'Faltante '],
    'Sobrante': ['Sobrante', 'Sobrantes', 'Sobrante '],
    'Stock Objetivo': ['Stock Objetivo', 'Stock Objetivo ']
}

# Normalizar los nombres de las columnas
for expected, variations in column_mapping.items():
    for variation in variations:
        if variation in df.columns:
            df.rename(columns={variation: expected}, inplace=True)

# Filtrar por POG = 'AP' y luego realizar el groupby por 'Ultimo Slabon'
# Sumar las columnas 'Faltante' y 'Sobrante'
Resumen_AP = df[df['POG'] == 'AP'].groupby('Ultimo Slabon')[['Faltante', 'Sobrante']].sum().reset_index()

# Mostrar el DataFrame 'Resumen AP'
print(Resumen_AP)
Resumen_AGP = df[df['POG'] == 'AGP'].groupby('Ultimo Slabon')[['Faltante', 'Sobrante']].sum().reset_index()


df_resultado['Material'] = df_resultado['Material'].astype(str).str.strip().str.lower()
Resumen_AP[Resumen_AP.columns[0]] = Resumen_AP[Resumen_AP.columns[0]].astype(str).str.strip().str.lower()
Resumen_AGP[Resumen_AGP.columns[0]] = Resumen_AGP[Resumen_AGP.columns[0]].astype(str).str.strip().str.lower()

merged_df = pd.merge(df_resultado, Resumen_AP, left_on='Material', right_on=Resumen_AP.columns[0], how='left')
merged_df.drop(Resumen_AP.columns[0], axis=1, inplace=True)
merged_df.rename(columns={'Faltante': 'Faltante AP', 'Sobrante': 'Sobrante AP'}, inplace=True)
# Reemplazando NaN por 0
merged_df['Faltante AP'] = merged_df['Faltante AP'].fillna(0)
merged_df['Sobrante AP'] = merged_df['Sobrante AP'].fillna(0)

# Segundo merge: Uniendo merged_df con Resumen_AGP
merged_df2 = pd.merge(merged_df, Resumen_AGP, left_on='Material', right_on=Resumen_AGP.columns[0], how='left')
merged_df2.drop(Resumen_AGP.columns[0], axis=1, inplace=True)
merged_df2.rename(columns={'Faltante': 'Faltante AGP', 'Sobrante': 'Sobrante AGP'}, inplace=True)
# Reemplazando NaN por 0
merged_df2['Faltante AGP'] = merged_df2['Faltante AGP'].fillna(0)
merged_df2['Sobrante AGP'] = merged_df2['Sobrante AGP'].fillna(0)


merged_df2['Faltantes'] = merged_df2['Faltante AP'] + merged_df2['Faltante AGP']
merged_df2['Sobrantes'] = merged_df2['Sobrante AP'] + merged_df2['Sobrante AGP']

columnas_existentes = [col for col in merged_df2.columns if col not in ['Material', 'Faltante AP', 'Faltante AGP', 'Sobrante AP', 'Sobrante AGP', 'Faltantes', 'Sobrantes']]
column_order = ['Material'] + columnas_existentes + ['Faltante AP', 'Faltante AGP', 'Faltantes', 'Sobrante AP', 'Sobrante AGP', 'Sobrantes']


merged_df2 = merged_df2[column_order]

#PLANOGRAMA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


# Obteniendo la primera columna de Resumen_AP y Resumen_AGP
primer_columna_AP = Resumen_AP.columns[0]
primer_columna_AGP = Resumen_AGP.columns[0]

# Creando la columna 'AP' en Merged_df2
merged_df2['AP'] = merged_df2['Material'].isin(Resumen_AP[primer_columna_AP]).astype(int)

# Creando la columna 'AGP' en Merged_df2
merged_df2['AGP'] = merged_df2['Material'].isin(Resumen_AGP[primer_columna_AGP]).astype(int)
merged_df2['Total Planograma'] = ((merged_df2['AGP'] + merged_df2['AP']) != 0).astype(int)
#CRUCE STOCK TIENDAS ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

import pandas as pd
ruta_destino_stock = os.path.join(carpeta_destino_semana, archivo_stock_tiendas)
# Cargar el DataFrame desde el archivo Excel
df123 = pd.read_excel(ruta_destino_stock, sheet_name="Sheet1", usecols=['Material', 'Libre utilización', 'Almacén'])


# Convertir 'Material' a entero (si es posible), luego a cadena para eliminar '.0'
df123['Material'] = pd.to_numeric(df123['Material'], errors='coerce').fillna(0).astype(int).astype(str)

# Convertir 'Libre utilización' y 'Almacén' a numérico
df123['Libre utilización'] = pd.to_numeric(df123['Libre utilización'], errors='coerce')
df123['Almacén'] = pd.to_numeric(df123['Almacén'], errors='coerce')

# Filtrar por 'Almacén' para quedarse con 1100.0 y NaN
# filtered_df = df123[(df123['Almacén'] == 1100.0) | (df123['Almacén'].isna())]
filtered_df = df123

# Realizar el cruce (merge) entre los DataFrames
TiendaUE = pd.merge(filtered_df, df_cod_actual, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')

# Rellenar los NaN en 'Cod_Actual_1' con los valores de 'Material' original
TiendaUE['Cod_Actual_1'] = TiendaUE['Cod_Actual_1'].fillna(TiendaUE['Material'])

# Eliminar la columna 'Material' y renombrar 'Cod_Actual_1' a 'Material'
TiendaUE = TiendaUE.drop('Material', axis=1)
TiendaUE.rename(columns={'Cod_Actual_1': 'Material'}, inplace=True)

# Aquí se ajustan los valores de 'Libre utilización' menores a 0 a 0
TiendaUE['Libre utilización'] = TiendaUE['Libre utilización'].apply(lambda x: 0 if x < 0 else x)

# Agrupar por 'Material' y sumar 'Libre utilización'
df_stock = TiendaUE.groupby('Material')['Libre utilización'].sum().reset_index()

# Ajustar valores menores a 1 a 0
df_stock['Libre utilización'] = df_stock['Libre utilización'].apply(lambda x: 0 if x < 1 else x)

# Convertir 'Libre utilización' a entero para eliminar decimales no necesarios
df_stock['Libre utilización'] = df_stock['Libre utilización'].astype(int)

duplicados = df_stock['Material'].duplicated()

# Para ver si hay al menos un duplicado en esa columna
hay_duplicados = duplicados.any()
print("¿Hay duplicados?:", hay_duplicados)

# Para contar el número total de duplicados en esa columna
numero_duplicados = duplicados.sum()
print("Número de duplicados:", numero_duplicados)

# Para obtener un DataFrame con todos los duplicados en esa columna
df_duplicados = df_stock[df_stock['Material'].duplicated(keep=False)]
print(df_duplicados)

df_stock.rename(columns={'Libre utilización': 'Stock Tiendas'}, inplace=True)
df_stock['Material'] = df_stock['Material'].fillna(0).astype(int)
df_stock['Stock Tiendas'] = df_stock['Stock Tiendas'].fillna(0).astype(int)

#STOCK CD ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
import pandas as pd
import numpy as np

# Especifica las columnas que necesitas
# cols_to_use3 = ['Material', 'Libre utilización', 'Alm.', 'Ce.', 'Insp. Calidad', 'Traslado', 'Ult. Eslabon']
cols_to_use3 = ['Ult. Eslabon', 'Libre utilización', 'Almacén', 'Centro', 'Inspecc.de calidad', 'Trans./Trasl.']
# Convertir la columna 'Material' a tipo object (string) para evitar el .0
# df_filtered1.rename(columns={'Inspecc.de calidad': 'Material'}, inplace=True)

# Convertir las columnas 'Material' y 'COD ACTUAL' a tipo object
if 'Material' in df.columns:
    df['Material'] = df['Material'].astype(str)
# Si la columna 'Material' no está, buscar 'Ultimo Slabon'
elif 'Ultimo Slabon' in df.columns:
    df['Ultimo Slabon'] = df['Ultimo Slabon'].astype(str)
# Si ninguna de las dos columnas está presente, no hacer nada
else:
    pass
df_agrupado['COD ACTUAL'] = df_agrupado['COD ACTUAL'].astype(str)
ruta_destino_stock2 = os.path.join(carpeta_destino_semana, archivo_stock_tiendas2)
# Carga solo las columnas que necesitas desde la hoja "Sheet1"
dfleerstock = pd.read_excel(ruta_destino_stock2, sheet_name='Sheet1', usecols=cols_to_use3)
dfleerstock.head()


# Convierte 'Alm.' a entero, maneja NaN y luego convierte a cadena
# dfleerstock['Alm.'] = pd.to_numeric(dfleerstock['Alm.'], errors='coerce').fillna(0).astype(int).astype(str)
dfleerstock['Almacén'] = pd.to_numeric(dfleerstock['Almacén'], errors='coerce').fillna(0).astype(int).astype(str)

# Convierte 'Ce.' a cadena
dfleerstock['Centro'] = dfleerstock['Centro'].astype(str)
# Convertir 'Ce.' a cadena y eliminar '.0'
dfleerstock['Centro'] = dfleerstock['Centro'].astype(str).str.replace('.0', '', regex=False)


# Define los valores permitidos para 'Alm.' y 'Ce.'
# allowed_alm = ['1100', '1600']
allowed_centers = ['710', '712', '713', '714', '0714', '0710', '0712', '0713']
print("Valores únicos en Alm.:", dfleerstock['Almacén'].unique())
print("Valores únicos en Ce.:", dfleerstock['Centro'].unique())

df_filtered1 = dfleerstock[(dfleerstock['Centro'].isin(allowed_centers))].copy()

# Convertir la columna 'Material' a tipo object (string) para evitar el .0

df_filtered1.rename(columns={'Ult. Eslabon': 'Material'}, inplace=True)
df_filtered1['Material'] = df_filtered1['Material'].astype(str)
df_filtered1.head()
FilteredUE = pd.merge(df_filtered1, df_cod_actual, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')

# Rellenar los NaN en 'Cod_Actual_1' con los valores de 'Material' original
FilteredUE['Cod_Actual_1'] = FilteredUE['Cod_Actual_1'].fillna(FilteredUE['Material'])
# Eliminar la columna 'Material' y renombrar 'Cod_Actual_1' a 'Material'
FilteredUE = FilteredUE.drop('Material', axis=1)


# FilteredUE = FilteredUE.drop('Ult. Eslabon', axis=1)

FilteredUE.rename(columns={'Inspecc.de calidad': 'Insp. Calidad', 'Trans./Trasl.': 'Traslado'}, inplace=True)

FilteredUE.rename(columns={'Cod_Actual_1': 'Ult. Eslabon'}, inplace=True)
# Aquí se ajustan los valores de 'Libre utilización' menores a 0 a 0
FilteredUE['Libre utilización'] = FilteredUE['Libre utilización'].apply(lambda x: 0 if x < 0 else x)

# Verifica si hay algún registro después del filtrado
if FilteredUE.empty:
    print("No hay registros después del filtrado. Verifique los criterios de filtrado y los datos de entrada.")
else:
    df_stock4 = FilteredUE.groupby(['Ult. Eslabon']).agg({
        'Libre utilización': 'sum',
        'Insp. Calidad': 'sum',
        'Traslado': 'sum'
    }).reset_index()

    df_stock4['Ult. Eslabon'] = df_stock4['Ult. Eslabon'].astype(str)
    df_stock4['Libre utilización'] = df_stock4['Libre utilización'].fillna(0).astype(int)
    df_stock4['Traslado'] = df_stock4['Traslado'].fillna(0).astype(int)
    df_stock4['Insp. Calidad'] = df_stock4['Insp. Calidad'].fillna(0).astype(int)

    # Agregar una columna 'Control' sumando 'Insp. Calidad' y 'Traslado'
    df_stock4['Control'] = df_stock4['Insp. Calidad'] + df_stock4['Traslado']
    df_stock4.rename(columns={'Ult. Eslabon': 'Material'}, inplace=True)
    FilteredUE = FilteredUE.rename(columns={'Ult. Eslabon': 'Material'})

duplicados = df_stock4['Material'].duplicated()

# Para ver si hay al menos un duplicado en esa columna
hay_duplicados = duplicados.any()
print("¿Hay duplicados?:", hay_duplicados)

# Para contar el número total de duplicados en esa columna
numero_duplicados = duplicados.sum()
print("Número de duplicados:", numero_duplicados)


df_duplicados = df_stock4[df_stock4['Material'].duplicated(keep=False)]
print(df_duplicados)
df_stock4 = df_stock4.drop_duplicates(keep='first')

duplicados = FilteredUE['Material'].duplicated()

# Para ver si hay al menos un duplicado en esa columna
hay_duplicados = duplicados.any()
print("¿Hay duplicados?:", hay_duplicados)

# Para contar el número total de duplicados en esa columna
numero_duplicados = duplicados.sum()
print("Número de duplicados:", numero_duplicados)

# Para obtener un DataFrame con todos los duplicados en esa columna
df_duplicados = FilteredUE[FilteredUE['Material'].duplicated(keep=False)]
print(df_duplicados)

# STOCK 711
# Especifica las columnas que necesitas
cols_to_use4 = ['Material', 'Libre utilización', 'Almacén', 'Centro', 'Inspecc.de calidad', 'Trans./Trasl.', 'Ult. Eslabon']

# Carga solo las columnas que necesitas desde la hoja "Sheet1"
dfleerstock2 = pd.read_excel(ruta_destino_stock2, sheet_name='Sheet1', usecols=cols_to_use4)
dfleerstock2.rename(columns={'Almacén': 'Alm.', 'Centro': 'Ce.', 'Inspecc.de calidad': 'Insp. Calidad', 'Trans./Trasl.': 'Traslado'}, inplace=True)
dfleerstock.rename(columns={'Centro': 'Ce.'}, inplace=True)


# Convierte 'Alm.' y 'Ce.' a cadena para garantizar la consistencia en el filtrado
dfleerstock2['Alm.'] = dfleerstock2['Alm.'].astype(str)
dfleerstock2['Ce.'] = dfleerstock2['Ce.'].astype(str)
dfleerstock['Ce.'] = dfleerstock['Ce.'].astype(str).str.replace('.0', '', regex=False)

# Filtrar por 'Centro'
allowed_centers2 = ['0711', '711', '711.0']
df_filtered2 = dfleerstock2[dfleerstock2['Ce.'].isin(allowed_centers2)]
df_filtered2['Material'] = df_filtered2['Material'].astype(str)
 # Convertir la columna 'Material' a tipo object (string) para evitar el .0
FilteredUE711 = pd.merge(df_filtered2, df_cod_actual, left_on='Material', right_on='Nro_pieza_fabricante_1', how='left')
FilteredUE711['Cod_Actual_1'] = FilteredUE711['Cod_Actual_1'].fillna(FilteredUE711['Material'])

# Eliminar la columna 'Material' y renombrar 'Cod_Actual_1' a 'Material'
FilteredUE711 = FilteredUE711.drop('Material', axis=1)
FilteredUE711 = FilteredUE711.drop('Ult. Eslabon', axis=1)
FilteredUE711.rename(columns={'Cod_Actual_1': 'Ult. Eslabon'}, inplace=True)

# Verifica si hay algún registro después del filtrado
if FilteredUE711.empty:
    print("No hay registros después del filtrado. Verifique los criterios de filtrado y los datos de entrada.")
else:
    # Agrupar por 'Material' y sumar las columnas especificadas
    FilteredUE711 = FilteredUE711.groupby('Ult. Eslabon').agg({
        'Libre utilización': 'sum',
        'Insp. Calidad': 'sum',
        'Traslado': 'sum'
    }).reset_index()

    # Convertir la columna 'Material' a tipo object (string) para evitar el .0
    FilteredUE711['Ult. Eslabon'] = FilteredUE711['Ult. Eslabon'].astype(str)

    # Añadir la columna 'Control' que es la suma de 'Insp. Calidad' y 'Traslado'
    FilteredUE711['Control'] = FilteredUE711['Insp. Calidad'] + FilteredUE711['Traslado']

    # Asumiendo que deseas rellenar los NaN con 0 y convertir a entero
    FilteredUE711['Libre utilización'] = FilteredUE711['Libre utilización'].fillna(0).astype(int)
    FilteredUE711['Insp. Calidad'] = FilteredUE711['Insp. Calidad'].fillna(0).astype(int)
    FilteredUE711['Traslado'] = FilteredUE711['Traslado'].fillna(0).astype(int)
    FilteredUE711['Control'] = FilteredUE711['Control'].fillna(0).astype(int)
    FilteredUE711.rename(columns={'Ult. Eslabon': 'Material'}, inplace=True)

    FilteredUE711.rename(columns={'Libre utilización': 'Stock 711'}, inplace=True)
    #MERGE STOCK TIENDA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#MERGE STOCK TIENDA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

df_stock['Material'] = df_stock['Material'].astype(str)
merged_df2 = pd.merge(merged_df2, df_stock[['Material', 'Stock Tiendas']], on='Material', how='left')
merged_df2 = pd.merge(merged_df2, df_stock4[['Material', 'Libre utilización', 'Control']], on='Material', how='left')

# Llenar NaN con 0 en la columna 'Libre utilización'
merged_df2['Libre utilización'] = merged_df2['Libre utilización'].fillna(0)
# Convertir 'Material' y 'Stock Tiendas' a enteros, asegurándose de que no haya NaNs
merged_df2.rename(columns={'Libre utilización': 'Stock CD'}, inplace=True)
merged_df2['Stock Tiendas'] = merged_df2['Stock Tiendas'].fillna(0)

# Hacer un merge 
merged_df2 = pd.merge(merged_df2, FilteredUE711[['Material', 'Stock 711']], on='Material', how='left')

merged_df2['Stock 711'] = merged_df2['Stock 711'].fillna(0)
merged_df2['Control'] = merged_df2['Control'].fillna(0)
merged_df2.head()


# MANIPULAR FC ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

import pandas as pd
from datetime import datetime

# Obtener los últimos dos dígitos del año actual y del siguiente año
año_actual = datetime.now().year
ultimo_digito_año_actual = str(año_actual)[-2:]
ultimo_digito_año_siguiente = str(año_actual + 1)[-2:]

# Leer el archivo Excel
xlinbound = pd.ExcelFile(ruta_completa_archivo)

# Buscar nombres de hojas que contengan la palabra 'Inbound' (en cualquier combinación de mayúsculas y minúsculas)
hojas_inbound = [hoja for hoja in xlinbound.sheet_names if 'inbound' in hoja.lower()]

# Leer la hoja de Excel que contiene 'Inbound' en el nombre, empezando por la segunda fila
if hojas_inbound:
    # El parámetro 'header=1' indica que la segunda fila contiene los nombres de las columnas
    inbound_dataframe = pd.read_excel(ruta_completa_archivo, sheet_name=hojas_inbound[0], header=1)
    
    # Filtrar columnas que contengan los últimos dos dígitos del año actual o del siguiente en el nombre
    columnas_filtradas = [
    col for col in inbound_dataframe.columns
    if ultimo_digito_año_actual in str(col) or ultimo_digito_año_siguiente in str(col)
]
    inbound_dataframe_filtrado = inbound_dataframe[columnas_filtradas + ['Ult. Eslabón']]
    
    # Convertir las columnas pertinentes a strings para el merge
    merged_df2['Material'] = merged_df2['Material'].astype(str)
    inbound_dataframe_filtrado = inbound_dataframe_filtrado.copy()
    inbound_dataframe_filtrado['Ult. Eslabón'] = inbound_dataframe_filtrado['Ult. Eslabón'].astype(str)


import pandas as pd
from datetime import datetime
month_to_number = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sept': 9, 'oct': 10, 'nov': 11, 'dic': 12
}


from collections import Counter
from datetime import datetime, timedelta

# Crear un nuevo DataFrame para los valores semanales
weekly_values_df = pd.DataFrame()
inbound_dataframe_filtrado['Ult. Eslabón'] = inbound_dataframe_filtrado['Ult. Eslabón'].astype(str)
inbound_dataframe_filtrado['Ult. Eslabón'] = inbound_dataframe_filtrado['Ult. Eslabón'].astype(str)
inbound_dataframe_filtrado.head()
month_to_number = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sept': 9, 'oct': 10, 'nov': 11, 'dic': 12
}


from collections import Counter
from datetime import datetime, timedelta

# Crear un nuevo DataFrame para los valores semanales
weekly_values_df = pd.DataFrame()

for column in inbound_dataframe_filtrado.columns:
    if column == 'Ult. Eslabón':
        continue

    mes, año = column.split('-')
    mes_numero = month_to_number[mes.lower()]
    año_numero = int('20' + año)

    valor_mensual = inbound_dataframe_filtrado[column].astype(float)

    # Dividir el valor mensual por el número aproximado de semanas
    valor_semanal = valor_mensual / 4.33

    # Determinar el rango de fechas de la semana para el mes
    primera_fecha = datetime(año_numero, mes_numero, 1)
    ultima_fecha = primera_fecha + pd.offsets.MonthEnd()

    # Contar los días que pertenecen a cada mes
    dias_en_semana = Counter()
    fecha_actual = primera_fecha
    while fecha_actual <= ultima_fecha:
        semana_del_año = fecha_actual.isocalendar()[1]
        dias_en_semana[semana_del_año] += 1
        fecha_actual += timedelta(days=1)

    # Asignar semanas a los meses correspondientes
    for semana, dias in dias_en_semana.items():
        # Determinar si la semana se debe asignar a este mes
        if dias >= 4:  # Si al menos 4 días de la semana caen en este mes, asignar la semana a este mes
            nombre_columna_semana = f'Sem {semana} ({año})'
            weekly_values_df[nombre_columna_semana] = valor_semanal

# Agregar la columna 'Ult. Eslabón' al nuevo DataFrame, asegurándote de que el índice coincida
weekly_values_df['Ult. Eslabón'] = inbound_dataframe_filtrado['Ult. Eslabón'].values
weekly_values_df.rename(columns={'Ult. Eslabón': 'Material'}, inplace=True)
merged_df2.fillna(0, inplace=True)
weekly_values_df
#TRANSITO ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

import pandas as pd
from datetime import datetime, timedelta

# Leer la hoja que se llama 'TR Final' en un DataFrame.
try:
    # Intenta cargar el archivo con el nombre de hoja 'Hoja1'
    df_transito_original = pd.read_excel(ruta_destino_stock3, sheet_name='Sheet1', usecols='A:Q')
except Exception as e:
    # Si falla, imprime un mensaje de advertencia y carga el archivo con el nombre de hoja 'Sheet1'
    print(f"No se pudo cargar el archivo con el nombre de hoja 'Hoja1'. Error: {e}")
    try:
        df_transito_original = pd.read_excel(ruta_destino_stock3, sheet_name='Sheet1')
    except Exception as e:
        # Si tampoco puede cargarlo con el nombre de hoja 'Sheet1', imprime un mensaje de error
        print(f"No se pudo cargar el archivo con el nombre de hoja 'Sheet1'. Error: {e}")


# Normalizar los nombres de las columnas (en caso de que haya inconsistencias en mayúsculas y minúsculas).
# Normalizar los nombres de las columnas (en caso de que haya inconsistencias en mayúsculas y minúsculas).
df_transito_original.columns = [col.title() if isinstance(col, str) else col for col in df_transito_original.columns]
df_transito_original['Fecha'] = pd.to_datetime(df_transito_original['Fecha'])


# Verificar que las columnas existan
columnas_requeridas = ['Material', 'Cantidad', 'Fecha']
for col in columnas_requeridas:
    if col not in df_transito_original.columns:
        raise ValueError(f"La columna requerida '{col}' no está en el DataFrame.")
    
    
# Convertir todas las columnas a formato string (alfanumérico) excepto la columna de fecha
for column in df_transito_original.columns:
    if column != 'Fecha':  # Excluyendo la columna de fecha
        df_transito_original[column] = df_transito_original[column].astype(str)



df_transito_original

df_transito_original['Semana_Año'] = df_transito_original.apply(lambda x: str(x['Semana']) + '-' + str(x['Año']), axis=1)
# Función para obtener la semana correspondiente al mes de la fecha
def asignar_semana_al_mes(fecha):
    # Encuentra el primer día del mes de la fecha dada
    primer_dia_del_mes = fecha.replace(day=1)
    # Encuentra el último día de la semana en que cae el primer día del mes
    ultimo_dia_semana = primer_dia_del_mes + timedelta(days=6 - primer_dia_del_mes.weekday())
    # Si la fecha dada es mayor al último día de la primera semana completa del mes, usar la fecha dada
    # De lo contrario, usa el primer día del mes
    fecha_referencia = fecha if fecha > ultimo_dia_semana else primer_dia_del_mes
    # Encuentra el año y la semana de la fecha de referencia
    año, semana, dia_semana = fecha_referencia.isocalendar()
    return semana, año

# Aplicar la función a la columna 'Fecha' para obtener la nueva semana y año
df_transito_original['Semana_Asignada'], df_transito_original['Año_Asignado'] = zip(*df_transito_original['Fecha'].apply(asignar_semana_al_mes))


# Crear la columna 'Semana_Año_Asignada'
df_transito_original['Semana_Año_Asignada'] = df_transito_original['Semana_Asignada'].astype(str) + '-' + df_transito_original['Año_Asignado'].astype(str)

df_transito_original['Cantidad'] = pd.to_numeric(df_transito_original['Cantidad'], errors='coerce')

df_transito = df_transito_original.pivot_table(
    index='Material',
    columns='Semana_Año_Asignada',
    values='Cantidad',
    aggfunc='sum',  # Sumará los valores numéricos
    fill_value=0  # Llena con ceros si no hay valores
).reset_index()


# Renombrar las columnas para que sigan el formato 'Sem X (YY)'
nuevos_nombres_columnas = {col: f"Sem {int(col.split('-')[0])} ({col.split('-')[1][2:]})" for col in df_transito.columns if '-' in col}
df_transito.rename(columns=nuevos_nombres_columnas, inplace=True)
inbound_dataframe_filtrado.rename(columns={'Ult. Eslabón': 'Material'}, inplace=True)

df_transito['Material'] = df_transito['Material'].astype(str)
df_transito.columns
import pandas as pd

# Asumiendo que df_transito ya está definido y cargado

# # Sumar los valores de la columna 3 y 4, y luego sumarle ese resultado a la columna 5
# # Sumar los valores de la columna 3 y 4, y luego sumarle ese resultado a la columna 5
df_transito.iloc[:, 2] = df_transito.iloc[:, 1] + df_transito.iloc[:, 2]

# Mostrar el DataFrame resultante
print(df_transito)
# Poner los valores de las columnas tercera y cuarta (posiciones 2 y 3) a 0
df_transito.iloc[:, 1] = 0
# df_transito.iloc[:, 2] = 0





# Ahora puedes realizar la conversión a string sin la advertencia
inbound_dataframe_filtrado['Material'] = inbound_dataframe_filtrado['Material'].astype(str)

#forecast anterior!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

import os
import shutil
from datetime import datetime, timedelta
import locale

# Ruta donde se encuentran las carpetas por año y mes
directorio_forecast_modificado = os.path.join(user_home_dir, 'Inchcape', 'Planificación y abastecimiento AFM - Documentos', 'Planificación y Compras AFM', 'Forecast Inbound')

# Mapeo de nombres de meses en inglés a español
meses_en_espanol_modificado = {
    "January": "Enero",
    "February": "Febrero",
    "March": "Marzo",
    "April": "Abril",
    "May": "Mayo",
    "June": "Junio",
    "July": "Julio",
    "August": "Agosto",
    "September": "Septiembre",
    "October": "Octubre",
    "November": "Noviembre",
    "December": "Diciembre"
}

# Establece el entorno local a inglés para obtener el mes en inglés
locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')

# Función para obtener el nombre de la carpeta objetivo
def obtener_nombre_carpeta_modificado(fecha):
    año_modificado = fecha.strftime("%Y")
    mes_inglés_modificado = fecha.strftime("%B")
    mes_español_modificado = meses_en_espanol_modificado[mes_inglés_modificado]
    mes_número_modificado = fecha.strftime("%m")
    return f"{año_modificado}-{mes_número_modificado} {mes_español_modificado}"

# Calcular la fecha del mes anterior
fecha_actual_modificado = datetime.now()
fecha_mes_anterior_modificado = fecha_actual_modificado - timedelta(days=fecha_actual_modificado.day)
nombre_carpeta_mes_anterior_modificado = obtener_nombre_carpeta_modificado(fecha_mes_anterior_modificado)

# Buscar la carpeta del mes anterior
ruta_carpeta_encontrada_modificado = ""
for nombre_carpeta in os.listdir(directorio_forecast_modificado):
    if nombre_carpeta_mes_anterior_modificado in nombre_carpeta:
        ruta_carpeta_encontrada_modificado = os.path.join(directorio_forecast_modificado, nombre_carpeta)
        break

# Si se encontró la carpeta del mes anterior
if ruta_carpeta_encontrada_modificado:
    # Buscar el archivo con la palabra "Inbound" más reciente
    lista_archivos_modificado = [archivo for archivo in os.listdir(ruta_carpeta_encontrada_modificado) if "Inbound" in archivo]
    lista_archivos_modificado.sort(key=lambda archivo: os.path.getmtime(os.path.join(ruta_carpeta_encontrada_modificado, archivo)), reverse=True)
    
    if lista_archivos_modificado:
        archivo_reciente_modificado = lista_archivos_modificado[0]
        ruta_completa_archivo_modificado = os.path.join(ruta_carpeta_encontrada_modificado, archivo_reciente_modificado)
        
        # Copiar el archivo más reciente a la carpeta destino
        carpeta_destino_semana_modificado = os.path.join(carpeta_destino_principal, f'Sem {semana_actual}')
        shutil.copy(ruta_completa_archivo_modificado, carpeta_destino_semana_modificado)
        print(f"El archivo {archivo_reciente_modificado} ha sido copiado a {carpeta_destino_semana_modificado} exitosamente.")
    else:
        print(f"No se encontraron archivos con la palabra 'Inbound' en {ruta_carpeta_encontrada_modificado}.")
else:
    print(f"No se encontró la carpeta para el mes anterior.")


import pandas as pd
from datetime import datetime
from collections import Counter
from datetime import timedelta

# Continuación del código con nombres de variables modificados

# Obtener los últimos dos dígitos del año actual y del siguiente año
año_actual_modificado = datetime.now().year
ultimo_digito_año_actual_modificado = str(año_actual_modificado)[-2:]
ultimo_digito_año_siguiente_modificado = str(año_actual_modificado + 1)[-2:]

# Leer el archivo Excel
xlinbound_modificado = pd.ExcelFile(ruta_completa_archivo_modificado)

# Buscar nombres de hojas que contengan la palabra 'Inbound'
hojas_inbound_modificado = [hoja for hoja in xlinbound_modificado.sheet_names if 'inbound' in hoja.lower()]



# Leer la hoja de Excel que contiene 'Inbound' en el nombre, empezando por la segunda fila
if hojas_inbound_modificado:
    inbound_dataframe_modificado = pd.read_excel(ruta_completa_archivo_modificado, sheet_name=hojas_inbound_modificado[0], header=1)
    
    # Filtrar columnas que contengan los últimos dos dígitos del año actual o del siguiente en el nombre
    columnas_filtradas_modificado = [col for col in inbound_dataframe_modificado.columns if ultimo_digito_año_actual_modificado in col or ultimo_digito_año_siguiente_modificado in col]
    inbound_dataframe_filtrado_modificado = inbound_dataframe_modificado[columnas_filtradas_modificado + ['Ult. Eslabón']]
    # Usar .loc[] para realizar la modificación directamente en el DataFrame
    inbound_dataframe_filtrado_modificado.loc[:, 'Ult. Eslabón'] = inbound_dataframe_filtrado_modificado['Ult. Eslabón'].astype(str)
 

# Mapeo de meses a números
month_to_number_modificado = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'ago': 8, 'sept': 9, 'oct': 10, 'nov': 11, 'dic': 12
}

# Crear un nuevo DataFrame para los valores semanales
weekly_values_df_modificado = pd.DataFrame()

for column in inbound_dataframe_filtrado_modificado.columns:
    if column == 'Ult. Eslabón':
        continue

    mes_modificado, año_modificado = column.split('-')
    mes_numero_modificado = month_to_number_modificado[mes_modificado.lower()]
    año_numero_modificado = int('20' + año_modificado)

    valor_mensual_modificado = inbound_dataframe_filtrado_modificado[column].astype(float)

    # Dividir el valor mensual por el número aproximado de semanas
    valor_semanal_modificado = valor_mensual_modificado / 4.33

    # Determinar el rango de fechas de la semana para el mes
    primera_fecha_modificada = datetime(año_numero_modificado, mes_numero_modificado, 1)
    ultima_fecha_modificada = primera_fecha_modificada + pd.offsets.MonthEnd()

    # Contar los días que pertenecen a cada mes
    dias_en_semana_modificado = Counter()
    fecha_actual_modificada = primera_fecha_modificada
    while fecha_actual_modificada <= ultima_fecha_modificada:
        semana_del_año_modificada = fecha_actual_modificada.isocalendar()[1]
        dias_en_semana_modificado[semana_del_año_modificada] += 1
        fecha_actual_modificada += timedelta(days=1)

    # Asignar semanas a los meses correspondientes
    for semana, dias in dias_en_semana_modificado.items():
        if dias >= 4:  # Si al menos 4 días de la semana caen en este mes, asignar la semana a este mes
            nombre_columna_semana_modificada = f'Sem {semana} ({año_modificado})'
            weekly_values_df_modificado[nombre_columna_semana_modificada] = valor_semanal_modificado

# Agregar la columna 'Ult. Eslabón' al nuevo DataFrame, asegurándote de que el índice coincida
weekly_values_df_modificado['Ult. Eslabón'] = inbound_dataframe_filtrado_modificado['Ult. Eslabón'].values
weekly_values_df_modificado.rename(columns={'Ult. Eslabón': 'Material'}, inplace=True)

ruta_completa_archivo = os.path.join(carpeta_destino_semana, archivo_encontrado)
UE = pd.read_excel(ruta_completa_archivo)
UE.head()
# Lista de columnas a eliminar


# Ahora puedes eliminar las columnas
columnas_a_eliminar = ["Clave PIC", "Nº sec.", "Nro_pieza_fabricante", "Cod_actual", "Vál.desde", "Estado"]
UE = UE.drop(columns=columnas_a_eliminar)

# Realizar el merge
merged_df = pd.merge(weekly_values_df_modificado, UE, 
                     left_on='Material', 
                     right_on='Nro_pieza_fabricante_1',
                     how='left')

# Reemplazar valores NaN en 'Cod_Actual_1' con los valores de 'Material'
merged_df['Cod_Actual_1'] = merged_df['Cod_Actual_1'].fillna(merged_df['Material'])

# Ahora, merged_df tiene 'Cod_Actual_1' con los valores de 'Material' donde no hubo coincidencias

# Ahora puedes eliminar las columnas
columnas_a_eliminar = ["Nro_pieza_fabricante_1", "Material"]
merged_df = merged_df.drop(columns=columnas_a_eliminar)
# Cambiar el nombre de la columna 'Cod_Actual_1' a 'Material'
merged_df = merged_df.rename(columns={'Cod_Actual_1': 'Material'})
merged_df.head()
import calendar
import datetime

# Obtener el mes y año actual
mes_actual23 = datetime.datetime.now().month
año_actual23 = datetime.datetime.now().year

# Obtener el número de semanas en el mes actual
semanas_en_mes23 = calendar.monthcalendar(año_actual23, mes_actual23)

# Determinar si el mes tiene 5 semanas o 4 semanas
if len(semanas_en_mes23) == 5:
    indice_columna = [2, 3, 4]  # Si hay 5 semanas, considera las columnas 2 y 3
else:
    indice_columna = [2, 3]  # Si hay 4 semanas, considera solo la columna 3

# Construir la lista de columnas que comienzan con "Sem" y seleccionar las columnas según el índice determinado
columnas_sem = [col for col in weekly_values_df_modificado.columns if col.startswith("Sem")]
cuarta_columna_sem = [columnas_sem[i] for i in indice_columna]
# # Fusionar los DataFrames
# Convertir las columnas seleccionadas en una lista de cadenas
cuarta_columna_sem_str = [str(col) for col in cuarta_columna_sem]
# _________

# Seleccionar solo estas columnas en weekly_values_df_modificado
columnas_a_conservar = ["Material"] + cuarta_columna_sem_str
weekly_values_df_modificado = weekly_values_df_modificado[columnas_a_conservar]

# Verificar las columnas seleccionadas
print(weekly_values_df_modificado.columns)
#_______________
duplicate_columns = list(set(weekly_values_df.columns).intersection(weekly_values_df_modificado.columns) - {'Material'})

# Realizar el merge usando 'suffixes' para manejar columnas duplicadas
weekly_values_df = pd.merge(weekly_values_df, weekly_values_df_modificado, 
                            on='Material', 
                            how='left', 
                            suffixes=('', '_mod'))

# Para cada columna duplicada, eliminar la del dataframe original y renombrar la del modificado
for column in duplicate_columns:
    weekly_values_df.drop(column, axis=1, inplace=True)  # Elimina la columna original
    weekly_values_df.rename(columns={f'{column}_mod': column}, inplace=True)  # Renombra la columna modificada

# Primero, asegurarte de que las columnas estén en el formato correcto (por si acaso)
weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)

# Luego, puedes usar loc para modificar solo las filas que cumplen con la condición
# (i.e., que tienen un valor no vacío en la columna 'Material')
weekly_values_df.loc[weekly_values_df['Material'] != '', :] = weekly_values_df.loc[weekly_values_df['Material'] != '', :].fillna(0)

# Verifica el resultado
print(weekly_values_df)


#______
import calendar
import datetime

# Obtener el mes y año actual
mes_actual32 = datetime.datetime.now().month
año_actual32 = datetime.datetime.now().year

# Obtener el número de semanas en el mes actual
semanas_en_mes32 = calendar.monthcalendar(año_actual32, mes_actual32)

# Determinar si el mes tiene 5 semanas o 4 semanas
if len(semanas_en_mes32) == 5:
    indice_ultima_semana = [-3,-2,-1]
else:
    indice_ultima_semana = [-2, -1]

#SI FUERA LA SEMANA 4 
# # Determinar si el mes tiene 5 semanas o 4 semanas
# if len(semanas_en_mes32) == 5:
#     indice_ultima_semana = [-2,-1]
# else:
#     indice_ultima_semana = -1
# Obtener la última columna del DataFrame con el índice ajustado
# Obtener la última columna del DataFrame con los índices ajustados
ultimas_columnas = [weekly_values_df.columns[indice] for indice in indice_ultima_semana]

# Crear una nueva lista de nombres de columnas con las últimas columnas al principio
nuevas_columnas = ultimas_columnas + [col for col in weekly_values_df.columns if col not in ultimas_columnas]
ultimas_columnas


# Reordenar las columnas del DataFrame
weekly_values_df = weekly_values_df[nuevas_columnas]

weekly_values_df.columns

import datetime
import numpy as np
import pandas as pd

import datetime
# Obtener la fecha actual
fecha_actual = datetime.datetime.now()
# Establecer la fecha de referencia en el día 15 del mes actual
fecha_referencia = fecha_actual.replace(day=17)

# Encontrar el número de semana del año para el día 15 del mes
numero_semana_quincena = fecha_referencia.isocalendar()[1]

nombre_columna_semana = f"Sem {numero_semana_quincena} ({fecha_actual.strftime('%y')})"

merged_df2[nombre_columna_semana] = np.nan  # Agregar la nueva columna con el nombre de la semana calculada
# todo numerico
merged_df2['Stock Tiendas'] = pd.to_numeric(merged_df2['Stock Tiendas'], errors='coerce').fillna(0)
merged_df2['Stock CD'] = pd.to_numeric(merged_df2['Stock CD'], errors='coerce').fillna(0)
merged_df2['Faltante AP'] = pd.to_numeric(merged_df2['Faltante AP'], errors='coerce').fillna(0)
merged_df2['Faltantes'] = pd.to_numeric(merged_df2['Faltantes'], errors='coerce').fillna(0)

# Realizar el cálculo y asignarlo a la nueva columna
merged_df2[nombre_columna_semana] = (merged_df2['Stock Tiendas'] + merged_df2['Stock CD']) - (merged_df2['Faltantes'])
print(nombre_columna_semana)

merged_df2['Material'] = merged_df2['Material'].astype(str)
df_transito['Material'] = df_transito['Material'].astype(str)
weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)

semana_actual =numero_semana_quincena
semana_actual

merged_df2.columns
# fecha_referencia = datetime.now().replace(day=1)
# numero_semana_quincena = fecha_referencia.isocalendar()[1]

# Calcular la semana siguiente y ajustar el año si es necesario
semana_siguiente = semana_actual + 1
año_actual = fecha_referencia.year

if semana_siguiente > 52:
    semana_siguiente = 1
    año_actual += 1

# Crear el nombre de la columna para la semana siguiente
nombre_columna_semana_siguiente = f"Sem {semana_siguiente} ({str(año_actual)[-2:]})"

# Inicializar la columna de la semana siguiente en merged_df4
merged_df2[nombre_columna_semana_siguiente] = merged_df2[nombre_columna_semana]

if nombre_columna_semana in merged_df2.columns:
    merged_df2[nombre_columna_semana_siguiente] = pd.to_numeric(merged_df2[nombre_columna_semana], errors='coerce').fillna(0)

# Asegurarse de que la columna 'Control' es numérica
merged_df2['Control'] = pd.to_numeric(merged_df2['Control'], errors='coerce').fillna(0)

# Sumar los valores de la columna 'Control' a la nueva columna de la semana siguiente
merged_df2[nombre_columna_semana_siguiente] += merged_df2['Control']

merged_df2['Material'] = merged_df2['Material'].astype(str)
df_transito['Material'] = df_transito['Material'].astype(str)
weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)
merged_df2.columns
merged_df2 = merged_df2.merge(df_transito[['Material', nombre_columna_semana_siguiente]], on='Material', how='left', suffixes=('', '_transito'))
tipo_de_dato = merged_df2[nombre_columna_semana_siguiente].dtype

print(f"El tipo de dato de la columna 'nombre_columna_semana_siguiente' es: {tipo_de_dato}")
merged_df2[nombre_columna_semana_siguiente] = merged_df2[nombre_columna_semana_siguiente] + merged_df2[nombre_columna_semana_siguiente + '_transito'].fillna(0)
# Combinación (merge) con weekly_values_df
merged_df2 = merged_df2.merge(weekly_values_df[['Material', nombre_columna_semana_siguiente]], on='Material', how='left', suffixes=('', '_weekly'))
# Condición para realizar la resta solo si el valor después de sumar con transito es mayor a 0
merged_df2[nombre_columna_semana_siguiente] = merged_df2.apply(
    lambda row: row[nombre_columna_semana_siguiente] - row[nombre_columna_semana_siguiente + '_weekly']
    if row[nombre_columna_semana_siguiente] > 0 else row[nombre_columna_semana_siguiente], axis=1
).fillna(0)

# Eliminar las columnas auxiliares del merge
merged_df2.drop(columns=[nombre_columna_semana_siguiente + '_transito', nombre_columna_semana_siguiente + '_weekly'], inplace=True)

# Imprimir el DataFrame para verificar los resultados
print(nombre_columna_semana_siguiente)


#TERCERA SEMANA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# Convierte todo a strng
merged_df2['Material'] = merged_df2['Material'].astype(str)
df_transito['Material'] = df_transito['Material'].astype(str)
weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)

semana_actual = int(nombre_columna_semana_siguiente.split(' ')[1].split('(')[0])
año_actual = int(nombre_columna_semana_siguiente.split('(')[1].split(')')[0])
semana_nueva = semana_actual + 1
año_nuevo = año_actual
if semana_nueva > 52:
    semana_nueva = 1
    año_nuevo += 1
nombre_columna_semana_nueva = f"Sem {semana_nueva} ({str(año_nuevo)[-2:]})"

# Inicializa la columna de la nueva semana en merged_df4.
merged_df2[nombre_columna_semana_nueva] = merged_df2[nombre_columna_semana_siguiente]

# Asegúrate de que la columna 'Stock 711' es numérica.
merged_df2['Stock 711'] = pd.to_numeric(merged_df2['Stock 711'], errors='coerce').fillna(0)

# Suma los valores de la columna 'Stock 711' a la nueva columna de la semana.
merged_df2[nombre_columna_semana_nueva] += merged_df2['Stock 711']

# Realiza el merge con 'df_transito' y suma los valores correspondientes a la nueva semana.
merged_df2 = merged_df2.merge(
    df_transito[['Material', nombre_columna_semana_nueva]],
    on='Material',
    how='left',
    suffixes=('', '_transito_nueva')
)
merged_df2[nombre_columna_semana_nueva] += merged_df2[nombre_columna_semana_nueva + '_transito_nueva'].fillna(0)

# Merge con el forecast
merged_df2 = merged_df2.merge(
    weekly_values_df[['Material', nombre_columna_semana_nueva]],
    on='Material',
    how='left',
    suffixes=('', '_forecast_nueva')
)

## Restar solo si el resultado del merge con 'df_transito' es igual o mayor a 0
merged_df2[nombre_columna_semana_nueva] = merged_df2.apply(
    lambda row: row[nombre_columna_semana_nueva] - row[nombre_columna_semana_nueva + '_forecast_nueva']
    if row[nombre_columna_semana_nueva] > 0 else row[nombre_columna_semana_nueva], axis=1
).fillna(0)

# Eliminar las columnas auxiliares del merge
merged_df2.drop(columns=[nombre_columna_semana_nueva + '_transito_nueva', nombre_columna_semana_nueva + '_forecast_nueva'], inplace=True)
print(nombre_columna_semana_nueva)

#sers\Etorres\Desktop\FC6.xlsx', index=False)
#46 SEMANAS ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


semana_actual = int(nombre_columna_semana_siguiente.split(' ')[1].split('(')[0])
año_actual = int(nombre_columna_semana_siguiente.split('(')[1].split(')')[0])
semana_siguiente = semana_actual + 1
año_siguiente = año_actual
if semana_siguiente > 52:
    semana_siguiente = 1
    año_siguiente += 1
nombre_columna_semana_siguiente = f"Sem {semana_siguiente} ({str(año_siguiente)[-2:]})"

for i in range(43):  # Repetir el proceso 46 veces
    # Calcula el nombre de la columna de la nueva semana.
    semana_actual = int(nombre_columna_semana_siguiente.split(' ')[1].split('(')[0])
    año_actual = int(nombre_columna_semana_siguiente.split('(')[1].split(')')[0])
    semana_nueva = semana_actual + 1
    año_nuevo = año_actual
    if semana_nueva > 52:
        semana_nueva = 1
        año_nuevo += 1
    nombre_columna_semana_nueva = f"Sem {semana_nueva} ({str(año_nuevo)[-2:]})"

    # Inicializa la columna de la nueva semana en merged_df4.
    merged_df2[nombre_columna_semana_nueva] = merged_df2[nombre_columna_semana_siguiente]

    # Realiza el merge con 'df_transito' y suma los valores correspondientes a la nueva semana.
    if nombre_columna_semana_nueva in df_transito.columns:
        merged_df2 = merged_df2.merge(
            df_transito[['Material', nombre_columna_semana_nueva]],
            on='Material',
            how='left',
            suffixes=('', '_transito_nueva')
        )
        merged_df2[nombre_columna_semana_nueva] += merged_df2[nombre_columna_semana_nueva + '_transito_nueva'].fillna(0)
        # Elimina la columna auxiliar del merge
        merged_df2.drop(columns=[nombre_columna_semana_nueva + '_transito_nueva'], inplace=True)

        # Realiza el merge con 'weekly_values_df' y resta el forecast correspondiente a la nueva semana solo si el total es mayor a 0
        if nombre_columna_semana_nueva in weekly_values_df.columns:
            merged_df2 = merged_df2.merge(
                weekly_values_df[['Material', nombre_columna_semana_nueva]],
                on='Material',
                how='left',
                suffixes=('', '_forecast_nueva')
            )
            merged_df2[nombre_columna_semana_nueva] = merged_df2.apply(
                lambda row: max(row[nombre_columna_semana_nueva] - row[nombre_columna_semana_nueva + '_forecast_nueva'], 0)
                if row[nombre_columna_semana_nueva] > 0 else row[nombre_columna_semana_nueva], axis=1
            ).fillna(0)
            merged_df2.drop(columns=[nombre_columna_semana_nueva + '_forecast_nueva'], inplace=True)

    # Actualiza el nombre de la columna para la próxima iteración
    nombre_columna_semana_siguiente = nombre_columna_semana_nueva

#PORCENTAJES DE DISPONIBILIDAD

columnas_semanas = [col for col in merged_df2.columns if col.startswith("Sem ") and "D" not in col]

# Realizar un merge (fusión) para alinear las filas según el 'Material'
merged_df2 = merged_df2.merge(weekly_values_df[['Material'] + columnas_semanas], on='Material', how='left', suffixes=('', '_weekly'))

# Diccionario para almacenar las nuevas columnas
nuevas_columnas = {}

# Iterar sobre las columnas de semanas originales
for col in columnas_semanas:
    nueva_col = col + "D"  # Crear el nombre de la nueva columna añadiendo "D" al final

    # Realizar el cálculo basado en las condiciones
    condiciones = [
        merged_df2[col] < 0,
        (merged_df2[col] == 0) & (merged_df2[col + '_weekly'] > 0),
        (merged_df2[col] == 0) & (merged_df2[col + '_weekly'] == 0),
        (merged_df2[col] > 0) & (merged_df2[col + '_weekly'] == 0),
        (merged_df2[col] > 0) & (merged_df2[col + '_weekly'] > 0),
        (merged_df2[col] >= 0) & (merged_df2[col + '_weekly'].isnull() | (merged_df2[col + '_weekly'] == 0))
    ]

    # Resultados basados en las condiciones
    resultados = [
        0,  # Semana < 0
        0,  # Semana = 0 y weekly > 0
        100,  # Semana = 0 y weekly = 0
        100,  # Semana > 0 y weekly = 0
        merged_df2[col] / merged_df2[col + '_weekly'] * 100,  # Semana > 0 y weekly > 0
        100  # Semana >= 0 y no hay weekly o weekly = 0
    ]

    # Aplicar las condiciones y almacenar en el diccionario
    columna_temporal = pd.Series(0, index=merged_df2.index).astype(float)  # Inicializar con 0
    for cond, res in zip(condiciones, resultados):
        columna_temporal[cond] = res.clip(upper=100) if not isinstance(res, int) else res
    
    nuevas_columnas[nueva_col] = columna_temporal

# Concatenar todas las nuevas columnas a merged_df2
merged_df2 = pd.concat([merged_df2, pd.DataFrame(nuevas_columnas)], axis=1)
# Eliminar las columnas que terminan en "_weekly"
columnas_weekly = [col for col in merged_df2.columns if col.endswith('_weekly')]
merged_df2.drop(columns=columnas_weekly, inplace=True)
merged_df2.columns
desktop_dir = os.path.join(os.path.expanduser('~'), 'OneDrive - Inchcape', 'Escritorio')
merged_df2.to_excel(os.path.join(desktop_dir, 'DispoFutura.xlsx'), index=False)
df_transito.to_excel(os.path.join(desktop_dir, 'TR_Semanal.xlsx'), index=False)
weekly_values_df.to_excel(os.path.join(desktop_dir, 'FC_Semanal.xlsx'), index=False)
weekly_values_df_modificado.to_excel(os.path.join(desktop_dir, 'FC_Semanal2.xlsx'), index=False)
df_stock.to_excel(os.path.join(desktop_dir, 'Stock_Tiendas.xlsx'), index=False)
df_stock4.to_excel(os.path.join(desktop_dir, 'Stock_CD.xlsx'), index=False)
FilteredUE711.to_excel(os.path.join(desktop_dir, 'Stock_711.xlsx'), index=False)