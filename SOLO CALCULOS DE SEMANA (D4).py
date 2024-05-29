




    


#SEMANA 1 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# Obtener la fecha actual
fecha_actual = datetime.now()

# Establecer la fecha de referencia en el día 15 del mes actual
fecha_referencia = fecha_actual.replace(day=22)

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
#SEMANA 2 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# todo string
merged_df2['Material'] = merged_df2['Material'].astype(str)
df_transito['Material'] = df_transito['Material'].astype(str)
weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)

# Obtener la fecha actual y establecerla en el día 15 del mes actual
fecha_referencia = datetime.now().replace(day=22)

# Calcular el número de semana para el día 15
numero_semana_quincena = fecha_referencia.isocalendar()[1]

# Calcular la semana siguiente y ajustar el año si es necesario
semana_siguiente = numero_semana_quincena + 1
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

merged_df2 = merged_df2.merge(df_transito[['Material', nombre_columna_semana_siguiente]], on='Material', how='left', suffixes=('', '_transito'))
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
# #SEMANA 1 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# # Obtener la fecha actual
# fecha_actual = datetime.now()

# # Establecer la fecha de referencia en el día 15 del mes actual
# fecha_referencia = fecha_actual.replace(day=15)

# # Encontrar el número de semana del año para el día 15 del mes
# numero_semana_quincena = fecha_referencia.isocalendar()[1]

# nombre_columna_semana = f"Sem {numero_semana_quincena} ({fecha_actual.strftime('%y')})"

# merged_df2[nombre_columna_semana] = np.nan  # Agregar la nueva columna con el nombre de la semana calculada
# # todo numerico
# merged_df2['Stock Tiendas'] = pd.to_numeric(merged_df2['Stock Tiendas'], errors='coerce').fillna(0)
# merged_df2['Stock CD'] = pd.to_numeric(merged_df2['Stock CD'], errors='coerce').fillna(0)
# merged_df2['Faltante AP'] = pd.to_numeric(merged_df2['Faltante AP'], errors='coerce').fillna(0)
# merged_df2['Faltantes'] = pd.to_numeric(merged_df2['Faltantes'], errors='coerce').fillna(0)

# # Realizar el cálculo y asignarlo a la nueva columna
# merged_df2[nombre_columna_semana] = (merged_df2['Stock Tiendas'] + merged_df2['Stock CD']) - (merged_df2['Faltantes'])
# print(nombre_columna_semana)
# #SEMANA 2 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# # todo string
# merged_df2['Material'] = merged_df2['Material'].astype(str)
# df_transito['Material'] = df_transito['Material'].astype(str)
# weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)
# weekly_values_df_modificado['Material'] = weekly_values_df_modificado['Material'].astype(str)

# # Obtener la fecha actual y establecerla en el día 15 del mes actual
# fecha_referencia = datetime.now().replace(day=15)

# # Calcular el número de semana para el día 15
# numero_semana_quincena = fecha_referencia.isocalendar()[1]

# # Calcular la semana siguiente y ajustar el año si es necesario
# semana_siguiente = numero_semana_quincena + 1
# año_actual = fecha_referencia.year

# if semana_siguiente > 52:
#     semana_siguiente = 1
#     año_actual += 1

# # Crear el nombre de la columna para la semana siguiente
# nombre_columna_semana_siguiente = f"Sem {semana_siguiente} ({str(año_actual)[-2:]})"

# # Inicializar la columna de la semana siguiente en merged_df4
# merged_df2[nombre_columna_semana_siguiente] = merged_df2[nombre_columna_semana]

# if nombre_columna_semana in merged_df2.columns:
#     merged_df2[nombre_columna_semana_siguiente] = pd.to_numeric(merged_df2[nombre_columna_semana], errors='coerce').fillna(0)

# # Asegurarse de que la columna 'Control' es numérica
# merged_df2['Control'] = pd.to_numeric(merged_df2['Control'], errors='coerce').fillna(0)

# # Sumar los valores de la columna 'Control' a la nueva columna de la semana siguiente
# merged_df2[nombre_columna_semana_siguiente] += merged_df2['Control']

# merged_df2['Material'] = merged_df2['Material'].astype(str)
# df_transito['Material'] = df_transito['Material'].astype(str)
# weekly_values_df['Material'] = weekly_values_df['Material'].astype(str)
# weekly_values_df_modificado['Material'] = weekly_values_df_modificado['Material'].astype(str)

# merged_df2 = merged_df2.merge(df_transito[['Material', nombre_columna_semana_siguiente]], on='Material', how='left', suffixes=('', '_transito'))
# merged_df2[nombre_columna_semana_siguiente] = merged_df2[nombre_columna_semana_siguiente] + merged_df2[nombre_columna_semana_siguiente + '_transito'].fillna(0)

# # Combinación (merge) con weekly_values_df
# merged_df2 = merged_df2.merge(weekly_values_df_modificado[['Material', nombre_columna_semana_siguiente]], on='Material', how='left', suffixes=('', '_weekly'))

# # Condición para realizar la resta solo si el valor después de sumar con transito es mayor a 0
# merged_df2[nombre_columna_semana_siguiente] = merged_df2.apply(
#     lambda row: row[nombre_columna_semana_siguiente] - row[nombre_columna_semana_siguiente + '_weekly']
#     if row[nombre_columna_semana_siguiente] > 0 else row[nombre_columna_semana_siguiente], axis=1
# ).fillna(0)

# # Eliminar las columnas auxiliares del merge
# merged_df2.drop(columns=[nombre_columna_semana_siguiente + '_transito', nombre_columna_semana_siguiente + '_weekly'], inplace=True)

# # Imprimir el DataFrame para verificar los resultados
# print(nombre_columna_semana_siguiente)
# print(weekly_values_df_modificado)

# #TERCERA SEMANA ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# # Convierte todo a strng
# merged_df2['Material'] = merged_df2['Material'].astype(str)
# df_transito['Material'] = df_transito['Material'].astype(str)
# weekly_values_df_modificado['Material'] = weekly_values_df_modificado['Material'].astype(str)

# semana_actual = int(nombre_columna_semana_siguiente.split(' ')[1].split('(')[0])
# año_actual = int(nombre_columna_semana_siguiente.split('(')[1].split(')')[0])
# semana_nueva = semana_actual + 1
# año_nuevo = año_actual
# if semana_nueva > 52:
#     semana_nueva = 1
#     año_nuevo += 1
# nombre_columna_semana_nueva = f"Sem {semana_nueva} ({str(año_nuevo)[-2:]})"

# # Inicializa la columna de la nueva semana en merged_df4.
# merged_df2[nombre_columna_semana_nueva] = merged_df2[nombre_columna_semana_siguiente]

# # Asegúrate de que la columna 'Stock 711' es numérica.
# merged_df2['Stock 711'] = pd.to_numeric(merged_df2['Stock 711'], errors='coerce').fillna(0)

# # Suma los valores de la columna 'Stock 711' a la nueva columna de la semana.
# merged_df2[nombre_columna_semana_nueva] += merged_df2['Stock 711']

# # Realiza el merge con 'df_transito' y suma los valores correspondientes a la nueva semana.
# merged_df2 = merged_df2.merge(
#     df_transito[['Material', nombre_columna_semana_nueva]],
#     on='Material',
#     how='left',
#     suffixes=('', '_transito_nueva')
# )
# merged_df2[nombre_columna_semana_nueva] += merged_df2[nombre_columna_semana_nueva + '_transito_nueva'].fillna(0)

# # Merge con el forecast
# merged_df2 = merged_df2.merge(
#     weekly_values_df_modificado[['Material', nombre_columna_semana_nueva]],
#     on='Material',
#     how='left',
#     suffixes=('', '_forecast_nueva')
# )

# ## Restar solo si el resultado del merge con 'df_transito' es igual o mayor a 0
# merged_df2[nombre_columna_semana_nueva] = merged_df2.apply(
#     lambda row: row[nombre_columna_semana_nueva] - row[nombre_columna_semana_nueva + '_forecast_nueva']
#     if row[nombre_columna_semana_nueva] > 0 else row[nombre_columna_semana_nueva], axis=1
# ).fillna(0)

# # Eliminar las columnas auxiliares del merge
# merged_df2.drop(columns=[nombre_columna_semana_nueva + '_transito_nueva', nombre_columna_semana_nueva + '_forecast_nueva'], inplace=True)
# print(nombre_columna_semana_nueva)
#46 SEMANAS ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


semana_actual = int(nombre_columna_semana_siguiente.split(' ')[1].split('(')[0])
año_actual = int(nombre_columna_semana_siguiente.split('(')[1].split(')')[0])
semana_siguiente = semana_actual + 1
año_siguiente = año_actual
if semana_siguiente > 52:
    semana_siguiente = 1
    año_siguiente += 1
nombre_columna_semana_siguiente = f"Sem {semana_siguiente} ({str(año_siguiente)[-2:]})"

for i in range(46):  # Repetir el proceso 46 veces
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


# with pd.ExcelWriter(r'C:\Users\Etorres\Desktop\merged.xlsx', mode='a', engine='openpyxl') as writer:
#     tabla_dinamica.to_excel(writer, sheet_name='Nueva_Hoja')

# merged_df2.to_excel(r'C:\Users\Etorres\Desktop\merged7.xlsx', index=False)

#df_transito.to_excel(r'C:\Users\Etorres\Desktop\tRANSITo6.xlsx', index=False)
#weekly_values_df.to_excel(r'C:\Users\Etorres\Desktop\FC6.xlsx', index=False)

desktop_dir = os.path.join(os.path.expanduser('~'), 'OneDrive - Inchcape', 'Escritorio')
merged_df2.to_excel(os.path.join(desktop_dir, 'DispoFutura.xlsx'), index=False)
df_transito.to_excel(os.path.join(desktop_dir, 'TR_Semanal.xlsx'), index=False)
weekly_values_df.to_excel(os.path.join(desktop_dir, 'FC_Semanal.xlsx'), index=False)
weekly_values_df_modificado.to_excel(os.path.join(desktop_dir, 'FC_Semanal2.xlsx'), index=False)
df_stock.to_excel(os.path.join(desktop_dir, 'Stock_Tiendas.xlsx'), index=False)
df_stock4.to_excel(os.path.join(desktop_dir, 'Stock_CD.xlsx'), index=False)
FilteredUE711.to_excel(os.path.join(desktop_dir, 'Stock_711.xlsx'), index=False)