#--------------------------------------------------------------------------------------------------------------------------
# DASHBOARD DEL UPVIC

# AUTOR: IMAD RIFAI

# OCTUBRE-NOVIEMBRE 2024

#--------------------------------------------------------------------------------------------------------------------------

# OBJETIVO: Crear un excel con varias hojas (tablas) fácil de manipular y estructurado para su uso en POWER BI a partir del report que devuelve IBKR. 
# Hacer código sencillo,comentado y con funciones para que se pueda modificar en un futuro.
# Se saca toda la información posible del IBKR, aunque no se vaya a usar toda en el POWER BI, por si alguien quiere darle un uso diferente.

#--------------------------------------------------------------------------------------------------------------------------

# INSTRUCCIONES: Poner nombre del report (archivo excel devuelto por IBKR), el report tiene que encontrarse en la misma carpeta que este ejecutable, 
# luego ejecutar y usar el nuevo excel devuelto.

#--------------------------------------------------------------------------------------------------------------------------

# RESUMEN DE LOS DATOS EXTRAÍDOS
#
# Estadísticas generales: resumen de los retornos del portfolio

# Posiciones Abiertas: resumen de las posiciones abiertas con otros atributos, ver abajo.

# Historical Performance Benchmark Comparison: : comparar el rendimiento de diferentes índices bursátiles y nuestra cartera

# Concentracion: De esta parte del report_16 (antes del preprocessing se saca las siguientes tablas):

    #Holdings (Posiciones): Detalle de las posiciones actuales en el portafolio, incluyendo el valor nominal de cada posición y su peso porcentual.
    #Asset Class Allocation (Asignación por Clase de Activo): Distribución del portafolio entre las principales clases de activos (acciones, efectivo, bonos).
    #Sector Allocation (Asignación por Sector): Cómo se distribuyen las inversiones entre diferentes sectores económicos.
    #Region Allocation (Asignación por Región): Distribución geográfica de las inversiones por regiones clave (Norteamérica, Europa, Asia).
    #Country Allocation (Asignación por País): Desglose de la inversión por país de origen.
    #Financial Instrument Allocation (Asignación por Instrumento Financiero): Descripción de la distribución del portafolio entre acciones, efectivo y bonos.
    #Exposure (Exposición): Nivel de exposición del portafolio, en términos de posiciones largas (optimismo) y cortas (pesimismo), y el porcentaje neto de la exposición.


# Allocation by asset: Contiene información sobre la asignación de valores en diferentes clases de activos (equities/renta variable, fixed income/renta fija, y cash/efectivo) en cada fecha.

# Allocation by Financial Instrulent: Similar al anterior, pero esta vez la información está dividida por instrumento financiero (bonos, acciones, efectivo) en lugar de clases de activos.

# Time Period Benchmark Comparison: Tabla de comparación de benchmarks a lo largo del tiempo, donde se evalúan varias métricas de rendimiento (retornos) de distintos índices de referencia con nuestra cartera.

# Cumulative Performance Statistics: Muestra el rendimiento acumulativo de nuestra cartera a lo largo del tiempo.

# Risk Measures: Describe la métrica de riesgo específica que se está evaluando. Cada columna es una métrica

# Performance by Asset: Rendimiento por activo

# Perfomance by Symbol: Rendimiento por símbolo

# Fixed Income Summary: Resumen de los ingresos de renta fija 

# Trade summary: Resumen de todas las operaciones comerciales


#--------------------------------------------------------------------------------------------------------------------------

#Las librerías que se usan para llevar a cabo este proyecto
import pandas as pd
from typing import Optional

# Lectura del excel bajado de IBKR
nombre_report = 'report_16may'

df = pd.read_excel(f"{nombre_report}.xlsx",header = None)
df.set_index(df.columns[0],inplace = True)




#Funciones python con el objetivo de facilitar la legibilidad del código, y evitar la repetición de código

def separar_en_dfs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Divide el DataFrame en varios subconjuntos basados en la aparición de 'Header' en la columna 1.
    Cada vez que encuentra 'Header', comienza un nuevo DataFrame.
    """
    dfs = []  # Lista para almacenar los dfs
    current_df = []  # Lista temporal para almacenar las filas del df en el que se este iterando

    for index, row in df.iterrows():
        if row[1] == 'Header':  # Si la columna 1 contiene 'Header'
            if current_df:  # Si hay filas acumuladas en el df actual, guardar el DataFrame actual
                dfs.append(pd.DataFrame(current_df)) 
                current_df = []  

        current_df.append(row)  # Añadir la fila actual al df temporal

    #Añadimos el ult df
    if current_df:
        dfs.append(pd.DataFrame(current_df))

    return dfs
        

def punto_partida(df : pd.DataFrame, nombre : Optional[str] = None) -> pd.DataFrame:
    """
    Esta acción se repite siempre que buscamos un nuevo dataset en el report, lo que hace es localizar el df buscado, eliminar la primera fila y columna y renombrar las columnas
    """
    #Si vamos a crear el df localizando una variable
    if nombre:
        df_res = df.loc[nombre] #Localizamos la tabla que queremos seleccionar
        df_res = df_res.iloc[1:, 1:] #Eliminamos la primera fila y la primera columna
        df_res.rename(columns = df_res.iloc[0,:], inplace = True) #Renombramos las columnas por los nombres de la primera fila
        
        # Eliminar columnas completamente NaN
        df_res = df_res.dropna(axis=1, how='all')

    #Si vamos a hacer cambios en el df partiendo del metodo separar por header (no hace falta localizar el df por nombre)
    else:
        df_res = df.iloc[:, 1:] #Eliminamos la primera columna
        df_res.rename(columns = df_res.iloc[0,:], inplace = True) #Renombramos las columnas por los nombres de la primera fila
        df_res = df_res.dropna(axis=1, how='all')
        # Eliminar filas que contienen "Total" en cualquier columna
        df_res = df_res[~df_res.apply(lambda row: row.astype(str).str.contains('Total', case=False)).any(axis=1)]

    return df_res


def convertir_fecha(df: pd.DataFrame, columna: str, formato_actual : str) -> pd.DataFrame:
    """
    Convertir las columnas a fechas, teniendo en cuenta el formato que nos da el report.
    Asegura que la columna se devuelva solo con la fecha, sin horas, y en el formato dd/mm/yyyy.
    """
    df = df.copy() #Para evitar el SettingWithCopyWarning, para que no haya efectos secundarios no deseados en df original
    # Eliminar filas donde xolumna sea NaN 
    df = df[df[columna].notna()]

    # Convertir la columna a formato fecha dependiendo del formato
    df[columna] = pd.to_datetime(df[columna], format=formato_actual, errors = 'coerce')
    #Independientemente de en que formato recibamos la fecha, devolvemos el df con el siguiente formato:
    df[columna] = df[columna].dt.strftime('%d/%m/%Y')

    return df

#OJO: Hay casos en el df en el que para una misma tabla ponen formatos diferentes para las fechas, es decir, igual te encuentras las fechas en este formato '%m/%d/%y',
# que en este otro '%Y-%d-%m', de eso me he dado cuenta en el POWER BI, ya que me aparecían datos de septiembre y octubre, cuando el report es de mayo
def convertir_fecha_por_filas(fecha):
    if '/' in fecha:
        return pd.to_datetime(fecha, format='%m/%d/%y')
    elif '-' in fecha:
        fecha = fecha.split()[0]
        return pd.to_datetime(fecha, format='%Y-%d-%m')

def convertir_nums(df : pd.DataFrame, cols_num : list):
    #Para cada columna numerica de un dataframe cambiar el tipo de dato a float64
    for col in cols_num:
        df[col] = pd.to_numeric(df[col], errors='coerce') # Usar errors='coerce' para convertir a NaN los valores no convertibles
    return df


def unpivot(df : pd.DataFrame, col_fija : list, vars_con_valor : list ,nom_nueva_var : str, nom_col_valores : str):
    #Si tenemos una tabla con muchas columnas, y queremos convertir varias columnas en filas por ej:
    #Tenemos Date, Bonds, Fixed Income, Stocks
    # Queremos convertir la tabla en Date, Instrumento Financiero, Capital(EUR)
    df = pd.melt(df, id_vars= col_fija , value_vars= vars_con_valor, var_name= nom_nueva_var, value_name = nom_col_valores)
    #               Columnas fijas    Variables con valores      Nombre nueva col con variables  Nombre nueva col con valores
    return df



# KEY STATISTICS

df_key_stats = punto_partida(df, 'Key Statistics')


nuevos_nombres = {
    'BeginningNAV': 'NAV Inicial',
    'EndingNAV': 'NAV Final',
    'CumulativeReturn': 'Rentabilidad Acumulada',
    '5DayReturn': 'Rentabilidad a 5 Días',
    '5DayReturnDateRange': 'Rango de Fechas Rentabilidad a 5 Días',
    '10DayReturn': 'Rentabilidad a 10 Días',
    '10DayReturnDateRange': 'Rango de Fechas Rentabilidad a 10 Días',
    'BestReturn': 'Mejor Rentabilidad',
    'BestReturnDate': 'Fecha de Mejor Rentabilidad',
    'WorstReturn': 'Peor Rentabilidad',
    'WorstReturnDate': 'Fecha de Peor Rentabilidad',
    'MTM': 'Valor de Mercado',
    'Deposits & Withdrawals': 'Depósitos y Retiros',
    'Dividends': 'Dividendos',
    'Interest': 'Intereses',
    'Fees & Commissions': 'Comisiones y Tarifas',
    'Other': 'Otros',
    'ChangeInNAV': 'Cambio en NAV'
}
df_key_stats.rename(columns = nuevos_nombres, inplace = True) #Cambiamos el nombre de los headers

# Convertir todas las columnas a numérico después del renombramiento
df_key_stats = convertir_nums(df_key_stats, ['NAV Inicial', 'NAV Final', 'Rentabilidad Acumulada', 'Rentabilidad a 5 Días', 'Rango de Fechas Rentabilidad a 5 Días', 'Rentabilidad a 10 Días', 'Rango de Fechas Rentabilidad a 10 Días', 'Mejor Rentabilidad','Peor Rentabilidad', 'Valor de Mercado', 'Depósitos y Retiros', 'Dividendos', 'Intereses', 'Comisiones y Tarifas', 'Otros', 'Cambio en NAV'])

#Eliminamos primera fila
df_key_stats = df_key_stats.iloc[1:,:].reset_index(drop = True).dropna(axis = 1)


#Convertimos los valores a fechas
df_key_stats = convertir_fecha(df_key_stats, 'Fecha de Mejor Rentabilidad', '%Y%m%d')
df_key_stats = convertir_fecha(df_key_stats, 'Fecha de Peor Rentabilidad', '%Y%m%d')



df_key_stats['NAV Inicial'] = 6814     #Es el valor neto de la cartera al comienzo, en este caso es una constante
df_key_stats['P&L'] = df_key_stats['NAV Final']-df_key_stats['NAV Inicial']   #ProfitOrLoss en valor neto del activo 
df_key_stats['P&L %'] = ((df_key_stats['NAV Final'] - df_key_stats['NAV Inicial'])/ df_key_stats['NAV Inicial'])*100 #Porcentaje de Ganancias o Pérdidas, teniendo en cuenta el punto inicial y el final
df_key_stats['Rentabilidad Acumulada'] = df_key_stats['Rentabilidad Acumulada']/100 

# Guardar en un archivo Excel
#df_key_stats.to_excel('estadisticas.xlsx', index=False)




# POSICIONES ABIERTAS

df_open_position = punto_partida(df, 'Open Position Summary')


#Eliminamos de la columna fecha los totales que va poniendo el reporte descargado, ya que en Power Bi trabajaremos con gráficos con opciones de ese tipo de filtrado
df_open_position = df_open_position[df_open_position['Date'] != 'Total']

#Cambiamos los nombres de las columnas al castellano

nuevos_nombres = {
    'Date': 'Date',
    'FinancialInstrument': 'FinancialInstrument',
    'Symbol': 'Symbol', #En que activo se invierte
    'Description': 'Description',
    'Sector': 'Sector', #Sector económico
    'Currency': 'Currency',
    'Quantity': 'Cantidad', #Cantidad que se adquiere 
    'ClosePrice': 'Precio Cierre', #El precio de cierre
    'Value': 'Valor', #Cantidad que se adquiere X Precio de cierre
    'Cost Basis': 'Costo Base', #Precio que se pagó inicialmente por el activo
    'UnrealizedP&L': 'Ganancia/Pérdida', #Calculo a partir del precio de cierre con el costo base
    'FXRateToBase': 'Tipo de Cambio' 
}

df_open_position.rename(columns = nuevos_nombres, inplace = True) #Cambiamos el nombre de los headers



#Eliminamos el antiguo header que coincide con la primera fila
df_open_position = df_open_position[1:]

# Convertir tipos de datos

#Columnas categóricas y númericas respectivamente
col_num = ['Cantidad','Precio Cierre','Valor','Costo Base','Ganancia/Pérdida','Tipo de Cambio']
df_open_position = convertir_nums(df_open_position, col_num)

#Llamamos a la función convertir fecha
df_open_position = convertir_fecha(df_open_position, 'Date', '%m/%d/%Y')


#Calculo de la ganacia perdida porcentual
df_open_position['UnrealizedP&L (%)'] = df_open_position['Ganancia/Pérdida']/df_open_position['Costo Base'] 

#Calculo del Valor en Euros de los activos
df_open_position['Valor en EUR'] = df_open_position['Valor']*df_open_position['Tipo de Cambio']

#Caclulamos la ganancia o perdida en euros
df_open_position['Ganancia/Pérdida EUR'] = df_open_position['Ganancia/Pérdida']*df_open_position['Tipo de Cambio']

#Calculamos el costo base por unidad
df_open_position['Costo Base por Ud'] = (df_open_position['Costo Base']/ df_open_position['Cantidad'])*df_open_position['Tipo de Cambio']


df_open_position = convertir_nums(df_open_position, ['Valor en EUR', 'Ganancia/Pérdida EUR', 'Costo Base por Ud'])
#df_open_position.to_excel('posiciones_abiertas.xlsx', index=False)









# HISTORICAL PERFORMANCE BENCHMARK COMPARISON

# Guardar las filas filtradas en un nuevo DataFrame

df_comparison = df.loc['Historical Performance Benchmark Comparison']  #Filtra todas las filas por la columna que tenga de etiqueta Open Position Summary y omite esta columna

comparison_dfs = separar_en_dfs(df_comparison)

# Guardar cada DataFrame en un archivo Excel separado
lista_nombres_comparaciones = ['Resumen', 'Mensual', 'Trimestral' ,'YTD']

del comparison_dfs[0] #El primer df no sirve

#Crear data sets para tratarlos a partir del data set de comparaciones en general
df_comparison_resumen = comparison_dfs[0]
df_comparison_mensual = comparison_dfs[1]


#df_comparison_resumen:

df_comparison_resumen = punto_partida(df_comparison_resumen)

col_num = ['MTD','QTD','YTD','1 Year','Since Inception',]

#Cambio el identificador de la cuenta por un nombre más representativo
df_comparison_resumen.replace('U11259084', 'UPVIC', inplace=True)

df_comparison_resumen = convertir_nums(df_comparison_resumen, col_num)

df_comparison_resumen = df_comparison_resumen.iloc[1:] # Eliminar la primera fila usando indexado

#df_comparison_resumen.to_excel('comparison_resumen_XTD.xlsx', index=False)


#df_comparison_mensual, 
#De este se puede sacar la informacion trimestral, anual, y demás con POWER BI
#por eso omitiremos los demas dataset obtenidos anteriormente :

df_comparison_mensual = punto_partida(df_comparison_mensual)
df_comparison_mensual = convertir_nums(df_comparison_mensual,['Month'])


# Convertir el DataFrame en el formato (Fecha, Índice, Retorno)
# Paso 1: Repetir la columna 'Month' para cada índice
df_melted = unpivot(df_comparison_mensual, ['Month'], ['BM1', 'BM2', 'BM3', 'Account'], 'Indice', 'Indice_Nombre')

# Paso 2: Repetir el DataFrame para agregar los retornos correspondientes
retornos = unpivot(df_comparison_mensual, ['Month'], ['BM1Return', 'BM2Return', 'BM3Return', 'AccountReturn'],'Retorno','Valor_Retorno')

# Paso 3: Combinar ambos DataFrames
df_final = df_melted.copy()
df_final['Retorno'] = retornos['Valor_Retorno']

# Paso 4: Eliminar las columnas innecesarias
df_final = df_final.drop(columns=['Indice'])

#Cambiamos nombres a las columnas
df_final.columns = ['Date', 'Indice', 'Retorno']

# Paso 5: Reemplazar los valores '-' por NaN si es necesario
df_final.replace('-', None, inplace=True)

df_final['Retorno'] = df_final['Retorno'].fillna(0) # Reemplazar NaN en la columna 'Retorno' con 0
df_final['Indice'] = df_final['Indice'].replace('U11259084', 'UPVIC') # Reemplazar U11259084 con UPVIC en el DataFrame

# Eliminar filas donde 'Indice' tenga los valores BM1, BM2, BM3 o Account
df_final = df_final[~df_final['Indice'].isin(['BM1', 'BM2', 'BM3', 'Account'])]

#Llamamos a la función convertir fecha
df_final = convertir_fecha(df_final, 'Date', '%Y%m')

#df_final.to_excel('comparacion_indice_retorno.xlsx', index=False)




# CONCENTRACION:


df_concentracion = df.loc['Concentration']  #Filtra todas las filas por la columna que tenga de etiqueta Concentration y omite esta columna

concentracion_dfs = separar_en_dfs(df_concentracion)
del concentracion_dfs[0] #El primer df no sirve

#Crear data sets para tratarlos a partir del data set de comparaciones en general

#Guardamos en un excel cada subconjunto y los limpiamos un poco (despues cada uno tendra que tratarse de una forma)
for i, subconjunto in enumerate(concentracion_dfs):
    subconjunto = punto_partida(subconjunto)
    subconjunto = subconjunto.iloc[1:,1:]
    # Eliminar la columna '+/-' si existe en el subconjunto
    if '+/-' in subconjunto.columns:
        subconjunto = subconjunto.drop(columns=['+/-'])
    concentracion_dfs[i] = subconjunto


#Hago esto para poder tratar las peculiaridades de cada data set
df_concentracion_holdings = concentracion_dfs[0]
df_concentracion_claseAct = concentracion_dfs[1]
df_concentracion_sector = concentracion_dfs[2]
df_concentracion_region = concentracion_dfs[3]
df_concentracion_pais = concentracion_dfs[4]
df_concentracion_instrumentoFin = concentracion_dfs[5]
df_concentracion_exposure = concentracion_dfs[6]

df_concentracion_holdings__num_cols = ['Value', 'ParsedWeight']
df_concentracion_claseAct_num_cols =  ['LongWeight', 'LongParsedWeight']
df_concentracion_sector_num_cols = ['LongWeight', 'LongParsedWeight']
df_concentracion_region_num_cols =['LongWeight', 'LongParsedWeight'] 
df_concentracion_pais_num_cols = ['LongWeight', 'LongParsedWeight']
df_concentracion_instrumentoFin_num_cols = ['LongWeight', 'LongParsedWeight']
df_concentracion_exposure_num_cols = ['LongAmount', 'LongPercent', 'ShortAmount', 'ShortPercent', 'GrossAmount', 'GrossPercent', 'NetAmount', 'NetPercent']

cols_num =[df_concentracion_holdings__num_cols, 
           df_concentracion_claseAct_num_cols,
           df_concentracion_sector_num_cols, 
           df_concentracion_region_num_cols,
           df_concentracion_pais_num_cols,
           df_concentracion_instrumentoFin_num_cols, 
           df_concentracion_exposure_num_cols,]


# Crear un diccionario donde las claves son índices de los DataFrames
dic_dfs_concentraciones = {}

# Asocia cada subconjunto con su respectiva lista de columnas numéricas usando el índice como clave
for i, subconjunto in enumerate(concentracion_dfs):
    dic_dfs_concentraciones[i] = cols_num[i]

# Iteramos sobre el diccionario, accediendo a los DataFrames a través de su índice
for i, num_col in dic_dfs_concentraciones.items():
    subconjunto = concentracion_dfs[i]  # Acceder al subconjunto por su índice
    # Convertir columnas numéricas a float
    subconjunto = convertir_nums(subconjunto,num_col)
    







#ALLOCATION BY ASSET CLASS
df_allbyClass = punto_partida(df, 'Allocation by Asset Class')



# Aplicar str.isnumeric() solo a los valores no nulos
df_allbyClass = df_allbyClass[df_allbyClass['Date'].apply(lambda x: str(x).isnumeric())]

# Convertir la columna 'Date' al formato de fecha adecuado
df_allbyClass = convertir_fecha(df_allbyClass,'Date','%Y%m%d')


#Columnas
df_allbyClass_cols = ['Date','Equities', 'Fixed Income','Cash','NAV']
allbyClass_cols_num = ['Equities', 'Fixed Income','Cash','NAV']

#Convertimos todos los datos númericos(que realmente estan en string) a float
df_allbyClass = convertir_nums(df_allbyClass, allbyClass_cols_num)


#Unpivot para tener una tabla con el siguiente formato: Date | Clase de Activo | Capital (EUR)
df_allbyClass = unpivot(df_allbyClass, ['Date'], ['Equities', 'Fixed Income', 'Cash', 'NAV'], 'Clase de Activo', 'Capital (EUR)')


#df_allbyClass.to_excel(f'Allocation By Asset Class.xlsx', index=False)










#ALLOCATION BY FINANCIAL INSTRUMENT
df_allbyFin = punto_partida(df, 'Allocation by Financial Instrument')

# Aplicar str.isnumeric() solo a los valores no nulos
df_allbyFin = df_allbyFin[df_allbyFin['Date'].apply(lambda x: str(x).isnumeric())]

# 4. Convertir la columna 'Date' al formato de fecha adecuado
df_allbyFin = convertir_fecha(df_allbyFin,'Date','%Y%m%d')


#Columnas
df_allbyFin_cols_num = ['Bonds','Stocks','Cash', 'NAV']
df_allbyFin = convertir_nums(df_allbyFin, df_allbyFin_cols_num)


#Unpivot para tener una tabla con el siguiente formato: Date | Clase de Instrumento Financiero | Capital (EUR)
df_allbyFin = unpivot(df_allbyFin, ['Date'], ['Bonds', 'Stocks', 'Cash', 'NAV'],'Clase de Instrumento Financiero','Capital (EUR)')

#df_allbyFin.to_excel(f'Allocation by Financial Instrument.xlsx', index=False)









# Time Period Benchmark Comparison: Comparación con indices

df_benchmark = punto_partida(df,'Time Period Benchmark Comparison')

# Convertir a formato de 3 columnas (Date, Indice, Retorno)
df_melted = unpivot(df_benchmark, ['Date'], ['BM1', 'BM2', 'BM3', 'U11259084'],'Indice','IndiceNombre')

# Separar los valores de retorno asociados a cada índice
retornos = unpivot(df_benchmark,['Date'],['BM1Return', 'BM2Return', 'BM3Return', 'U11259084Return'],'IndiceReturn','RetornoValue')

# Concatenar los DataFrames para formar el formato deseado
df_melted['Retorno'] = retornos['RetornoValue']

df_benchmark = df_melted

#Cambiamos el nombre de nuestra cartera a UPVIC
df_benchmark.replace('U11259084','UPVIC' ,inplace =True)

#Eliminamos las filas en las que pone Date para la columna Date (error que se genera con el metodo melt)
df_benchmark = df_benchmark[df_benchmark['Date'] != 'Date']


# Convertir la columna 'Retorno' a valores numéricos
df_benchmark.loc[:, 'Retorno'] = pd.to_numeric(df_benchmark['Retorno'], errors='coerce')

df_benchmark = df_benchmark.drop('Indice', axis=1)
df_benchmark = df_benchmark.rename(columns={'IndiceNombre': 'Indice'})

# Aplicar la función a la columna 'Date'
df_benchmark['Date'] = df_benchmark['Date'].astype(str)  # Asegura que todos son strings
df_benchmark['Date'] = df_benchmark['Date'].apply(convertir_fecha_por_filas)

# Formatear las fechas a 'DD/MM/YYYY' una vez convertidas
df_benchmark['Date'] = df_benchmark['Date'].dt.strftime('%d/%m/%Y')


# Convertir la columna 'Retorno' a valores numéricos
df_benchmark['Retorno'] = pd.to_numeric(df_benchmark['Retorno'], errors='coerce')

# Verificar si hay valores NaN en 'Retorno' después de la conversión
# Puedes eliminarlos o imputarlos según sea necesario
df_benchmark = df_benchmark.dropna(subset=['Retorno'])

# Ahora aplica el cálculo de retorno acumulativo
df_benchmark['RetornoAcumulativo'] = df_benchmark.groupby('Indice')['Retorno'].cumsum()



#df_benchmark.to_excel(f'Time Period Benchmark Comparison.xlsx', index=False)





# Cumulative Performance Statistics


df_performance = punto_partida(df, 'Cumulative Performance Statistics')

# Convertir la columna 'Retorno' a valores numéricos
df_performance.loc[:, 'Return'] = pd.to_numeric(df_performance['Return'], errors='coerce')

#Cambiamos el nombre de nuestra cartera a UPVIC
df_performance.replace('U11259084','UPVIC' ,inplace =True)

# Eliminar filas con valores nulos
df_performance.dropna(inplace=True)

#Convertimos el tipo de dato a Date
df_performance['Date'] = df_performance['Date'].astype(str)
df_performance['Date'] = df_performance['Date'].apply(convertir_fecha_por_filas)

#df_performance.to_excel(f'Rendimiento Cartera.xlsx', index=False)





#Risk Measures:

df_risk = punto_partida(df,'Risk Measures')

# Eliminar la columna "Account"
df_risk.drop(columns=['Account'], inplace=True)


#Transponemos el data set para que cada metrica sea una columna
df_risk = df_risk.set_index('RiskRatio').T



nuevos_nombres = {
    "Ending VAMI:": "VAMI final",  # Valor de la inversión en activos monetarios al final del período.
    "Max Drawdown:": "Máxima caída",  # La máxima pérdida porcentual desde un pico hasta el valle más bajo durante un período.
    "Peak-To-Valley:": "Pico a valle",  # Período entre el punto más alto (pico) y el punto más bajo (valle) de una inversión.
    "Recovery:": "Recuperación",  # Tiempo que toma recuperar la inversión después de una caída.
    "Sharpe Ratio:": "Ratio de Sharpe",  # Mide el rendimiento adicional por unidad de riesgo en comparación con la tasa libre de riesgo.
    "Sortino Ratio:": "Ratio de Sortino",  # Similar al Ratio de Sharpe, pero considera solo la volatilidad negativa.
    "Standard Deviation:": "Desviación estándar",  # Mide la variabilidad de los retornos de la inversión.
    "Downside Deviation:": "Desviación a la baja",  # Mide solo la variabilidad de los retornos negativos.
    "Turnover:": "Rotación",  # Frecuencia de compra y venta de activos en una cartera.
    "Mean Return:": "Rentabilidad media",  # Rentabilidad promedio de la inversión durante el período analizado.
    "Positive Periods:": "Períodos positivos",  # Número de períodos en los que la inversión tuvo rendimientos positivos.
    "Negative Periods:": "Períodos negativos"  # Número de períodos en los que la inversión tuvo rendimientos negativos.
}

df_risk.rename(columns = nuevos_nombres, inplace = True) #Cambiamos el nombre de los headers


#df_risk.to_excel(f'Métricas de Riesgo.xlsx', index=False)







#Performance by Asset Class
df_perf_asset = punto_partida(df, 'Performance by Asset Class')



# Aplicar str.isnumeric() solo a los valores no nulos
df_perf_asset = df_perf_asset[df_perf_asset['Date'].apply(lambda x: str(x).isnumeric())]

# Convertir la columna 'Date' especificando el formato correcto
df_perf_asset = convertir_fecha(df_perf_asset, 'Date','%Y%m%d')

#Convertimos todos los datos númericos(que realmente estan en string) a float
perf_asset_col_num = ['Equities', 'Fixed Income','Cash']
df_perf_asset = convertir_nums(df_perf_asset, perf_asset_col_num)


# Calcular el acumulado para cada clase de activo
for col in perf_asset_col_num:
    df_perf_asset[f'{col} Acumulado'] = df_perf_asset[col].cumsum()

df_perf_asset = convertir_nums(df_perf_asset, df_perf_asset.columns[1:])

columnas = ['Date', 'Equities', 'Fixed Income', 'Cash', 'Equities Acumulado', 'Fixed Income Acumulado', 'Cash Acumulado']
df_perf_asset.columns = ['Date', 'Equities 0', 'Fixed Income 0', 'Cash 0', 'Equities', 'Fixed Income', 'Cash']

#Unpivot para tener tres columnas, Date, Clase de Activo, Retorno Acumulado
df_perf_asset = unpivot(df_perf_asset, ['Date'],['Equities', 'Fixed Income', 'Cash'], 'Clase de Activo','Retorno Acumulado')

#df_perf_asset.to_excel(f'Rendimiento Por Activo.xlsx', index=False)



#Performance by Symbol
df_perf_symbol = punto_partida(df,'Performance by Symbol')
df_perf_symbol = df_perf_symbol.iloc[1:,:]

# Eliminar filas donde la columna 'FinancialInstrument' es nula
df_perf_symbol = df_perf_symbol.dropna(subset=['FinancialInstrument']) #Para quitarnos la suma del total que nos da IBKR, las calcularemos después en POWERBI si hace falta

# Convertir todas las columnas en lista_num a numéricas
lista_num = ['AvgWeight', 'Return', 'Contribution', 'Unrealized_P&L', 'Realized_P&L']

df_perf_symbol = convertir_nums(df_perf_symbol, lista_num)

#AvgWeight: Mide el peso promedio de un activo en la cartera.
#Return: Indica la rentabilidad de un activo durante un período.
#Contribution: Muestra cuánto ha contribuido ese activo al rendimiento total de la cartera.
#Unrealized_P&L: Refleja las ganancias o pérdidas potenciales (no realizadas) de un activo.
#Realized_P&L: Refleja las ganancias o pérdidas efectivas tras la venta de un activo.


#df_perf_symbol.to_excel(f'Rendimiento Por Símbolo.xlsx', index=False)




# Fixed Income Summary


df_fixed_inc = df.loc['Fixed Income']  #Filtra todas las filas por la columna que tenga de etiqueta Fixed Income Summary y omite esta columna


fix_inc_dfs = separar_en_dfs(df_fixed_inc)

df_fix_inc_summary = fix_inc_dfs[1] #Del método anterior, solo nos interesa este data set

df_fix_inc_summary = punto_partida(df_fix_inc_summary)


df_fix_inc_summary = df_fix_inc_summary.set_index('Fixed Income Summary').T #Transponemos

df_fix_inc_summary = df_fix_inc_summary.iloc[:, 1:] #Eliminamos la primera columna

nuevos_nombres_fixed_income = {
    'Total Market Value': 'Valor de Mercado Total',  # Valor total de todos los activos de renta fija
    'Total Face Value': 'Valor Nominal Total',        # Valor nominal total de los bonos
    'Average Maturity': 'Madurez Promedio',           # Madurez promedio de los activos
    'Average Duration': 'Duración Promedio',          # Duración promedio en años
    'Average YTM': 'Rendimiento Promedio al Vencimiento',  # Rendimiento promedio al vencimiento
    'Average Coupon': 'Cupón Promedio'                # Tasa de cupón promedio
}



df_fix_inc_summary.rename(columns = nuevos_nombres_fixed_income, inplace = True) #Cambiamos el nombre de los headers


#df_fix_inc_summary.to_excel(f'Ingresos Renta Fija Resumen.xlsx', index=False)



# Trade summary

df_trade_summary = punto_partida(df,'Trade Summary')
df_trade_summary = df_trade_summary.iloc[1:,:]

trade_col_num = ['Quantity Bought', 'Average Price Bought', 'Proceeds Bought',
       'Proceeds Bought in Base', 'Quantity Sold', 'Average Price Sold',
       'Proceeds Sold', 'Proceeds Sold in Base']

df_trade_summary = convertir_nums(df_trade_summary, trade_col_num)

traducciones = {
    'Quantity Bought': 'Cantidad Comprada',
    'Average Price Bought': 'Precio Medio Comprado',
    'Proceeds Bought': 'Ingresos Comprados',
    'Proceeds Bought in Base': 'Ingresos Comprados en Base',
    'Quantity Sold': 'Cantidad Vendida',
    'Average Price Sold': 'Precio Medio Vendido',
    'Proceeds Sold': 'Ingresos Vendidos',
    'Proceeds Sold in Base': 'Ingresos Vendidos en Base'
}

df_trade_summary.rename(columns=traducciones, inplace=True)

 
#df_trade_summary.to_excel(f'Resumen de Operaciones.xlsx', index=False)


#Guardar cada df en una hoja en un excel, veréis que en el código arriba comento los df.to_excel, eso lo utilice para ir viendo los cambios que hacia de forma rápida
#Pero el resultado esperado de este programa es devolver un excel solo.
dic_dfs_guardar = {
    'Estadísticas Generales': df_key_stats,
    'Posiciones Abiertas': df_open_position,
    'Comparación Resumen': df_comparison_resumen,
    'Comparación Mensual': df_final,
    'conc_Holdings(Posiciones)': df_concentracion_holdings,
    'conc_Clase de Activos': df_concentracion_claseAct,
    'conc_Sector': df_concentracion_sector,
    'conc_Región': df_concentracion_region,
    'conc_País': df_concentracion_pais,
    'conc_Instrumento Financiero': df_concentracion_instrumentoFin,
    'conc_Exposición de la Cartera': df_concentracion_exposure,
    'Distr por Activos': df_allbyClass,
    'Distr por Instrumento Fin': df_allbyFin,
    'Comparación con índices': df_benchmark,
    'Rendimiento Por Símbolo': df_perf_symbol,
    'Rendimiento Por Activo': df_perf_asset,
    'Métricas de Riesgo': df_risk,
    'Rendimiento Cartera': df_performance,
    'Resumen de Operaciones': df_trade_summary,
    'Ingresos Renta Fija Resumen': df_fix_inc_summary
}


# Crear archivo de Excel y escribir en cada hoja
with pd.ExcelWriter('excel_ibkr_upvic_limpio.xlsx') as writer:
    for nombre, df in dic_dfs_guardar.items():
        df.to_excel(writer, sheet_name=nombre, index=False)