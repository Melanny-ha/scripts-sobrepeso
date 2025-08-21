def main_Script_Env_Sol_v2():
    #!/usr/bin/env python
    # coding: utf-8

    # **Script compacto sobrepeso - Envase Soluble**

    # In[1]:


    import os
    import numpy as np
    import pandas as pd
    import warnings
    import funciones_tpm as fc


    # In[2]:


    #Ocultar advertencias/warnings mas no los borra
    warnings.filterwarnings("ignore", category=pd.errors.DtypeWarning)


    # In[3]:


    #Año evaluado
    Añoeval = 2025


    # **Variables globales de las rutas necesarias**

    # In[4]:


    #Rutas del archivo de la COOISPI
    Ruta_COOISPI = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\COOISPI\Datos_COOISPI_Env_Sol.csv'                                    # Ruta COOISPI en csv

    #Rutas del archivo del EGE
    Ruta_Archivo_EGE = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Backup_BD_Salones\Datos_Env_Soluble_2021.csv'                     # Ruta Envase Soluble

    #Rutas del archivo de Novedades
    Ruta_Novedades = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Env_Soluble_Novedades_V2.xlsx'  # Ruta del archivo de Novedades

    #Rutas del archivo de Consolidados
    Ruta_Consolidados = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Consolidado_V2.xlsx'        # Ruta del archivo de Consolidados

    # #Ruta de salida
    # Ruta_Env_Sol = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble'                                 # DB_Envase_Soluble_Out_vf.csv

    #Ruta - Metas de Sobrepeso - Env Soluble
    Ruta_Metas_Sobrepeso = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Metas_Env_Sol_2025.xlsx'  #Ruta del archivo metas sobrepeso


    #Validacion de existencia de las rutas
    for nombre, ruta in {
        "Ruta_COOISPI": Ruta_COOISPI,
        "Ruta_Archivo_EGE": Ruta_Archivo_EGE,
        "Ruta_Novedades": Ruta_Novedades,
        "Ruta_Consolidados": Ruta_Consolidados,
        # "Ruta_Env_Sol": Ruta_Env_Sol,
        "Ruta_Metas_Sobrepeso": Ruta_Metas_Sobrepeso,
    }.items():
        if not os.path.exists(ruta):
            print(f"La ruta '{nombre}' no fue encontrada.")
        else:
            print(f"Ruta '{nombre}' encontrada.")


    # In[5]:


    #leer_archivo COOISPI csv
    df_COOISPI_COMB = fc.leer_archivo(Ruta_COOISPI)


    # In[6]:


    #Convertir columna 'Codigo' a str para futuros procesor (merge y tabla pivote)
    df_COOISPI_COMB['Codigo'] = df_COOISPI_COMB['Codigo'].astype(str)


    # In[7]:


    #Funcion para leer archivo el EGE/TPM
    df_Env_Sol = fc.leer_archivo(Ruta_Archivo_EGE, 'Datos')
    print(df_Env_Sol.shape)


    # In[8]:


    df_Env_Sol_Metas = fc.leer_archivo(Ruta_Metas_Sobrepeso, 'Hoja1')
    print(df_Env_Sol_Metas.head())


    # In[9]:


    df_Env_Sol.head()


    # In[10]:


    #Reemplazar el nombre de la columna "Día:" y "Mes_N" para estandarizar los scripts
    df_Env_Sol.rename(columns={"Dia ":"Dia", "Mes_N":"IdMes"}, inplace=True)


    # In[11]:


    #eliminar_filas_vacias
    numeros_vacios_filtrados = fc.eliminar_filas_vacias(df_Env_Sol, 'Número')
    numeros_vacios_filtrados.head()


    # In[12]:


    df_Env_Sol.shape


    # In[13]:


    #crear_fecha
    df_Env_Sol = fc.crear_fecha(df_Env_Sol)
    print(df_Env_Sol.shape)
    df_Env_Sol.head(2)


    # In[14]:


    #Tomar solo las columnas de interes del informe de Envase Soluble para el sobrepeso
    columnas=['Número','Dia','Fecha','Mes:','IdMes','Año:','Máquina / Equipo:','Semana:','Turno:','Código y Descrip. / Producto:','Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):','Gramaje (K):']
    df_Env_Sol = df_Env_Sol[columnas]
    df_Env_Sol.columns


    # In[15]:


    df_Env_Sol.dtypes


    # In[16]:


    #filtrar_desde_anio
    df_Env_Sol = fc.filtrar_desde_anio(df_Env_Sol, 2024)
    print(df_Env_Sol.shape)
    df_Env_Sol.head(2)


    # In[17]:


    # Especificar las columnas en las que deseas aplicar el filtro para eliminar los valores nulos de las columnas de Gramaje, Unidades producidas y Peso promedio
    columnas_filtrar_nulos = ['Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):']


    # In[18]:


    #filtrar_nulos
    df_Env_Sol_nan_ceros, df_Env_Sol = fc.filtrar_nulos(df_Env_Sol, columnas_filtrar_nulos)
    df_Env_Sol_nan_ceros.shape
    df_Env_Sol.shape


    # In[19]:


    df_Env_Sol_nan_ceros.head()


    # **Modificación de los tipos de datos del archivo de TPM de Envase Soluble**

    # In[20]:


    #filtrar_nulos
    df_non_conver, df_Env_Sol = fc.validar_numericos(df_Env_Sol, 'Peso Promedio de la unidad (K):')
    print(df_Env_Sol.shape)
    print(df_non_conver.shape)


    # In[21]:


    df_non_conver.head()


    # In[22]:


    #extraer_codigo
    df_Env_Sol = fc.extraer_codigo(df_Env_Sol)
    df_Env_Sol['Codigo'].unique()


    # In[23]:


    df_Env_Sol_nan_ceros.head()


    # **Generar un dataframe con los códigos y descripción de códigos**

    # In[24]:


    #Se genera un dataframe con codigos y descripcion de codigos en el cual si hay descripciones repetidas las agrupa en una lista para verificaciones futuras
    df_pivot_Codigos = df_Env_Sol.groupby('Codigo')['Código y Descrip. / Producto:'].unique().reset_index() #Para cada Codigo, toma los valores unicos de cod/descrip y coloca los indices en una columna
    df_pivot_Codigos.head()


    # **Identificación de Codigos con mas de una descripción**

    # In[25]:


    #Agrupar por Código y contar descripciones únicas
    #df_con_multiples_desc = df_pivot_Codigos[df_pivot_Codigos['Código y Descrip. / Producto:'].apply(len) > 1]
    df_con_multiples_desc = df_Env_Sol.groupby('Codigo')['Código y Descrip. / Producto:'].nunique().reset_index() #.nunique() trae CUANTAS descrip unicas hay por codigo
    df_con_multiples_desc = df_con_multiples_desc[df_con_multiples_desc['Código y Descrip. / Producto:'] > 1] #filtra los codigos en el que la descrip (que ya es un numero) sea > 1
    df_con_multiples_desc.head()


    # In[26]:


    df_pivot_Codigos.dtypes


    # **Identificación y supresión de aquellos códigos con valores nulos**

    # In[27]:


    #eliminar_codigos_nan
    df_Env_Sol, df_Env_Sol_Codigos_nan = fc.eliminar_codigos_nan(df_Env_Sol)
    print(df_Env_Sol.shape)
    print(df_Env_Sol_Codigos_nan.shape)


    # In[28]:


    df_Env_Sol_Codigos_nan = df_Env_Sol_Codigos_nan[df_Env_Sol_Codigos_nan['Año:'] == Añoeval]
    df_Env_Sol_Codigos_nan.head()


    # In[29]:


    df_Env_Sol.dtypes


    # **Por temas de sobrepeso se borrará el Codigo 1055015 del archivo de TPM**

    # In[30]:


    #Por temas de Sobrepeso se borrará el código 1055015 del archivo de TPM
    indices_to_drop = df_Env_Sol[(df_Env_Sol['Codigo'] == "1055015")].index
    df_Env_Sol.drop(indices_to_drop, inplace=True)  #inplace=True asegura que la eliminación se haga sobre df_Env_Sol


    # In[31]:


    #columnas_a_filtrar = 'Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):'
    filtro_nan_ceros = df_Env_Sol[columnas_filtrar_nulos].isna().any(axis=1) | (df_Env_Sol[columnas_filtrar_nulos] == 0).any(axis=1)
    df_Env_Sol_nan_ceros = df_Env_Sol[filtro_nan_ceros]
    df_Env_Sol = df_Env_Sol[~filtro_nan_ceros]
    df_Env_Sol_nan_ceros.head()


    # In[32]:


    df_Env_Sol_nan_ceros.dtypes


    # In[33]:


    df_Env_Sol_nan_ceros.shape


    # In[34]:


    #calcular_columnas
    df_Env_Sol = fc.calcular_columnas(df_Env_Sol)
    df_Env_Sol.head(2)


    # **Generación de Novedades**

    # In[35]:


    Sobrepeso_Novedades = 0.5  #50% de sobrepeso


    # In[36]:


    #generar_novedades
    df_Env_Sol_Nov_Sobrepeso = fc.generar_novedades(df_Env_Sol, Sobrepeso_Novedades)
    df_Env_Sol_Nov_Sobrepeso.head()


    # In[37]:


    #Agregar columna 'Costo/kg'
    df_Env_Sol['Costo/kg'] = 0.0
    df_Env_Sol.head(2)


    # **Tabla de Consolidado Mensual**

    # In[38]:


    #pivote_consolidado
    df_Env_Sol_Mes = fc.pivote_consolidado(df_Env_Sol)
    df_Env_Sol_Mes.head()


    # In[39]:


    #Tomar datos desde el año 2024 en adelante
    df_Env_Sol_Mes = fc.filtrar_desde_anio(df_Env_Sol_Mes, 2024)
    df_Env_Sol_Mes.head()


    # In[40]:


    #Visualizacion de EJEMPLO de los datos del consolidado del Mes para el año 2024 y codigo 1003350
    df_Env_Sol_Mes[(df_Env_Sol_Mes['Año:'] == 2024) & (df_Env_Sol_Mes['Codigo'] == "1003350")].head()


    # In[41]:


    #Verificacion de los meses en el DataFrame
    df_Env_Sol_Mes['Mes:'].unique()


    # **Creación de la agrupación Mensual del archivo de semielaborados**

    # In[42]:


    df_costo_semi = df_COOISPI_COMB.copy(deep=True)


    # In[43]:


    df_costo_semi_Mes = fc.pivote_semielaborados(df_costo_semi)


    # In[44]:


    df_costo_semi_Mes[df_costo_semi_Mes['Mes:'] == "Enero"].head()


    # In[45]:


    df_costo_semi_Mes[df_costo_semi_Mes['Año:'] == 2025].head()


    # **Generación del nuevo archivo consolidado V2**

    # In[46]:


    df_Env_Sol_2 = df_Env_Sol.copy(deep=True)


    # In[47]:


    df_Env_Sol_2.head(2)


    # In[48]:


    #Traer la columna 'Costo/kg' al dataframe original
    df_Env_Sol_2 = fc.merge_costo(df_Env_Sol_2, df_costo_semi_Mes)
    df_Env_Sol_2.head(2)


    # In[49]:


    #Eliminar la columna Costo/kg antigua y renombrar la nueva
    df_Env_Sol_2 = fc.modificar_costo(df_Env_Sol_2)
    df_Env_Sol_2.head(2)


    # **Asignación de Meta de Envase Soluble por referencia**

    # In[50]:


    df_Env_Sol_Metas.head(3)


    # In[51]:


    df_Env_Sol_Metas['Codigo'] = df_Env_Sol_Metas['Codigo'].astype(str)
    df_Env_Sol_Metas = df_Env_Sol_Metas[['Codigo', 'Meta Sobrepeso']].copy()
    df_Env_Sol_Metas.head(3)


    # In[52]:


    valor_fijo_2024 = 0.015

    #Renombrar la columna para evitar conflicto temporalmente
    df_Env_Sol_Metas = df_Env_Sol_Metas.rename(columns={'Meta Sobrepeso': 'Meta_nueva'})

    #Hacer merge para obtener Meta_nueva segun codigo
    df_Env_Sol_2 = df_Env_Sol_2.merge(
        df_Env_Sol_Metas,
        on='Codigo',
        how='left'
    )

    #Asignar la meta segun la condicion del año
    df_Env_Sol_2['Meta'] = np.where(
        df_Env_Sol_2['Año:'] == 2024,
        valor_fijo_2024,                #Si es =2024 queda el valor fijo
        df_Env_Sol_2['Meta_nueva']      #Sino valor desde metas_sobrepeso para otros años
    )

    #Eliminar la columna temporal
    df_Env_Sol_2.drop(columns=['Meta_nueva'], inplace=True)
    df_Env_Sol_2.head(3)


    # In[53]:


    df_Env_Sol_2[df_Env_Sol_2['Meta'] != valor_fijo_2024].head(3)


    # **Calculo de la columna 'Costo/kg' para el año 2025**

    # In[54]:


    df_Costos_Mp = fc.leer_archivo(r"\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Tablas_Costos_Productos_MP.xlsx", "Recetas_solubles_2025")
    df_Financiera_Mp = fc.leer_archivo(r"\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Tablas_Costos_Productos_MP.xlsx", "Datos_Financiera_Consolidados")
    df_Costos_Mp.head(2)


    # In[55]:


    df_Financiera_Mp.head(2)


    # In[56]:


    df_Costos_Mp.dtypes


    # In[57]:


    #.bfill() (rellenar "" o NaN hacia arriba) ------ .ffill() (relleno "" o NaN hacia abajo)
    df_Costos_Mp['Codigo'] = df_Costos_Mp['Codigo'].ffill().astype(int).astype(str)
    df_Costos_Mp['Materia_Prima'] = df_Costos_Mp['Materia_Prima'].astype(int).astype(str)
    df_Costos_Mp.dtypes


    # In[58]:


    df_Costos_Mp.tail()


    # In[59]:


    #Dividir metodo validar_numericos
    df_Costos_Mp['Proporción'] = (df_Costos_Mp['Proporción']
        .astype(str)                         # Convertir a string por si hay valores mezclados
        .str.replace(',', '.', regex=False)  # Reemplazar coma por punto
        .str.strip()                         # Eliminar espacios
        .astype(float)                       # Convertir finalmente a float
    )
    df_Costos_Mp.head()


    # In[60]:


    df_Costos_merged = df_Env_Sol_2.merge(df_Costos_Mp, on='Codigo', how='left')
    df_Costos_merged.head(2)


    # In[61]:


    # df_Costos_merged.to_excel("df_Costos_merged_V2.xlsx", index=False)


    # In[62]:


    df_Financiera_Mp.dtypes


    # In[63]:


    #Paso 1: cambiar a formato largo para facilitar join(id_vars=columnas fijas, var_name=nombre columna en base a las columnas originales, value_name: el nombre de la columna que contendrá los valores de las originales)
    df_Financiera_Mp_melted = df_Financiera_Mp.melt(id_vars=['Año', 'Descripción', 'Materia_Prima'],
                                                    var_name='Mes',
                                                    value_name='Costo_materia_kg')

    df_Financiera_Mp_melted['Mes'] = df_Financiera_Mp_melted['Mes'].str.capitalize()
    df_Financiera_Mp_melted['Materia_Prima'] = df_Financiera_Mp_melted['Materia_Prima'].astype(int).astype(str)
    df_Financiera_Mp_melted.rename(columns={'Año': 'Año:', 'Mes': 'Mes:'}, inplace=True)
    df_Financiera_Mp_melted.head(3)


    # In[64]:


    df_Financiera_Mp_melted.dtypes


    # In[65]:


    #Paso 2: hacer un merge entre df_Costos_merged y df_Financiera_Mp_melted
    df_Cost_Finan_Mp = df_Costos_merged.merge(
        df_Financiera_Mp_melted,
        on=['Año:', 'Materia_Prima', 'Mes:'],
        how='left'
    )


    # In[66]:


    # #Paso 3: exportar df_Cost_Finan_Mp como archivo excel
    # df_Cost_Finan_Mp.to_excel("df_Cost_Finan_Mp.xlsx", index=False)
    # df_Cost_Finan_Mp.head(3)


    # In[67]:


    #Paso 4: calcular el costo proporcional por materia prima
    df_Cost_Finan_Mp['Costo_proporcional'] = df_Cost_Finan_Mp['Proporción'] * df_Cost_Finan_Mp['Costo_materia_kg']
    df_Cost_Finan_Mp.head(3)


    # In[68]:


    df_Costo_Mp_Env_Sol = (
        df_Cost_Finan_Mp
        .groupby(['Número','Año:', 'Mes:', 'Codigo'])['Costo_proporcional']
        .sum()
        .reset_index()
        .rename(columns={'Costo_proporcional': 'Costo/kg'})
    )
    df_Costo_Mp_Env_Sol.head(3)


    # In[69]:


    df_Costo_Mp_Env_Sol.loc[(df_Costo_Mp_Env_Sol['Mes:'] == "Marzo") & (df_Costo_Mp_Env_Sol['Año:'] == 2025)].head()


    # In[70]:


    orden_meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


    # In[71]:


    #Paso 3: hacer merge completo sin sobrescribir 'Costo/kg'
    df_mergeado = df_Env_Sol_2.merge(
        df_Costo_Mp_Env_Sol[['Número', 'Año:', 'Mes:', 'Codigo', 'Costo/kg']],
        on=['Número', 'Año:', 'Mes:', 'Codigo'],
        how='left',
        suffixes=('', '_nuevo')
    )


    # In[72]:


    #Paso 4: Actualizar 'Costo/kg' solo si Año es 2025 y Mes desde Marzo
    condicion_actualizacion = (
        (df_mergeado['Año:'] == 2025) &
        (df_mergeado['Mes:'].isin(orden_meses[2:])) &
        (df_mergeado['Costo/kg_nuevo'].notna())
    )


    # In[73]:


    df_mergeado.loc[condicion_actualizacion, 'Costo/kg'] = df_mergeado.loc[condicion_actualizacion, 'Costo/kg_nuevo']


    # In[74]:


    # Paso 5: Eliminar columna auxiliar
    df_mergeado.drop(columns=['Costo/kg_nuevo'], inplace=True)


    # In[75]:


    df_mergeado.loc[(df_mergeado['Mes:'] == "Marzo")&(df_mergeado['Año:'] == 2025)].head(5)


    # In[76]:


    df_Env_Sol_2 = df_mergeado.copy()


    # In[77]:


    df_Env_Sol_2.columns


    # In[78]:


    # df_Costo_Mp_Env_Sol.to_excel("df_Costo_Mp_Env_Sol.xlsx", index=False)


    # In[79]:


    df_Env_Sol_2['Ahorros/Perdidas'] = (df_Env_Sol_2['Meta'] - df_Env_Sol_2['Sobrepeso_calculado']) * df_Env_Sol_2['Real_empaque_calculado'] * df_Env_Sol_2['Costo/kg']
    df_Env_Sol_2.head(3)


    # **Unión de los DataFrames en un DataFrame totalizado**

    # In[80]:


    df_totalizado_Mes = df_Env_Sol_Mes.copy(deep=True)


    # In[81]:


    df_totalizado_Mes[df_totalizado_Mes['Año:'] == 2025].head()


    # In[82]:


    df_costo_semi_Mes.loc[df_costo_semi_Mes['Año:'] == 2025].head()


    # In[83]:


    fc.filas_repetidas(df_totalizado_Mes)


    # In[84]:


    df_totalizado_Mes = fc.unir_dataframes(df_totalizado_Mes, df_costo_semi_Mes)


    # In[85]:


    df_totalizado_Mes[df_totalizado_Mes['Año:'] == 2025].head()


    # In[86]:


    df_totalizado_Mes.head()


    # In[87]:


    df_totalizado_Mes = fc.calcular_columnas_totalizado(df_totalizado_Mes, 0.015)


    # In[88]:


    df_totalizado_Mes.head()


    # In[89]:


    #Eliminar el Codigo 1055015 tomando su indice
    indices_to_drop = df_totalizado_Mes[(df_totalizado_Mes['Codigo'] == "1055015")].index
    df_totalizado_Mes.drop(indices_to_drop, inplace=True)


    # In[90]:


    df_totalizado_Mes.head()


    # **Codigo para generar archivo de busqueda para Envase Soluble**

    # In[91]:


    #eliminar_duplicados
    df_Codigos_Env_Sol = fc.eliminar_duplicados_columna(df_Env_Sol, 'Codigo')
    df_Codigos_Env_Sol['Codigo'].unique()


    # In[92]:


    df_Codigos_Env_Sol.head()


    # In[93]:


    df_Maquinas_Env_Sol = fc.eliminar_duplicados_columna(df_Env_Sol, 'Máquina / Equipo:')
    df_Maquinas_Env_Sol.head()


    # In[94]:


    df_Codigos_Env_Sol_COOISPI = fc.eliminar_duplicados_columna(df_COOISPI_COMB, 'Codigo')
    df_Codigos_Env_Sol_COOISPI['Codigo'].unique()


    # **Codigos que están en TPM pero no están en la COOISPI**

    # In[95]:


    df_Env_Sol_Codigos_unicos = df_Codigos_Env_Sol[~df_Codigos_Env_Sol['Codigo'].isin(df_Codigos_Env_Sol_COOISPI['Codigo'])]
    df_Env_Sol_Codigos_unicos.head()


    # **Generación de un archivo de fechas automáticas**

    # In[96]:


    df_Fechas = fc.generar_fechas("2024-01-01")
    df_Fechas.tail()


    # **Rutas de códigos totales para definición de metas**

    # In[97]:


    df_pivot_Codigos.to_excel(r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Envase_Soluble\Codigos_Totales.xlsx', sheet_name='Hoja1', index=False)

    # In[98]:


    # # print(df_Env_Sol.shape)
    # # print(df_COOISPI_COMB.shape)
    # # print(df_Env_Sol_Mes.shape)
    # print(df_Codigos_Env_Sol.shape)
    # print(df_Maquinas_Env_Sol.shape)
    # # print(df_totalizado_Mes.shape)
    # print(df_Fechas.shape)
    # print(df_Env_Sol_2.shape)
    # print()
    # print(df_Env_Sol_Nov_Sobrepeso.shape)
    # print(df_Env_Sol_Codigos_nan.shape)
    # print(df_Env_Sol_Codigos_unicos.shape)
    # print(df_Env_Sol_nan_ceros.shape)
    # print(df_non_conver.shape)
    # print()
    # print(df_Env_Sol_2.shape)
    # print(df_Costo_Mp_Env_Sol.shape)
    # print(df_Cost_Finan_Mp.shape)
    # print(df_Costos_merged.shape)
    # print(df_Financiera_Mp_melted.shape)
    # print(df_Financiera_Mp.shape)
    # print(df_Costos_Mp.shape)


    # **Consolidado csv**

    # In[99]:


    # #Ruta para consolidado df_Env_Sol_2
    # df_Env_Sol_2.to_csv(Ruta_Env_Sol + "DB_Envase_Soluble_Out_vf.csv", index=False)


    # In[100]:


    #Diccionario de hojas para el archivo de consolidado
    dataframes_consolidado = {
        # "Hoja1": df_Env_Sol,
        # "Hoja2": df_COOISPI_COMB,
        # "Hoja3": df_Env_Sol_Mes,
        "Hoja4": df_Codigos_Env_Sol,
        "Hoja5": df_Maquinas_Env_Sol,
        # "Hoja6": df_totalizado_Mes,
        "Hoja7": df_Fechas,
        "Hoja8": df_Env_Sol_2
    }
    #Actualizar hojas del excel de Consolidado_V2
    fc.actualizar_hojas_excel(Ruta_Consolidados, dataframes_consolidado)


    # **Archivo de novedades: este archivo compila las diferentes novedades presentes en los archivos de TPM**

    # In[101]:


    #Diccionario de hojas para el archivo de novedades
    dataframes_novedades = {
        "Sobrepeso mayor  50%": df_Env_Sol_Nov_Sobrepeso,          #Novedades de sobrepeso
        "Codigos Nulo TPM": df_Env_Sol_Codigos_nan,                #Novedades codigos nulos
        "Codigos unicos en TPM": df_Env_Sol_Codigos_unicos,        #Codigos presentes es TPM y no en la COOISPI
        "Codigos NA,Vacio, 0": df_Env_Sol_nan_ceros,               #Valores nulos o ceros para gramaje, unidades producidas y promedio peso unidad
        "Errores": df_non_conver,                                  #Valores no convertibles (NaN)
    }
    #Guardar el archivo de Novedades
    fc.crear_archivo_novedades(Ruta_Novedades, dataframes_novedades)


if __name__ == '__main__':
    main_Script_Env_Sol_v2()
