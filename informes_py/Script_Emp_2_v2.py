def main_Script_Emp_2_v2():
    #!/usr/bin/env python
    # coding: utf-8

    # In[1]:


    import os
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
    Ruta_COOISPI = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\COOISPI\Datos_COOISPI_Empaques2.csv'                           # Ruta COOISPI en csv

    #Rutas del archivo del EGE
    Ruta_Archivo_EGE = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Backup_BD_Salones\Datos_Emp2_2021.csv'                     # Ruta Empaques2

    #Rutas del archivo de Novedades
    Ruta_Novedades = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Empaques_2\Empaques_2_Novedades_V2.xlsx'   # Ruta del archivo de Novedades

    #Rutas del archivo de Consolidados
    Ruta_Consolidados = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Empaques_2\Consolidado_V2.xlsx'      # Ruta del archivo de Consolidados

    #Ruta de salida
    # Ruta_Emp_2 = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Empaques_2'                               # DB_Emp2_Out_vf.csv

    #Validacion de existencia de las rutas
    for nombre, ruta in {
      "Ruta_COOISPI": Ruta_COOISPI,
      "Ruta_Archivo_EGE": Ruta_Archivo_EGE,
      "Ruta_Novedades": Ruta_Novedades,
      "Ruta_Consolidados": Ruta_Consolidados,
      # "Ruta_Emp_2": Ruta_Emp_2,
    }.items():
      if not os.path.exists(ruta):
        raise FileNotFoundError(f"La ruta '{nombre}' no fue encontrada.")
      else:
        print(f"Ruta '{nombre}' encontrada.")


    # In[5]:


    #Leer_archivo COOISPI csv 
    df_COOISPI_COMB = fc.leer_archivo(Ruta_COOISPI)


    # In[6]:


    #Convertir columna 'Codigo' a str para futuros procesor (merge y tabla pivote)
    df_COOISPI_COMB['Codigo'] = df_COOISPI_COMB['Codigo'].astype(str)


    # In[7]:


    #Funcion para leer archivo el EGE/TPM
    df_Emp_2 = fc.leer_archivo(Ruta_Archivo_EGE, 'Datos')


    # In[8]:


    #Reemplazar el nombre de la columna "Día:" para estandarizar los scripts
    df_Emp_2.rename(columns={"Día:":"Dia"}, inplace=True)


    # In[9]:


    #Eliminar_filas_vacias
    numeros_vacios_filtrados = fc.eliminar_filas_vacias(df_Emp_2, 'Número')


    # In[10]:


    # #Filtrando el DataFrame según las condiciones dadas
    # df_filtrado = df_Emp_2.loc[
    #     (df_Emp_2['Año:'] == 2024) &
    #     (df_Emp_2['Mes:'] == "Abril") &
    #     (df_Emp_2['Código y Descrip. / Producto:'] == "1040596 Cafe COLCAFE Clasic Grnel 2500X1.5g esp")
    # ]

    # #Filtrando el DataFrame según las condiciones dadas
    # df_filtrado = df_Emp_2.loc[
    #     (df_Emp_2['Año:'] == 2024) &
    #     (df_Emp_2['Mes:'] == "Enero") &
    #     (df_Emp_2['Código y Descrip. / Producto:'] == "1056678 Cafe COLCAFE clasico 1.5g 48sob 30ple")
    # ]
    #
    # # Calculando la suma de la columna específica
    # suma_unidades_producidas = df_filtrado['Unidades Producidas (Conformes) :'].sum()
    #
    # # Mostrando la suma
    # print(suma_unidades_producidas)


    # In[11]:


    #Crear_fecha
    df_Emp_2 = fc.crear_fecha(df_Emp_2)
    df_Emp_2.head()


    # In[12]:


    #Tomar solo las columnas de interes del informe de Empaques2 para el sobrepeso
    Columnas=['Número','Dia','Fecha','Mes:','IdMes','Año:','Máquina / Equipo:','Semana:','Turno:','Código y Descrip. / Producto:','Unidades Producidas (Conformes) :','Peso Promedio de la unidad (K):','Gramaje (K):']
    df_Emp_2 = df_Emp_2[Columnas]


    # In[13]:


    #Filtrar_desde_añio
    df_Emp_2 = fc.filtrar_desde_anio(df_Emp_2, 2024)


    # In[14]:


    #Especificar las columnas en las que deseas aplicar el filtro para eliminar los valores nulos de las columnas de Gramaje, Unidades producidas y Peso promedio
    columnas_filtrar_nulos = ['Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):']


    # In[15]:


    #Filtrar_nulos
    df_Emp_2_nan_ceros, df_Emp_2 = fc.filtrar_nulos(df_Emp_2, columnas_filtrar_nulos)


    # **Modificación de los tipos de datos del archivo de TPM de EMPAQUES 2**

    # In[16]:


    #Validar_numericos
    df_non_conver_a, df_Emp_2 = fc.validar_numericos(df_Emp_2, 'Peso Promedio de la unidad (K):')


    # In[17]:


    #Validar_numericos
    df_non_conver_b, df_Emp_2 = fc.validar_numericos(df_Emp_2, 'Gramaje (K):')


    # In[18]:


    #Extraer_codigo
    df_Emp_2 = fc.extraer_codigo(df_Emp_2)


    # **Identificación y supresión de aquellos códigos con valores nulos**

    # In[19]:


    #Eliminar_codigos_nan
    df_Emp_2, df_Emp_2_Codigos_nan = fc.eliminar_codigos_nan(df_Emp_2)


    # In[20]:


    filtro_nan_ceros = df_Emp_2[columnas_filtrar_nulos].isna().any(axis=1) | (df_Emp_2[columnas_filtrar_nulos] == 0).any(axis=1)
    df_Emp_2_nan_ceros = df_Emp_2[filtro_nan_ceros]
    df_Emp_2 = df_Emp_2[~filtro_nan_ceros]


    # In[21]:


    #Calcular_columnas
    df_Emp_2 = fc.calcular_columnas(df_Emp_2)


    # **Generación de Novedades**

    # In[22]:


    Sobrepeso_Novedades = 0.1  #10% de Sobrepeso


    # In[23]:


    #Generar_novedades
    df_Emp_2_Nov_Sobrepeso = fc.generar_novedades(df_Emp_2, Sobrepeso_Novedades)
    # df_Emp_2_Nov_Sobrepeso.head()


    # In[24]:


    #Agregar columna 'Costo/kg'
    df_Emp_2['Costo/kg'] = 0.0


    # **Tabla de Consolidado Mensual**

    # In[25]:


    #Pivote_consolidado
    df_Emp_2_Mes = fc.pivote_consolidado(df_Emp_2)


    # In[26]:


    #Tomar datos desde el año 2024 en adelante
    df_Emp_2_Mes = fc.filtrar_desde_anio(df_Emp_2_Mes, 2024)


    # **Creación de la agrupación Mensual del archivo de semielaborados**

    # In[27]:


    df_costo_semi = df_COOISPI_COMB.copy(deep=True)


    # In[28]:


    #Creacion de la agrupacion mensual del archivo semielaborado
    df_costo_semi_Mes = fc.pivote_semielaborados(df_costo_semi)


    # **Generación del nuevo archivo consolidado V2**

    # In[29]:


    df_Emp2_2 = df_Emp_2.copy(deep=True)


    # In[30]:


    #Traer la columna 'Costo/kg' al dataframe original
    df_Emp2_2 = fc.merge_costo(df_Emp2_2, df_costo_semi_Mes)


    # In[31]:


    #Eliminar la columna Costo/kg antigua y renombrar la nueva
    df_Emp2_2 = fc.modificar_costo(df_Emp2_2)


    # In[32]:


    #Definir meta sobrepeso
    df_Emp2_2['Meta'] = 0.015


    # In[33]:


    #Calcular columna 'Ahorros/Perdidas'
    df_Emp2_2 = fc.calcular_ahorros_perdidas(df_Emp2_2)
    # df_Emp2_2.head()


    # In[34]:


    # df_Emp2_2.dtypes


    # **Unión de los DataFrames en un DataFrame totalizado**

    # In[35]:


    #Crear copia de df_Emp_2_Mes
    df_totalizado_Mes = df_Emp_2_Mes.copy(deep=True)


    # In[36]:


    #Validar que no hay filas repetidas en df_totalizado_Mes
    fc.filas_repetidas(df_totalizado_Mes)


    # In[37]:


    #Fusionar los DataFrames en base a 'Año:', 'Codigo' y 'Mes:'
    df_totalizado_Mes = fc.unir_dataframes(df_totalizado_Mes, df_costo_semi_Mes)


    # In[38]:


    #Metodo calcular_columnas_totalizado con meta fija
    df_totalizado_Mes = fc.calcular_columnas_totalizado(df_totalizado_Mes, 0.015)
    # df_totalizado_Mes.head()


    # In[39]:


    #Obtener Codigos unicos | Cod. Descrip. | Maquina con un merge
    df_Codigos_Emp_2 = fc.obtener_cod_descrip(df_costo_semi, df_Emp_2_Mes)


    # In[40]:


    #Obtener Maquinas únicas de TPM o df_codigos_desc_maq_Mezclas
    df_Maquinas_Emp_2 = fc.eliminar_duplicados_columna(df_Codigos_Emp_2, 'Máquina / Equipo:')


    # In[41]:


    #Unión de los dataframes de no convertibles (validar_y_convertir_datos)
    df_non_conver = pd.concat([df_non_conver_a, df_non_conver_b], ignore_index=True)


    # **Generación de un archivo de fechas Automáticas**

    # In[42]:


    #Generar dataframe con fechas dese 2024-01-01 hasta el dia actual
    df_Fechas = fc.generar_fechas("2024-01-01")


    # In[43]:


    # print(df_Emp_2.shape)
    # print(df_COOISPI_COMB.shape)
    # print(df_Emp_2_Mes.shape)
    # print(df_Codigos_Emp_2.shape)
    # print(df_Maquinas_Emp_2.shape)
    # print(df_totalizado_Mes.shape)
    # print(df_Fechas.shape)
    # print(df_Emp2_2.shape)
    # print()
    # print(df_Emp_2_Nov_Sobrepeso.shape)
    # print(df_Emp_2_Codigos_nan.shape)
    # #print(df_Mol_Codigos_unicos.shape)
    # print(df_Emp_2_nan_ceros.shape)
    # print(df_non_conver.shape)


    # **Consolidado csv**

    # In[44]:


    # Ruta para consolidado df_Emp2_2
    # df_Emp2_2.to_csv(Ruta_Emp_2 + "DB_Emp2_Out_vf.csv", index=False)


    # **Archivo de consolidado**

    # In[45]:


    #Diccionario de hojas para el archivo de consolidado
    dataframes_consolidado = {
        "Hoja1": df_Emp_2,
        "Hoja2": df_COOISPI_COMB,
        "Hoja3": df_Emp_2_Mes,
        "Hoja4": df_Codigos_Emp_2,  #df_merged
        "Hoja5": df_Maquinas_Emp_2,
        "Hoja6": df_totalizado_Mes,
        "Hoja7": df_Fechas,
        "Hoja8": df_Emp2_2,
    }
    #Actualizar hojas del excel de Consolidado_V2
    fc.actualizar_hojas_excel(Ruta_Consolidados, dataframes_consolidado)


    # **Archivo de novedades : Este archivo compila las diferentes novedades presentes en los archivos de TPM**

    # In[46]:


    #Diccionario de hojas para el archivo de novedades
    dataframes_novedades = {
        "Sobrepeso mayor 10%": df_Emp_2_Nov_Sobrepeso,            #Novedades sobrepeso
        "Codigos Nulo TPM": df_Emp_2_Codigos_nan,                 #Novedades codigos nulos
        #"Codigos unicos en TPM": df_Emp_2_Codigos_unicos,        #Codigos presentes en TPM y no en COOISPI
        "Codigos NA,Vacio, 0": df_Emp_2_nan_ceros,                #Valores nulos o ceros para gramaje, unidades producidas y promedio peso unidad
        "Errores": df_non_conver                                  #Valores no convertibles (NaN)
    }
    #Guardar el archivo de Novedades
    fc.crear_archivo_novedades(Ruta_Novedades, dataframes_novedades)


if __name__ == '__main__':
    main_Script_Emp_2_v2()
