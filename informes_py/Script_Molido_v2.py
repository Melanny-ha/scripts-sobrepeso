def main_Script_Molido_v2():
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
    Ruta_COOISPI = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\COOISPI\Datos_COOISPI_Molido.csv'                         # Ruta COOISPI en csv

    #Rutas del archivo del EGE
    Ruta_Archivo_EGE = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Backup_BD_Salones\Datos_Molido_2021.csv'              # Ruta Molido

    #Rutas del archivo de Novedades
    Ruta_Novedades = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Molido\Molido_Novedades_V2.xlsx'   # Ruta del archivo de Novedades

    #Rutas del archivo de Consolidados
    Ruta_Consolidados = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Molido\Consolidado_V2.xlsx'     # Ruta del archivo de Consolidados

    # Ruta de salida
    # Ruta_Mol = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Molido'                                # DB_Molido_Out_vf.csv

    #Validacion de existencia de las rutas
    for nombre, ruta in {
      "Ruta_COOISPI": Ruta_COOISPI,
      "Ruta_Archivo_EGE": Ruta_Archivo_EGE,
      "Ruta_Novedades": Ruta_Novedades,
      "Ruta_Consolidados": Ruta_Consolidados,
      # "Ruta_Mol": Ruta_Mol,
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
    df_Mol = fc.leer_archivo(Ruta_Archivo_EGE, 'Datos')


    # In[8]:


    #Reemplazar el nombre de la columna "Día:" para estandarizar los scripts
    df_Mol.rename(columns={"Día:":"Dia"}, inplace=True)


    # In[9]:


    #Eliminar_filas_vacias
    numeros_vacios_filtrados = fc.eliminar_filas_vacias(df_Mol, 'Número')


    # In[10]:


    #crear_fecha
    df_Mol = fc.crear_fecha(df_Mol)


    # In[11]:


    #Tomar solo las columnas de interes del informe de Molido para el sobrepeso
    columnas=['Número','Dia','Fecha','Mes:','IdMes','Año:','Máquina / Equipo:','Semana:','Turno:','Código y Descrip. / Producto:','Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):','Gramaje (K):']
    df_Mol = df_Mol[columnas]


    # In[12]:


    #Filtrar_desde_anio
    df_Mol = fc.filtrar_desde_anio(df_Mol, 2024)


    # In[13]:


    #Especificar las columnas en las que deseas aplicar el filtro para eliminar los valores nulos de las columnas de Gramaje, Unidades producidas y Peso promedio
    columnas_filtrar_nulos = ['Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):']


    # In[14]:


    #Filtrar_nulos
    df_Mol_nan_ceros, df_Mol = fc.filtrar_nulos(df_Mol, columnas_filtrar_nulos)
    df_Mol_nan_ceros = df_Mol_nan_ceros[df_Mol_nan_ceros['Año:']==Añoeval] # Buscando solo los del 2025 en adelante


    # **Modificación de los tipos de datos del archivo de Molido de TPM**

    # In[15]:


    #Validar_numericos
    df_non_conver, df_Mol = fc.validar_numericos(df_Mol, 'Peso Promedio de la unidad (K):')


    # In[16]:


    #Extraer_codigo
    df_Mol = fc.extraer_codigo(df_Mol)


    # **Identificación y supresión de aquellos códigos con valores nulos**

    # In[17]:


    #Eliminar_codigos_nan
    df_Mol, df_Mol_Codigos_nan = fc.eliminar_codigos_nan(df_Mol)
    df_Mol_Codigos_nan = df_Mol_Codigos_nan[df_Mol_Codigos_nan['Año:'] == Añoeval]


    # In[18]:


    #Calcular_columnas
    df_Mol = fc.calcular_columnas(df_Mol)


    # **Generación de Novedades**

    # In[19]:


    Sobrepeso_Novedades = 0.20  #20% de Sobrepeso


    # In[20]:


    #Generar_novedades
    df_Mol_Nov_Sobrepeso = fc.generar_novedades(df_Mol, Sobrepeso_Novedades)


    # In[21]:


    #Agregar columna 'Costo/kg'
    df_Mol['Costo/kg'] = 0.0


    # **Tabla de Consolidado Mensual**

    # In[22]:


    #Pivote_consolidado
    df_Mol_Mes = fc.pivote_consolidado(df_Mol)


    # In[23]:


    #Tomar datos desde el año 2024 en adelante
    df_Mol_Mes = fc.filtrar_desde_anio(df_Mol_Mes, 2024)


    # **Creación de la agrupación Mensual del archivo de semielaborados**

    # In[24]:


    df_costo_semi = df_COOISPI_COMB.copy(deep=True)


    # In[25]:


    #Creacion de la agrupacion mensual del archivo semielaborado
    df_costo_semi_Mes = fc.pivote_semielaborados(df_costo_semi)


    # **Generación del nuevo archivo consolidado v2**

    # In[26]:


    df_Mol_2 = df_Mol.copy(deep=True)


    # In[27]:


    #Traer la columna 'Costo/kg' al dataframe original
    df_Mol_2 = fc.merge_costo(df_Mol_2, df_costo_semi_Mes)


    # In[28]:


    #Eliminar la columna Costo/kg antigua y renombrar la nueva
    df_Mol_2 = fc.modificar_costo(df_Mol_2)


    # In[29]:


    #Definir meta
    df_Mol_2['Meta'] = 0.005


    # In[30]:


    #Calcular columna 'Ahorros/Perdidas'
    df_Mol_2 = fc.calcular_ahorros_perdidas(df_Mol_2)
    # df_Mol_2.head()


    # **Unión de los DataFrames en un DataFrame totalizado**

    # In[31]:


    df_totalizado_Mes = df_Mol_Mes.copy(deep=True)


    # In[32]:


    #Validar que no hay filas repetidas en df_totalizado_Mes
    fc.filas_repetidas(df_totalizado_Mes)


    # In[33]:


    #Fusionar los DataFrames en base a 'Año:', 'Codigo' y 'Mes:'
    df_totalizado_Mes = fc.unir_dataframes(df_totalizado_Mes, df_costo_semi_Mes)


    # In[34]:


    #Metodo calcular_columnas_totalizado con meta fija
    df_totalizado_Mes = fc.calcular_columnas_totalizado(df_totalizado_Mes, 0.005)
    # df_totalizado_Mes.head()


    # In[35]:


    #Obtener Codigos unicos | Cod. Descrip. | Maquina con un merge
    df_Codigos_Mol = fc.obtener_cod_descrip(df_costo_semi, df_Mol_Mes)


    # In[36]:


    #Eliminar_duplicados_columna
    df_Maquinas_Mol = fc.eliminar_duplicados_columna(df_Codigos_Mol, 'Máquina / Equipo:')


    # **Generación de un archivo de fechas Automáticas**

    # In[37]:


    #Generar dataframe con fechas dese 2024-01-01 hasta el dia actual
    df_Fechas = fc.generar_fechas("2024-01-01")


    # In[38]:


    # print(df_Mol.shape)
    # print(df_COOISPI_COMB.shape)
    # print(df_Mol_Mes.shape)
    # print(df_Codigos_Mol.shape)
    # print(df_Maquinas_Mol.shape)
    # print(df_totalizado_Mes.shape)
    # print(df_Fechas.shape)
    # print(df_Mol_2.shape)
    # print()
    # print(df_Mol_Nov_Sobrepeso.shape)
    # print(df_Mol_Codigos_nan.shape)
    # #print(df_Env_Sol_Codigos_unicos.shape)
    # print(df_Mol_nan_ceros.shape)
    # print(df_non_conver.shape)


    # **Consolidado csv**

    # In[39]:


    # Ruta para consolidado df_Mol_2
    # df_Mol_2.to_csv(Ruta_Mol + "DB_Molido_Out_vf.csv", index=True)


    # **Archivo de consolidado**

    # In[40]:


    #Diccionario de hojas para el archivo de consolidado
    dataframes_consolidado = {
        "Hoja1": df_Mol,
        "Hoja2": df_COOISPI_COMB,
        "Hoja3": df_Mol_Mes,
        "Hoja4": df_Codigos_Mol,         #df_merged
        "Hoja5": df_Maquinas_Mol,
        "Hoja6": df_totalizado_Mes,
        "Hoja7": df_Fechas,
        "Hoja8": df_Mol_2,
    }
    #Actualizar hojas del excel de Consolidado_V2
    fc.actualizar_hojas_excel(Ruta_Consolidados, dataframes_consolidado)


    # **Archivo de novedades: este archivo compila las diferentes novedades presentes en los archivos de TPM**

    # In[41]:


    #Diccionario de hojas para el archivo de novedades
    dataframes_novedades = {
        "Sobrepeso mayor  20%": df_Mol_Nov_Sobrepeso,           #Novedades sobrepeso
        "Codigos Nulo TPM": df_Mol_Codigos_nan,                 #Novedades codigos nulos
        #"Codigos unicos en TPM": df_Mol_Codigos_unicos,        #Codigos presentes en TPM y no en COOISPI
        "Codigos NA,Vacio, 0": df_Mol_nan_ceros,                #Valores nulos o ceros para gramaje, unidades producidas y promedio peso unidad
        "Errores": df_non_conver                                #Valores no convertibles
    }
    #Guardar el archivo de Novedades
    fc.crear_archivo_novedades(Ruta_Novedades, dataframes_novedades)


if __name__ == '__main__':
    main_Script_Molido_v2()
