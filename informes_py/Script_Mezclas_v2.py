def main_Script_Mezclas_v2():
    #!/usr/bin/env python
    # coding: utf-8

    # In[1]:


    import os
    import pandas as pd
    import warnings
    import functions as fc


    # In[2]:


    #Ocultar advertencias/warnings mas no los borra
    warnings.filterwarnings("ignore", category=pd.errors.DtypeWarning)


    # In[3]:


    #Año evaluado
    Añoeval = 2025


    # **Variables globales de las rutas necesarias**

    # In[4]:


    #Rutas del archivo de la COOISPI
    Ruta_COOISPI = "G:/.shortcut-targets-by-id/1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN/SIRI_2024_INFORMES_SOBREPESO/DATOS/Datos_COOISPI_Mezclas.csv"                        # Ruta COOISPI en csv
    #Ruta_COOISPI_2024="/content/drive/MyDrive/Consolidado COOISPI/2024/Z09.xlsx"                                                                                      # Ruta COOISPI en el 2024
    #Ruta_COOISPI_2025="/content/drive/MyDrive/Consolidado COOISPI/2025/Z09.xlsx"  

    #Rutas del archivo del EGE
    Ruta_Archivo_EGE = "G:/.shortcut-targets-by-id/1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN/SIRI_2024_INFORMES_SOBREPESO/DATOS/Datos_Mezclas_2021.csv"
    #Ruta_Archivo_EGE = "/content/drive/MyDrive/SIRI_2024_INFORMES_SOBREPESO/DATOS/Datos_Mezclas_2021.csv"                                                             # En formato csv en el drive de Melanny Herrera
    #Ruta_Archivo_EGE = "/content/drive/MyDrive/appsheet/Coolab/Documentos/Mezclas/DB Empaque de Mezclas 2021_Copia.xlsm"                                              # Ruta EGE pruebas Melanny Herrera
    #Ruta_Archivo_EGE = "/content/drive/MyDrive/EGE Colcafé Medellín 2018 - 2024/Data historica/2021-2024/DB Empaque de Mezclas 2021.xlsm"                             # Ruta EGE en el drive de Jesse Mauricio Beltran

    #Rutas del archivo de Novedades
    Ruta_Novedades = "G:/.shortcut-targets-by-id/1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Mezclas_Novedades_V2.xlsx"
    #Ruta_Novedades = "/content/drive/MyDrive/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Mezclas_Novedades_V2.xlsx"                                         # Ruta del archivo de Novedades en el drive de Melanny Herrera
    #Ruta_Novedades = '/content/drive/MyDrive/Proyectos/2024/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Mezclas_Novedades.xlsx'                             # Ruta del archivo de Novedades en el drive de Jesse Mauricio Beltran

    #Rutas del archivo de Consolidados
    Ruta_Consolidados = "G:/.shortcut-targets-by-id/1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Consolidado_V2.xlsx"
    #Ruta_Consolidados ="/content/drive/MyDrive/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Consolidado_V2.xlsx"                                             # Ruta del archivo de Consolidados en el drive de Melanny Herrera
    #Ruta_Consolidados ='/content/drive/MyDrive/Proyectos/2024/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/Consolidado.xlsx'                                 # Ruta del archivo de Novedades en el drive de Jesse Mauricio Beltran

    #Ruta de salida
    Ruta_Mezclas = "G:/.shortcut-targets-by-id/1rqpfbdZ6z51epFv6ZwognhckW7HqBMjN/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/"
    #Ruta_Mezclas = "/content/drive/MyDrive/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/"                                                                    # Ruta del archivo consolidado csv en el drive de Melanny Herrera
    #Ruta_Mezclas = "/content/drive/MyDrive/Proyectos/2024/SIRI_2024_INFORMES_SOBREPESO/INFORME_MEZCLAS_MEDELLIN/"                                                     # Ruta del archivo consolidado csv en el drive de Jesse Mauricio Beltran

    #Validacion de existencia de las rutas
    for nombre, ruta in {
      "Ruta_COOISPI": Ruta_COOISPI,
      "Ruta_Archivo_EGE": Ruta_Archivo_EGE,
      "Ruta_Novedades": Ruta_Novedades,
      "Ruta_Consolidados": Ruta_Consolidados,
      "Ruta_Mezclas": Ruta_Mezclas,
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
    df_Mezclas = fc.leer_archivo(Ruta_Archivo_EGE, 'Datos')


    # In[8]:


    #Reemplazar el nombre de la columna "Día:" y "Mes_N" para estandarizar los scripts
    df_Mezclas.rename(columns={"Día:":"Dia", "Mes_N":"IdMes"}, inplace=True)


    # In[9]:


    #Eliminar_filas_vacias
    numeros_vacios_filtrados = fc.eliminar_filas_vacias(df_Mezclas, 'Número')


    # In[10]:


    #Crear_fecha
    df_Mezclas = fc.crear_fecha(df_Mezclas)


    # In[11]:


    #Tomaré solo las columnas de interes del informe de mezclas para el sobrepeso
    columnas=['Número','Dia','Fecha','Mes:', 'IdMes', 'Año:', 'Máquina / Equipo:', 'Semana:', 'Turno:', 'Código y Descrip. / Producto:', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):','Gramaje (K):']
    df_Mezclas = df_Mezclas[columnas]


    # In[12]:


    #Filtrar_desde_anio
    df_Mezclas = fc.filtrar_desde_anio(df_Mezclas, 2024)


    # In[13]:


    #Especificar las columnas en las que deseas aplicar el filtro para eliminar los valores nulos de las columnas de Gramaje, Unidades producidas y Peso promedio
    columnas_filtrar_nulos = ['Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):']


    # In[14]:


    #Eliminar_valores_nulos
    df_Mezclas_nan_ceros, df_Mezclas = fc.filtrar_nulos(df_Mezclas, columnas_filtrar_nulos)


    # **Modificación de los tipos de datos del archivo de TPM de Mezclas**

    # In[15]:


    #Validar_y_convertir_datos
    df_non_conver, df_Mezclas = fc.validar_numericos(df_Mezclas, 'Peso Promedio de la unidad (K):')


    # In[16]:


    #Extraer_codigo
    df_Mezclas = fc.extraer_codigo(df_Mezclas)


    # In[17]:


    #Convertir a numerico sino asignar NaN
    df_Mezclas['Gramaje (K):'] = pd.to_numeric(df_Mezclas['Gramaje (K):'], errors='coerce')


    # In[18]:


    #Eliminar_codigos_nan
    df_Mezclas, df_Mezclas_Codigos_nan = fc.eliminar_codigos_nan(df_Mezclas)


    # In[19]:


    #Calcular_columnas
    df_Mezclas = fc.calcular_columnas(df_Mezclas)


    # **Generación de Novedades**

    # In[20]:


    Sobrepeso_Novedades = 0.05 #5% de Sobrepeso


    # In[21]:


    #Generar_novedades
    df_Mezclas_Nov_Sobrepeso = fc.generar_novedades(df_Mezclas, Sobrepeso_Novedades)


    # In[22]:


    #Agregar columna 'Costo/kg'
    df_Mezclas['Costo/kg'] = 0.0


    # **Tabla de Consolidado Mensual**

    # In[23]:


    #Pivote_consolidado
    df_Mezclas_Mes = fc.pivote_consolidado(df_Mezclas)


    # In[24]:


    #Tomar datos desde el año 2024 en adelante
    df_Mezclas_Mes = fc.filtrar_desde_anio(df_Mezclas_Mes, 2024)


    # **Creación de la agrupación Mensual del archivo de semielaborados**

    # In[25]:


    df_costo_semi = df_COOISPI_COMB.copy(deep=True)


    # In[26]:


    #Creacion de la agrupacion mensual del archivo semielaborado
    df_costo_semi_Mes = fc.pivote_semielaborados(df_costo_semi)


    # **Generación del nuevo archivo consolidado V2**

    # In[27]:


    df_Mezclas_2 = df_Mezclas.copy(deep=True)


    # In[28]:


    #Traer la columna 'Costo/kg' al dataframe original
    df_Mezclas_2 = fc.merge_costo(df_Mezclas_2, df_costo_semi_Mes)


    # In[29]:


    #Eliminar la columna Costo/kg antigua y renombrar la nueva
    df_Mezclas_2 = fc.modificar_costo(df_Mezclas_2)


    # In[30]:


    #Definir meta
    df_Mezclas_2['Meta'] = df_Mezclas_2['Máquina / Equipo:'].map({'ENFLEX': 0.0056,
                                                                  'TRANSPACK 2': 0.0033,
                                                                  'INGEPACK': 0.0072,
                                                                  'ROVEMA':0.0078,
                                                                  'LANIC':0.0055,
                                                                  'TOYO 7':0.0065,
                                                                  'TOYO':0.0088,
                                                                  'ENCAPSULADORA':0.006,
                                                                  'INGEPACK 2':0.0068})


    # In[31]:


    #Calcular columna 'Ahorros/Perdidas'
    df_Mezclas_2 = fc.calcular_ahorros_perdidas(df_Mezclas_2)
    # df_Mezclas_2.head()


    # In[32]:


    #Eliminar registro con Número = 81187
    df_Mezclas_2.drop(df_Mezclas_2.loc[df_Mezclas_2['Número'] == 81187].index, inplace=True)


    # **Unión de los dataframes en un dataframe totalizado**

    # In[33]:


    df_totalizado_Mes = df_Mezclas_Mes.copy(deep=True)


    # In[34]:


    #Validar que no hay filas repetidas en df_totalizado_Mes
    fc.filas_repetidas(df_totalizado_Mes)


    # In[35]:


    #Fusionar los DataFrames en base a 'Año:', 'Codigo' y 'Mes:'
    df_totalizado_Mes = fc.unir_dataframes(df_totalizado_Mes, df_costo_semi_Mes)


    # In[36]:


    #Metodo diferente a los demas calcular_columnas_totalizado ya que sus metas de sobrepeso no son fijas
    df_totalizado_Mes = fc.calcular_columnas_totalizado_mezclas(df_totalizado_Mes)
    # df_totalizado_Mes.head()


    # In[37]:


    # Obtener Codigos unicos de df_Mezclas
    df_Codigos_Mezclas = fc.eliminar_duplicados_columna(df_Mezclas, 'Codigo')


    # In[38]:


    # Obtener Maquinas únicas de TPM o df_Codigos_Mezclas
    df_Maquinas_Mezclas = fc.eliminar_duplicados_columna(df_Mezclas, 'Máquina / Equipo:')


    # In[39]:


    # Obtener Codigos únicos de COOISPI
    df_Codigos_Mezclas_COOISPI = fc.eliminar_duplicados_columna(df_COOISPI_COMB, 'Codigo')


    # **Codigos que están en TPM pero no están en la COOISPI**

    # In[40]:


    #Obtener Codigos que están en TPM (df_Codigos_Mezclas) pero NO en COOISPI(df_Codigos_Mezclas_COOISPI)
    df_Mezclas_Codigos_unicos = df_Codigos_Mezclas[~df_Codigos_Mezclas['Codigo'].isin(df_Codigos_Mezclas_COOISPI['Codigo'])]


    # **Generación de un archivo de fechas Automáticas**

    # In[41]:


    #Generar dataframe con fechas dese 2024-01-01 hasta el dia actual
    df_Fechas = fc.generar_fechas("2024-01-01")


    # In[42]:


    # print(df_Mezclas.shape)
    # print(df_COOISPI_COMB.shape)
    # print(df_Mezclas_Mes.shape)
    # print(df_Codigos_Mezclas.shape)
    # print(df_Maquinas_Mezclas.shape)
    # print(df_totalizado_Mes.shape)
    # print(df_Fechas.shape)
    # print(df_Mezclas_2.shape)
    # print()
    # print(df_Mezclas_Nov_Sobrepeso.shape)
    # print(df_Mezclas_Codigos_nan.shape)
    # print(df_Mezclas_Codigos_unicos.shape)
    # print(df_Mezclas_nan_ceros.shape)
    # print(df_non_conver.shape)


    # **Consolidado csv**

    # In[43]:


    #Ruta para consolidado df_Mezclas_2
    df_Mezclas_2.to_csv(Ruta_Mezclas + "DB_Mezclas_Out_vf.csv", index=False)


    # **Archivo de consolidado**

    # In[44]:


    #Diccionario de hojas para el archivo de Consolidado_V2
    dataframes_consolidado = {
        "Hoja1": df_Mezclas,
        "Hoja2": df_COOISPI_COMB,
        "Hoja3": df_Mezclas_Mes,
        "Hoja4": df_Codigos_Mezclas,
        "Hoja5": df_Maquinas_Mezclas,
        "Hoja6": df_totalizado_Mes,
        "Hoja7": df_Fechas,
        "Hoja8": df_Mezclas_2,
    }
    #Actualizar hojas del excel de Consolidado_V2
    fc.actualizar_hojas_excel(Ruta_Consolidados, dataframes_consolidado)


    # **Archivo de novedades: Este archivo compila las diferentes novedades presentes en los archivos de TPM**

    # In[45]:


    #Diccionario de hojas para el archivo de novedades
    dataframes_novedades = {
        "Sobrepeso mayor  5%": df_Mezclas_Nov_Sobrepeso,
        "Codigos Nulo TPM": df_Mezclas_Codigos_nan,
        "Codigos unicos en TPM": df_Mezclas_Codigos_unicos,
        "Codigos NA,Vacio, 0": df_Mezclas_nan_ceros,
        "Errores": df_non_conver
    }
    #Guardar el archivo de Novedades
    fc.crear_archivo_novedades(Ruta_Novedades, dataframes_novedades)


if __name__ == '__main__':
    main_Script_Mezclas_v2()
