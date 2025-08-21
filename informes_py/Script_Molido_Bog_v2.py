def main_Script_Molido_Bog_v2():
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
    Ruta_COOISPI = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\COOISPI\Datos_COOISPI_Molido_Bog.csv'

    #Rutas del archivo del EGE
    df_Sobrepeso_2024 = r"G:\.shortcut-targets-by-id\1bHqTZkqHanfKuGXtnbSPppx42ejuve5I\SOBREPESO BOGOTÁ\Informe control de proceso 2024.xlsx"
    # df_Sobrepeso_2024_S1 = "/content/drive/MyDrive/SOBREPESO BOGOTÁ/Informe control de proceso Primer Semestre 2024.xlsx"
    # df_Sobrepeso_2024_S2 = "/content/drive/MyDrive/SOBREPESO BOGOTÁ/Informe control de proceso Segundo Semestre 2024.xlsm"

    df_Sobrepeso_2025 = r"G:\.shortcut-targets-by-id\1bHqTZkqHanfKuGXtnbSPppx42ejuve5I\SOBREPESO BOGOTÁ\Informes control de proceso 2025.xlsm"

    #Rutas del archivo de Novedades
    Ruta_Novedades = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Molido_Bog\Molido_Bog_Novedades_V2.xlsx'   # Ruta del archivo de Novedades

    #Rutas del archivo de Consolidados
    Ruta_Consolidados = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Consolidados_Salones\Molido_Bog\Consolidado_V2.xlsx'      # Ruta del archivo de Consolidados


    #Validacion de existencia de las rutas
    for nombre, ruta in {
        "Ruta_COOISPI": Ruta_COOISPI,
        "df_Sobrepeso_2024": df_Sobrepeso_2024,
        "df_Sobrepeso_2025": df_Sobrepeso_2025,
        "Ruta_Novedades": Ruta_Novedades,
        "Ruta_Consolidados": Ruta_Consolidados
    }.items():
      if not os.path.exists(ruta):
        print(f"La ruta '{nombre}' no fue encontrada.")
      else:
        print(f"Ruta '{nombre}' encontrada.")


    # In[5]:


    #leer_archivo COOISPI csv
    df_COOISPI_COMB = fc.leer_archivo(Ruta_COOISPI)

    df_Sobrepeso_2024_VPKS = fc.leer_archivo(df_Sobrepeso_2024, 'VPK')
    df_Sobrepeso_2024_LINEAS = fc.leer_archivo(df_Sobrepeso_2024, 'LINEAS')

    df_Sobrepeso_2025_VPKS = fc.leer_archivo(df_Sobrepeso_2025, 'VPKS')
    df_Sobrepeso_2025_LINEAS = fc.leer_archivo(df_Sobrepeso_2025, 'LINEAS')


    # **Validación de las columnas de cada dataframe VPKS**

    # In[6]:


    # df_Sobrepeso_2024_VPKS.columns.to_list()


    # In[7]:


    # df_Sobrepeso_2024_LINEAS.columns.to_list()


    # In[8]:


    #Equipo = Máquina
    #Material = Referencia
    columnas_df_Sobrepeso_2024_VPKS = set(df_Sobrepeso_2024_VPKS.columns)
    columnas_df_Sobrepeso_2025_VPKS = set(df_Sobrepeso_2025_VPKS.columns)

    iguales = columnas_df_Sobrepeso_2024_VPKS & columnas_df_Sobrepeso_2025_VPKS  #Intersección
    diferentes_df1 = columnas_df_Sobrepeso_2024_VPKS - columnas_df_Sobrepeso_2025_VPKS  #En df1 pero no en df2
    diferentes_df2 = columnas_df_Sobrepeso_2025_VPKS - columnas_df_Sobrepeso_2024_VPKS  #En df2 pero no en df1

    # print("Columnas iguales: ", iguales)
    # print("Columnas en df1 pero no en df2: ", diferentes_df1)
    # print("Columnas en df2 pero no es df1: ", diferentes_df2)


    # In[9]:


    df_Sobrepeso_2024_VPKS.rename(columns={'Descripcioón':'Descripción'}, inplace=True)


    # In[10]:


    df_Sobrepeso_2025_VPKS.rename(columns={'Referencia': 'Material',
                                           'Máquina':'Equipo'}, inplace=True)


    # In[11]:


    # df_Sobrepeso_2024_VPKS.columns


    # In[12]:


    # df_Sobrepeso_2025_VPKS.columns


    # In[13]:


    df_Mol_VPKS = pd.concat([df_Sobrepeso_2024_VPKS, df_Sobrepeso_2025_VPKS], ignore_index=True)


    # In[14]:


    # df_Mol_VPKS.shape


    # In[15]:


    # fc.filas_repetidas(df_Mol_VPKS)


    # In[16]:


    # df_Mol_VPKS.columns


    # **---------------------------------------------------------------------------------------------------------------------------------------------------------**

    # In[17]:


    #Equipo = Máquina
    #Material = Referencia
    columnas_df_Sobrepeso_2024_LINEAS = set(df_Sobrepeso_2024_LINEAS.columns)
    columnas_df_Sobrepeso_2025_LINEAS = set(df_Sobrepeso_2025_LINEAS.columns)

    iguales = columnas_df_Sobrepeso_2024_LINEAS & columnas_df_Sobrepeso_2025_LINEAS  #Intersección
    diferentes_df1 = columnas_df_Sobrepeso_2024_LINEAS - columnas_df_Sobrepeso_2025_LINEAS  #En df1 pero no en df2
    diferentes_df2 = columnas_df_Sobrepeso_2025_LINEAS - columnas_df_Sobrepeso_2024_LINEAS  #En df2 pero no en df1

    # print("Columnas iguales: ", iguales)
    # print("Columnas en df1 pero no en df2: ", diferentes_df1)
    # print("Columnas en df2 pero no es df1: ", diferentes_df2)


    # In[18]:


    df_Sobrepeso_2025_LINEAS.rename(columns={'Referencia': 'Material',
                                             'Máquina':'Equipo'}, inplace=True)


    # In[19]:


    # df_Sobrepeso_2024_LINEAS.columns


    # In[20]:


    # df_Sobrepeso_2025_LINEAS.columns


    # In[21]:


    df_Mol_LINEAS = pd.concat([df_Sobrepeso_2024_LINEAS, df_Sobrepeso_2025_LINEAS], ignore_index=True)


    # In[22]:


    # df_Mol_LINEAS.shape #4243


    # In[23]:


    # fc.filas_repetidas(df_Mol_VPKS)


    # In[24]:


    # df_Mol_LINEAS.columns


    # **Lectura de los archivos de la COOISPI- Archivos de COOISPI**

    # In[25]:


    #leer_archivo COOISPI csv
    df_COOISPI_COMB = fc.leer_archivo(Ruta_COOISPI)


    # In[26]:


    #Convertir columna 'Codigo' a str para futuros procesor (merge y tabla pivote)
    df_COOISPI_COMB['Codigo'] = df_COOISPI_COMB['Codigo'].astype(str)


    # **Modificaciones sobre el archivo de VPKS**

    # In[27]:


    # df_Mol_VPKS.dtypes


    # In[28]:


    df_Mol_VPKS.rename(columns={'Mes':'Mes:',
                                'Equipo':'Máquina / Equipo:',
                                'Descripción':'Código y Descrip. / Producto:',
                                'Semana':'Semana:',
                                'Turno ':'Turno:',
                                'Gramaje':'Gramaje (K):',
                                'Media':'Peso Promedio de la unidad (K):',
                                'Material':'Codigo',
                                'Unidades Conformes':'Unidades Producidas (Conformes) :',
                               },inplace=True)


    # In[29]:


    # df_Mol_VPKS.columns


    # In[30]:


    df_Mol_VPKS = fc.estandarizar_mes(df_Mol_VPKS)
    df_Mol_VPKS['Dia'] = df_Mol_VPKS['Fecha'].dt.day
    df_Mol_VPKS['Año:'] = df_Mol_VPKS['Fecha'].dt.year
    df_Mol_VPKS['IdMes'] = df_Mol_VPKS['Mes:'].map(fc.mes_a_numero())

    # df_Mol_VPKS['Mes'] = df_Mol_VPKS['Fecha'].dt.strftime('%B')
    # df_Mol_VPKS['Mes'] = df_Mol_VPKS['Mes'].str.capitalize()
    # df_Mol_VPKS['Mes'] = df_Mol_VPKS['Mes'].map(meses_eng_esp)


    # In[31]:


    df_Mol_VPKS['Número'] = range(1, len(df_Mol_VPKS) + 1)


    # In[32]:


    # df_Mol_VPKS.head(2)


    # **Modificaciones sobre el archivo de LINEAS**

    # In[33]:


    # df_Mol_LINEAS.dtypes


    # In[34]:


    df_Mol_LINEAS.rename(columns={'Mes':'Mes:',
                                  'Equipo':'Máquina / Equipo:',
                                  'Descripción':'Código y Descrip. / Producto:',
                                  'Semana':'Semana:',
                                  'Turno ':'Turno:',
                                  'Gramaje':'Gramaje (K):',
                                  'Media':'Peso Promedio de la unidad (K):',
                                  'Material':'Codigo',
                                  'Unidades Conformes':'Unidades Producidas (Conformes) :',
                               },inplace=True)


    # In[35]:


    df_Mol_LINEAS = fc.estandarizar_mes(df_Mol_LINEAS)
    df_Mol_LINEAS['Dia'] = df_Mol_LINEAS['Fecha'].dt.day
    df_Mol_LINEAS['Año:'] = df_Mol_LINEAS['Fecha'].dt.year
    df_Mol_LINEAS['IdMes'] = df_Mol_LINEAS['Mes:'].map(fc.mes_a_numero())

    # df_Mol_LINEAS['Mes'] = df_Mol_LINEAS['Fecha'].dt.strftime('%B')
    # df_Mol_LINEAS['Mes'] = df_Mol_LINEAS['Mes'].str.capitalize()
    # df_Mol_LINEAS['Mes'] = df_Mol_LINEAS['Mes'].map(meses_eng_esp)


    # In[36]:


    df_Mol_LINEAS['Número'] = range(1000, len(df_Mol_LINEAS) + 1000)


    # In[37]:


    # df_Mol_LINEAS.head(2)


    # In[38]:


    # df_Mol_LINEAS.columns


    # **Unión del los archivos de Lineas y VPKS**

    # In[39]:


    # df_Mol_LINEAS['IdMes'] = df_Mol_LINEAS['Mes:'].map(fc.mes_a_numero())
    # df_Mol_VPKS['IdMes'] = df_Mol_VPKS['Mes:'].map(fc.mes_a_numero())


    # **Generación de DB_Molido_Bog**

    # In[40]:


    #Concaenar backup 2024 y 2025 con VPKS y LINEAS
    df_Mol = pd.concat([df_Mol_LINEAS, df_Mol_VPKS], axis=0, ignore_index=True)


    # In[41]:


    # fc.filas_repetidas(df_Mol)


    # In[42]:


    # df_Mol.shape


    # In[43]:


    #Generar backup 2024 y 2025 con VPKS y LINEAS
    df_Mol.to_csv(r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Backup_BD_Salones/Datos_Molido_Bog_2024.csv')


    # **Generación de Consolidado y novedades**

    # In[44]:


    columnas=['Número','Dia','Fecha','Mes:','IdMes','Año:','Máquina / Equipo:','Semana:','Turno:','Código y Descrip. / Producto:','Unidades Producidas (Conformes) :',
              'Peso Promedio de la unidad (K):','Gramaje (K):']
    df_Mol = df_Mol[columnas]
    # print(df_Mol.columns)
    # print(df_Mol.shape)


    # In[45]:


    df_Mol['Número'] = range(1, len(df_Mol) + 1)


    # In[46]:


    # df_Mol.head(2)


    # **Elimino aquellos valores de "Código y Descrip./Producto" que no estén debidamente diligenciados: Por ejemplo; "Sin datos de peso"**

    # In[47]:


    df_Mol.drop(df_Mol[df_Mol['Código y Descrip. / Producto:'] == "Sin datos de peso"].index, inplace=True)


    # In[48]:


    # df_Mol.dtypes


    # **Modificación de los tipos de datos del archivo de Molido Bogotá de TPM (Garantizar que las variables de Peso promedio por unidad y Gramaje estén correctas)**

    # In[49]:


    columnas_filtrar_nulos = ['Gramaje (K):', 'Unidades Producidas (Conformes) :', 'Peso Promedio de la unidad (K):']


    # In[50]:


    #filtrar_nulos
    df_Mol_nan_ceros, df_Mol = fc.filtrar_nulos(df_Mol, columnas_filtrar_nulos)


    # In[51]:


    # df_Mol_nan_ceros.head()


    # In[52]:


    df_non_conver, df_Mol = fc.validar_numericos(df_Mol, 'Peso Promedio de la unidad (K):')


    # In[53]:


    # df_non_conver.head()


    # In[54]:


    df_Mol = fc.extraer_codigo(df_Mol)


    # In[55]:


    # df_Mol.head(2)


    # In[56]:


    df_Mol['Gramaje (K):'] = pd.to_numeric(df_Mol['Gramaje (K):'], errors='coerce')


    # In[57]:


    # df_Mol.loc[df_Mol['Gramaje (K):'].isna()]


    # **Identificación y supresión de aquellos códigos con valores nulos**

    # In[58]:


    df_Mol, df_Mol_Codigos_nan = fc.eliminar_codigos_nan(df_Mol)


    # In[59]:


    # df_Mol_Codigos_nan.head()


    # In[60]:


    # df_Mol.shape #quedan m-2 registros en los que Unidades Producidas (Conformes) : eran igual a 0.0


    # **Calculo en el archivo Original de TPM de Real empacado, Debe ser, Diferencia y Sobrepeso**

    # In[61]:


    df_Mol = fc.calcular_columnas_bog(df_Mol)


    # In[62]:


    # df_Mol.head(2)


    # In[63]:


    # df_Mol.dtypes


    # In[64]:


    # df_Mol.shape


    # **Generación de novedades: Se filtrarán aquellos valores de sobrepeso calculados mayores al 10%**

    # In[65]:


    Sobrepeso_Novedades = 0.1   #10% de Sobrepeso


    # In[66]:


    df_Mol_Nov_Sobrepeso = fc.generar_novedades(df_Mol, Sobrepeso_Novedades)


    # In[67]:


    # df_Mol_Nov_Sobrepeso.head(5)


    # In[68]:


    #Agregar columna de Costos
    df_Mol['Costo/kg'] = 0.0


    # In[69]:


    # df_Mol.head(2)


    # In[70]:


    # df_Mol.shape


    # **Tabla de Consolidado Mensual**

    # In[71]:


    df_Mol_Mes = fc.pivote_consolidado(df_Mol)


    # In[72]:


    # df_Mol_Mes.head()


    # **Visualización de ejemplo de los datos del consolidado del MES para el año 2024 y el código 1003294**

    # In[73]:


    # df_Mol_Mes.loc[(df_Mol_Mes['Año:'] == 2024) & (df_Mol_Mes['Codigo'] == "1003298")].head()


    # In[74]:


    #Tomar datos desde el año 2024 en adelante
    df_Mol_Mes = fc.filtrar_desde_anio(df_Mol_Mes, 2024)


    # In[75]:


    # #Verificación de los meses en el Dataframe
    # df_Mol_Mes['Mes:'].unique()


    # **Creación de la agrupación Mensual del archivo de semielaborados**

    # In[76]:


    df_costo_semi = df_COOISPI_COMB.copy(deep=True)


    # In[77]:


    # df_costo_semi.head(2)


    # In[78]:


    df_costo_semi_Mes = fc.pivote_semielaborados(df_costo_semi)


    # In[79]:


    # df_costo_semi_Mes.loc[df_costo_semi_Mes['Mes:']=="Agosto"].head()


    # **Generación del nuevo archivo consolidado V2**

    # In[80]:


    df_Mol_2 = df_Mol.copy(deep=True)


    # In[81]:


    # df_Mol_2.head(2)


    # In[82]:


    # df_Mol_2.dtypes


    # In[83]:


    # df_costo_semi_Mes.dtypes


    # In[84]:


    df_Mol_2 = fc.merge_costo(df_Mol_2, df_costo_semi_Mes)


    # In[85]:


    # df_Mol_2.head(2)


    # In[86]:


    df_Mol_2 = fc.modificar_costo(df_Mol_2)


    # In[87]:


    # df_Mol_2.head(2)


    # In[88]:


    #Definir meta sobrepeso
    df_Mol_2['Meta'] = 0.003


    # In[89]:


    # df_Mol_2.head(2)


    # In[90]:


    df_Mol_2 = fc.calcular_ahorros_perdidas(df_Mol_2)


    # In[91]:


    # df_Mol_2.head(2)


    # **Unión de los dataframes en un dataframe totalizado**

    # In[92]:


    df_totalizado_Mes = df_Mol_Mes.copy(deep=True)


    # In[93]:


    # print(df_Mol_Mes.columns)
    # print(df_Mol_Mes.shape)


    # In[94]:


    # print(df_costo_semi_Mes.columns)
    # print(df_totalizado_Mes.shape)


    # In[95]:


    # fc.filas_repetidas(df_totalizado_Mes)


    # In[96]:


    #Merge para traer a df_totalizado_Mes las columnas Ctd. y Impte.ML
    df_totalizado_Mes = fc.unir_dataframes(df_totalizado_Mes, df_costo_semi_Mes)


    # In[97]:


    # df_totalizado_Mes.head(2)


    # In[98]:


    df_totalizado_Mes = fc.calcular_columnas_totalizado_bog(df_totalizado_Mes, 0.003)


    # In[99]:


    # df_totalizado_Mes.loc[df_totalizado_Mes['Mes:']=="Julio"].head()   #se genera duplicidad de columnas


    # **Codigo para generar archivo de busqueda para Molido - Bogotá**

    # In[100]:


    df_Codigos_Mol = fc.eliminar_duplicados_columna(df_Mol, 'Codigo')


    # In[101]:


    # fc.filas_repetidas(df_Codigos_Mol)


    # In[102]:


    df_Maquinas_Mol = fc.eliminar_duplicados_columna(df_Mol, 'Máquina / Equipo:')


    # In[103]:


    # fc.filas_repetidas(df_Maquinas_Mol)


    # In[104]:


    df_Fechas = fc.generar_fechas("2024-01-01")
    # df_Fechas.tail()


    # In[105]:


    # print(df_Mol.shape)
    # print(df_COOISPI_COMB.shape)
    # print(df_Mol_Mes.shape)
    # print(df_Codigos_Mol.shape)
    # print(df_Maquinas_Mol.shape)
    # print(df_totalizado_Mes.shape)
    # print(df_Fechas.shape)
    # print(df_Mol_2.shape)  #-2 que tenian unidades producidas conformes 0.0
    # print()
    # print(df_Mol_Nov_Sobrepeso.shape)
    # print(df_Mol_Codigos_nan.shape)
    # #print(df_Mol_Codigos_unicos.shape)
    # print(df_Mol_nan_ceros.shape)   #2 de unidades producidas conformes 0.0
    # print(df_non_conver.shape)


    # **Archivo de consolidado**

    # In[106]:


    #Diccionario de hojas para el archivo de novedades
    dataframes_consolidado = {
        "Hoja1": df_Mol,
        "Hoja2": df_COOISPI_COMB,
        "Hoja3": df_Mol_Mes,
        "Hoja4": df_Codigos_Mol,
        "Hoja5": df_Maquinas_Mol,
        "Hoja6": df_totalizado_Mes,
        "Hoja7": df_Fechas,
        "Hoja8": df_Mol_2,
    }
    fc.actualizar_hojas_excel(Ruta_Consolidados, dataframes_consolidado)


    # **Archivo de novedades : Este archivo compila las diferentes novedades presentes en los archivos de TPM**

    # In[107]:


    #Diccionario de hojas para el archivo de consolidado
    dataframes_novedades = {
        "Sobrepeso mayor 10%": df_Mol_Nov_Sobrepeso,
        "Codigos Nulo TPM": df_Mol_Codigos_nan,
        # "Hoja3": df_Mol_Codigos_unicos,
        "Codigos NA,Vacio,0": df_Mol_nan_ceros,
        "Errores": df_non_conver,
    }
    fc.crear_archivo_novedades(Ruta_Novedades, dataframes_novedades)


if __name__ == '__main__':
    main_Script_Molido_Bog_v2()
