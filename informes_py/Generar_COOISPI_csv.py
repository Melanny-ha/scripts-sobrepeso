def main_Generar_COOISPI_csv():
    #!/usr/bin/env python
    # coding: utf-8

    # In[1]:


    import pandas as pd
    import os
    from openpyxl import load_workbook


    # In[2]:


    #Funcion para definir numero del mes segun nombre
    def mes_a_numero():
      return {
      "Enero": 1,
      "Febrero": 2,
      "Marzo": 3,
      "Abril": 4,
      "Mayo": 5,
      "Junio": 6,
      "Julio": 7,
      "Agosto": 8,
      "Septiembre": 9,
      "Octubre": 10,
      "Noviembre": 11,
      "Diciembre": 12
      }


    # In[3]:


    #Funcion para convertir mes de ingles a español
    def mes_eng_esp():
      return {
      "January":"Enero",
      "February":"Febrero",
      "March":"Marzo",
      "April":"Abril",
      "May":"Mayo",
      "June":"Junio",
      "July":"Julio",
      "August":"Agosto",
      "September":"Septiembre",
      "October":"Octubre",
      "November":"Noviembre",
      "December":"Diciembre"
      }


    # In[4]:


    #Funcion para leer los archivos de la COOISPI
    def leer_archivo(ruta, hoja=None):
      try:
        if not os.path.exists(ruta):  #Validar si la ruta existe
          raise FileNotFoundError(f"La ruta '{ruta}' no existe.")

        if ruta.lower().endswith('.csv'):  #Si la ruta es un csv
          df = pd.read_csv(ruta)
          return df
        elif ruta.lower().endswith(('.xlsx', 'xls')):  #Sino, si es un xlsx o xls
          excel = pd.ExcelFile(ruta)

          if hoja not in excel.sheet_names:  #Verificar que la hoja solicitada exista
            raise ValueError(f"La hoja '{hoja}' no existe en el archivo. Hojas disponibles: {excel.sheet_names}")

          df = pd.read_excel(excel, sheet_name=hoja)  #Leer directamente la hoja
          return df
        else:
          raise ValueError("Formato no soportado. Solo se permiten archivos CSV y Excel.")
      except Exception as e:
        raise RuntimeError(f"Error al leer archivo: {e}")


    # In[5]:


    #Funcion para renombrar las columnas de la COOISPI
    def renombrar_cooispi(df):
      if df.empty:
        print("El DataFrame está vacío.")
        return df

      #Definir posibles diccionarios de renombrado
      opcion_1 = {
        'Unnamed: 0': 'Orden',
        'Unnamed: 1': 'Material',
        'Unnamed: 2': 'Texto de material',
        'Unnamed: 3': 'Ctd.',
        'Unnamed: 4': 'Lote',
        'Unnamed: 5': 'Almacén',
        'Unnamed: 6': 'Cl.mov.',
        'Unnamed: 7': 'Unidad',
        'Unnamed: 8': 'Doc.mat.',
        'Unnamed: 9': 'Pto.desc.',
        'Unnamed: 10': 'Fecha',
        'Unnamed: 11': 'ctd.UME',
        'Unnamed: 12': 'Impte.ML'
      }

      opcion_2 = {
        'Orden': 'Orden',
        'Material': 'Material',
        'Texto de material': 'Texto de material',
        'Ctd.en UM base y signo +/- (MEINS)': 'Ctd.',
        'Lote': 'Lote',
        'Almacén': 'Almacén',
        'Clase de movimiento': 'Cl.mov.',
        'Unidad medida base (=MEINS)': 'Unidad',
        'Documento material': 'Doc.mat.',
        'Puesto descarga': 'Pto.desc.',
        'Fe.contabilización': 'Fecha',
        'Ctd.en UM entrada (/CWM/ERFME)': 'ctd.UME',
        'Importe ML (WAERS)': 'Impte.ML'
      }

      df_COOISPI = df.copy(deep=True) #Crear una copia completa de todo df

      #Identificar las columnas actuales del DataFrame
      columnas_actuales = df_COOISPI.columns

      #Seleccionar el diccionario de renombrado adecuado
      if set(opcion_1.keys()).issubset(columnas_actuales):
        df_COOISPI.rename(columns=opcion_1, inplace=True)
      elif set(opcion_2.keys()).issubset(columnas_actuales):
        df_COOISPI.rename(columns=opcion_2, inplace=True)
      else:
        print("No se encontraron columnas coincidentes para renombrar.")

      return df_COOISPI


    # In[6]:


    #Funcion para normalizar dataframes
    def normalizar_dataframe(df):
      if df.empty:
        print("El DataFrame está vacío. No es posible normalizar.")
        return df
      for col in df.columns:
        #Convertir a str y eliminar espacios en columnas de tipo 'object'/strings
        if df[col].dtype == 'object':
          df[col] = df[col].astype(str).str.strip()
      return df


    # In[7]:


    #Funcion para limpiar y unir dos DataFrames
    def procesar_cooispi(df1, df2):
      if df1.empty or df2.empty:
        raise ValueError(f"Uno o ambos DataFrame están vacíos. Se requieren ambos para el procesamiento.")

      #Cambia los tipos de datos de las columnas de 2025 para que coincidan con los de 2024
      #dtype= muestra tipo de dato, astype=metodo para convertir tipo de dato
      df2 = df2.astype(df1.dtypes)

      # Aplicar normalización(limpieza de columnas) a ambos DataFrames
      df1 = normalizar_dataframe(df1)
      df2 = normalizar_dataframe(df2)

      if list(df1.columns) != list(df2.columns):
        raise ValueError(f"Las columnas de {df1} y {df2} no coinciden. No se puede continuar.")

      #Une ambos DataFrames verticalmente
      df_COOISPI = pd.concat([df1, df2], ignore_index=True)

      #Crea una copia completa de df_COOISPI
      df_COOISPI = df_COOISPI.copy(deep=True)

      return df_COOISPI


    # In[8]:


    #Funcion para estandarizar meses en ingles
    def estandarizar_mes(df):
      if 'Fecha' not in df.columns:
        print(f"La columna 'Fecha' no existe en el DataFrame.")
        return df
      df = df.copy()
      df['Mes:'] = df['Fecha'].dt.strftime('%B').str.capitalize().map(mes_eng_esp())
      df['Año:'] = df['Fecha'].dt.year.astype(int)
      return df


    # In[9]:


    #Funcion para filtrar columna por valor especifico
    def filtrar_columna(df, columna, valores):
      if columna not in df.columns:
        print(f"La columna '{columna}' no existe en el DataFrame.")
        return df
      return df[df[columna].isin(valores)]


    # In[10]:


    #Funcion para quitar los Materiales que empiezan en 8
    def filtrar_material(df, prefijo):
      if 'Material' not in df.columns:
        print("La columna 'Material' no existe en el DataFrame.")
        return df
      return df[~df['Material'].astype(str).str.startswith(str(prefijo))]


    # In[11]:


    #Funcion para realizar un "left join" entre df_COOISPI_261_262 y df_COOISPI_101_102
    def merge_cooispi(df2, df1):
      df_COOISPI_COMB = pd.merge(
        df2,                                              #DataFrame base
        df1[['Orden', 'Material', 'Texto de material']],  #DataFrame a unir
        on='Orden',                                       #Llave de unión
        how='left',                                       #Se quedan todas las filas de la izquierda(df2) y solo las coincidentes de la derecha(df1)
        suffixes=('', '_producto_term')                   #Si hay columnas con el mismo nombre, las del segundo DataFrame tendrán el sufijo _producto_term
      )
      #Eliminar filas duplicadas
      df_COOISPI_COMB.drop_duplicates(inplace=True)
      return df_COOISPI_COMB


    # In[12]:


    def rutas(num_cooispi):
      #Rutas del archivo de la COOISPI
      Ruta_COOISPI_2024 = f"/content/drive/MyDrive/Consolidado COOISPI/2024/{num_cooispi}.xlsx"                                                            # Ruta COOISPI en 2024
      Ruta_COOISPI_2025 = f"/content/drive/MyDrive/Consolidado COOISPI/2025/{num_cooispi}.xlsx"                                                            # Ruta COOISPI en 2025

      #Validacion de existencia de las rutas
      for nombre, ruta in {
        "Ruta_COOISPI_2024": Ruta_COOISPI_2024,
        "Ruta_COOISPI_2025": Ruta_COOISPI_2025,
      }.items():
        if not os.path.exists(ruta):
          print(f"La ruta '{nombre}' no fue encontrada.")
        else:
          print(f"Ruta '{nombre}' encontrada.")
      return Ruta_COOISPI_2024, Ruta_COOISPI_2025


    # **------------------------------------------------------------------------------------------------------------------------------------------------------------**

    # In[13]:


    def ejecutar_COOISPI(num_arch_cooispi, nombre_arch_salida, almacenes):
      Ruta_COOISPI_2024, Ruta_COOISPI_2025 = rutas(num_arch_cooispi)

      #leer_archivos_cooispi
      df_COOISPI_2024 = leer_archivo(Ruta_COOISPI_2024, 'Sheet1')
      df_COOISPI_2025 = leer_archivo(Ruta_COOISPI_2025, 'Sheet1')

      #renombrar_cooispi
      df_COOISPI_2024 = renombrar_cooispi(df_COOISPI_2024)
      df_COOISPI_2025 = renombrar_cooispi(df_COOISPI_2025)

      #procesar_cooispi
      df_COOISPI = procesar_cooispi(df_COOISPI_2024, df_COOISPI_2025)

      #estandarizar_mes
      df_COOISPI = estandarizar_mes(df_COOISPI)

      #filtrar_columna, tomar el Almacen para cada informe de TPM
      df_COOISPI = filtrar_columna(df_COOISPI, 'Almacén', almacenes)

      #filtrar_material que comienzan por 8 y eliminarlos
      df_COOISPI = filtrar_material(df_COOISPI, '8')

      df_COOISPI['Cl.mov.']=df_COOISPI['Cl.mov.'].astype('int64')

      #filtrar_columna, tomar las clases con movimiento 101, 102(Producto terminado), 261, 262(Producto semielaborado)
      df_COOISPI_101_102 = filtrar_columna(df_COOISPI, 'Cl.mov.', [101, 102])
      df_COOISPI_261_262 = filtrar_columna(df_COOISPI, 'Cl.mov.', [261, 262])

      #merge_cooispi/merge
      df_COOISPI_COMB = merge_cooispi(df_COOISPI_261_262, df_COOISPI_101_102)

      #Renombrar columnas para quedar igual al archivo de Excel de Luz Indira
      df_COOISPI_COMB.rename(columns={'Material_producto_term':'Codigo', 'Texto de material_producto_term':'TEXTO PT'},inplace=True)

      #Asignarle el nombre del centro
      df_COOISPI_COMB['Centro']='CM10'

      #Definir columnas para df_COOISPI_COMB y asignarlas
      Columnas=['Orden', 'Codigo','TEXTO PT','Doc.mat.','Cl.mov.','Material','Texto de material','Almacén','Lote','Centro','Unidad','Ctd.','Impte.ML','Fecha']
      df_COOISPI_COMB=df_COOISPI_COMB[Columnas]

      #Quitar Codigos que no empiezen o sean con un valor numerico
      df_COOISPI_COMB = df_COOISPI_COMB[df_COOISPI_COMB['Codigo'].astype(str).str.isnumeric()]

      #estandarizar la fecha creando la columna Mes y Año a partir de Fecha
      df_COOISPI_COMB = estandarizar_mes(df_COOISPI_COMB)

      #Forzar el cambio de tipo de dato para un merge futuro en scripts principales
      df_COOISPI_COMB['Codigo'] = df_COOISPI_COMB['Codigo'].astype(str)

      return df_COOISPI_COMB.to_csv(f"/content/drive/MyDrive/SIRI_2024_INFORMES_SOBREPESO/DATOS/{nombre_arch_salida}.csv", index=False)


    # In[14]:


    #iniciar_credenciales, si algo falla aquí, el script se detiene
    #gc = iniciar_credenciales()

    #Diccionario de datos de entrada por informe de TPM
    archivos = [
      {"codigo_ruta": "Z10", "nombre_arch": "Datos_COOISPI_Empaques2", "almacenes": [1028, 1001]},
      {"codigo_ruta": "Z06", "nombre_arch": "Datos_COOISPI_Env_Sol", "almacenes": [1035, 1001]},
      {"codigo_ruta": "Z09", "nombre_arch": "Datos_COOISPI_Mezclas", "almacenes": [1038, 1001]},
      {"codigo_ruta": "Z02", "nombre_arch": "Datos_COOISPI_Molido", "almacenes": [1030, 1001]},
    ]

    #Ejecucion de todos los informes
    for archivo in archivos:
      print(f"Procesando archivo '{archivo['nombre_arch']}'...")
      try:
        ejecutar_COOISPI(
          num_arch_cooispi = archivo["codigo_ruta"],
          nombre_arch_salida = archivo["nombre_arch"],
          almacenes = archivo["almacenes"]
        )
        print(f"✅ Archivo '{archivo['nombre_arch']}' procesado correctamente.\n")
      except Exception as e:
        print(f"Error al procesar '{archivo['nombre_arch']}': {e}\n")


if __name__ == '__main__':
    main_Generar_COOISPI_csv()
