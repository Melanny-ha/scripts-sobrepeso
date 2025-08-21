def main_Generar_COOISPI_csv():
    #!/usr/bin/env python
    # coding: utf-8

    # In[1]:


    import os
    import funciones_tpm as fc


    # In[2]:


    def rutas(num_cooispi):
      #Rutas del archivo de la COOISPI
      Ruta_COOISPI_2024 = f"G:/.shortcut-targets-by-id/1lzhI3V8qzGYIWjuRcohkkLwKGTzfbmsY/Consolidado COOISPI/2024/{num_cooispi}.xlsx"                                                            # Ruta COOISPI en 2024
      Ruta_COOISPI_2025 = f"G:/.shortcut-targets-by-id/1lzhI3V8qzGYIWjuRcohkkLwKGTzfbmsY/Consolidado COOISPI/2025/{num_cooispi}.xlsx"                                                            # Ruta COOISPI en 2025

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

    # In[3]:


    def ejecutar_COOISPI(num_arch_cooispi, nombre_arch_salida, almacenes, material=None):
      Ruta_COOISPI_2024, Ruta_COOISPI_2025 = rutas(num_arch_cooispi)

      #leer_archivos_cooispi
      df_COOISPI_2024 = fc.leer_archivo(Ruta_COOISPI_2024, 'Sheet1')
      df_COOISPI_2025 = fc.leer_archivo(Ruta_COOISPI_2025, 'Sheet1')

      #renombrar_cooispi
      df_COOISPI_2024 = fc.renombrar_cooispi(df_COOISPI_2024)
      df_COOISPI_2025 = fc.renombrar_cooispi(df_COOISPI_2025)

      #procesar_cooispi
      df_COOISPI = fc.procesar_cooispi(df_COOISPI_2024, df_COOISPI_2025)

      #estandarizar_mes
      df_COOISPI = fc.estandarizar_mes(df_COOISPI)

      #filtrar_columna, tomar el Almacen para cada informe de TPM
      df_COOISPI = fc.filtrar_columna(df_COOISPI, 'Almacén', almacenes)

      #filtrar_material que comienzan por 8 y eliminarlos
      df_COOISPI = fc.filtrar_material(df_COOISPI, '8')

      df_COOISPI['Cl.mov.']=df_COOISPI['Cl.mov.'].astype('int64')

      #filtrar_columna, tomar las clases con movimiento 101, 102(Producto terminado), 261, 262(Producto semielaborado)
      df_COOISPI_101_102 = fc.filtrar_columna(df_COOISPI, 'Cl.mov.', [101, 102])
      df_COOISPI_261_262 = fc.filtrar_columna(df_COOISPI, 'Cl.mov.', [261, 262])

      #merge_cooispi/merge
      df_COOISPI_COMB = fc.merge_cooispi(df_COOISPI_261_262, df_COOISPI_101_102)

      #Filtrar materiales que comiencen por un número específico (solo si se especificó un valor en 'material') osea para Molido - Bogotá
      if material is not None:
        #Filtrar materiales que comiencen por 3
        df_COOISPI_COMB = df_COOISPI_COMB[df_COOISPI_COMB['Material'].str.startswith(f"{material}")]
        #Asignarle el nombre del centro
        df_COOISPI_COMB['Centro']='CM15'
      else:
        #Asignarle el nombre del centro
        df_COOISPI_COMB['Centro']='CM10'

      #Renombrar columnas para quedar igual al archivo de Excel de Luz Indira
      df_COOISPI_COMB.rename(columns={'Material_producto_term':'Codigo', 'Texto de material_producto_term':'TEXTO PT'},inplace=True)

      #Definir columnas para df_COOISPI_COMB y asignarlas
      Columnas=['Orden', 'Codigo','TEXTO PT','Doc.mat.','Cl.mov.','Material','Texto de material','Almacén','Lote','Centro','Unidad','Ctd.','Impte.ML','Fecha']
      df_COOISPI_COMB=df_COOISPI_COMB[Columnas]

      #Quitar Codigos que no empiezen o sean con un valor numerico
      df_COOISPI_COMB = df_COOISPI_COMB[df_COOISPI_COMB['Codigo'].astype(str).str.isnumeric()]

      #estandarizar la fecha creando la columna Mes y Año a partir de Fecha
      df_COOISPI_COMB = fc.estandarizar_mes(df_COOISPI_COMB)

      #Forzar el cambio de tipo de dato para un merge futuro en scripts principales
      df_COOISPI_COMB['Codigo'] = df_COOISPI_COMB['Codigo'].astype(str)

      #Rutas de destino
      ruta_destino = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\COOISPI'

      df_COOISPI_COMB.to_csv(os.path.join(ruta_destino, f"{nombre_arch_salida}.csv"), index=False)

      return df_COOISPI_COMB

    # In[4]:


    #iniciar_credenciales, si algo falla aquí, el script se detiene
    #gc = iniciar_credenciales()


    #Diccionario de datos de entrada por informe de TPM
    archivos = [
      {"codigo_ruta": "Z10", "nombre_arch": "Datos_COOISPI_Empaques2", "almacenes": [1028, 1001], "material": None},
      {"codigo_ruta": "Z06", "nombre_arch": "Datos_COOISPI_Env_Sol", "almacenes": [1035, 1001], "material": None},
      {"codigo_ruta": "Z09", "nombre_arch": "Datos_COOISPI_Mezclas", "almacenes": [1038, 1001], "material": None},
      {"codigo_ruta": "Z02", "nombre_arch": "Datos_COOISPI_Molido", "almacenes": [1030, 1001], "material": None},
      {"codigo_ruta": "Z02", "nombre_arch": "Datos_COOISPI_Molido_Bog", "almacenes": [1530, 1501], "material": 3},
    ]

    #Ejecucion de todos los informes
    resultados = {}

    for archivo in archivos:
      print(f"Procesando archivo '{archivo['nombre_arch']}'...")
      try:
        df_resultado = ejecutar_COOISPI(
          num_arch_cooispi = archivo["codigo_ruta"],
          nombre_arch_salida = archivo["nombre_arch"],
          almacenes = archivo["almacenes"],
          material = archivo["material"],
        )
        resultados[archivo["nombre_arch"]] = df_resultado
        print(f"✅ Archivo '{archivo['nombre_arch']}' procesado correctamente. Registros: {df_resultado.shape[0]}\n")
      except Exception as e:
        print(f"❌ Error al procesar '{archivo['nombre_arch']}': {e}\n")


if __name__ == '__main__':
    main_Generar_COOISPI_csv()
