import pandas as pd
import os


################## --------------Conexión con el IC ----------------------##################3

ruta_destino = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\BD Sobrepeso\Backup_BD_Salones'
os.makedirs(ruta_destino, exist_ok=True)

################## ---------------------Rutas--------------################################3

Ruta_Env_Soluble_2021 = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\Envase Soluble\DB Envase de Soluble 2021.xlsm'
Ruta_Molido_2021 = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\Tostado y Molido\DB Empaque de Molido 2021 MES.xlsm'
Ruta_Mezclas_2021 = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\Mezclas\DB Empaque de Mezclas 2021.xlsm'
Ruta_Emp2_2021 = r'\\10.28.5.232\s3-1colcafeci-servicios-jtc\TPM\Colcafé Formularios\Salon de Empaque 2\DB Empaque de Estrella 2021.xlsm'

#################### -----------------Asignación a dataframes de asignación de la lectura de archivo de EGE

Archivos = {
    "Env_Soluble_2021": Ruta_Env_Soluble_2021,
    "Molido_2021": Ruta_Molido_2021,
    "Mezclas_2021": Ruta_Mezclas_2021,
    "Emp2_2021": Ruta_Emp2_2021
}

#################### -------------Lectura de los Archivos ----------------##############################3

for nombre, ruta in Archivos.items():
    try:
        df = pd.read_excel(ruta, sheet_name='Datos', engine='openpyxl')
        df.to_csv(os.path.join(ruta_destino, f'Datos_{nombre}.csv'), index=False)
    except Exception as e:
        print(f"Error con {nombre}: {e}")

#df_Env_Soluble_2021 = pd.read_excel(Ruta_Env_Soluble_2021,sheet_name='Datos', engine='openpyxl')
#df_Molido_2021 = pd.read_excel(Ruta_Molido_2021,sheet_name='Datos', engine='openpyxl')
#df_Mezclas_2021 = pd.read_excel(Ruta_Mezclas_2021,sheet_name='Datos', engine='openpyxl')
#df_Emp2_2021 = pd.read_excel(Ruta_Emp2_2021,sheet_name='Datos', engine='openpyxl')
#df_Mezclas_2021=pd.read_excel(Ruta_Mezclas_2021,sheet_name='Datos', engine='openpyxl')
