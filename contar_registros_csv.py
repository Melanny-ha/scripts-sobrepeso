import csv

def contar_registros_csv(nombre_archivo):
    with open(nombre_archivo, 'r', encoding='utf-8') as archivo:
        lector_csv = csv.reader(archivo)
        return len(list(lector_csv))

archivo = r'Y:\BD Sobrepeso\Backup_BD_Salones\Datos_Molido_2021.csv'

try:
    numero_de_registros = contar_registros_csv(archivo)
    print(f"El archivo CSV tiene {numero_de_registros} registros.")
except UnicodeDecodeError:
    print("Error con UTF-8")
    with open(archivo, 'r', encoding='latin-1') as archivo:
        lector_csv = csv.reader(archivo)
        numero_de_registros = len(list(lector_csv))
        print(f"El archivo CSV tiene {numero_de_registros} registros.")
