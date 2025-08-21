import os                                                 #para rutas y archivos
import nbformat                                           #para leer archivos .ipynb
from nbconvert import ScriptExporter                      #ScriptExporter convierte un notebook a .py
from nbconvert.preprocessors import ExecutePreprocessor   #ejecuta internamente el notebook para ver si tiene errores
from datetime import datetime
import asyncio

asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy()) # usar WindowsSelectorEventLoopPolicy para ejecución con libreria Jupyter y/o notebooks y no generar "RuntimeWarning: Proactor event loop..."

# carpeta_notebooks = "C:/mha/scripts_sobrepeso/informes_ipynb"       #carpeta de archivos .ipynb
# carpeta_salida_py = "C:/mha/scripts_sobrepeso/informes_py"          #carpeta de archivos .py

carpeta_notebooks = "C:/mha/scripts_ambiental/scripts_ipynb"       #carpeta de archivos .ipynb
carpeta_salida_py = "C:/mha/scripts_ambiental/scripts_py"          #carpeta de archivos .py

os.makedirs(carpeta_salida_py, exist_ok=True)    #crea la carpeta para .py si no existe

exporter = ScriptExporter()                   #crea un objeto exporter de la clase ScriptExporter que convierte los notebooks a scripts
errores = []                                  #lista para guardar errores si algo falla

def envolver_en_funcion(nombre_funcion, codigo_py):
    # Agrega sangría a cada línea del código
    sangria = "\n".join("    " + linea if linea.strip() else "" for linea in codigo_py.splitlines())
    # Arma la función con el bloque if deseado
    return (
        f"def {nombre_funcion}():\n"
        f"{sangria}\n\n"
        f"if __name__ == '__main__':\n"
        f"    {nombre_funcion}()\n"
    )

for archivo in os.listdir(carpeta_notebooks):                     #recorre todos los archivos de la carpeta
    if archivo.endswith(".ipynb"):                                #solo procesa los que terminan en .ipynb
        ruta_ipynb = os.path.join(carpeta_notebooks, archivo)
        nombre_py = archivo.replace(".ipynb", ".py")
        ruta_py = os.path.join(carpeta_salida_py, nombre_py)

        print(f"🔍 Validando: {archivo}")
        try:
            with open(ruta_ipynb, 'r', encoding='utf-8') as f:    #Abre el notebook como texto
                notebook = nbformat.read(f, as_version=4)

            ep = ExecutePreprocessor(timeout=300, kernel_name="python3")
            ep.preprocess(notebook, {'metadata': {'path': carpeta_notebooks}})

            codigo_py, _ = exporter.from_notebook_node(notebook)   #Si no hubo errores, convierte el notebook ejecutado a .py

            # Genera el nombre de la función según el nombre del archivo
            nombre_funcion = "main_" + nombre_py.replace(".py", "")

            contenido_final = envolver_en_funcion(nombre_funcion, codigo_py)

            if os.path.exists(ruta_py):
                with open(ruta_py, 'r', encoding='utf-8') as f:   #Si el archivo ya existe, verifica si cambió
                    contenido_existente = f.read()

                if contenido_existente == codigo_py:               #Si no hay cambios, no lo sobreescribe (para eficiencia)
                    print(f"✅ Sin cambios: {nombre_py}")
                    continue

            with open(ruta_py, 'w', encoding='utf-8') as f:       #Si cambió o no existía, guarda el nuevo archivo .py
                f.write(contenido_final)
            print(f"✅ Actualizado: {nombre_py}")

        except Exception as e:                                    #Si falla algo en la ejecución del notebook, lo captura y guarda el error sin detener todoy qué
            errores.append((archivo, str(e)))
            print(f"❌ Error en '{archivo}': {e}")
    elif archivo.endswith(".py"):
        ruta_py_origen = os.path.join(carpeta_notebooks, archivo)
        ruta_py_destino = os.path.join(carpeta_salida_py, archivo)

        try:
            with open(ruta_py_origen, 'r', encoding='utf-8') as f_src:
                contenido_origen = f_src.read()

            if os.path.exists(ruta_py_destino):
                with open(ruta_py_destino, 'r', encoding='utf-8') as f_dst:
                    contenido_destino = f_dst.read()
                if contenido_origen == contenido_destino:
                    print(f"📄 Sin cambios en archivo auxiliar: {archivo}")
                    continue

            with open(ruta_py_destino, 'w', encoding='utf-8') as f_dst:
                f_dst.write(contenido_origen)
            print(f"📄 Copiado archivo auxiliar: {archivo}")

        except Exception as e:
            errores.append((archivo, str(e)))
            print(f"❌ Error al copiar '{archivo}': {e}")

print("🎉 Conversión finalizada.")
if errores:
    print("\n🛑 Errores detectados:")
    ruta_errores = os.path.join(carpeta_salida_py, "errors.txt")
    with open(ruta_errores, 'w', encoding='utf-8') as log:
        for archivo, error in errores:
            print(f"  - {archivo}: {error}")
            log.write(f"-------------------------Registro de errores - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}-------------------------\n\n")
            log.write(f"❌ Error en '{archivo}': {error}\n")
