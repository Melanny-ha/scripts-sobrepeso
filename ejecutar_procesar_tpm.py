import subprocess
import os
from datetime import datetime
import traceback
from enviar_correo_errores import enviar_error_por_correo

carpeta_scripts = "C:/mha/scripts_sobrepeso/informes_py"

try:
    for archivo in os.listdir(carpeta_scripts):
        if archivo.endswith(".py"):
            ruta = os.path.join(carpeta_scripts, archivo)
            subprocess.run(["python", ruta], check=True)
except Exception as e:
    #Registrar el error en un .txt
    error_txt = "C:/mha/scripts_sobrepeso/errores_TPM.txt"
    with open(error_txt, "w", encoding='utf-8') as f:
        f.write(f"-----------------------------------Registro de errores - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}-----------------------------------\n\n")
        f.write(f"🚨{traceback.format_exc()}")
    #Enviar el correo con el .txt adjunto
    enviar_error_por_correo(
        asunto="🚨Error en ejecución de scripts TPM",
        cuerpo="Se ha producido un error durante la ejecución del script. Revisar el archivo adjunto.",
        archivo_adjunto = error_txt
    )