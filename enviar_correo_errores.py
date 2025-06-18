import smtplib
from email.message import EmailMessage
import os

def enviar_error_por_correo(asunto, cuerpo, archivo_adjunto=None):
    remitente = "mherreraa@colcafe.com.co"
    destinatario1 = "mherreraa@colcafe.com.co"
    # destinatario = "jmbeltran@colcafe.com.co"
    contraseña = "rpctrldxyjrqhuib"

    mensaje = EmailMessage()
    mensaje["From"] = remitente
    mensaje["To"] = f"{destinatario1}"
    # mensaje["To"] = f"{destinatario1}, {destinatario2}"
    mensaje["Subject"] = asunto
    mensaje.set_content(cuerpo)

    if archivo_adjunto and os.path.exists(archivo_adjunto):
        with open(archivo_adjunto, "rb") as f:
            mensaje.add_attachment(f.read(), filename=os.path.basename(archivo_adjunto), maintype="text", subtype="plain")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(remitente, contraseña)
            smtp.send_message(mensaje)
    except Exception as e:
        print(f"No se pudo enviar el correo: {e}")
