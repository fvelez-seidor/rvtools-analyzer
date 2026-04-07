import argparse
import smtplib
from email.message import EmailMessage


def send_email_via_relay(relay_host, sender, receiver, body, file_path=None):
    # 1. Configuración del Relay
    # Reemplaza con la IP o hostname de tu relay y el puerto (usualmente 25 o 587)
    relay_port = 25

    # 2. Crear el mensaje
    msg = EmailMessage()
    msg["Subject"] = "Notificación de Sistema"
    msg["From"] = sender
    msg["To"] = receiver
    msg.set_content(body)
    # add an existent test.txt file in the same directory as this script to test the attachment functionality
    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = f.name
        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="octet-stream",
            filename=file_name,
        )

    # 3. Enviar sin login
    try:
        # Conexión directa al host y puerto especificados
        with smtplib.SMTP(relay_host, relay_port) as server:
            # Nota: No llamamos a server.login() porque el relay confía en tu IP
            server.send_message(msg)
        print("Correo enviado correctamente a través del relay.")
    except Exception as e:
        print(f"Error al conectar con el relay: {e}")


parser = argparse.ArgumentParser(
    description="Analiza reportes RVTools y detecta anomalías"
)

parser.add_argument(
    "--email-receiver",
    default="destino@ejemplo.com",
    help="Dirección de correo del destinatario (usado con --send-email).",
)

args = parser.parse_args()


send_email_via_relay(
    "10.245.12.22",
    "noreply@seidor.net",
    args.email_receiver,
    "Este es un mensaje de prueba.",
)
