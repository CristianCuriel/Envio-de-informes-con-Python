import smtplib
from email.message import EmailMessage
import os
from config.setting import EMAIL_ADDRESS, EMAIL_PASSWORD, SERVER_EMAIL, PORT_EMAIL

class EmailSender:
    def __init__(self):
        self.server = SERVER_EMAIL
        self.port = PORT_EMAIL
        self.username = EMAIL_ADDRESS
        self.password = EMAIL_PASSWORD

    def send_email(self, to_email: str, subject: str, body: str, file_path: str) -> tuple[bool, str]:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = self.username
        msg["To"] = to_email
        msg.add_alternative(body, subtype="html")

        # Adjuntar archivo
        with open(file_path, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(file_path)
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

        try:
            # Enviar
            with smtplib.SMTP(self.server, self.port) as smtp:
                smtp.starttls()
                smtp.login(self.username, self.password)
                smtp.send_message(msg)
            status = f"Correo enviado a {to_email} con el archivo adjunto {file_name}"
            return True , status
        except smtplib.SMTPException as e:
            status = f"Error al enviar el correo: {e}"
        except smtplib.SMTPAuthenticationError:
            status = "Error de autenticación. Verifica tu nombre de usuario y contraseña."
        except smtplib.SMTPConnectError:
            status = "Error de conexión. Verifica tu conexión a Internet y la configuración del servidor SMTP."
        except smtplib.SMTPRecipientsRefused:
            status = f"El destinatario {to_email} fue rechazado por el servidor SMTP."
        except Exception as e:
            status = f"Error inesperado: {e}"
        return False, status