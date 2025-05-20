
# -*- coding: utf-8 -*-
import os
from infrastructure.send_email import EmailSender
from config.setting import EMAIL_SEDE_MAP


class SendReportUseCase:
    def __init__(self, email_sender: EmailSender):
        self.email_sender = email_sender

    def enviar_archivos_por_sede(self) -> list[tuple[bool, str]]:
        ruta_archivos = "Excel/"
        status_send = []

        for archivo in os.listdir(ruta_archivos):
            if archivo.endswith(".xlsx"):
                sede = archivo.replace(".xlsx", "")
                correo_destino = EMAIL_SEDE_MAP.get(sede, [])
                if correo_destino:
                   status =  self.email_sender.send_email(
                        to_email=correo_destino,
                        subject=f"Reporte de garantias - {sede}",
                        body = _cuerpo_msg(sede = sede),
                        file_path=os.path.join(ruta_archivos, archivo)
                    )
                   status_send.append(status)   
                else:
                    status_send.append((False, f"❌ No se encontró correo para la sede {sede}")) 
        return status_send    
    
def _cuerpo_msg(sede:str) -> str:
    # Cuerpo del mensaje

    body_msg = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                <h2 style="color: #2e6c80;">📊 REPORTE DE GARANTIAS - SEDE HLTB {sede.upper()}</h2>

                <p>Estimado equipo de ventas,</p>

                <p>
                    Adjunto encontrarán el reporte actualizado de las garantías registradas en el formulario oficial de la empresa. 
                    Este informe corresponde a las solicitudes recibidas durante los últimos dos meses para la sede <b>{sede}</b>.
                </p>

                <p>
                    Les agradecemos revisar el documento adjunto y realizar el seguimiento correspondiente.
                    Si tienen alguna observación o requieren soporte adicional, no duden en contactarme.
                </p>

                <p style="font-style: italic; color: #555;">
                    Por favor, revisar el archivo adjunto para más detalles.
                </p>

                <br>

                <p>Saludos cordiales,</p>

                <p>
                    <b>Cristian Curiel Camargo</b><br>
                    Líder de garantías en Heliteb SAS<br>
                    📞 317 218 5359<br>
                    ✉️ <a href="mailto:garantias@heliteb.com.co">garantias@heliteb.com.co</a>
                </p>
            </body>
            </html>

        """

    return body_msg
     
    

