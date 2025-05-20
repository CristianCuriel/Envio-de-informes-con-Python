
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
                <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">

                        <h2 style="color: #2e6c80;">📊 REPORTE DE GARANTÍAS - SEDE HLTB {sede}</h2>

                        <p>Estimado equipo de ventas,</p>

                        <p>
                            Adjunto encontrarán el reporte actualizado de las garantías registradas en el formulario oficial de la empresa. 
                            Este informe contiene los registros de los últimos <b>dos meses</b> correspondientes a la sede <b>{sede}</b>.
                        </p>

                        <p>
                            A continuación, se describen las categorías de estado de las garantías representadas por colores en el reporte:
                        </p>

                        <ul style="list-style: none; padding-left: 0;">
                            <li>
                                <span style="display: inline-block; width: 12px; height: 12px; background-color: #c6efce; border: 1px solid #ccc ;border-radius: 50%; margin-right: 8px;"></span>
                                <strong>Verde:</strong> El caso ha sido cerrado de manera exitosa. (Aplicó a garantía)
                            </li>
                            <li>
                                <span style="display: inline-block; width: 12px; height: 12px; background-color: #fff2cc; border: 1px solid #ccc ;border-radius: 50%; margin-right: 8px;"></span>
                                <strong>Amarillo:</strong> El caso está en proceso o trámite con el proveedor. (En espera de RMA o respuesta de casa matriz)
                            </li>
                            <li>
                                <span style="display: inline-block; width: 12px; height: 12px; background-color: #e85c5d; border: 1px solid #ccc; border-radius: 50%; margin-right: 8px;"></span>
                                <strong>Rojo:</strong> El caso aún no ha sido tramitado.
                            </li>
                            <li>
                                <span style="display: inline-block; width: 12px; height: 12px; background-color: #ffffff; border: 1px solid #ccc; border-radius: 50%; margin-right: 8px;"></span>
                                <strong>Blanco:</strong> El caso aún no ha sido tramitado.
                            </li>
                        </ul>

                        <p>
                            Les agradecemos revisar el documento adjunto y realizar el seguimiento correspondiente. 
                            Si tienen alguna observación o requieren soporte adicional, no duden en comunicarse conmigo.
                        </p>

                        <br>

                        <p>Saludos cordiales,</p>

                        <p>
                            <strong>Cristian Curiel Camargo</strong><br>
                            Líder de garantías en Heliteb SAS<br>
                            📞 317 218 5359<br>
                            ✉️ <a href="mailto:garantias@heliteb.com.co">garantias@heliteb.com.co</a>
                        </p>

                    </body>

        """

    return body_msg
     
    

