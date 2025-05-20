from use_cases.extract_data import ExtractData
from use_cases.send_report import SendReportUseCase
from infrastructure.google_sheet import GoogleSheet
from infrastructure.send_email import EmailSender
from use_cases.process_data import ProcessData
from use_cases.export_excel import ExportDataUseCase
from infrastructure.excel_report import ExcelReportExporter
from use_cases.status_reporter import StatusReporter

def run():

    #Exractar datos de Google Sheets
    print("Analizando datos de Google Sheets...")
    google_sheet = GoogleSheet()
    extract_values = ExtractData(google_sheet).extraccion_datos()

    #Procesamos los datos segun el caso
    data_procees = ProcessData(extract_values).filtrar_ultimos2meses() #Filtramos los datos de los ultimos 2 meses
    data_procees = ProcessData(data_procees).organizar_por_sede() #Organizamos los datos por sede


    #Exportar a Excel
    exporter = ExcelReportExporter()
    report_export = ExportDataUseCase(exporter).execute(data_procees)
    StatusReporter(report_export).view_satus_report()

    #Enviar correo
    email_sender = EmailSender()
    send_report = SendReportUseCase(email_sender).enviar_archivos_por_sede()
    StatusReporter(send_report).view_satus_report()
