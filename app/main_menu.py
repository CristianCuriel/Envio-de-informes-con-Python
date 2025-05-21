from use_cases.extract_data import ExtractData
from use_cases.send_report import SendReportUseCase
from infrastructure.google_sheet import GoogleSheet
from infrastructure.send_email import EmailSender
from use_cases.process_data import ProcessData
from use_cases.export_excel import ExportDataUseCase
from infrastructure.excel_report import ExcelReportExporter
from use_cases.status_reporter import StatusReporter

class Menu():

    def __init__(self):
        self.opciones = {
            "1": self.__reportes_todos,
            "2": self.__reportes_filtrados_ultimos2meses,
            "3": self.salir
        }
        self.en_ejecucion = True
        
    def mostrar_menu(self):
        print("\n--- MENÚ PRINCIPAL ---")
        print("1. Enviar reportes completos")
        print("2. Enviar repotes filtrados - 2 meses")
        print("3. Salir")

    
    def ejecutar_opcion(self):
        while self.en_ejecucion:
            self.mostrar_menu()
            opcion = input("Seleccione una opción: ")
            if opcion in self.opciones:
                self.opciones[opcion]()
            else:
                print("Opción no válida. Intente nuevamente.")

    def __reportes_todos(self):
        #Procesamos los datos segun el caso
        data_procees = ProcessData(self.__extraer_datos()).organizar_por_sede() #Organizamos los datos por sede
        #Exportar a Excel
        self.__exportar_excel(data_procees)
        #Enviar correo
        self.__enviar_correo(2)
        self.en_ejecucion = False
    
    def __reportes_filtrados_ultimos2meses(self):
        #Procesamos los datos segun el caso
        data_procees = ProcessData(self.__extraer_datos()).filtrar_ultimos2meses() #Filtramos los datos de los ultimos 2 meses
        data_procees = ProcessData(data_procees).organizar_por_sede() #Organizamos los datos por sede
        #Exportar a Excel
        self.__exportar_excel(data_procees)
        #Enviar correo
        self.__enviar_correo(1)
        self.en_ejecucion = False

    def __extraer_datos(self): 
        google_sheet = GoogleSheet()
        return ExtractData(google_sheet).extraccion_datos()

    def __exportar_excel(self, data_procees:dict):    
        #Exportar a Excel
        exporter = ExcelReportExporter()
        report_export = ExportDataUseCase(exporter).execute(data_procees)
        StatusReporter(report_export).view_satus_report()
    
    def __enviar_correo(self, op:int):
        #Enviar correo
        email_sender = EmailSender()
        send_report = SendReportUseCase(email_sender).enviar_archivos_por_sede(op)
        StatusReporter(send_report).view_satus_report()



    def salir(self):
        print("Saliendo del programa...")
        self.en_ejecucion = False


