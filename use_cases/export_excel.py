# use_cases/export_data_use_case.py

from infrastructure.excel_report import ExcelReportExporter

class ExportDataUseCase:

    def __init__(self, exporter: ExcelReportExporter):
        self.exporter = exporter

    def execute(self, data: dict)-> list[tuple[bool, str]]:
        status_exported = []

        for nombre_sede, df_sede in data.items():
            status_exported.append(self.exporter.exportar(nombre_sede, df_sede))
        
        return status_exported
