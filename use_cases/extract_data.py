import pandas as pd
from infrastructure.google_sheet import GoogleSheet


class ExtractData:
    def __init__(self, google_sheet: GoogleSheet):
        self.google_sheet = google_sheet

    def extraccion_datos(self):
        # Usa la primera fila como encabezado y el resto como datos
        self.values = self.google_sheet.get_values()
        if not self.values:
            raise ValueError("No se encontraron datos en la hoja de cálculo.")
        else:
            df = pd.DataFrame(self.values[1:], columns=self.values[0])
            return df