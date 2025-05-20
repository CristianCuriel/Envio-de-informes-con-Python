# -*- coding: utf-8 -*-

import pandas as pd

class ProcessData:
    def __init__(self, df):
        self.df = df

    def organizar_por_sede(self):

        sedes_dict = {}

        if 'Sede' not in self.df.columns:
            raise ValueError("La columna 'Sede' no existe en el DataFrame")

        for sede, grupo in self.df.groupby('Sede'):
            sedes_dict[sede] = grupo.reset_index(drop=True)
            
        return sedes_dict
    
    def filtrar_por_fecha(self, fecha_inicio, fecha_fin):
        if 'Fecha de recepcion' not in self.df.columns:
            raise ValueError("La columna 'Fecha' no existe en el DataFrame")

        self.df['Fecha de recepcion'] = pd.to_datetime(self.df['Fecha de recepcion'], format='%d/%m/%Y', errors='coerce')
        self.df = self.df[(self.df['Fecha de recepcion'] >= fecha_inicio) & (self.df['Fecha de recepcion'] <= fecha_fin)]
        
        return self.df

    def filtrar_por_fecha_y_sede(self, fecha_inicio, fecha_fin, sede):
        if 'Fecha de recepcion' not in self.df.columns:
            raise ValueError("La columna 'Fecha' no existe en el DataFrame")
        if 'Sede' not in self.df.columns:
            raise ValueError("La columna 'Sede' no existe en el DataFrame")

        self.df['Fecha de recepcion'] = pd.to_datetime(self.df['Fecha de recepcion'], format='%d/%m/%Y', errors='coerce')
        self.df = self.df[(self.df['Fecha de recepcion'] >= fecha_inicio) & (self.df['Fecha de recepcion'] <= fecha_fin) & (self.df['Sede'] == sede)]
        
        return self.df

    def filtrar_ultimos2meses(self):
        if 'Fecha de recepcion' not in self.df.columns:
            raise ValueError("La columna 'Fecha' no existe en el DataFrame")
        
        self.df["Fecha de recepcion"] = pd.to_datetime(self.df["Fecha de recepcion"], dayfirst=True, errors="coerce")
        # 2. Calcula la fecha límite (hoy menos 2 meses)
        hoy = pd.Timestamp.today()
        limite = hoy - pd.DateOffset(months=2)
        # 3. Filtra el DataFrame
        self.df = self.df[self.df["Fecha de recepcion"] >= limite]
        # 4. Resetea el índice
        self.df.reset_index(drop=True, inplace=True)
        # 5. Devuelve el DataFrame filtrado
        return self.df









