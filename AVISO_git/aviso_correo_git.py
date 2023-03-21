# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 16:28:24 2023

@author: sgalvez
"""

import os
import pandas as pd
import win32com.client as win32
from datetime import datetime

# Lee el archivo excel
file_path = r"C:\Users\YOUR_USER\TU_RUTA" # usa el directorio correcto dependiendo de donde guardes el archivo excel (en este git acompaño una planilla sencilla que puedes usar).
df = pd.read_excel(file_path, parse_dates=['rango_ini', 'rango_fin'])

# Obten la fecha de hoy
today = datetime.today().date()

# Filtra la tabla de datos en base a la condicion
filtered_df = df[(df['rango_ini'].dt.date <= today) & (today <= df['rango_fin'].dt.date)]

# Crea una instancia de Outlook. Enviará un correo desde tu dirección outlook.
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

for _, row in filtered_df.iterrows():
    # obtiene los correos desde las columnas B, C, and D
    email_address_1 = row['mail_1']
    email_address_2 = row['mail_2']
    email_address_3 = row['mail_3']
    client_name = row['cliente']

    # Crea y envia el primer correo
    mail1 = outlook.CreateItem(0)
    mail1.Subject = f"TEST - Aviso Gestión Carta Aviso {client_name}"
    contract_end_date = row['fin'].strftime('%d/%m/%Y')
    due_date = row['limite'].strftime('%d/%m/%Y')
    mail1.Body = (f"Estimad@, el contrato del cliente {client_name} vence el día {contract_end_date}. "
                  f"Se debe gestionar carta de aviso y envío por correo certificado con límite fecha {due_date}")
    mail1.To = f"{email_address_1}; {email_address_2}"
    mail1.Send()
    print(f"Email sent to {email_address_1} and {email_address_2}")

    # Crea y envia el segundo correo
    mail2 = outlook.CreateItem(0)
    mail2.Subject = f"TEST - Aviso Gestión Carta Aviso {client_name}"
    mail2.Body = (f"Estimad@, el contrato del cliente {client_name} vence el día {contract_end_date}. "
                  f"Por favor, te pedimos ayuda gestionando con fiscalía una carta de aviso de no renovación. "
                  f"El envío de la carta por correo certificado tiene como límite el día {due_date}. "
                  f"Este correo ha sido generado automáticamente.")
    mail2.To = email_address_3
    mail2.Send()
    print(f"Email sent to {email_address_3}")
