# -*- coding: utf-8 -*-
"""
Created on Tue Jan 25 15:49:34 2022

@author: hlopez
"""

from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import pandas as pd


def manda_correo(encabezado, cuerpo, remitente, destinatario, destinatario_copia):
    msg = MIMEMultipart()
    msg['Subject'] = encabezado
    msg['To'] = destinatario
    msg['Cc'] = destinatario_copia
    
    if ';' in destinatario: rcpt = destinatario.split(';') + destinatario_copia.split(';')
    else: rcpt = destinatario.split(',') + destinatario_copia.split(',')
    
    # Define y agrega el cuerpo del mensaje
    texto = MIMEText(cuerpo , 'html')
    msg.attach(texto)
    
    # Carga y agrega el archivo a enviar
    archivo = MIMEApplication(open(r'DOCUMENT.docx', 'rb').read())
    archivo.add_header('Content-Disposition', 'attachment', filename='DOCUMENT.docx')
    msg.attach(archivo)
    
    # Envia el correo
    server = SMTP(SMTP_HOST, SMTP_PORT)
    server.sendmail(remitente, rcpt, msg.as_string())

# Port y host para poder enviar mails
SMTP_HOST = 'smtp.'
SMTP_PORT = '25'
remitente = ''
destinatario_copia = ''

# Lee el archivo de destinatarios
tabla_destinatarios = pd.read_excel(r'lista_correo.xlsx')

# Definicion encabezado y cuerpo
encabezado_correo = 'ENCABEZADO CON DESTIONO A | {}'
cuerpo_correo = '''
                <html>
                    <body>
                        <p> Buenos días, </p>
                        <p> <strong> COSAS EN NEGRITA </strong>
                            LINEA DE ESPACIO <br>
                            METE AQUI ALGO {}
                        </p>
                        <p> Muchas gracias. <br>
                            Un saludo. <br>
                        </p>
                    </body>
                </html>
                '''

# Bucle por destinatario
for sociedad, contacto in zip(tabla_destinatarios['Sociedad'][2:], tabla_destinatarios['Contacto'][2:]):
    destinatario = contacto

    encabezado = encabezado_correo.format(sociedad)
    cuerpo = cuerpo_correo.format(sociedad)
    
    print(sociedad)
    
    # Manda el correo
    manda_correo(encabezado, cuerpo, remitente, destinatario, destinatario_copia)
