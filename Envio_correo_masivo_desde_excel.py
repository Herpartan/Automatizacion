# -*- coding: utf-8 -*-
"""
Created on Wed Jan 26 11:02:11 2022

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
    archivo = MIMEApplication(open(r'\DOC.docx', 'rb').read())
    archivo.add_header('Content-Disposition', 'attachment', filename='DOC.docx')
    msg.attach(archivo)
    
    # Envia el correo
    server = SMTP(SMTP_HOST, SMTP_PORT)
    server.sendmail(remitente, rcpt, msg.as_string())

# Carpetas
carpeta_correo = r''
archivo_contactos = '\Contactos.xlsx'
archivo_texto = '\Texto.xlsx'

# Port y host para poder enviar mails
SMTP_HOST='smtp.'
SMTP_PORT='25'
remitente = ''
destinatario_copia = ''

# Lee el archivo de destinatarios
tabla_destinatarios = pd.read_excel(carpeta_correo+archivo_contactos)
remitente = tabla_destinatarios['Remitente'].dropna()[0]
destinatarios = tabla_destinatarios['Destinatario'].dropna().to_list()
destinatarios_copia = tabla_destinatarios['Destinatario copia'].dropna().to_list()

# Lee el archivo de texto y genera el encabezado y cuerpo
tabla_texto = pd.read_excel(carpeta_correo+archivo_texto, index_col='Tipo')

encabezado_correo = tabla_texto['Texto']['Asunto']
saludo = '<br>'.join([tabla_texto['Texto'][s] for s in tabla_texto.index if 'saludo' in s.lower()])
cuerpo = '<br>'.join([tabla_texto['Texto'][c] for c in tabla_texto.index if 'cuerpo' in c.lower()])
despedida = '<br>'.join([tabla_texto['Texto'][d] for d in tabla_texto.index if 'despedida' in d.lower()])

cuerpo_correo = '''
                    <html>
                        <body>
                            <p> {saludo} </p>
                            <p> {cuerpo} </p>
                            <p> {despedida} </p>
                        </body>
                    </html>
                '''.format(saludo=saludo, cuerpo=cuerpo, despedida=despedida)

# Bucle por destinatario
for destinatario in destinatarios:
    # Manda el correo
    manda_correo(encabezado_correo, cuerpo_correo, remitente, destinatario, destinatario_copia)
