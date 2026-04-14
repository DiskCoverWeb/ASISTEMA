#!/usr/bin/python3
# -*- coding: utf-8 -*-

import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP
from datetime import datetime
import ssl

msg = MIMEMultipart()
message = "Prueba de correo IMAP SERVER 2024"
msg['From'] = "CORREO DESDE 192.168.115.120 USANDO RELAYHOST IMAP <admin@imap.diskcoversystem.com>"
msg['To'] = "actualizar@diskcoversystem.com"
msg['Cc'] = "diskcover.system@yahoo.com"

msg['Subject'] = f"Correo test new server IMAP {datetime.now()}"
msg.attach(MIMEText(message, 'plain'))

context = ssl.create_default_context()
server = SMTP('imap.diskcoversystem.com', 26)
server.ehlo('imap.diskcoversystem.com')
#server.set_debuglevel(True)
server.connect('imap.diskcoversystem.com', 26)
server.ehlo('imap.diskcoversystem.com')

#server.starttls()
server.login('admin', 'Admin@2023')
server.sendmail(msg['From'], msg['To'].split(',') + msg['Cc'].split(','), msg.as_string())
print(f"Enviado a: {msg['To'].split(',')} con copia a:  {msg['Cc'].split(',')}")
server.quit()

