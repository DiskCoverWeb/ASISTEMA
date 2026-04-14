#!/usr/bin/python3
# -*- coding: utf-8 -*-

import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP
from datetime import datetime

msg = MIMEMultipart()
message = "Una Prueba de correo Puerto 26 - Test"
msg['From'] = "<admin@smtp.diskcoversystem.com>"
#msg['To'] = "colecturiaalausi@gmail.com"
#msg['To'] = "diskcover.system@yahoo.com"
msg['To'] = "soporte@smtech.com.ec"
msg['Cc'] = "ronramiro2007@gmail.com"

msg['Subject'] = f"Correo test port 26 {datetime.now()}"

msg.attach(MIMEText(message, 'plain'))
server = SMTP('smtp.diskcoversystem.com', 26)
#server = SMTP('smtp.diskcoversystem.com')
#server.set_debuglevel(True)
server.connect('smtp.diskcoversystem.com', 26)
#server.starttls()
server.ehlo("smtp.diskcoversystem.com")
server.login('admin@smtp.diskcoversystem.com', 'Admin@2023')
server.sendmail(msg['From'], msg['To'].split(',') + msg['Cc'].split(','), msg.as_string())
print(f"Enviado desde: {msg['From']}\n  a: {msg['To'].split(',')} con copia a:  {msg['Cc'].split(',')}")
server.quit()
