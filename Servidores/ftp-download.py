from ftplib import FTP, error_perm
import sys
from time import time


ftp = FTP("ftpds.diskcoversystem.com")
ftp.login("ftpuser", "ftp2023User")

ftp.cwd('/files/SISTEMA/')

ftp.retrlines('LIST')

def getFile(ftp, filename):
    try:
        ftp.retrbinary('RETR '+ filename , open(filename, 'wb').write)

    except error_perm as e:
        print("FTP error:", e)

        if str(e).split(None, 1)[0] == '530':
            print("Check your username and password.")

        elif str(e).split(None, 1)[0] == '550':
            print("File not found or no access.")

start = time()
print("iniciando conteo")
#getFile(ftp,'CajaChica.exe')
getFile(ftp,'Contabilidad.exe')
#getFile(ftp,'Facturacion.exe')
#getFile(ftp,'Farmacia.exe')
#getFile(ftp,'RolPagos.exe')
#getFile(ftp,'inventario.exe')
#getFile(ftp,'Seteos.exe')

end= time()
print("Terminando conteo")
print(end - start)

ftp.quit()
