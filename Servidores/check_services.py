#!/usr/bin/python3
#-*- coding: utf-8 -*-

import socket
import requests
import os
import platform

def limpiar_pantalla():
  # Verifica el sistema operativo
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

def isOpen(ip,port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect((ip, port))
        s.shutdown(2)
        return True

    except:
        return False

def url_checker(url, amb, comp):
    try:
                #Get Url
        get = requests.get(url)
                # if the request succeeds
        if get.status_code == 200:
            return(f"SRI {amb} / {comp}: Esta Funcionando!")
        else:
            pass
    except:
        return(f"SRI {amb} / {comp}: NO ESTA ACTIVO")

def main():
    limpiar_pantalla()

    if (isOpen('104.237.138.52',8010)) == True:
        print('\nserver API DiskCover System Port 8010: OK')
    else:
        print('\nServer API DiskCover System Port 8010: is Down!')
    print('-------------------------------------------------')
    if (isOpen('104.237.138.52',587)) == True:
        print('\nserver IMAP DiskCover System Port 587: OK')
    else:
        print('\nServer IMAP DiskCover System Port 587: is Down!')
    print('-------------------------------------------------')
    if (isOpen('194.195.222.54',11433)) == True:
        print('\nserver DB SQL DiskCover System Port 11433: OK')
    else:
        print('\nServer DB SQL DiskCover System Port 11433: is Down!')
    print('-------------------------------------------------')
    if (isOpen('69.164.192.53',13306)) == True:
        print('\nserver DB MySQL DiskCover System Port 13306: OK')
    else:
        print('\nServer DB MySQL DiskCover System Port 13306: is Down!')

if __name__ == '__main__':
    main()
    url_produccion_rc = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantesOffline?wsdl"
    url_pruebas_rc = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantesOffline?wsdl"
    url_produccion_ac = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
    url_pruebas_ac = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
    print('================================================')
    print(url_checker(url_produccion_rc,'PRODUCCION', 'RECEPCION COMPROBANTES'))
    print(url_checker(url_produccion_ac,'PRODUCCION', 'AUTORIZACION COMPROBANTES'),'\n')
    print('================================================')
    print(url_checker(url_pruebas_rc,'PRUEBAS', 'RECEPCION COMPROBANTES'))
    print(url_checker(url_pruebas_ac,'PRUEBAS', 'AUTORIZACION COMPROBANTES'),'\n')

