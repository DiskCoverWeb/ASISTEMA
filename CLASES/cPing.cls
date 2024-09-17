VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Clase para hacer ping desde Visual Basic"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales que contienen valores de propiedad
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Respuesta As String
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private mvarIPDestino As String 'copia local
Private mvarLongitudDatos As Long 'copia local
Private mvarTimeOut As Long 'copia local
Private mvarEstado As Long 'copia local
Private mvarDescripcion As String 'copia local
Private mvarTiempo As Long 'copia local

'códigos de error
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const PING_TIMEOUT = 200
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

'Para saber si hay internet
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Private dwFlags As Long

'estructuras
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private ICMPOPT As ICMP_OPTIONS

Private Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

'funciones del api
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, _
                                                      ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, _
                                                      ByVal ReplySize As Long, ByVal Timeout As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long

'Funcion Api que obtiene información sobre el estado de Red
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Funciones API: URLMON
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                                                                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'Funciones API: IPHlpApi para obtener IP
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Private Function AddressStringToLong(ByVal tmp As String) As Long
Dim I As Integer
Dim parts(1 To 4) As String
I = 0
'extrae cada parte del string '123.456.789.123
'delimitado por puntos
While InStr(tmp, ".") > 0
    I = I + 1
    parts(I) = Mid(tmp, 1, InStr(tmp, ".") - 1)
    If Trim(parts(I)) = "" Then parts(I) = "0"
    tmp = Mid(tmp, InStr(tmp, ".") + 1)
Wend
I = I + 1
parts(I) = tmp
If Trim(parts(I)) = "" Then parts(I) = "0"
If I <> 4 Then
    AddressStringToLong = 0
    Exit Function
End If
'build the long value out of the  'hex of the extracted strings
AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
End Function

Private Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Private Function SocketsCleanup() As Boolean
Dim X As Long
X = WSACleanup()
If X <> 0 Then
    MsgBox "Windows Sockets error " & Trim$(Str$(X)) & " occurred in Cleanup.", vbExclamation
    SocketsCleanup = False
Else
    SocketsCleanup = True
End If
End Function

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim X As Integer
Dim szLoByte As String, szHiByte As String, szBuf As String
X = WSAStartup(WS_VERSION_REQD, WSAD)
If X <> 0 Then
    MsgBox "Windows Sockets for 32 bit Windows " & "environments is not successfully responding."
    SocketsInitialize = False
    Exit Function
End If
If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    szHiByte = Trim$(Str$(HiByte(WSAD.wVersion)))
    szLoByte = Trim$(Str$(LoByte(WSAD.wVersion)))
    szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
    szBuf = szBuf & " is not supported by Windows " & _
                        "Sockets for 32 bit Windows environments."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    szBuf = "This application requires a minimum of " & _
    Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
End If
SocketsInitialize = True
End Function


Private Function GetStatusCode(Status As Long) As String
Dim msg As String

Select Case Status
    Case IP_SUCCESS:               msg = "Servidor en linea"
    Case IP_BUF_TOO_SMALL:         msg = "buffer demasiado pequeño"
    Case IP_DEST_NET_UNREACHABLE:  msg = "red de destino no encontrada"
    Case IP_DEST_HOST_UNREACHABLE: msg = "host destino no encontrado"
    Case IP_DEST_PROT_UNREACHABLE: msg = "dest prot unreachable"
    Case IP_DEST_PORT_UNREACHABLE: msg = "puerto destino no encontrado"
    Case IP_NO_RESOURCES:          msg = "sin recursos"
    Case IP_BAD_OPTION:            msg = "opción inválida"
    Case IP_HW_ERROR:              msg = "error hardware"
    Case IP_PACKET_TOO_BIG:        msg = "paquete demasiado grande"
    Case IP_REQ_TIMED_OUT:         msg = "timeout"
    Case IP_BAD_REQ:               msg = "respuesta incorrecta"
    Case IP_BAD_ROUTE:             msg = "ruta inválida"
    Case IP_TTL_EXPIRED_TRANSIT:   msg = "ttl finalizado"
    Case IP_TTL_EXPIRED_REASSEM:   msg = "ttl expired reassem"
    Case IP_PARAM_PROBLEM:         msg = "error en parámetros"
    Case IP_SOURCE_QUENCH:         msg = "source quench"
    Case IP_OPTION_TOO_BIG:        msg = "opción demasiado grande"
    Case IP_BAD_DESTINATION:       msg = "destino incorrecto"
    Case IP_ADDR_DELETED:          msg = "addr deleted"
    Case IP_SPEC_MTU_CHANGE:       msg = "spec mtu change"
    Case IP_MTU_CHANGE:            msg = "ip mtu_change"
    Case IP_UNLOAD:                msg = "unload"
    Case IP_ADDR_ADDED:            msg = "addr added"
    Case IP_GENERAL_FAILURE:       msg = "fallo general"
    Case IP_PENDING:               msg = "pendiente"
    Case PING_TIMEOUT:             msg = "ping timeout"
    Case INTERNET_CONNECTION_OFFLINE: msg = "Sin conexion a internet"
    Case Else:                     msg = "recibido mensaje desconocido"
End Select
GetStatusCode = msg
End Function

Public Property Get Tiempo() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Tiempo
    Tiempo = mvarTiempo
End Property

Public Property Get Descripcion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Descripcion
    Descripcion = mvarDescripcion
End Property

Public Property Get Estado() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Estado
    Respuesta = Respuesta & Descripcion & vbCrLf
    Estado = mvarEstado
End Property

Public Property Let Timeout(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.TimeOut = 5
    mvarTimeOut = vData
    If mvarTimeOut < 1 Then mvarTimeOut = 1
End Property

Public Property Get Timeout() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.TimeOut
    Timeout = mvarTimeOut
End Property

Public Property Let LongitudDatos(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.LongitudDatos = 5
    mvarLongitudDatos = vData
    If mvarLongitudDatos > 250 Then mvarLongitudDatos = 250
    If mvarLongitudDatos < 1 Then mvarLongitudDatos = 1
End Property

Public Property Get LongitudDatos() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.LongitudDatos
    LongitudDatos = mvarLongitudDatos
End Property

Public Property Let IPDestino(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.IPDestino = 5
    mvarIPDestino = vData
End Property

Public Property Get IPDestino() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.IPDestino
    IPDestino = mvarIPDestino
End Property

Public Function RecuperarIP(ByVal Nombre_Host As String) As String
Dim ErrorNo As String
Dim lHost As Long, T_Host As HOSTENT, ipDir As Long
Dim tIP() As Byte, I As Integer, sIP As String
      
    If Not SocketsInitialize() Then
        RecuperarIP = ""
        Exit Function
    End If
    Nombre_Host = Trim(Nombre_Host)
    lHost = gethostbyname(Nombre_Host)
    If lHost = 0 Then
        RecuperarIP = ""
        ErrorNo = "Error: " & Nombre_Host & ", fuera de linea o sin internet"
        'Remover_Socket
        Exit Function
    End If
    CopyMemoryIP T_Host, lHost, Len(T_Host)
    CopyMemoryIP ipDir, T_Host.hAddrList, 4
    ReDim tIP(1 To T_Host.hLen)
    CopyMemoryIP tIP(1), ipDir, T_Host.hLen
     
    For I = 1 To T_Host.hLen
        sIP = sIP & tIP(I) & "."
    Next
    RecuperarIP = Mid(sIP, 1, Len(sIP) - 1)
    'Remover_Socket
End Function

Private Sub Class_Initialize()
mvarLongitudDatos = 32
mvarTimeOut = 1500
Respuesta = ""
End Sub

'Funciones Publicas de la clase
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Función Booleana para saber si se está conectado a internet
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConexionInternet() As Boolean
Dim Result As Boolean
   Result = InternetGetConnectedState(0&, 0&)
   If Result Then Respuesta = "Conectado al Internet" & vbCrLf
   ConexionInternet = Result
End Function

Public Sub ping()
    Dim hPort As Long
    Dim dwAddress As Long
    Dim sDataToSend As String
    Dim iOpt As Long
    Dim ECHO As ICMP_ECHO_REPLY
    Dim res As Boolean
    
    res = SocketsInitialize
    If res Then
        If ConexionInternet Then
           sDataToSend = String$(mvarLongitudDatos, "A")
           dwAddress = AddressStringToLong(mvarIPDestino)
           hPort = IcmpCreateFile()
    
           If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), mvarTimeOut) Then
                'the ping succeeded,.Status will be 0
                '.RoundTripTime is the time in ms for the ping to complete,
                '.Data is the data returned (NULL terminated)
                '.Address is the Ip address that actually replied
                '.DataSize is the size of the string in .Data
                 mvarEstado = 0
                 mvarTiempo = ECHO.RoundTripTime
           Else
                mvarTiempo = 0
                mvarEstado = Abs(ECHO.Status)
                If mvarEstado = 0 Then mvarEstado = IP_GENERAL_FAILURE
           End If
           Call IcmpCloseHandle(hPort)
           res = SocketsCleanup
        Else
           mvarEstado = INTERNET_CONNECTION_OFFLINE
        End If
    Else
        mvarEstado = IP_GENERAL_FAILURE
        mvarTiempo = 0
    End If
    mvarDescripcion = GetStatusCode(mvarEstado)
End Sub

Public Function IP_Publica() As String
Dim cTemp As String
Dim arTemp() As String
Dim URL As String
Dim IP As String
Dim PosIpI As Long
Dim PosIpF As Long
    
    cTemp = "0.0.0.0"
    If ConexionInternet Then
       IP = RutaSysBases & "\TEMP\IP_" & CodigoUsuario & ".txt"
        URL = "https://whatismyipaddress.com/"
        If Dir(IP) <> "" Then Kill IP
        Call URLDownloadToFile(0, URL, IP, 0, 0)
        If Dir(IP) <> "" Then
           cTemp = CreateObject("Scripting.FileSystemObject").OpenTextFile(IP).ReadAll
           cTemp = Replace(cTemp, """", "'")
           PosIpI = InStr(cTemp, "https://whatismyipaddress.com/ip/")
            If PosIpI > 0 Then
               cTemp = MidStrg(cTemp, PosIpI, Len(cTemp))
               PosIpF = InStr(cTemp, "'")
               cTemp = MidStrg(cTemp, 1, PosIpF - 1)
               cTemp = Replace(cTemp, "https://whatismyipaddress.com/ip/", "")
               cTemp = TrimStrg(MidStrg(cTemp, 1, 15))
            End If
            Kill IP
        End If
    End If
    IP_Publica = cTemp
End Function

Public Function Nombre_PC() As String
Dim dwLen As Long
Dim strString As String

   'Averiguamos el nombre del PC
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    GetComputerName strString, dwLen
    strString = LeftStrg(strString, dwLen)
    Nombre_PC = strString
End Function

Public Function IP_Del_PC() As String
Dim TempIP As String
Dim L3 As String
Dim ret As Long
Dim Tel As Long
Dim RSeg As Boolean
Dim Listing As MIB_IPADDRTABLE

   'Determinamos la IP del Ordenador
    TempIP = "0.0.0.0"
    RSeg = True
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then
       IP_Del_PC = TempIP
       Exit Function
    End If
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    GetIpAddrTable bBytes(0), ret, False
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
        L3 = LeftStrg(TempList(Tel), 3)
        If L3 <> "169" And L3 <> "127" And RSeg Then
           TempIP = TempList(Tel)
           RSeg = False
        End If
    Next Tel
    'IP_PC.MAC_PC = Mi_MAC_Local
    IP_Del_PC = TempIP
End Function

Public Function MAC_Del_PC() As String
Dim colNetAdapters, objWMIService, objitem As Object
Dim strComputer As String
Dim tMAC As String
    tMAC = "00:00:00:00:00:00"
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    For Each objitem In colNetAdapters
        tMAC = objitem.MACAddress
    Next
    MAC_Del_PC = tMAC
End Function
