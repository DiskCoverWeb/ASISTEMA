VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Path mas largo
Private Const MAX_PATH = 260

'Constante para el atributo de directorio
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
        intYear         As Integer
        intMonth        As Integer
        intDayOfWeek    As Integer
        intDay          As Integer
        intHour         As Integer
        intMinute       As Integer
        intSecond       As Integer
        intMilliSeconds As Integer
End Type

'--- tipos de archivos --- para el Upload y Download
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

'Puerto FTP
Private Const INTERNET_SERVICE_FTP = 1

' Modo de conexión FTP
Private Const INTERNET_FLAG_PASSIVE = &H8000000
'Private Const INTERNET_FLAG_PASSIVE = New Binary(0x08000000)

Private Const PassiveConnection As Boolean = True

'--- formas de entrar en internet ---
' usa config del registro
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
' directo a internetnet
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
' via  proxy
Private Const INTERNET_OPEN_TYPE_PROXY = 3
' prevent using java/script/INS
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4

'Type para atributos de fecha y hora de archivos
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Otros atributos de archivo tamaño, nombre, fecha etc..
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' Declaraciones Apis
'***************************************************************
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Establece una conexión a internet para poder iniciar seción Ftp
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
  (ByVal hInternetSession As Long, _
   ByVal sServerName As String, _
   ByVal nServerPort As Integer, _
   ByVal sUserName As String, _
   ByVal sPassword As String, _
   ByVal lService As Long, ByVal lFlags As Long, _
   ByVal lContext As Long) As Long
'Conecta al Ftp
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal sAgent As String, ByVal lAccessType As Long, _
    ByVal sProxyName As String, _
    ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long
'Establece el path corriente
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
   (ByVal hFtpSession As Long, _
    ByVal lpszDirectory As String) As Boolean
'Recupera el path corriente
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
  (ByVal hFtpSession As Long, _
   ByVal lpszCurrentDirectory As String, _
   lpdwCurrentDirectory As Long) As Long
'Crea un directorio
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" _
   (ByVal hFtpSession As Long, _
   ByVal lpszDirectory As String) As Boolean
'Elimina un directorio del FTP
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" _
  (ByVal hFtpSession As Long, _
  ByVal lpszDirectory As String) As Boolean
'Borra un fichero
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
  (ByVal hFtpSession As Long, _
  ByVal lpszFileName As String) As Boolean
'Renombra un fichero
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
  (ByVal hFtpSession As Long, _
  ByVal lpszExisting As String, _
  ByVal lpszNew As String) As Boolean
'Recupera un archivo
Private Declare Function FtpGetFile Lib "wininet.dll" _
   Alias "FtpGetFileA" (ByVal hConnect As Long, _
  ByVal lpszRemoteFile As String, _
  ByVal lpszNewFile As String, ByVal fFailIfExists As Long, _
  ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
  ByRef dwContext As Long) As Boolean
'Escribe un archivo
Private Declare Function FtpPutFile Lib "wininet.dll" _
   Alias "FtpPutFileA" (ByVal hConnect As Long, _
  ByVal lpszLocalFile As String, _
  ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, _
  ByVal dwContext As Long) As Boolean
'Api Para los errores
Private Declare Function InternetGetLastResponseInfo _
  Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" _
  (lpdwError As Long, ByVal lpszBuffer As String, _
   lpdwBufferLength As Long) As Boolean
'Busca el primer archivo de un path
Private Declare Function FtpFindFirstFile Lib "wininet.dll" _
   Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, _
   ByVal lpszSearchFile As String, _
   lpFindFileData As WIN32_FIND_DATA, _
   ByVal dwFlags As Long, ByVal dwContent As Long) As Long
'api para buscar el siguiente archivo
Private Declare Function InternetFindNextFile Lib "wininet.dll" _
   Alias "InternetFindNextFileA" (ByVal hFind As Long, _
   lpvFindData As WIN32_FIND_DATA) As Long

Public Enum e_TipoTransferencia
[ BINARIO ] = FTP_TRANSFER_TYPE_BINARY
[ ASCII ] = FTP_TRANSFER_TYPE_ASCII
[ DESCONOCIDO ] = FTP_TRANSFER_TYPE_UNKNOWN
End Enum

    'Handle de la conexión Ftp
    Dim HandleConect As Long
    'Handle de la conexión a Internet
    Dim hOpen As Long
    
'Variables locales
Private INTERNET_DEFAULT_FTP_PORT As Integer
Private m_DirectorioActual As String
Private m_Usuario As String
Private m_PassWord As String
Private m_Servidor As String
Private m_DirAnterior As String
Private m_listView As ListView
Private m_Puerto As Integer
Private m_TipoTransferencia As Long
Private m_form As Form
Private ctrl As Object

'Funciones Varias para el manejo de archivos y carpetas en el servidor Ftp
'***********************************************************************

'Rutina que conecta al Servidor Ftp
Public Function ConectarFtp(Optional ControlStatus As Object = Nothing) As Boolean
    
'Conectamos al servidor FTP, con un puerto determinado
 If m_Puerto = 0 Then INTERNET_DEFAULT_FTP_PORT = 21 Else INTERNET_DEFAULT_FTP_PORT = m_Puerto

'Verificamos que los datos de la cuenta estén establecidas, si no mostramos un _
 mensaje y salimos
If m_Usuario = "" Or m_Servidor = "" Or m_PassWord = "" Then
   MsgBox "No se puede conectar. Verifique el Nombre de usuario," _
   & "El nombre del Servidor y la contraseña que estén establecidas", vbCritical
   ConectarFtp = False
   Exit Function
End If
    
   Set ctrl = ControlStatus
   Status "...Intentando conectar a: " & m_Servidor
   m_form.MousePointer = vbHourglass
'Abrimos una conección a Internet
'
'hOpen = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
hOpen = InternetOpen(Modulo, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
If hOpen = 0 Then
   Status "Error en la conexión a internet, compruebe la conexión"
   m_form.MousePointer = vbDefault
   ConectarFtp = False
   Exit Function
End If
'Conectamos al servidor FTP, pasandole los datos: login y servidor
 HandleConect = InternetConnect(hOpen, m_Servidor, INTERNET_DEFAULT_FTP_PORT, m_Usuario, m_PassWord, _
                INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
If HandleConect = 0 Then
   m_form.MousePointer = vbDefault
   Status "Error (" & Err.Number & ")" & Err.Description & ". Compruebe los datos del servidor Ftp sin son correctos"
   ConectarFtp = False
   Exit Function
End If

Status "Conectado a: " & m_Servidor
m_form.MousePointer = vbDefault
ConectarFtp = True
End Function

'Desconecta del servidor FTP
'**************************************************
Public Sub Desconectar()
    Dim ret As Long
    'cierra la conexion FTP
    ret = InternetCloseHandle(HandleConect)
    If ret = 0 Then Status "Error al desconectar": Exit Sub
    'cierra la conexion a internet
    ret = InternetCloseHandle(hOpen)
    If ret = 0 Then Status "Error al desconectar": Exit Sub
    
    Status "Desconectado de: " & m_Servidor
    Class_Terminate
End Sub

'Recupera el directorio actual donde estamos ubicados
'*****************************************************
Public Function GetDirectorioActual() As String
    'Crea un buffer
    m_DirectorioActual = String(MAX_PATH, 0)
    'Obtenemos el directorio actual
     ret = FtpGetCurrentDirectory(HandleConect, m_DirectorioActual, Len(m_DirectorioActual))
     GetDirectorioActual = m_DirectorioActual
End Function

'Establecemos el Directorio Actual
'****************************************************
Public Sub CambiarDirectorio(PathDirectorio As String)
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long  'handle usado para buscar fichs en FTP
    Dim ret As Long
    Dim strDir As String
    
    strDir = Replace(m_DirectorioActual, Chr(0), "")
    
    If strDir = "/" And PathDirectorio = "../Subir un nivel" Then: Exit Sub
    RatonReloj
    
    If PathDirectorio = "../Subir un nivel" Then
      pos = InStrRev(strDir, "/")
      strDir = Left(strDir, pos)
        
      'Cambia al Directorio Ftp especificado
       ret = FtpSetCurrentDirectory(HandleConect, strDir)
       
       If ret = 0 Then
          Status "Error al cambiar de directorio."
       End If
       RatonNormal
       Exit Sub
    End If
    
   'Cambia al Directorio especificado
    ret = FtpSetCurrentDirectory(HandleConect, ftpDir & PathDirectorio)   ' & "/"
    'MsgBox ret
    If ret = 0 Then
       Status "Error al cambiar de directorio"
    End If
    RatonNormal
End Sub

'Crea un nuevo directorio
'*******************************************
Public Sub CrearDirectorio(NameDirectorio As String)

'Creamos un nuevo  directorio ('testing')
ret = FtpCreateDirectory(HandleConect, NameDirectorio)
If Not ret Then
   Status "Error al crear el directorio, compruebe el nombre que sea válido"
Else
   m_listView.ListItems.Add , , NameDirectorio, , "carpeta"
   m_listView.ListItems(m_listView.ListItems.Count).Selected = True
   m_listView.SetFocus
End If
   
End Sub

'Crea o sube un nuevo Archivo.
'********************************************
Public Sub SubirArchivo(localArchivo As String, NombreArchivoRemoto As String, Optional NoAddListView As Boolean)
'manda fichero al servidor FTP
ret = FtpPutFile(HandleConect, localArchivo, NombreArchivoRemoto, m_TipoTransferencia, 0)
If ret Then
   If Not NoAddListView Then
      m_listView.ListItems.Add , , NombreArchivoRemoto, , "archivo"
      m_listView.ListItems(m_listView.ListItems.Count).Selected = True
      m_listView.SetFocus
   End If
Else
   Status "Error al subir el fichero:" & NombreArchivoRemoto
End If
End Sub
        
'Renombra un archivo en el directorio Ftp corriente
'****************************************************
Public Sub RenombrarArchivo(Archivo As String, nuevoNombre As String)

'renombra 'test.htm' to 'apiguide.htm'
ret = FtpRenameFile(HandleConect, Archivo, nuevoNombre)

If ret Then
   m_listView.SelectedItem.Text = nuevoNombre
   m_listView.SetFocus
Else
   Status "Error al renombrar el fichero:" & nuevoNombre
End If

End Sub

Public Sub ObtenerArchivo(ArchivoRemoto As String, ArchivoLocal As String, _
           Optional SobreEscribir As Boolean = False)
'recupera fichero del servidor FTP: ArchivoRemoto es el nombre del archivo remoto
'ArchivoLocal es el nombre y ruta donde se colocará el archivo en local
ret = FtpGetFile(HandleConect, ArchivoRemoto, ArchivoLocal, SobreEscribir, 0, m_TipoTransferencia, 0)

If ret Then
   Status "Fichero: " & ArchivoRemoto & " descargado"
   'm_listView.SetFocus
Else
   Status "Error al descargar el fichero: " & ArchivoRemoto
End If

End Sub

'Eliminar Archivo del servidor Ftp
Public Sub EliminarArchivo(Archivo As String)
 'elimina el fichero del servidor FTP
 ret = FtpDeleteFile(HandleConect, Archivo)
 If Not ret Then
    Status "Error. No se pudo eliminar el archivo: " & Archivo
 End If
End Sub

Public Sub EliminarDirectorio(Directorio As String)
'elimina el directorio
ret = FtpRemoveDirectory(HandleConect, Directorio)
 If Not ret Then
    Status "Error. No se pudo eliminar el Directorio: " & Directorio
 End If
End Sub

Private Sub Status(Mensaje As String)
On Error GoTo SubError
ctrl = Mensaje
Exit Sub
SubError:
If Err.Number = 91 Then Resume Next
End Sub

Public Sub ListarArchivos()
Dim Item As ListItem
Dim pData As WIN32_FIND_DATA
Dim hFind As Long  'handle usado para buscar fichs en FTP
Dim ret As Long   'valor devuelto por API
    
    m_form.MousePointer = vbHourglass
    
    'crea buffer
    pData.cFileName = String(MAX_PATH, 0)
    
    'busca el primer fichero
    hFind = FtpFindFirstFile(HandleConect, "*.*", pData, 0, 0)
    
    m_listView.ListItems.Clear
    
    'Si Hfind vale 0 es porque no hay archivos ni directorios
    If hFind = 0 Then
       Set Item = m_listView.ListItems.Add(, , "../Subir un nivel", , "carpeta")
       Item.SubItems(2) = getFecha(pData)
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    
    Set Item = m_listView.ListItems.Add(, , "../Subir un nivel", , "carpeta")
    
    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "carpeta")
       Item.SubItems(2) = getFecha(pData)
    Else
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "archivo")
       Item.SubItems(1) = Round((pData.nFileSizeLow / 1024), 2) & " Kb"
       Item.SubItems(2) = getFecha(pData)
    End If
    
    'si no hay mas Archivos sale
    If hFind = 0 Then
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    Do
        'crea buffer
        pData.cFileName = String(MAX_PATH, 0) 'se llena con nulos
        'find the next file
        ret = InternetFindNextFile(hFind, pData)
        'si no hay ficheros, no sigue
        If ret = 0 Then Exit Do

        If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or _
                                   pData.dwFileAttributes = 0 Then
           'Agrega el nombre del directorio
           Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "carpeta")
           Item.SubItems(2) = getFecha(pData)
        Else
           'agrega el archivo y Muestra el tamaño del mismo en el LV
           Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "archivo")
           Item.SubItems(1) = Round((pData.nFileSizeLow / 1024), 2) & " Kb"
           Item.SubItems(2) = getFecha(pData)
        End If
    Loop
    'Cerramos el handle de búsqueda
    InternetCloseHandle hFind
    m_listView.Sorted = True
    m_form.MousePointer = vbDefault
End Sub


'Actualiza la lista de Archivos y directorios en el ListView
'************************************************************
Public Sub Actualizar()
 
Dim pData As WIN32_FIND_DATA
Dim hFind As Long  'handle usado para buscar fichs en FTP
Dim ret As Long   'valor devuelto por API
Dim Item As ListItem
    
    
    m_form.MousePointer = vbHourglass
    
    'crea buffer
    pData.cFileName = String(MAX_PATH, 0)
    
    'busca el primer fichero
    hFind = FtpFindFirstFile(HandleConect, "*.*", pData, 0, 0)
    
    m_listView.ListItems.Clear
    
    If hFind = 0 Then
       Set Item = m_listView.ListItems.Add(, , "../Subir un nivel", , "carpeta")
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
        
    Set Item = m_listView.ListItems.Add(, , "../Subir un nivel", , "carpeta")
    
    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "carpeta")
       Item.SubItems(2) = getFecha(pData)
    Else
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "archivo")
       Item.SubItems(1) = Round((pData.nFileSizeLow / 1024), 2) & " Kb"
       Item.SubItems(2) = getFecha(pData)
    End If
    
    'si no hay mas Archivos sale
    If hFind = 0 Then
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    Do
        'crea buffer
        pData.cFileName = String(MAX_PATH, 0) 'se llena con nulos
        'find the next file
        ret = InternetFindNextFile(hFind, pData)
        'si no hay ficheros, no sigue
        If ret = 0 Then Exit Do
        'Archivo

    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or pData.dwFileAttributes = 0 Then
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "carpeta")
       Item.SubItems(2) = getFecha(pData)
       
    Else
       Set Item = m_listView.ListItems.Add(, , pData.cFileName, , "archivo")

       Item.SubItems(1) = Round((pData.nFileSizeLow / 1024), 2) & " Kb"
       Item.SubItems(2) = getFecha(pData)
    End If
    
    Loop
    'Cerramos el handle de búsqueda
    InternetCloseHandle hFind
    m_listView.Sorted = True
    m_form.MousePointer = vbDefault
End Sub

Private Function getFecha(pData As WIN32_FIND_DATA) As Date
Dim stSystemTime As SYSTEMTIME

    If FileTimeToSystemTime(pData.ftLastWriteTime, stSystemTime) > 0 Then
        VBATime = DateSerial(stSystemTime.intYear, _
         stSystemTime.intMonth, _
        stSystemTime.intDay) + TimeSerial(stSystemTime.intHour, _
        stSystemTime.intMinute, stSystemTime.intSecond)
    End If
    getFecha = VBATime
End Function

Public Sub Inicializar(Formulario As Form)
    Set m_form = Formulario
End Sub

'Para mostrar los errores
'************************************************+
Private Sub ShowError()
Dim lngNumError As Long
Dim strMemoError As String
Dim lngTamBuffer As Long
    '-----------------------------
    'Tamaño del buffer
    InternetGetLastResponseInfo lngNumError, strMemoError, lngTamBuffer
    'crea buffer
    strMemoError = String(lngTamBuffer, 0)
    'Recupera informacion del error
    InternetGetLastResponseInfo lngNumError, strMemoError, lngTamBuffer
    'Mostrar el error en msgbox
    MsgBox "Error " & CStr(lngNumError) & ": " & strMemoError, vbOKOnly Or vbCritical
End Sub

'Nombre de usuario de la cuenta Ftp
'**********************************
Public Property Get Puerto() As String
Puerto = m_Puerto
End Property

Public Property Let Puerto(ByVal vNewValue As String)
m_Puerto = vNewValue
End Property

'Nombre de usuario de la cuenta Ftp
'**********************************
Public Property Get Usuario() As String
Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal vNewValue As String)
m_Usuario = vNewValue
End Property

'Nombre del servidor Ftp
'***********************
Public Property Get servidor() As String
servidor = m_Servidor
End Property

Public Property Let servidor(ByVal vNewValue As String)
m_Servidor = vNewValue
End Property

'Contraseña de la cuenta FTP
'***************************
Public Property Get Password() As String
Password = m_PassWord
End Property

Public Property Let Password(ByVal vNewValue As String)
m_PassWord = vNewValue
End Property

'Establece el ListView donde listar los ficheros
'***********************************************
Public Property Get ListView() As ListView
Set ListView = m_listView
End Property

Public Property Set ListView(ByVal vNewValue As ListView)
Set m_listView = vNewValue
End Property

Public Property Get ListBox() As ListBox
Set ListBox = m_listBox
End Property

Public Property Set ListBox(ByVal vNewValue As ListBox)
Set m_listBox = vNewValue
End Property

'Modo de Transferencia
'**********************************************
Public Property Get TipoTransferencia() As e_TipoTransferencia
    TipoTransferencia = m_TipoTransferencia
End Property

Public Property Let TipoTransferencia(NewData As e_TipoTransferencia)
    m_TipoTransferencia = NewData
End Property

Private Sub Class_Terminate()
On Local Error Resume Next
  'Cerramos la cesión FTP y la conexión a internet
  InternetCloseHandle HandleConect
  InternetCloseHandle hOpen
  'Eliminamos las variables de objeto
  Set ctrl = Nothing
  Set ListView = Nothing
  Set m_form = Nothing
End Sub

Public Sub Mostar_Estado_FTP(ProgressBarEstado As ProgressBar, LBEstado As ListBox)
Dim LongCad As Integer
Dim UnPunto As Integer
Dim AddEstado As String
Dim CadProgreso As String
Dim Porcentaje As Long

    RatonReloj
    With Progreso_Barra
         LongCad = Len(.Mensaje_Box)
         UnPunto = InStr(.Mensaje_Box, ".")
         
         AddEstado = ""
         If MiTiempo > 0 Then AddEstado = "[" & Format(Time - MiTiempo, FormatoTimes) & "] "
         If UnPunto > 0 Then
            AddEstado = AddEstado & MidStrg(.Mensaje_Box, 1, LongCad - 10) & "****" & MidStrg(.Mensaje_Box, LongCad - 5, LongCad - UnPunto - 1)
         Else
            AddEstado = AddEstado & .Mensaje_Box
         End If
         LBEstado.AddItem AddEstado
         LBEstado.Text = AddEstado
         LBEstado.Refresh
         DoEvents
    
         CadProgreso = ""
         If .Puntos < 0 Then .Puntos = 1
         If .Puntos > 20 Then .Puntos = 1
         If .color > 15 Then .color = 0
         If .Valor_Maximo <= 0 Then .Valor_Maximo = 1
         If .Incremento < .Valor_Maximo Then Porcentaje = .Incremento / .Valor_Maximo Else Porcentaje = 1
         If Porcentaje > 0 Then CadProgreso = "[" & Format(Porcentaje, "00%") & "] "
         If Len(.Mensaje_Box) > 1 Then
            CadProgreso = CadProgreso & AddEstado
         Else
            CadProgreso = CadProgreso & "Espere un momento, procesando"
         End If
         CadProgreso = CadProgreso & String(.Puntos, ".")
         If Procesando > 10 Then Procesando = 0
       ' Establece el color del progress
         Color_Progreso ProgressBarEstado.hwnd, QBColor(.color)
         ProgressBarEstado.Max = .Valor_Maximo
         If .Incremento <= .Valor_Maximo Then ProgressBarEstado.value = .Incremento
        .Incremento = .Incremento + 1
        .Puntos = .Puntos + 1
        .color = .color + 1
         Procesando = Procesando + 1
     End With
End Sub

