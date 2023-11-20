VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDOmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'025150475

'Variables Locales
Private mServidor As String
Private mEhlo As String
Private mDe As String
Private mPara As String
Private mAsunto As String
Private mMensaje As String
Private mMensajeHTML As String
Private mAdjunto As String
Private mPuerto As Integer ' Variant
Private mUsuario As String
Private mPassword As String
Private mUseAuntentificacion As Boolean
Private mSSL As Boolean
Private mTLS As Boolean

Public Event Error(Descripcion As String, Numero As Variant)
Public Event EnvioCompleto()

Function Enviar_Backup() As Boolean
 On Error GoTo GestionaError
   'Variable de objeto Cdo.Message
    Dim oCDO As Object
    Dim archivoAdjunto As String
    Dim posPuntoComa As String
    Dim vAdjunto As String

   'chequea si hay conexión
    If InternetGetConnectedState(0&, 0&) = False Then
       RaiseEvent Error("No se puede enviar el correo. " & "Verificar la conexión a internet si está disponible", 0)
       Exit Function
    End If

   'chequea que el puerto sea un número, o que no esté vacío
    If Not IsNumeric(Puerto) Then
       RaiseEvent Error("No se ha indicado el puerto del servidor", 0)
       Exit Function
    End If

   'Crea un Nuevo objeto CDO.Message
    Set oCDO = CreateObject("CDO.Message")
    With oCDO.configuration
        'SMTP server IP or Name
       .fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").value = mServidor
        'Indica el servidor Smtp para poder enviar el Mail (puede ser el nombre del servidor o su dirección IP)
       .fields("http://schemas.microsoft.com/cdo/configuration/sendusing").value = 2
        'Server port (typically 25, 465, 587) 'The next line is commented out because of the use of SSL encryption
        'Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usa el puerto 465
       .fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").value = mPuerto    'Commonly -> 465 SSL, 587 TLS
        'Enviar un mensaje al servdor
        If Len(mEhlo) > 1 Then .fields("http://schemas.microsoft.com/cdo/configuration/ehlo").value = mEhlo
        'Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
        'Use SSL for the connection (False or True) -> If using SSL, do not specify the Port above
        If mSSL Then .fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl").value = mSSL
        'Use TLS for the connection (False or True)
        If mTLS Then .fields("http://schemas.microsoft.com/cdo/configuration/sendtls").value = mTLS
        'Configura las opciones para el login en el SMTP
        'SMTP Account User ID
        If mUseAuntentificacion Then
          'Id de usuario del servidor Smtp (en el caso de gmail, debe ser la dirección de correro mas el @gmail.com)
         .fields("http://schemas.microsoft.com/cdo/configuration/sendusername").value = mUsuario
          'Password de la cuenta
         .fields("http://schemas.microsoft.com/cdo/configuration/sendpassword").value = mPassword
        End If
        'Type of authentication, NONE, Basic (Base64 encoded), NTLM
        'Indica el tipo de autentificación con el servidor de correo.
        'El valor 0 no requiere autentificarse, el valor 1 es con autentificación
       .fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").value = Abs(mUseAuntentificacion)
        'Tiempo máximo de espera en segundos para la conexión
        'Number of seconds to wait for a response from the server before aborting
       .fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout").value = 10
        'Cuerpo del mensaje
       .fields("http://schemas.microsoft.com/cdo/configuration/BodyFormat").value = 2
        'Actualiza los datos antes de enviar
       .fields.Update
    End With
   'Estructura del mail
   '''''''''''''''''''''''''''''''''''''''''''''''
   'Dirección del Destinatario
    oCDO.To = mPara

   'Dirección del remitente
    oCDO.From = mDe

   'Asunto del mensaje
    oCDO.Subject = mAsunto

   'Cuerpo del mensaje
    If Len(mMensajeHTML) > 10 Then
       oCDO.HTMLBody = mMensajeHTML
       oCDO.TextBody = ""
    Else
       oCDO.TextBody = mMensaje
    End If
   'Ruta del archivo adjunto
   'MsgBox mSSL      'mAdjunto
    vAdjunto = TrimStrg(mAdjunto)
    If Len(FirmaDigital) > 1 Then vAdjunto = vAdjunto & ";" & FirmaDigital
   'MsgBox vAdjunto & vbCrLf & FirmaDigital
    If vAdjunto <> "" Then
       Do While Len(vAdjunto) >= 3
          posPuntoComa = InStr(vAdjunto, ";")
          If posPuntoComa = 0 Then posPuntoComa = Len(vAdjunto) + 1
          archivoAdjunto = TrimStrg(MidStrg(vAdjunto, 1, posPuntoComa - 1))
         'MsgBox "Varios: " & archivoAdjunto
         '..lo agrega
          If Len(Dir(archivoAdjunto)) <> 0 Then oCDO.AddAttachment (archivoAdjunto)
          vAdjunto = TrimStrg(MidStrg(vAdjunto, posPuntoComa + 1, Len(vAdjunto)))
       Loop
    End If
   
    On Error Resume Next
    Screen.MousePointer = vbHourglass

   'Envía el email
    If Len(mServidor) > 1 And Len(mUsuario) > 1 And Len(mPassword) > 1 And Len(mDe) > 1 And Len(mPara) > 1 Then
       Cadena = "SERVIDOR: " & mServidor & vbCrLf & vbCrLf _
              & "USER: " & mUsuario & vbCrLf & vbCrLf _
              & "PWD: " & mPassword & vbCrLf & vbCrLf _
              & "PUERTO: " & mPuerto & vbCrLf & vbCrLf _
              & "DE: " & mDe & vbCrLf & vbCrLf _
              & "PARA: " & mPara
      'MsgBox Cadena
       oCDO.send
GestionaError:
      'MsgBox Err.Number
       ' .. si no hubo error
       If Err.Number = 0 Then
          Enviar_Backup = True
          RaiseEvent EnvioCompleto
       Else
          Select Case Err.Number
            Case -2147220973
                 Screen.MousePointer = vbNormal
                 RaiseEvent Error("Posible error : nombre del Servidor incorrecto o número de puerto incorrecto", Err.Number)
            Case -2147220975
                 Screen.MousePointer = vbNormal
                 RaiseEvent Error("Posible error : error en el nombre de usuario, o en el password ", Err.Number)
            Case Else
                 Screen.MousePointer = vbNormal
                 RaiseEvent Error(Err.Description, Err.Number)
                 Resume Next
          End Select
       End If
       
     ' Descarga la referencia
       If Not oCDO Is Nothing Then Set oCDO = Nothing
       Err.Clear
    End If
    Screen.MousePointer = vbNormal
End Function

' propiedades
'''''''''''''''''''''
Property Get servidor() As String
    servidor = mServidor
End Property

Property Let servidor(value As String)
    mServidor = value
End Property

Property Get ehlo() As String
    servidor = mEhlo
End Property

Property Let ehlo(value As String)
    mEhlo = value
End Property

Property Get para() As String
    para = mPara
End Property

Property Let para(value As String)
    mPara = value
End Property

Property Get de() As String
    de = mDe
End Property

Property Let de(value As String)
    mDe = value
End Property

Property Get Asunto() As String
    Asunto = mAsunto
End Property

Property Let Asunto(value As String)
    mAsunto = value
End Property

Property Get Mensaje() As String
    Mensaje = mMensaje
End Property

Property Let Mensaje(value As String)
    mMensaje = value
End Property

Property Get MensajeHTML() As String
    MensajeHTML = mMensajeHTML
End Property

Property Let MensajeHTML(value As String)
    mMensajeHTML = value
End Property

Property Get Adjunto() As String
    Adjunto = mAdjunto
End Property

Property Let Adjunto(value As String)
    mAdjunto = value
End Property

Property Get Puerto() As Variant
    Puerto = mPuerto
End Property

Property Let Puerto(value As Variant)
    mPuerto = value
End Property

Property Get Usuario() As String
    Usuario = mUsuario
End Property

Property Let Usuario(value As String)
    mUsuario = value
End Property

Property Get Password() As String
    Password = mPassword
End Property

Property Let Password(value As String)
    mPassword = value
End Property

Property Get useAuntentificacion() As Boolean
    useAuntentificacion = mUseAuntentificacion
End Property

Property Let useAuntentificacion(value As Boolean)
    mUseAuntentificacion = value
End Property

Property Get ssl() As Boolean
    ssl = mSSL
End Property

Property Let ssl(value As Boolean)
    mSSL = value
End Property

Property Get tls() As Boolean
    tls = mTLS
End Property

Property Let tls(value As Boolean)
    mTLS = value
End Property

