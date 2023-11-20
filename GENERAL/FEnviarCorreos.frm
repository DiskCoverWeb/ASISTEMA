VERSION 5.00
Begin VB.Form FEnviarCorreos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "."
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   9780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FEnviarCorreos.frx":0000
   ScaleHeight     =   1200
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   1260
      TabIndex        =   0
      Top             =   105
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   210
      Picture         =   "FEnviarCorreos.frx":40B7
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "FEnviarCorreos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Dim AdoSMTP As ADODB.Recordset

Dim NumFile As Long
Dim nFrames As Long
Dim Indx As Long
Dim IndD As Long

Dim LineaFile As Integer

Dim AnchoMaxForm As Single

Dim TextFile(8) As String
Dim RutaFile As String
Dim Temp As Variant

Dim Si_Enviar As Boolean

'Dim Parpadear As Boolean

'Empieza enviar correos
Private Sub Form_Activate()
Dim Tiempo_Espera As Single
Dim MiTiempo_Espera As Single
Dim MsgAux As String
Dim EMailPara As String
Dim Emails As String
Dim Contactos As String
Dim posPuntoComa As String

 If Si_Enviar Then
   ''Timer1_Timer
   'Activamos la clase para envios de los mails
    Set oMail = New clsCDOmail
    If TMail.Asunto = "" Then TMail.Asunto = "Sin asunto"
    
   'Datos para enviar
    oMail.servidor = TMail.servidor                       ' smtp.gmail.com
    oMail.Usuario = TMail.Usuario
    oMail.Password = TMail.Password
    oMail.Puerto = TMail.Puerto                           ' 465
    oMail.useAuntentificacion = TMail.useAuntentificacion ' True
    oMail.ssl = TMail.ssl                                 ' True
    oMail.tls = TMail.tls                                 ' True
   '---------------------------------------------------------------
    oMail.Asunto = TMail.Asunto
    oMail.Adjunto = TMail.Adjunto
    oMail.de = TMail.de
   'Verificamos que el mail no este vacio
    With TMail
        'Mensaje automatizado de
         Contactos = ""
         Insertar_Cadena Contactos, Telefono1
         Insertar_Cadena Contactos, Telefono2
         If Len(RazonSocial) > 1 Then MsgAux = RazonSocial Else MsgAux = Empresa
         
         'MsgBox InStr(.Mensaje, "Este correo electrónico fue generado automáticamente del Sistema Financiero")
         If InStr(.Mensaje, "Este correo electronico fue generado automaticamente del Sistema Financiero") = 0 Then
           .Mensaje = .Mensaje & vbCrLf & MensajeAutomatizado
           .Mensaje = Replace(.Mensaje, "Nombre_Usuario", NombreUsuario)
           .Mensaje = Replace(.Mensaje, "Mensaje_Comunicado", ComunicadoEntidad)
           .Mensaje = Replace(.Mensaje, "Representante_Legal", NombreGerente)
           .Mensaje = Replace(.Mensaje, "Numero_Telefono", Contactos)
           .Mensaje = Replace(.Mensaje, "Emails", EmailProcesos)
           .Mensaje = Replace(.Mensaje, "Razon_Social", MsgAux)
        End If
     End With

     'Timer1_Timer
    'Empezamos a enviar el mails
     With oMail
         'Datos para enviar
         .Mensaje = TMail.Mensaje
         .MensajeHTML = TMail.MensajeHTML
          Emails = TMail.para
          
         'MsgBox "DE: " & oMail.de & vbCrLf & "PARA: " & Emails
          
          If MidStrg(Emails, Len(Emails), 1) <> ";" Then Emails = Emails & ";"
          If Len(Emails) > 3 Then
             Do While Len(Emails) > 3
                posPuntoComa = InStr(Emails, ";")
                EMailPara = MidStrg(Emails, 1, posPuntoComa - 1)
               'MsgBox "Lista: " & emails
                If EsUnEmail(EMailPara) Then
                   'Timer1_Timer
                   Label1.Caption = "Remitente: " & .de & vbCrLf _
                                  & "Para: " & EMailPara & vbCrLf _
                                  & "Asunto: " & .Asunto
                   Label1.Refresh
                   'MsgBox Label1.Caption
                  'MsgBox "Email: " & Email & vbCrLf & RutaXML
                  .para = EMailPara
                  'Metodo manda el mail
                  .Enviar_Backup
                End If
                Emails = MidStrg(Emails, posPuntoComa + 1, Len(Emails))
             Loop
          End If
     End With
     Set oMail = Nothing
     'Timer1_Timer
     RatonNormal
    'If Len(TMail.ListaError) > 1 Then MsgBox TMail.ListaError
     Unload FEnviarCorreos
     'MsgBox "Enviado"
 Else
     'Timer1_Timer
     Set oMail = Nothing
    'If Len(TMail.ListaError) > 1 Then MsgBox TMail.ListaError
     Unload FEnviarCorreos
 End If
   'MsgBox "..."
End Sub

Private Sub Form_Load()
Dim CadAncho As String

    RatonReloj
    CentrarForm FEnviarCorreos
    Redondear_Formulario FEnviarCorreos, 40
    Si_Enviar = False
    FEnviarCorreos.width = 2000 + FEnviarCorreos.TextWidth(String(Len(" Remitente: " & TMail.de & " "), "_"))
    FEnviarCorreos.Refresh
    AnchoMaxForm = FEnviarCorreos.width
    
    'CadAncho = "ASUNTO: " & TMail.Asunto & "__"
    'If AnchoMaxForm < FEnviarCorreos.TextWidth(CadAncho) Then AnchoMaxForm = FEnviarCorreos.TextWidth(CadAncho)
    'FEnviarCorreos.width = AnchoMaxForm
    'FEnviarCorreos.Refresh
    Label1.width = AnchoMaxForm - Label1.Left - 200
    
    Label1.Caption = "CONECTANDOSE AL SERVIDOR" & vbCrLf & vbCrLf _
                   & "DE CORREOS ELECTRONICOS"
    Label1.Refresh
    
    nFrames = Load_Gif(RutaSistema & "\FORMATOS\MAILS.gif", Image1)
    If nFrames > 0 Then
       FrameCount = 0
       Timer1.Interval = 20
       Timer1.Enabled = True
    End If
   'ErrorMails = ""
   'Determinamos si esta activado envio de correos
   'Timer1_Timer
   
    sSQL = "SELECT smtp_Servidor, smtp_Puerto, smtp_UseAuntentificacion, smtp_SSL, smtp_Secure, " _
         & "Email_Conexion, Email_Contraseña, Email_Conexion_CE, Email_Contraseña_CE, Email_CE_Copia " _
         & "FROM Empresas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND LEN(smtp_Servidor) > 1 " _
         & "AND smtp_Puerto > 0 "
    Select_AdoDB AdoSMTP, sSQL
   'MsgBox "..."
    With AdoSMTP
     If .RecordCount > 0 Then
         TMail.useAuntentificacion = CBool(.fields("smtp_UseAuntentificacion"))
         TMail.ssl = CBool(.fields("smtp_SSL"))
         TMail.tls = CBool(.fields("smtp_Secure"))
         TMail.Puerto = .fields("smtp_Puerto")
         Email_CE_Copia = CBool(.fields("Email_CE_Copia"))
         TMail.servidor = .fields("smtp_Servidor")
         If TMail.TipoDeEnvio = "CE" Then
            TMail.Usuario = .fields("Email_Conexion_CE")
            TMail.Password = .fields("Email_Contraseña_CE")
            TMail.de = .fields("Email_Conexion_CE")
         Else
            TMail.Usuario = .fields("Email_Conexion")
            TMail.Password = .fields("Email_Contraseña")
            TMail.de = .fields("Email_Conexion")
         End If
         If Email_CE_Copia Then Insertar_Mail Emails, EmailProcesos
         If TMail.de = "" And 0 <= TMail.ListaMail And TMail.ListaMail <= 6 Then TMail.de = Lista_De_Correos(TMail.ListaMail).Correo_Electronico
         TMail.de = Replace(UCase(Empresa), """", "") & " <" & TMail.de & ">"
        'Si utilizamos el correo de DiskCover System
         If TMail.servidor = "mail.diskcoversystem.com" Then
            TMail.servidor = "smtp.diskcoversystem.com"
            TMail.ehlo = "smtp.diskcoversystem.com"
            TMail.ssl = False
            TMail.tls = True
            TMail.Usuario = "admin"
            TMail.Password = "Admin@2023"
            TMail.Puerto = 26
            TMail.de = Replace(TMail.de, "diskcoversystem.com", "smtp.diskcoversystem.com")
            
'''         TMail.servidor = "relay.dnsexit.com"
'''         TMail.Usuario = "diskcoversystem"
'''         TMail.Password = "Dlcjvl1210@"
'''         TMail.puerto = 25                '25,26,80,587,940,2525,8001
         End If
     End If
    End With
    AdoSMTP.Close
    If TMail.de <> "" And TMail.para <> "" Then
       Si_Enviar = True
    Else
       TMail.ListaError = TMail.ListaError = ". Credenciales no asignadas para el envio de Correos electronicos, solicite ayuda al Administrador del Sistema"
    End If
End Sub

' envio completo
Private Sub oMail_EnvioCompleto()
Dim MiTiempo_Espera As Single
Dim Tiempo_Espera As Single
    RatonNormal
   'MsgBox "Mensaje enviado", vbInformation
    
    Control_Procesos "M", TMail.para, TMail.Destinatario & "- Mensaje enviado Correctamente"
   'MsgBox TMail.Destinatario & "..."
    If Len(TMail.Adjunto) <= 1 Then
       MiTiempo_Espera = 0
       Tiempo_Espera = Time
       Do While MiTiempo_Espera <= 0.03
          Minutos = Time
          Segundos = Second(Minutos - Tiempo_Espera)
          Minutos = Minute(Minutos - Tiempo_Espera)
          MiTiempo_Espera = CSng(Format$(Minutos, "00") & "." & Format$(Segundos, "00"))
         'MsgBox "[" & contadorEmail & "] Tiempo Espera: " & MiTiempo_Espera
       Loop
    End If
   'Unload FEnviarCorreos

End Sub

' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    RatonNormal
    'MsgBox Descripcion, vbCritical, Numero
    TMail.ListaError = TMail.ListaError & "Error: Para: " & TMail.para & "(" & TMail.Destinatario & "), " & Descripcion & vbCrLf
    Control_Procesos "M", TMail.para, TMail.Destinatario & "-" & Descripcion
   'Unload FEnviarCorreos
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim I As Long

If FrameCount < TotalFrames Then
   Image1(FrameCount).Visible = False
   FrameCount = FrameCount + 1
Else
   FrameCount = 0
   For I = 1 To Image1.Count - 1
       Image1(I).Visible = False
   Next I
End If
Image1(FrameCount).Visible = True
TMail.ContadorTiempo = TMail.ContadorTiempo + 1
If TMail.ContadorTiempo > 2 Then TMail.ContadorTiempo = 0
Select Case TMail.ContadorTiempo
  Case 0: Label1.ForeColor = Amarillo_Claro ' &HC00000
  Case 1: Label1.ForeColor = Azul
  Case 2: Label1.ForeColor = Blanco_Claro
End Select
Label1.Refresh
'MsgBox Label1.ForeColor

If Err Then Exit Sub
End Sub

