VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form FEnviarCorreos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "."
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   10650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FEnviarCorreos.frx":0000
   ScaleHeight     =   1335
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstStatud 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   105
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   6315
   End
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   6615
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgLstFTP"
      SmallIcons      =   "ImgLstFTP"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Archivos"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tama�o"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modificado"
         Object.Width           =   2646
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   8715
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":4D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":5082
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":539C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":56A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":59BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":5CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":5FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":67E2
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":6AFC
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":6E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":7054
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnviarCorreos.frx":736E
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   1065
      Left            =   1260
      TabIndex        =   0
      Top             =   105
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   210
      Picture         =   "FEnviarCorreos.frx":7688
      Top             =   315
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
Dim Si_Enviar As Boolean
Dim nFrames As Long
Dim ContMails As Long
Dim I As Integer
Dim Contactos As String
Dim MsgAux As String

Public Function ReplaceHTML(vExpression As String, vFind As String, vReplace As String) As String
    'Cursor_Img
    If Len(vReplace) > 1 Then
       ReplaceHTML = Replace(vExpression, vFind, vReplace)
    Else
       ReplaceHTML = Replace(vExpression, vFind, "")
    End If
End Function

'Empieza enviar correos
Private Sub Form_Activate()
Dim Tiempo_Espera As Single
Dim MiTiempo_Espera As Single

Dim EmailsCliente As String
Dim EMailPara As String
Dim File As String
Dim Files As String
Dim DirFiles() As String
Dim DirFilesFTP() As String
Dim TextoHTML As String
Dim sOrigen As String
Dim sDestino As String

Dim CaracPiloto As Integer
Dim ContFile As Integer

 If Si_Enviar Then
    With TMail
      If .Asunto = "" Then .Asunto = "Sin asunto"
      If .Remitente = "" Then .Remitente = Replace(Empresa, """", "")
      If Len(.MensajeHTML) > 1 Then
         TextoHTML = .MensajeHTML
      Else
         TextoHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_mail_basico.html")
         If Len(.Mensaje) > 1 Then html_Informacion_adicional = .Mensaje Else html_Informacion_adicional = ""
      End If
      
      If UCase(RazonSocial) = UCase(NombreComercial) Then MsgAux = "" Else MsgAux = NombreComercial
      
      If Contactos = "" Then
         Insertar_Cadena Contactos, Telefono1
         Insertar_Cadena Contactos, Telefono2
         Contactos = MidStrg(Contactos, 1, Len(Contactos) - 1)
      End If
      
      EmailsCliente = ""
      Insertar_Cadena EmailsCliente, FA.EmailC
      Insertar_Cadena EmailsCliente, FA.EmailC2
      Insertar_Cadena EmailsCliente, FA.EmailR
      EmailsCliente = MidStrg(EmailsCliente, 1, Len(EmailsCliente) - 1)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vInformacion_adicional", html_Informacion_adicional)
      TextoHTML = ReplaceHTML(TextoHTML, "vDetalle_adicional", html_Detalle_adicional)

      TextoHTML = ReplaceHTML(TextoHTML, "vMensajeFinal", MensajeAutomatizado)
      TextoHTML = ReplaceHTML(TextoHTML, "vMensajeDeboPagare", MensajeDeboPagare)
      TextoHTML = ReplaceHTML(TextoHTML, "vMensajeEmpresa", MensajeEmpresa)
      TextoHTML = ReplaceHTML(TextoHTML, "vMensaje_Comunicado", ComunicadoEntidad)
      TextoHTML = ReplaceHTML(TextoHTML, "vEmails_Copia", EmailProcesos)

      TextoHTML = ReplaceHTML(TextoHTML, "vLogoTipo", NLogoTipo)
      TextoHTML = ReplaceHTML(TextoHTML, "vNombre_Usuario", NombreUsuario)
      TextoHTML = ReplaceHTML(TextoHTML, "vRepresentante_Legal", NombreGerente)
      TextoHTML = ReplaceHTML(TextoHTML, "vNumero_Telefono", Contactos)
      TextoHTML = ReplaceHTML(TextoHTML, "vRUC_Empresa", RUC)

      TextoHTML = ReplaceHTML(TextoHTML, "vRazon_Social", RazonSocial)
      TextoHTML = ReplaceHTML(TextoHTML, "vNombre_Comercial", MsgAux)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vDireccion_Empresa", Direccion)
      TextoHTML = ReplaceHTML(TextoHTML, "vEmail_Empresa", EmailEmpresa)
      TextoHTML = ReplaceHTML(TextoHTML, "vObligado_Contabilidad", Obligado_Conta)
        
      TextoHTML = ReplaceHTML(TextoHTML, "vFecha_Reporte", Mifecha)
      TextoHTML = ReplaceHTML(TextoHTML, "vCiudad_Empresa", NombreCiudad)
      TextoHTML = ReplaceHTML(TextoHTML, "vPais_Empresa", NombrePais)
              
      TextoHTML = ReplaceHTML(TextoHTML, "vGrupo_Cliente", FA.Grupo)
      TextoHTML = ReplaceHTML(TextoHTML, "vNombre_Cliente", FA.Cliente)
      TextoHTML = ReplaceHTML(TextoHTML, "vNombre_Representante", FA.Razon_Social)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vFecha_FA_V", FA.Fecha_V)
      TextoHTML = ReplaceHTML(TextoHTML, "vFecha_FA", FA.Fecha)
      TextoHTML = ReplaceHTML(TextoHTML, "vFecha_RE", FA.Fecha_R)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vHora_FA", FA.Hora)
      TextoHTML = ReplaceHTML(TextoHTML, "vHora_RE", FA.Hora_R)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vRUC_Cliente", FA.CI_RUC)
      TextoHTML = ReplaceHTML(TextoHTML, "vRUC_Representante", FA.RUC_CI)
      TextoHTML = ReplaceHTML(TextoHTML, "vDireccion_Cliente", FA.DireccionC)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vSerie_Cliente", FA.Serie)
      TextoHTML = ReplaceHTML(TextoHTML, "vFactura_Cliente", Format(FA.Factura, "000000000"))
      
      TextoHTML = ReplaceHTML(TextoHTML, "vAutorizacion_Retencion", FA.Autorizacion_R)
      TextoHTML = ReplaceHTML(TextoHTML, "vAutorizacion_Factura", FA.Autorizacion)
      TextoHTML = ReplaceHTML(TextoHTML, "vContacto", FA.Contacto)
      TextoHTML = ReplaceHTML(TextoHTML, "vOrden_Compra", FA.Orden_Compra)
      TextoHTML = ReplaceHTML(TextoHTML, "vObservacion", FA.Observacion)
      TextoHTML = ReplaceHTML(TextoHTML, "vNota", FA.Nota)
      TextoHTML = ReplaceHTML(TextoHTML, "vEmails_Cliente", EmailsCliente)
      TextoHTML = ReplaceHTML(TextoHTML, "vTelefonos_Cliente", FA.TelefonoC)
      TextoHTML = ReplaceHTML(TextoHTML, "vRecibo_No", FA.Recibo_No)
      
      TextoHTML = ReplaceHTML(TextoHTML, "vTipo_Documento_RE", FA.TR)
      TextoHTML = ReplaceHTML(TextoHTML, "vTipo_Documento_FA", FA.TC)
      
      If FA.Factura > 0 Then
         TextoHTML = ReplaceHTML(TextoHTML, "vSerie_Documento_FA", FA.Serie & "-" & Format(FA.Factura, "000000000"))
      Else
         TextoHTML = ReplaceHTML(TextoHTML, "vSerie_Documento_FA", "")
      End If
      If FA.Retencion > 0 Then
         TextoHTML = ReplaceHTML(TextoHTML, "vSerie_Documento_RE", FA.Serie_R & "-" & Format(FA.Retencion, "000000000"))
      Else
         TextoHTML = ReplaceHTML(TextoHTML, "vSerie_Documento_RE", "")
      End If
      
      TextoHTML = ReplaceHTML(TextoHTML, "vValor_Total", Format(ValorTotal, "#,##0.00"))
            
      TextoHTML = ReplaceHTML(TextoHTML, vbCrLf, "<br>")
      TextoHTML = ReplaceHTML(TextoHTML, "<N>", "<strong>")
      TextoHTML = ReplaceHTML(TextoHTML, "</N>", "</strong>")
      Cursor_Img
     .MensajeHTML = TextoHTML
     .Mensaje = ""
      Cursor_Img
    End With
    TMail.Adjunto = Replace(TMail.Adjunto, vbCrLf, "")
    TMail.Adjunto = Replace(TMail.Adjunto, " ", "")
    Label1.Caption = "Remitente: " & TMail.de & String(86 - Len(TMail.de), " ") & "Cuota Diaria: " & Format(ContMails / 6000, "00%") & vbCrLf _
                   & "Para: " & Replace(TMail.para, ";", "; ") & vbCrLf & vbCrLf _
                   & "Asunto: " & TMail.Asunto
    Cursor_Img
    If TMail.servidor = "imap.diskcoversystem.com" Then
      'Obtenemos la ruta inicial de donde vienen los archivos
       Files = TMail.Adjunto
      'MsgBox "Ruta Origen: " & Files
       ContFile = -1
       If Len(Files) > 1 Then
          CaracPiloto = InStr(Files, ";")
          If CaracPiloto > 0 Then
            'Son varios archivos que se envian
             Do While Len(Files) > 1 And CaracPiloto > 0
                Cursor_Img
                File = MidStrg(Files, 1, CaracPiloto - 1)
                If Existe_File(File) Then
                   ContFile = ContFile + 1
                   ReDim Preserve DirFiles(ContFile) As String
                   ReDim Preserve DirFilesFTP(ContFile) As String
                   DirFiles(ContFile) = File
                   DirFilesFTP(ContFile) = MidStrg(File, InStrRev(File, "\") + 1, Len(File))
                End If
                Files = MidStrg(Files, CaracPiloto + 1, Len(Files))
                CaracPiloto = InStr(Files, ";")
                Cursor_Img
             Loop
             If Existe_File(Files) Then
                ContFile = ContFile + 1
                ReDim Preserve DirFiles(ContFile) As String
                ReDim Preserve DirFilesFTP(ContFile) As String
                DirFiles(ContFile) = Files
                DirFilesFTP(ContFile) = MidStrg(Files, InStrRev(Files, "\") + 1, Len(Files))
             End If
          Else
            'Es un solo archivo
             If Existe_File(Files) Then
                Cursor_Img
                ContFile = ContFile + 1
                ReDim Preserve DirFiles(ContFile) As String
                ReDim Preserve DirFilesFTP(ContFile) As String
                DirFiles(ContFile) = Files
                DirFilesFTP(ContFile) = MidStrg(Files, InStrRev(Files, "\") + 1, Len(Files))
             End If
          End If
       End If
       Cursor_Img
       TMail.Adjunto = ""
       If ContFile >= 0 Then
          For I = 0 To UBound(DirFiles)
              Cursor_Img
              TMail.Adjunto = TMail.Adjunto & DirFilesFTP(I) & ";"
          Next I
          TMail.Adjunto = MidStrg(TMail.Adjunto, 1, Len(TMail.Adjunto) - 1)
          Cursor_Img
          
         'Conectamos al servidor ERP
          With ftp
            If ContFile >= 0 Then
              .Inicializar FEnviarCorreos
              .Password = ftpPwr  'Le establecemos la contrase�a de la cuenta Ftp
              .Usuario = ftpUse   'Le establecemos el nombre de usuario de la cuenta
              .servidor = ftpSvr  'Establecesmo el nombre del Servidor FTP
               Set .ListView = LstVwFTP
              .ConStatus = True
               If Not .ConectarFtp(LstStatud) Then
                  TMail.Adjunto = ""
               Else
                 'Subiendo documento firmado para su autorizacion
                  For I = 0 To UBound(DirFilesFTP)
                      Cursor_Img
                      sOrigen = DirFiles(I)
                      sDestino = "/files/AddAttachment/" & DirFilesFTP(I)
                     ' MsgBox "Subiendo: |" & sOrigen & "|" & vbCrLf & "|" & sDestino & "|"
                     .SubirArchivo DirFiles(I), "/files/AddAttachment/" & DirFilesFTP(I), True
                  Next I
               End If
            End If
          End With
       End If
       If InStrRev(TMail.para, ";") > 0 Then TMail.para = MidStrg(TMail.para, 1, Len(TMail.para) - 1)
       URLHTTP = "https://erp.diskcoversystem.com/lib/phpmailer/EnvioEmailvisual.php?EnviarVisual"
       URLParams = "from=" & TMail.de & "" _
                 & "&fromName=" & TMail.Remitente & " <" & TMail.de & ">" _
                 & "&to=" & TMail.para & "" _
                 & "&subject=" & TMail.Asunto & "" _
                 & "&HTML=1" _
                 & "&body=" & Sin_Signos_XML(TMail.MensajeHTML) & "" _
                 & "&Archivo=" & TMail.Adjunto & "" _
                 & "&reply=&replyName=&debug=0 "
       Si_No = PostUrlSource(URLHTTP, URLParams)
      ' Clipboard.Clear
      ' Clipboard.SetText URLParams
      'MsgBox Si_No ' & vbCrLf & URLParams ' TMail.MensajeHTML
       If Si_No Then
          Cursor_Img
          EMailPara = TMail.para
          CaracPiloto = InStr(EMailPara, ";")
          If CaracPiloto > 0 Then
             Do While Len(EMailPara) > 1 And CaracPiloto > 0
                File = MidStrg(EMailPara, 1, CaracPiloto - 1)
                Control_Procesos "EF", "Email: " & TMail.de & " => " & File, "Asunto: " & TMail.Asunto
                EMailPara = MidStrg(EMailPara, CaracPiloto + 1, Len(EMailPara))
                CaracPiloto = InStr(EMailPara, ";")
                Cursor_Img
             Loop
             Control_Procesos "EM", "Email: " & TMail.de & " => " & EMailPara, "Asunto: " & TMail.Asunto
          Else
             Control_Procesos "EM", "Email: " & TMail.de & " => " & EMailPara, "Asunto: " & TMail.Asunto
          End If
       Else
          Control_Procesos "EM", "Email: " & TMail.de & " => " & TMail.para, "Asunto(Error): " & TMail.Asunto
       End If
       Cursor_Img
      'Eliminando archivos que se fueron con los correos
       If ContFile >= 0 Then
         'Le indicamos el ListView donde se listar�n los archivos
          For I = 0 To UBound(DirFilesFTP)
              Cursor_Img
              ftp.EliminarArchivo "/files/AddAttachment/" & DirFilesFTP(I)
          Next I
          Cursor_Img
          ftp.Desconectar
       End If
    Else
      'Activamos la clase para envios de los mails
       Set oMail = New clsCDOmail
      'Empezamos a enviar el mails
       Cursor_Img
       With oMail
           'Datos para enviar
           .servidor = TMail.servidor                       ' smtp.gmail.com
           .Usuario = TMail.Usuario
           .Password = TMail.Password
           .Puerto = TMail.Puerto                           ' 465
           .useAuntentificacion = TMail.useAuntentificacion ' True
           .ssl = TMail.ssl                                 ' True
           .tls = TMail.tls                                 ' True
           '.ehlo = TMail.ehlo
           '---------------------------------------------------------------
           .Asunto = TMail.Asunto
           .Adjunto = TMail.Adjunto
           .de = TMail.de
           .Mensaje = TMail.Mensaje
           .MensajeHTML = TMail.MensajeHTML
           .para = TMail.para
            If Len(TMail.para) > 3 Then
              .de = TMail.Remitente & " <" & TMail.de & ">"
              .Enviar_Backup
               Cursor_Img
               Control_Procesos "EX", "Email: " & TMail.de & " => " & TMail.para, "Asunto: " & TMail.Asunto
               Cursor_Img
            End If
       End With
       Set oMail = Nothing
    End If
   'If Len(TMail.ListaError) > 1 Then MsgBox TMail.ListaError
 End If
 Cursor_Img
 RatonNormal
'MsgBox "......."
 Unload FEnviarCorreos
End Sub

Private Sub Form_Load()
Dim AdoSMTP As ADODB.Recordset
Dim CadAncho As String
Dim CantCadAncho As Long
Dim IdImg As Integer

    RatonReloj
    CentrarForm FEnviarCorreos
    Redondear_Formulario FEnviarCorreos, 40
    nFrames = Load_Gif(RutaSistema & "\FORMATOS\MAILS.gif", Image1)
    If nFrames > 0 Then FrameCount = 0
    For IdImg = 0 To Image1.Count - 1
        Image1(IdImg).Visible = False
    Next IdImg
    Image1(FrameCount).Visible = True
    Image1(FrameCount).Refresh
    Cursor_Img
    MsgAux = ""
    Contactos = ""
   'Contamos cuantos mails se han enviado por medio de MySQL
    ContMails = 0
    sSQL = "SELECT Fecha, COUNT(Fecha) As ContMails " _
         & "FROM acceso_pcs " _
         & "WHERE Fecha = '" & BuscarFecha(FechaSistema) & "' " _
         & "AND ES IN ('EM') "
    Select_AdoDB_MySQL AdoSMTP, sSQL
    If AdoSMTP.RecordCount > 0 Then
       ContMails = AdoSMTP.fields("ContMails")
    End If
    AdoSMTP.Close
    Cursor_Img
    
    Si_Enviar = False
    Set ftp = New cFTP
        
    Label1.Caption = "CONECTANDOSE AL SERVIDOR" & vbCrLf & vbCrLf & "DE CORREOS ELECTRONICOS"
    Label1.Refresh
    TMail.ContadorTiempo = 0
   'ErrorMails = ""
    Cursor_Img
   'Determinamos si esta activado envio de correos
    sSQL = "SELECT smtp_Servidor, smtp_Puerto, smtp_UseAuntentificacion, smtp_SSL, smtp_Secure, " _
         & "Email_Conexion, Email_Contrase�a, Email_Conexion_CE, Email_Contrase�a_CE, Email_Procesos, Email_CE_Copia " _
         & "FROM Empresas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND LEN(smtp_Servidor) > 1 " _
         & "AND smtp_Puerto > 0 "
    Select_AdoDB AdoSMTP, sSQL
    Cursor_Img
    With AdoSMTP
     If .RecordCount > 0 Then
         EmailProcesos = .fields("Email_Procesos")
         TMail.useAuntentificacion = CBool(.fields("smtp_UseAuntentificacion"))
         TMail.ssl = CBool(.fields("smtp_SSL"))
         TMail.tls = CBool(.fields("smtp_Secure"))
         TMail.Puerto = .fields("smtp_Puerto")
         Email_CE_Copia = CBool(.fields("Email_CE_Copia"))
         TMail.servidor = .fields("smtp_Servidor")
         If TMail.TipoDeEnvio = "CE" Then
            TMail.Usuario = .fields("Email_Conexion_CE")
            TMail.Password = .fields("Email_Contrase�a_CE")
            TMail.de = .fields("Email_Conexion_CE")
         Else
            TMail.Usuario = .fields("Email_Conexion")
            TMail.Password = .fields("Email_Contrase�a")
            TMail.de = .fields("Email_Conexion")
         End If
         
         If Email_CE_Copia Then Insertar_Mail TMail.para, EmailProcesos
         If TMail.de = "" And 0 <= TMail.ListaMail And TMail.ListaMail <= 6 Then TMail.de = Lista_De_Correos(TMail.ListaMail).Correo_Electronico
        'Si utilizamos el correo de DiskCover System
         If TMail.servidor = "mail.diskcoversystem.com" Then
            TMail.servidor = "imap.diskcoversystem.com"
            TMail.de = Replace(TMail.de, "@diskcoversystem.com", "@imap.diskcoversystem.com")
         End If
     Else
        'Si enviamos mail desde el modulo de actualizacion se activa el servidor propio de DiskCover System
         If Modulo = "UPDATE" Then
            TMail.servidor = ServidorCorreos
            NombreUsuario = "Update DiskCover"
            ComunicadoEntidad = ""
            NombreGerente = "Walter Vaca Prieto"
            Contactos = "09-9965-4196/09-8910-5300"
            EmailProcesos = CorreoUpdate
            MsgAux = "DISKCOVER SYSTEM"
            TMail.de = "Actualizacion de DiskCover System" & " <" & CorreoDiskCover & ">"
            TMail.de = Replace(TMail.de, "@diskcoversystem.com", "@imap.diskcoversystem.com")
         End If
     End If
    End With
    AdoSMTP.Close
    Cursor_Img
    If TMail.de <> "" And TMail.para <> "" Then
       TMail.ListaError = ""
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
    TMail.ListaError = TMail.ListaError & "Error No. " & Numero & ": " & Descripcion & vbCrLf
    Control_Procesos "ER", "Email: " & TMail.de & " => " & TMail.para, "Error: " & TMail.ListaError
   'Unload FEnviarCorreos
End Sub

Public Sub Cursor_Img()
On Local Error Resume Next

    Image1(FrameCount).Visible = False
    If FrameCount < TotalFrames Then FrameCount = FrameCount + 1 Else FrameCount = 0
    Image1(FrameCount).Visible = True
    Image1(FrameCount).Refresh
    TMail.ContadorTiempo = TMail.ContadorTiempo + 1
    If TMail.ContadorTiempo > 3 Then TMail.ContadorTiempo = 0
    Select Case TMail.ContadorTiempo
      Case 0: Label1.ForeColor = Blanco_Claro ' &HC00000
      Case 1: Label1.ForeColor = Amarillo_Claro
      Case 2: Label1.ForeColor = Verde_Claro
      Case 3: Label1.ForeColor = Azul
    End Select
    If Err Then Exit Sub
End Sub

