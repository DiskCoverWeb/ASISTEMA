VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form FAutorizaXmlSRI 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7860
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
      Left            =   210
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.TextBox TxtConexion 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   735
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FAutorizaXmlSRI.frx":0000
      Top             =   210
      Width           =   6840
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   105
   End
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   630
      TabIndex        =   2
      Top             =   2940
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
         Text            =   "Tamaño"
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
   Begin VB.Image Image1 
      Height          =   510
      Index           =   0
      Left            =   105
      Picture         =   "FAutorizaXmlSRI.frx":0017
      Top             =   210
      Width           =   510
   End
End
Attribute VB_Name = "FAutorizaXmlSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nFrames As Long

Private Sub Form_Activate()
Dim AdoDBEmpresa As ADODB.Recordset
Dim obj As New Cls_FirmarXML
Dim ObjEnviar As New WS_Recepcion
Dim ObjAutori As New WS_Autorizacion
Dim URLRecepcion As String
Dim URLAutorizacion As String
Dim Resultado As Boolean

Dim ClaveCertificado As String
Dim RutaXML As String
Dim RutaXMLFirmado As String
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim Documento As String
Dim MensajeError As String
Dim ArrayRecepcion() As String
Dim ArrayAutorizacion() As String
Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim EsperaEspera As Integer

   'SRI_Autorizacion.Clave_De_Acceso: Viene desde afuera ya la clave de acceso que se va autorizar
    EsperaEspera = 3000
    
    Progreso_Barra.Mensaje_Box = "Conectandose al S.R.I..."
    TxtConexion = Progreso_Barra.Mensaje_Box & vbCrLf
    Progreso_Esperar True
    TxtConexion.Refresh
 
    Sleep EsperaEspera
    Set ftp = New cFTP
'''    For ContadorEstados = 0 To 4
'''        ArrayRecepcion(ContadorEstados) = ""
'''        ArrayAutorizacion(ContadorEstados) = ""
'''    Next ContadorEstados

   'Determinamos si esta activado envio de correos
    sSQL = "SELECT Ambiente, Codigo_Contribuyente_Especial, Obligado_Conta, Ruta_Certificado, Clave_Certificado, Web_SRI_Recepcion, Web_SRI_Autorizado " _
         & "FROM Empresas " _
         & "WHERE Item = '" & NumEmpresa & "' "
    Select_AdoDB AdoDBEmpresa, sSQL
    With AdoDBEmpresa
     If .RecordCount > 0 Then
         Ambiente = .fields("Ambiente")
         ContEspec = .fields("Codigo_Contribuyente_Especial")
         Obligado_Conta = .fields("Obligado_Conta")
        
        'Ruta del Certificado para Firmar el documento
         RutaCertificado = RutaSistema & "\CERTIFIC\" & .fields("Ruta_Certificado")
         ClaveCertificado = .fields("Clave_Certificado")
    
        'Pagina de Conexion con el SRI
         URLRecepcion = .fields("Web_SRI_Recepcion")
         URLAutorizacion = .fields("Web_SRI_Autorizado")
     End If
    End With
    AdoDBEmpresa.Close
   'Sleep EsperaEspera
    With SRI_Autorizacion
         Documento = MidStrg(.Clave_De_Acceso, 25, 15)
         Progreso_Barra.Mensaje_Box = "Determinando Carpetas de Conexion"
         TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
         Progreso_Esperar True
         TxtConexion.Refresh
        
        .Documento_XML = ""
        .Error_SRI = ""
    
         RutaXML = RutaDocumentos & "\Comprobantes Generados\" & .Clave_De_Acceso & ".xml"
         RutaXMLFirmado = RutaDocumentos & "\Comprobantes Firmados\" & .Clave_De_Acceso & ".xml"
         RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & .Clave_De_Acceso & ".xml"
         RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & .Clave_De_Acceso & ".xml"
     
         If Dir$(RutaXML) = "" Then
           .Estado_SRI = "CNG"
           .Error_SRI = "Error: Comprobante no generado"
            GoTo Fin_Autorizacion
         End If
         
         If Dir$(RutaXMLFirmado) = "" Then
           .Estado_SRI = "CNF"
           .Error_SRI = "Error: Comprobante Firmado"
         End If
         
         Progreso_Barra.Mensaje_Box = "Leer Certificado del Documento: " & Documento
         TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
         Progreso_Esperar True
         TxtConexion.Refresh
     
        'Verificar si el documento ya esta autorizado
         ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, .Clave_De_Acceso, RutaXMLAutorizado, RutaXMLRechazado)
         If ArrayAutorizacion(0) = "AUTORIZADO" Then
            Progreso_Barra.Mensaje_Box = "El documento: " & Documento & ", esta autorizado"
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
           .Estado_SRI = "OK"
           .Error_SRI = "OK"
           .Autorizacion = ArrayAutorizacion(1)
           .Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
           .Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
           .Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
            SRI_Actualizar_Documento_XML .Clave_De_Acceso
            Progreso_Barra.Mensaje_Box = "Actualizando el Documento: " & Documento & " en la base"
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
           'MsgBox ArrayAutorizacion(1) & ".."
            GoTo Fin_Autorizacion
         End If

        'MsgBox "Primer face: " & .Estado_SRI
'''         Select Case .Estado_SRI
'''           Case "CNA", "CNF": GoTo Volver_Firmar
'''           Case "ESC": GoTo Volver_Autorizar
'''           Case "CF", "CR", "ESI": GoTo Volver_Enviar
'''         End Select
         
        'Firmamos el documento
         Progreso_Barra.Mensaje_Box = "Firmando el Documento: " & Documento
         TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
         Progreso_Esperar True
         TxtConexion.Refresh
        'MsgBox "Firmar: " & RutaXML & vbCrLf & vbCrLf & RutaXMLFirmado & vbCrLf & vbCrLf & RutaCertificado & vbCrLf & vbCrLf & ClaveCertificado
         Resultado = obj.FirmarXML(RutaCertificado, ClaveCertificado, RutaXML, RutaXMLFirmado, MensajeError)
         If Not Resultado Then
            ReDim ArrayRecepcion(0 To 4) As String
            ArrayRecepcion(0) = "ERROR"
            ArrayRecepcion(1) = MensajeError
            ArrayRecepcion(2) = ""
            ArrayRecepcion(3) = ""
            ArrayRecepcion(4) = ""
         End If
        'MsgBox "Firmar: (" & Resultado & ") " & MensajeError
         If Resultado Then
           .Estado_SRI = "CF"
            ftp.Inicializar MDIFormulario
            ftp.Password = ftpPwr  'Le establecemos la contraseña de la cuenta Ftp
            ftp.Usuario = ftpUse   'Le establecemos el nombre de usuario de la cuenta
            ftp.servidor = ftpSvr  'Establecesmo el nombre del Servidor FTP
            Set ftp.ListView = LstVwFTP
            ftp.ConStatus = True
             'LstStatud
            If ftp.ConectarFtp(LstStatud) = False Then
                RatonNormal
                MsgBox "No se pudo conectar"
                GoTo Fin_Autorizacion
            End If
            Progreso_Barra.Mensaje_Box = "Enviando el Documento al S.R.I.: " & Documento
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
           'Subiendo documento firmado para su autorizacion
            ftp.SubirArchivo RutaXMLFirmado, "/files/ComprobantesElectronicos/" & .Clave_De_Acceso & ".xml", True
           .Documento_XML = Leer_Archivo_Texto(RutaXMLFirmado)
           '----------------------------------------------------------------------------------------------------------------------------
            Sleep EsperaEspera
            URLHTTP = "https://erp.diskcoversystem.com/php/comprobantes/SRI/autorizar_sri_visual.php?AutorizarXMLOnline=true"
            URLParams = "XML=" & .Clave_De_Acceso & ".xml"
            MensajeError = PostUrlSourceStr(URLHTTP, URLParams)
           'MsgBox MensajeError
           '----------------------------------------------------------------------------------------------------------------------------
           '5Le indicamos el ListView donde se listarán los archivos
            ftp.EliminarArchivo "/files/ComprobantesElectronicos/" & .Clave_De_Acceso & ".xml"
            ftp.Desconectar

            If MensajeError = "Autorizado" Then
               Progreso_Barra.Mensaje_Box = "Documento: " & Documento & ". Autorizado"
               TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               TxtConexion.Refresh
              .Estado_SRI = "OK"
            Else
               Progreso_Barra.Mensaje_Box = "Documento: " & Documento & ". No autorizado"
               TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               TxtConexion.Refresh

              .Estado_SRI = "Err"
              .Error_SRI = "Error al Autorizar: " & MensajeError
'''               For ContadorEstados = 0 To 4
'''                   If Len(ArrayRecepcion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayRecepcion(ContadorEstados) & "; "
'''               Next ContadorEstados
               GoTo Fin_Autorizacion
            End If
            Progreso_Barra.Mensaje_Box = "Cargando Documento: " & Documento & " Firmado y Autorizado"
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
              
           'Tiempo de Espera antes de averiguar al SRI de la autorizacion
            Sleep EsperaEspera
            'For Tiempo_Espera = 1 To 3
                ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, .Clave_De_Acceso, RutaXMLAutorizado, RutaXMLRechazado)
            'Next Tiempo_Espera
            'MsgBox ArrayAutorizacion(0)
            If ArrayAutorizacion(0) = "AUTORIZADO" Then
              'MsgBox "Ok Documento Firmado y Autorizado"
               Progreso_Barra.Mensaje_Box = "Extrayendo Documentos Autorizado: " & Documento
               TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               TxtConexion.Refresh
              
               RatonReloj
              .Estado_SRI = "OK"
              .Error_SRI = "OK"
              .Autorizacion = ArrayAutorizacion(1)
              .Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
              .Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
              .Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
               SRI_Actualizar_Documento_XML .Clave_De_Acceso
               Progreso_Barra.Mensaje_Box = "Grabando en la base el Documento: " & Documento
               TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               TxtConexion.Refresh
            Else
               Progreso_Barra.Mensaje_Box = "Error: CNA"
               TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               TxtConexion.Refresh
            
              .Estado_SRI = "CNA"
              .Error_SRI = "Error al Autorizar: "
               For ContadorEstados = 0 To 4
                   If Len(ArrayAutorizacion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayAutorizacion(ContadorEstados) & ", "
               Next ContadorEstados
            End If
         ElseIf ArrayRecepcion(0) = "ERROR" Then
            Progreso_Barra.Mensaje_Box = "Error: ESI"
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
            
           .Estado_SRI = "ESI"
           .Error_SRI = " Error al enviar: "
            For ContadorEstados = 0 To 4
                If Len(ArrayRecepcion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayRecepcion(ContadorEstados) & ", "
            Next ContadorEstados
         Else
            Progreso_Barra.Mensaje_Box = "Error: CNF"
            TxtConexion = TxtConexion & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            TxtConexion.Refresh
         
           .Estado_SRI = "CNF"
           .Error_SRI = MensajeError
           .Documento_XML = MensajeError
         End If
        .Error_SRI = TrimStrg(.Error_SRI)
    End With
Fin_Autorizacion:
    Progreso_Final
    RatonNormal
   ' MsgBox "......."
    Unload FAutorizaXmlSRI
End Sub

Private Sub Form_Load()
Dim AnchoMaxForm As Single

    RatonReloj
    CentrarForm FAutorizaXmlSRI
    Redondear_Formulario FAutorizaXmlSRI, 40
    Progreso_Barra.Mensaje_Box = "Conectandose al S.R.I..."
    Progreso_Esperar True
    
    Timer1.Interval = 1000
    Timer1.Enabled = True
    
    nFrames = Load_Gif(RutaSistema & "\FORMATOS\conexion.gif", Image1)
    If nFrames > 0 Then FrameCount = 0
    AnchoMaxForm = FAutorizaXmlSRI.width
    If AnchoMaxForm < FAutorizaXmlSRI.TextWidth(Progreso_Barra.Mensaje_Box) Then
       AnchoMaxForm = FAutorizaXmlSRI.TextWidth(Progreso_Barra.Mensaje_Box)
    End If
    AnchoMaxForm = AnchoMaxForm + 1500
    FAutorizaXmlSRI.width = AnchoMaxForm
    FAutorizaXmlSRI.TxtConexion.width = AnchoMaxForm - 1000
    FAutorizaXmlSRI.Refresh
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
Timer1.Interval = CLng(Image1(FrameCount).Tag)
TxtConexion.ForeColor = Azul   ' &HC00000
If Err Then Exit Sub
End Sub

