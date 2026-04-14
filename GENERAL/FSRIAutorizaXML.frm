VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FSRIAutorizaXML 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSRIAutorizaXML.frx":0000
   ScaleHeight     =   1200
   ScaleWidth      =   7305
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
      TabIndex        =   0
      Top             =   3150
      Visible         =   0   'False
      Width           =   6315
   End
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   525
      TabIndex        =   1
      Top             =   3885
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
   Begin VB.Label LblConexion 
      BackStyle       =   0  'Transparent
      Caption         =   "CONEXION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   1260
      TabIndex        =   2
      Top             =   210
      Width           =   5790
   End
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   3465
      Top             =   3780
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
            Picture         =   "FSRIAutorizaXML.frx":0E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":119C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":14B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":17BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":1AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":1DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":20E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":28FC
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":2C16
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":2F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":316E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSRIAutorizaXML.frx":3488
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FSRIAutorizaXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lectura_XML As Boolean

Dim MensajeError As String

Dim Inicio As Date
Dim fin As Date

Dim DocumentoXML As MSXML2.DOMDocument
Dim pJSON As Object

Private Sub Form_Activate()
   'SRI_Autorizacion.Clave_De_Acceso: Viene desde afuera ya la clave de acceso que se va autorizar
    With SRI_Autorizacion
        '---------------------------------------------------------------------
         RutaXML = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Generados\" & .Clave_De_Acceso & ".xml"
         RutaXMLFirmado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Firmados\" & .Clave_De_Acceso & ".xml"
         RutaXMLAutorizado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Autorizados\" & .Clave_De_Acceso & ".xml"
         RutaXMLRechazado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes no Autorizados\" & .Clave_De_Acceso & ".xml"
                  
        .Documento_XML = ""
        .Error_SRI = ""
        .Estado_SRI = ""
        .Error_SRI = ""
        .Fecha_Autorizacion = FechaSistema
        .Hora_Autorizacion = HoraSistema
        .Autorizacion = FA.Autorizacion
         Lectura_XML = DocumentoXML.Load(RutaXML)
         If Lectura_XML Then
            TextoXML = DocumentoXML.XML
           'Clave de Acceso a autorizar
            Progreso_Barra.Mensaje_Box = "Enviando el Comprobante al S.R.I. " & .Clave_De_Acceso
            LblConexion.Caption = "CONECTAR AL WEBSERVICE DEL S.R.I..." & vbCrLf _
                                & "Generando y enviando el Comprobante para Autorizar" & vbCrLf _
                                & "Documento (" & .Tipo_Doc_SRI & "): " & .Clave_De_Acceso
            Progreso_Esperar True
            LblConexion.Refresh
            RatonReloj
           '--------------------------------------------------------------------------------------------------------------------------------
            Inicio = Now
            URLHTTP = "https://erp.diskcoversystem.com/php/comprobantes/SRI/autorizar_sri_visual.php?EnviarAutorizarXMLOnline=true"
            URLParams = "Autorizacion=" & .Clave_De_Acceso & "&RUTA=" & NombreCertificado & "&PASS=" & ClaveCertificado & "&XML=" & TextoXML
            MensajeError = PostUrlSourceStr(URLHTTP, URLParams)
            fin = Now
           '--------------------------------------------------------------------------------------------------------------------------------
            Set pJSON = JSON.parse(MensajeError)
            If IsNull(pJSON.Item("mensaje")) Then .Error_SRI = "" Else .Error_SRI = pJSON.Item("mensaje")
            If pJSON.Item("respuesta") = "1" Then
              .Resultado = True
            Else
              .Resultado = False
               If pJSON.Item("respuesta") = "CNASS" Then
                 .Estado_SRI = "CNASS"
                 .Error_SRI = "CNASS: Servidor de Aprobacion se encuentra saturado, intente mas tarde."
               Else
                  .Estado_SRI = "CNA"
                  If Len(.Error_SRI) > 1 Then .Error_SRI = "CNA: " & .Error_SRI Else .Error_SRI = "CNA: Comprobante no autorizado."
               End If
            End If
            
           .Documento_XML = pJSON.Item("XML")
            DocumentoXML.loadXML .Documento_XML
           '-----------------------------------
           'Clipboard.Clear
           'Clipboard.SetText pJSON.Item("XML")
'           MsgBox "RESULTADO DEL S.R.I. [" & Format(fin - Inicio, "hh:mm:ss") & "]" & vbCrLf _
'                  & "Documento: " & .Clave_De_Acceso & vbCrLf _
'                  & "AUTORIZADO: " & .Resultado
           '-----------------------------------
            If .Resultado Then
               .Estado_SRI = "OK"
               .Fecha_Autorizacion = MidStrg(pJSON.Item("FechaAutorizacion"), 1, 10)
               .Hora_Autorizacion = MidStrg(pJSON.Item("FechaAutorizacion"), 12, 8)
                
                Progreso_Barra.Mensaje_Box = "Documento: " & .Clave_De_Acceso & " autorizado"
                LblConexion.Caption = "RESULTADO DEL S.R.I. [" & Format(fin - Inicio, "hh:mm:ss") & "]" & vbCrLf _
                                    & "Documento: " & .Clave_De_Acceso & vbCrLf _
                                    & "AUTORIZADO."
                Progreso_Esperar True
                LblConexion.Refresh
                Autorizar_Documento_XML_SP SRI_Autorizacion, FA
                TA.Autorizacion = FA.Autorizacion
                DocumentoXML.save RutaXMLAutorizado
               'MsgBox .Autorizacion
            Else
                fin = Now
                FSRIAutorizaXML.BackColor = &HFF&
                FSRIAutorizaXML.Refresh
                LblConexion.ForeColor = &H8000000F
                LblConexion.Refresh
                Progreso_Barra.Mensaje_Box = .Estado_SRI & ": Comprobante no autorizado"
                LblConexion.Caption = "RESULTADO DEL S.R.I. " & .Estado_SRI & ": [" & Format(fin - Inicio, "hh:mm:ss") & "]" & vbCrLf _
                                    & "Documento: " & .Clave_De_Acceso & vbCrLf _
                                    & "NO AUTORIZADO."
                Progreso_Esperar True
               .Error_SRI = .Error_SRI & ". " & SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion)
                DocumentoXML.save RutaXMLRechazado
            End If
         Else
           .Documento_XML = Replace(Replace(Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\sri_file_not_exist.xml"), "Clave_", .Clave_De_Acceso), "Error_", DocumentoXML.parseError.reason)
            fin = Now
            FSRIAutorizaXML.BackColor = &HFF&
            FSRIAutorizaXML.Refresh
            LblConexion.ForeColor = &H8000000F
            LblConexion.Refresh
            Progreso_Barra.Mensaje_Box = "Documento: " & .Clave_De_Acceso & " no autorizado"
            LblConexion.Caption = "RESULTADO DEL S.R.I. CNG: [" & Format(fin - Inicio, "hh:mm:ss") & "]" & vbCrLf _
                                & "Documento: " & .Clave_De_Acceso & vbCrLf _
                                & "NO AUTORIZADO."
            Progreso_Esperar True
            LblConexion.Refresh
           .Resultado = False
           .Estado_SRI = "CNG"
           .Error_SRI = "Error: Comprobante no generado: " & SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion)
            DocumentoXML.save RutaXMLRechazado
         End If
    End With
    RatonNormal
    Progreso_Barra.Mensaje_Box = ""
    Progreso_Esperar True
    Progreso_Final
   'MsgBox "Stop: " & SRI_Autorizacion.Estado_SRI & vbCrLf & SRI_Autorizacion.Error_SRI & vbCrLf & SRI_Autorizacion.Documento_XML
    Unload FSRIAutorizaXML
End Sub

Private Sub Form_Load()
    RatonReloj
    CentrarForm FSRIAutorizaXML
    Redondear_Formulario FSRIAutorizaXML, 40
    Progreso_Barra.Mensaje_Box = "CONECTAR AL WEBSERVICE DEL S.R.I..."
    LblConexion.Caption = Progreso_Barra.Mensaje_Box
    Set DocumentoXML = New DOMDocument
    Progreso_Esperar True
    LblConexion.Refresh
End Sub

