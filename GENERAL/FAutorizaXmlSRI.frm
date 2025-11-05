VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FAutorizaXmlSRI 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   9030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FAutorizaXmlSRI.frx":0000
   ScaleHeight     =   1740
   ScaleWidth      =   9030
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
      Height          =   1275
      Left            =   1575
      TabIndex        =   2
      Top             =   210
      Width           =   6945
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
            Picture         =   "FAutorizaXmlSRI.frx":4F3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":5259
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":5573
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":5879
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":5B93
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":5EAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":619F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":69B9
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":6CD3
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":6FED
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":722B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FAutorizaXmlSRI.frx":7545
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FAutorizaXmlSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
'Dim obj As New Cls_FirmarXML
'Dim ObjEnviar As New WS_Recepcion
'Dim ObjAutori As New WS_Autorizacion
Dim Resultado As Boolean
Dim success As Boolean

Dim doc As New MSXML2.DOMDocument
Dim nodeList As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode

Dim EstadoXML As String
Dim RutaFileError As String
Dim Documento As String
Dim MensajeError As String
Dim ArrayRecepcion() As String
'Dim ArrayAutorizacion() As String

Dim autoday As String
Dim autohour As String
Dim autominute As String
Dim automonth As String
Dim autosecond As String
Dim autotimezone As String
Dim autoyear As String

Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim EsperaEspera As Integer
Dim NumFile As Integer

   'SRI_Autorizacion.Clave_De_Acceso: Viene desde afuera ya la clave de acceso que se va autorizar
    EsperaEspera = 1000
    
    Progreso_Barra.Mensaje_Box = "CONECTAR AL WEBSERVICE DEL S.R.I..."
    LblConexion.Caption = Progreso_Barra.Mensaje_Box & vbCrLf
    Progreso_Esperar True
    LblConexion.Refresh
 
   'Sleep EsperaEspera
    Set ftp = New cFTP

   'Sleep EsperaEspera
    With SRI_Autorizacion
        'Prueba si existe el archivo
        '---------------------------------------------------------------------
        '.Clave_De_Acceso = "2404202501070216417900110010030000025261234567814"  'Generados
        '.Clave_De_Acceso = "2404202501070216417900110010030000025261234567814"  'Firmado
        '.Clave_De_Acceso = "2404202507179300609400120010050000074171234567817"  'Autorizados
        '.Clave_De_Acceso = "2504202504070216417900110010030000001131234567816"  'No Autorizados
        '---------------------------------------------------------------------
         Documento = MidStrg(.Clave_De_Acceso, 25, 15)
          
         Progreso_Barra.Mensaje_Box = "Generando Comprobante para Autorizar"
         LblConexion.Caption = "CONECTAR AL WEBSERVICE DEL S.R.I... -> Generando Comprobante para Autorizar"
         Progreso_Esperar True
         LblConexion.Refresh
         
         RutaXML = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Generados\" & .Clave_De_Acceso & ".xml"
         RutaXMLFirmado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Firmados\" & .Clave_De_Acceso & ".xml"
         RutaXMLAutorizado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Autorizados\" & .Clave_De_Acceso & ".xml"
         RutaXMLRechazado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes no Autorizados\" & .Clave_De_Acceso & ".xml"
         
         RutaFileError = RutaSysBases & "\CE\CE" & NumEmpresa & "\Error_"
         Select Case MidStrg(.Clave_De_Acceso, 9, 2)
           Case "01": RutaFileError = RutaFileError & "FA_"
           Case "03": RutaFileError = RutaFileError & "LC_"
           Case "04": RutaFileError = RutaFileError & "NC_"
           Case "06": RutaFileError = RutaFileError & "GR_"
           Case "07": RutaFileError = RutaFileError & "RE_"
           Case Else: RutaFileError = RutaFileError & "XX_"
         End Select
         RutaFileError = RutaFileError & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
         If Dir$(RutaFileError & ".txt") <> "" Then Kill RutaFileError & ".txt"
         
        .Documento_XML = ""
        .Error_SRI = ""
        .Estado_SRI = ""
        .Error_SRI = ""
        .Fecha_Autorizacion = FechaSistema
        .Hora_Autorizacion = HoraSistema
            
         If Dir$(RutaXML) = "" Then
           .Estado_SRI = "CNG"
           .Error_SRI = "Error: Comprobante no generado"
            GoTo Fin_Autorizacion
         End If
         
        'Clave de Acceso a autorizar
         Progreso_Barra.Mensaje_Box = .Clave_De_Acceso
         LblConexion.Caption = "CONECTAR AL WEBSERVICE DEL S.R.I... -> Generando Comprobante para Autorizar" & vbCrLf _
                             & "Documento: " & .Clave_De_Acceso & vbCrLf
         Progreso_Esperar True
         LblConexion.Refresh
         
         ftp.Inicializar MDIFormulario
         ftp.Password = ftpPwr  'Le establecemos la contraseña de la cuenta Ftp
         ftp.Usuario = ftpUse   'Le establecemos el nombre de usuario de la cuenta
         ftp.servidor = ftpSvr  'Establecesmo el nombre del Servidor FTP
         ftp.ConStatus = True
         If ftp.ConectarFtp(LstStatud) Then
           'MsgBox "Desktop Test: " & LstStatud
            RatonReloj
           'Seteamos variables y controles para subida y bajada de archivos del FTP
            Set ftp.ListView = LstVwFTP

           'Subimos el Archivo XML al FTP para ser aurotizado
            ftp.SubirArchivo RutaXML, "/files/ComprobantesElectronicos/" & .Clave_De_Acceso & ".xml", True
           'Firmando documento generado para su autorizacion
            Progreso_Barra.Mensaje_Box = "Firmando el Comprobante No. " & Documento
            Progreso_Esperar True
            
           'SRI_Obtener_Datos_Comprobantes_Electronicos
           'MsgBox "Desktop Test: Iniciar Autorizar"
           '-------------------------------------------------------------------------------------------------------------------
            Progreso_Barra.Mensaje_Box = "Enviando el Comprobante al S.R.I."
            LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            LblConexion.Refresh
            URLHTTP = "https://erp.diskcoversystem.com/php/comprobantes/SRI/autorizar_sri_visual.php?AutorizarXMLOnline=true"
            URLParams = "XML=" & .Clave_De_Acceso & ".xml&RUTA=" & NombreCertificado & "&PASS=" & ClaveCertificado
            MensajeError = PostUrlSourceStr(URLHTTP, URLParams)
           '-------------------------------------------------------------------------------------------------------------------
           'Le indicamos el ListView donde se listarán los archivos
           ' Clipboard.Clear
           ' Clipboard.SetText URLHTTP & vbCrLf & URLParams & vbCrLf & String(80, "-") & vbCrLf & "Resultado: " & MensajeError
           'MsgBox "Desktop Test: " & URLParams & vbCrLf & MensajeError
           .Estado_SRI = MensajeError
            If MensajeError = "Autorizado" Then
               Progreso_Barra.Mensaje_Box = String(120, "-")
               LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
               LblConexion.Refresh
               Progreso_Esperar True
               
               Progreso_Barra.Mensaje_Box = "COMPROBANTE ELECTRONICO AUTORIZADO"
               LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
               LblConexion.Refresh
               Progreso_Esperar True
                
               Progreso_Barra.Mensaje_Box = String(120, "-")
               LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               LblConexion.Refresh
                
              'MsgBox "--------"
               
               Progreso_Barra.Mensaje_Box = "Guardando Comprobante en la Base de Datos"
               Progreso_Esperar True
               
               Resultado = False
               ftp.CambiarDirectorio "/ComprobantesElectronicos/Firmados/"
               ftp.ListarArchivos
               For I = 1 To LstVwFTP.ListItems.Count
                   If .Clave_De_Acceso & ".xml" = LstVwFTP.ListItems(I) Then
                       Progreso_Barra.Mensaje_Box = "Obteniendo: " & LstVwFTP.ListItems(I)
                       Progreso_Esperar True
                       ftp.ObtenerArchivo LstVwFTP.ListItems(I), RutaXMLFirmado, True
                      .Estado_SRI = "CF"
                       Resultado = True
                       Exit For
                   End If
               Next I
               If Not Resultado Then
                  Progreso_Barra.Mensaje_Box = "Documento No Firmado"
                  Progreso_Esperar True
                 .Estado_SRI = "CNF"
                 'GoTo Fin_Autorizacion
               End If
               
               ftp.CambiarDirectorio "/ComprobantesElectronicos/Autorizados/"
               ftp.ListarArchivos
               For I = 1 To LstVwFTP.ListItems.Count
                   If .Clave_De_Acceso & ".xml" = LstVwFTP.ListItems(I) Then
                       Progreso_Barra.Mensaje_Box = "Obteniendo: " & LstVwFTP.ListItems(I)
                       Progreso_Esperar True
                       ftp.ObtenerArchivo LstVwFTP.ListItems(I), RutaXMLAutorizado, True
                      .Estado_SRI = "OK"
                       Exit For
                   End If
               Next I
              .Error_SRI = "OK"
               ftp.EliminarArchivo "/files/ComprobantesElectronicos/" & .Clave_De_Acceso & ".xml"
               ftp.EliminarArchivo "/files/ComprobantesElectronicos/Firmados/" & .Clave_De_Acceso & ".xml"
               ftp.EliminarArchivo "/files/ComprobantesElectronicos/Autorizados/" & .Clave_De_Acceso & ".xml"
               ftp.Desconectar

              'Capturando la Autorizacion y sus Fechas
              .Autorizacion = RUC
               success = doc.Load(RutaXMLAutorizado)
               If success Then
                  Set nodeList = doc.selectNodes("/autorizacion")
                  If Not nodeList Is Nothing Then
                     For Each node In nodeList
                         If UCaseStrg(node.selectSingleNode("estado").Text) = "AUTORIZADO" Then .Estado_SRI = "OK"
                        .Autorizacion = node.selectSingleNode("numeroAutorizacion").Text
                         Documento = node.selectSingleNode("fechaAutorizacion").Text
                         If IsDate(MidStrg(Documento, 1, 10)) Then
                           .Fecha_Autorizacion = SinEspaciosIzq(Documento)
                           .Hora_Autorizacion = MidStrg(SinEspaciosDer(Documento), 1, 8)
                         Else
                            Cadena = Documento
                            
                            autoday = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(autoday) + 1, Len(Cadena)))
                            
                            autohour = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(autohour) + 1, Len(Cadena)))
                            
                            autominute = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(autominute) + 1, Len(Cadena)))
                            
                            automonth = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(automonth) + 1, Len(Cadena)))
                            
                            autosecond = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(autosecond) + 1, Len(Cadena)))
                            
                            autotimezone = SinEspaciosIzq(Cadena)
                            Cadena = TrimStrg(MidStrg(Cadena, Len(autotimezone) + 1, Len(Cadena)))
                            
                            autoyear = Cadena
                           
                           .Fecha_Autorizacion = Format(Val(autoday), "00") & "/" & Format(Val(automonth), "00") & "/" & Format(Val(autoyear), "0000")
                           .Hora_Autorizacion = Format(Val(autohour), "00") & ":" & Format(Val(autominute), "00") & ":" & Format(Val(autosecond), "00")
                         End If
                     Next node
                  End If
               End If
              .Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
               SRI_Actualizar_Documento_XML SRI_Autorizacion
            ElseIf MensajeError = "CNASS" Then
              .Error_SRI = "CNASS: Servidor de Aprobacion se encuentra saturado, intente mas tade"
              .Estado_SRI = "ERR"
            Else
               Progreso_Barra.Mensaje_Box = "Documento Electronico No autorizado"
               LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               LblConexion.Refresh

               ftp.CambiarDirectorio "/ComprobantesElectronicos/No_Autorizados/"
               ftp.ListarArchivos
               For I = 1 To LstVwFTP.ListItems.Count
                   If .Clave_De_Acceso & ".xml" = LstVwFTP.ListItems(I) Then
                       Progreso_Barra.Mensaje_Box = "Obteniendo: " & LstVwFTP.ListItems(I)
                       Progreso_Esperar True
                      'MsgBox RutaXMLRechazado
                       ftp.ObtenerArchivo LstVwFTP.ListItems(I), RutaXMLRechazado, True
                       Exit For
                   End If
               Next I
               Progreso_Barra.Mensaje_Box = "Guardando Comprobante no Autorizado"
               LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
               Progreso_Esperar True
               LblConexion.Refresh
               
               ftp.EliminarArchivo "/files/ComprobantesElectronicos/" & .Clave_De_Acceso & ".xml"
               ftp.EliminarArchivo "/files/ComprobantesElectronicos/No_Autorizados/" & .Clave_De_Acceso & ".xml"
               ftp.Desconectar
               
              'Capturando el Error
              .Error_SRI = ""
              .Estado_SRI = "CNA"
               success = doc.Load(RutaXMLRechazado)
               If success Then
                  If InStr(doc.XML, "<estado>DEVUELTA</estado>") Then EstadoXML = "DEVUELTA"
                  
                  If InStr(doc.XML, "<autorizacion") Then
                     Set nodeList = doc.selectNodes("/autorizacion")
                     If Not nodeList Is Nothing Then
                        For Each node In nodeList
                            EstadoXML = node.selectSingleNode("estado").Text
                           .Error_SRI = .Error_SRI & "-" & EstadoXML & vbCrLf
                            'MsgBox "Desktop test: " & EstadoXML
                            Select Case EstadoXML
                              Case "RECHAZADA", "DEVUELTA"
                              
                              Case Else
                                   Documento = node.selectSingleNode("fechaAutorizacion").Text
                                  .Fecha_Autorizacion = SinEspaciosIzq(Documento)
                                  .Hora_Autorizacion = MidStrg(SinEspaciosDer(Documento), 1, 8)
                            End Select
                        Next node
                     End If
                     Set nodeList = doc.selectNodes("/autorizacion/mensajes/mensaje/mensaje")
                  ElseIf InStr(doc.XML, "<factura") Then
                     Set nodeList = doc.selectNodes("/factura/ns2:respuestaSolicitud/comprobantes/comprobante/mensajes/mensaje")
                  ElseIf InStr(doc.XML, "<comprobanteRetencion") Then
                     Set nodeList = doc.selectNodes("/comprobanteRetencion/ns2:respuestaSolicitud/comprobantes/comprobante/mensajes/mensaje")
                  ElseIf InStr(doc.XML, "<notaCredito") Then
                     Set nodeList = doc.selectNodes("/notaCredito/ns2:respuestaSolicitud/comprobantes/comprobante/mensajes/mensaje")
                  ElseIf InStr(doc.XML, "<guiaRemision") Then
                     Set nodeList = doc.selectNodes("/guiaRemision/ns2:respuestaSolicitud/comprobantes/comprobante/mensajes/mensaje")
                  End If
                  
                  If Not nodeList Is Nothing Then
                     For Each node In nodeList
                        .Error_SRI = .Error_SRI & node.selectSingleNode("tipo").Text
                        .Error_SRI = .Error_SRI & " - " & node.selectSingleNode("identificador").Text & ": "
                        .Error_SRI = .Error_SRI & node.selectSingleNode("mensaje").Text & "; "
                         Select Case EstadoXML
                           Case "RECHAZADA", "DEVUELTA"
                               .Estado_SRI = EstadoXML
                           Case Else
                               .Error_SRI = .Error_SRI & node.selectSingleNode("informacionAdicional").Text
                         End Select
                     Next node
                  End If
                  If Len(MensajeError) > 1 Then .Error_SRI = .Error_SRI & ", [" & MensajeError & "]" & vbCrLf
                  .Error_SRI = Replace(.Error_SRI, vbTab, " ")
                 .Error_SRI = Sin_Signos_Especiales(.Error_SRI)
               End If
              'MsgBox .Error_SRI
               If .Error_SRI = "" Then
                   RutaFileError = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes no Autorizados\Error_"
                   Select Case MidStrg(.Clave_De_Acceso, 9, 2)
                     Case "01": .Error_SRI = "FA_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                     Case "03": .Error_SRI = "LC_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                     Case "04": .Error_SRI = "NC_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                     Case "06": .Error_SRI = "GR_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                     Case "07": .Error_SRI = "RE_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                     Case Else: .Error_SRI = "XX_" & MidStrg(.Clave_De_Acceso, 25, 6) & "-" & MidStrg(.Clave_De_Acceso, 31, 9)
                   End Select
                  .Error_SRI = "Documento " & .Error_SRI & ", enviado al SRI sin respueta de aprobacion"
                  .Estado_SRI = "DES"
               End If
            End If
            RatonNormal
         Else
            Progreso_Barra.Mensaje_Box = "No se pudo conectar al servidor del S.R.I."
            LblConexion.Caption = LblConexion.Caption & Progreso_Barra.Mensaje_Box & vbCrLf
            Progreso_Esperar True
            LblConexion.Refresh
         
            RatonNormal
           .Error_SRI = "No se pudo conectar al servidor"
           .Estado_SRI = "ERR"
         End If
    End With
Fin_Autorizacion:
    SRI_Autorizacion.Error_SRI = TrimStrg(SRI_Autorizacion.Error_SRI)
    Progreso_Barra.Mensaje_Box = "Ok"
    Progreso_Esperar True
    Progreso_Final
    If SRI_Autorizacion.Error_SRI <> "OK" Then
       MensajeError = RutaFileError & vbCrLf & "Estado: " & SRI_Autorizacion.Estado_SRI & vbCrLf & SRI_Autorizacion.Error_SRI
       NumFile = FreeFile
       Open RutaFileError & ".txt" For Output As #NumFile ' Abre el archivo.
            Print #NumFile, MensajeError
       Close #NumFile
    End If
    RatonNormal
   'MsgBox SRI_Autorizacion.Estado_SRI & vbCrLf & SRI_Autorizacion.Error_SRI & vbCrLf & SRI_Autorizacion.Clave_De_Acceso
    Unload FAutorizaXmlSRI
End Sub

Private Sub Form_Load()
Dim AnchoMaxForm As Single

    RatonReloj
    CentrarForm FAutorizaXmlSRI
    Redondear_Formulario FAutorizaXmlSRI, 60
    Progreso_Barra.Mensaje_Box = "Conectandose al S.R.I..."
    Progreso_Esperar True
        
''    AnchoMaxForm = FAutorizaXmlSRI.width
''    If AnchoMaxForm < FAutorizaXmlSRI.TextWidth(Progreso_Barra.Mensaje_Box) Then
''       AnchoMaxForm = FAutorizaXmlSRI.TextWidth(Progreso_Barra.Mensaje_Box)
''    End If
''    AnchoMaxForm = AnchoMaxForm + 1500
''    FAutorizaXmlSRI.width = AnchoMaxForm
''    FAutorizaXmlSRI.LblConexion.width = AnchoMaxForm - 1000
''    FAutorizaXmlSRI.Refresh
End Sub

