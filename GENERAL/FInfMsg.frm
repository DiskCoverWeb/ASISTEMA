VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FInfoMensaje 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORMULARIO DE COMUNICADO A LA ENTIDAD"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11580
   DrawMode        =   12  'Nop
   Icon            =   "FInfMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Height          =   1065
      Left            =   10395
      Picture         =   "FInfMsg.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5145
      Visible         =   0   'False
      Width           =   1065
   End
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
      TabIndex        =   7
      Top             =   7245
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10395
      Picture         =   "FInfMsg.frx":1620
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Height          =   1065
      Left            =   10395
      Picture         =   "FInfMsg.frx":1CB6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3990
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10395
      Picture         =   "FInfMsg.frx":2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox TxtMensaje 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2955
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2310
      Width           =   10095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&PDF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10395
      Picture         =   "FInfMsg.frx":2BC0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1470
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoInfoError 
      Height          =   330
      Left            =   210
      Top             =   7875
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "InfoError"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   540
      Left            =   4515
      TabIndex        =   8
      Top             =   7875
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   953
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
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   2835
      Top             =   7875
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
            Picture         =   "FInfMsg.frx":3802
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":3B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":3E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":413C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":4456
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":4770
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":4A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":527C
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":5596
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":58B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":5AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FInfMsg.frx":5E08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   105
      TabIndex        =   4
      Top             =   5355
      Width           =   10200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2010
      Left            =   105
      TabIndex        =   3
      Top             =   210
      Width           =   10095
   End
End
Attribute VB_Name = "FInfoMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If ListaComunicados <> "" Then
       Titulo = "COMUNICADO A LA ENTIDAD"
       Mensajes = "DESEA SEGUIR RECIBIENDO ESTE COMUNICADO"
       If BoxMensaje = vbNo Then
         'Datos del destinatario de mails, Cod datos de privacidad
          TMail.Mensaje = Label2.Caption _
                        & TxtMensaje.Text _
                        & "NOTA: " & vbCrLf & Label3.Caption & vbCrLf & vbCrLf _
                        & "SERVIRLE ES NUESTRO OBJETIVO, DISFRUTARLO EL SUYO"
          TMail.ListaMail = 5
          TMail.Asunto = "Informativo enviado a Sr(a): " & RepresentanteLegal & ", representante de: " & NombreEntidad
          TMail.Adjunto = ""
          TMail.Credito_No = ""

         'Enviamos lista de mails
          TMail.para = ""
          Insertar_Mail TMail.para, CorreoDiskCover
          Insertar_Mail TMail.para, EmailContador
          Insertar_Mail TMail.para, EmailProcesos
         'MsgBox "Lista: " & emails

          FEnviarCorreos.Show 1

          sSQL = "UPDATE entidad " _
               & "SET Comunicado = '.' " _
               & "WHERE ID_Empresa = " & IDEntidad & " "
          Conectar_Ado_Execute_MySQL sSQL
          
          sSQL = "UPDATE lista_empresas " _
               & "SET Mensaje='.', Archivo_Incluido='.', Link_Canal='.' " _
               & "WHERE ID_Empresa = " & IDEntidad & " " _
               & "AND RUC_CI_NIC = '" & RUC & "' " _
               & "AND Item = '" & NumEmpresa & "' "
          Conectar_Ado_Execute_MySQL sSQL
       End If
    End If
    
    If FinalizarSistema Then End
    Unload Me
End Sub

Private Sub Command2_Click()
Dim PDF_Nombre_Documento As String
Dim PDF_Titulo As String
Dim PDF_TipoDeLetra As String
Dim PDF_VerDocumento As Boolean
Dim PosLineaTemp As Single

   RatonReloj
      PDF_Nombre_Documento = "Comunicado_Entidad"
      PDF_Titulo = "Comuncado para la entidad"
      PDF_TipoDeLetra = TipoCourierNew
     'Generamos el documento
      tPrint.TipoImpresion = Es_PDF
      tPrint.NombreArchivo = PDF_Nombre_Documento
      tPrint.TituloArchivo = PDF_Titulo
      tPrint.TipoLetra = PDF_TipoDeLetra
      tPrint.OrientacionPagina = Orientacion_Pagina
      tPrint.PaginaA4 = True
      tPrint.EsCampoCorto = False
      tPrint.VerDocumento = True
      Set cPrint = New cImpresion
      cPrint.iniciaImpresion
     
      cPrint.printImagen LogoTipo, 1, 1, 5, 2
      PosLinea = 3.5
      cPrint.colorDeLetra = Negro
      cPrint.tipoNegrilla = True
      If cPrint.anchoTexto(UCaseStrg(RazonSocial)) > 9 Then
         cPrint.letraTipo PDF_TipoDeLetra, 6
         cPrint.printTexto 1.1, PosLinea, UCaseStrg(RazonSocial)
         PosLinea = PosLinea + 0.35
      Else
         cPrint.printTexto 1.1, PosLinea, UCaseStrg(RazonSocial)
         PosLinea = PosLinea + 0.35
      End If
      PosLineaTemp = PosLinea
      PosLineaTemp = cPrint.printTextoMultiple(1.6, PosLineaTemp, Label2.Caption, 18)
      PosLineaTemp = PosLineaTemp + 0.35
      PosLineaTemp = cPrint.printTextoMultiple(1.6, PosLineaTemp, TxtMensaje.Text, 18)
      PosLineaTemp = PosLineaTemp + 0.35
      PosLineaTemp = cPrint.printTextoMultiple(1.6, PosLineaTemp, "NOTA: " & vbCrLf & Label3.Caption, 18)
      PosLineaTemp = PosLineaTemp + 0.7
      cPrint.printTexto 1.6, PosLineaTemp, "SERVIRLE ES NUESTRO OBJETIVO, DISFRUTARLO EL SUYO"
      PosLineaTemp = PosLineaTemp + 0.35
      cPrint.paginaNueva
      cPrint.finalizaImpresion
      RatonNormal
End Sub

Private Sub Command3_Click()
Dim Resultado As Long
Dim RutaArchivoWEB As String
    
    RatonNormal
    MsgBox "DEBE CERRAR TODOS LOS PROGRAMAS DE DISKCOVERSYSTEM, ANTES DE EMPEZAR LA ACTUALIZACION"
    RatonReloj
    sSQL = "UPDATE lista_empresas " _
         & "SET Version_Nueva=0 " _
         & "WHERE ID_Empresa = " & IDEntidad & " " _
         & "AND RUC_CI_NIC = '" & RUC & "' " _
         & "AND Item = '" & NumEmpresa & "' "
    Conectar_Ado_Execute_MySQL sSQL

  ' Ejecuta la acción "open" sobre el archivo
    RutaArchivoWEB = RutaSistema & "\UpdSystem.exe"
    Resultado = ShellExecute(0, "open", RutaArchivoWEB, vbNullString, vbNullString, SW_SHOWNORMAL)
  ' Si el resultado es menor o igual a 32, ocurrió un error
    RatonNormal
    If Resultado <= 32 Then
       MsgBox "No se pudo abrir el archivo. Verifica que la ruta exista o que haya una aplicación asociada.", vbExclamation, "Error"
    Else
       End
    End If

End Sub

Private Sub Command4_Click()
  ' Se construye la URL del video usando su ID
    ShellExecute 0, "open", Link_Canal, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Command5_Click()
Dim Resultado As Long
Dim RutaArchivoWEB As String

  ' Ejecuta la acción "open" sobre el archivo
    RutaArchivoWEB = RutaSysBases & "\TEMP\" & Archivo_Incluido
    Resultado = ShellExecute(0, "open", RutaArchivoWEB, vbNullString, vbNullString, SW_SHOWNORMAL)
     
  ' Si el resultado es menor o igual a 32, ocurrió un error
    If Resultado <= 32 Then MsgBox "No se pudo abrir el archivo. Verifica que la ruta exista o que haya una aplicación asociada.", vbExclamation, "Error"
    RatonNormal
End Sub

Private Sub Form_Activate()
    Label2.Caption = "Entidad: " & NombreEntidad & vbCrLf _
             & "Representante Legal: " & RepresentanteLegal & vbCrLf _
             & "Empresa: " & Empresa & vbCrLf _
             & "R.U.C.: " & RUC & " - " & UCaseStrg(MicroEmpresa) & vbCrLf _
             & "Telefono(s): " & Telefono1 & "/" & Telefono2 & ", Estado: " & DescripcionEstado & vbCrLf
    If Len(AgenteRetencion) > 1 Then Label2.Caption = Label2.Caption & " Resolución: " & AgenteRetencion & vbCrLf
    
    Label2.Caption = Label2.Caption _
             & "Renovacion Sistema: " & Fecha_CO & ", Comprobantes Electronicos: " & Fecha_CE & ", Certificado Electronico: " & Fecha_P12 & vbCrLf
    If Len(SerieFE) = 6 Then Label2.Caption = Label2.Caption & "Serie FA: " & SerieFE & " "
    If TipoPlan > 0 Then Label2.Caption = Label2.Caption & "Plan contratado: " & TipoPlan
    Label2.Caption = Label2.Caption & vbCrLf
    
    If ListaComunicados <> "" Then
       TxtMensaje.Text = "INFORMATIVO: Este es un mensaje automatico, enviado desde el Centro de " _
                    & "Control de Atención al Cliente, el informativo dice:" & vbCrLf _
                    & ListaComunicados & vbCrLf _
                    & "Por la atención que se de al presente quedamos de usted."
    End If
    Label3.Caption = "En caso de requerir atención personalizada por parte de un asesor de servicio " _
                   & "al cliente de DiskCover System, usted podrá solicitar ayuda mediante los canales de " _
                   & "atención al cliente oficiales que detallamos a continuación: " _
                   & "Telefonos: 099-965-4196 / 098-910-5300 / 099-467-8317."
                       
    If Len(Link_Canal) > 1 Then Command4.Visible = True
    If Len(Archivo_Incluido) > 1 Then Command5.Visible = True
    If Version_Nueva Then Command3.Visible = True
End Sub

Private Sub Form_Load()
    CentrarForm FInfoMensaje
    ConectarAdodc AdoInfoError
End Sub

