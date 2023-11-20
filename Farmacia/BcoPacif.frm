VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FBancoPacifico 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANCO PACIFICO"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "BcoPacif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recibir &Facturas del Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6195
      Picture         =   "BcoPacif.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   525
      Width           =   1590
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   3405
      Left            =   2940
      TabIndex        =   8
      Top             =   735
      Width           =   3165
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   4200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   735
      Width           =   2850
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6195
      Picture         =   "BcoPacif.frx":0CE8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1470
      Width           =   1590
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   105
      TabIndex        =   5
      Top             =   1050
      Width           =   2850
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2415
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   315
      Top             =   2730
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Query"
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
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   315
      Top             =   3045
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Act"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2940
      TabIndex        =   6
      Top             =   525
      Width           =   3165
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Width           =   2850
   End
End
Attribute VB_Name = "FBancoPacifico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal

Public Sub TipoProcesos(Opciones As String)
  NombreArchivo = ""
  Select Case Opciones
    Case "DESCUENT"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "FACTURAS"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "NOMINA"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case Else
         Dir1.Path = RutaBackup & "\"
         File1.FileName = Dir1.Path & "\*.*"
  End Select
  Dir1.Refresh
End Sub

Private Sub Command1_Click()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  'MsgBox RutaGeneraFile
  sSQL = "SELECT Codigo,CI_RUC,Cliente " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoQuery, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL = True " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Cta_Cobrar = AdoAux.Recordset.Fields("CxC")
     NivelNo = AdoAux.Recordset.Fields("Codigo")
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC = 'P'  " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Cta_Ventas = AdoAux.Recordset.Fields("Cta_Ventas")
     CodigoInv = AdoAux.Recordset.Fields("Codigo_Inv")
     Producto = AdoAux.Recordset.Fields("Producto")
  End If
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaI.Text)
  ProgBarra.Value = 0
  ProgBarra.Min = 0
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
       Line Input #NumFile, Cod_Field
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          CodigoCorresp = Mid(Cod_Field, 1, 3)
          If CodigoCorresp <> "ZZZ" Then
             If Contador <= 0 Then
                FechaInicial = Mid(Cod_Field, 27, 2) & "/" & Mid(Cod_Field, 25, 2) & "/" & Mid(Cod_Field, 21, 4)
                If FechaInicial = "00/00/0000" Then FechaInicial = FechaSistema
                FechaFinal = FechaInicial
             End If
             MiFecha = Mid(Cod_Field, 27, 2) & "/" & Mid(Cod_Field, 25, 2) & "/" & Mid(Cod_Field, 21, 4)
             If CFechaLong(MiFecha) <= CFechaLong(FechaInicial) Then FechaInicial = MiFecha
             If CFechaLong(MiFecha) >= CFechaLong(FechaFinal) Then FechaFinal = MiFecha
          End If
          Contador = Contador + 1
      Loop
  Close #NumFile
  ProgBarra.Max = Contador + 1
  FechaIni = BuscarFecha(FechaInicial)
  FechaFin = BuscarFecha(FechaFinal)
               sSQL = "DELETE * " _
                  & "FROM Facturas " _
                  & "WHERE Fecha = #" & FechaIni & "# Between #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             ConectarAdoExecute sSQL
             sSQL = "DELETE * " _
                  & "FROM Trans_Abonos " _
                  & "WHERE Fecha = #" & FechaIni & "# Between #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             ConectarAdoExecute sSQL
             sSQL = "DELETE * " _
                  & "FROM Detalle_Factura " _
                  & "WHERE Fecha = #" & FechaIni & "# Between #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
 ConectarAdoExecute sSQL
 '& "WHERE Factura = " & Factura_No & " "
  Contador = 0
  Open RutaGeneraFile For Input As #NumFile
       Line Input #NumFile, Cod_Field
       MiFecha = Format(Mid(Cod_Field, 119, 10), "dd/mm/yyyy")
       FechaTexto = MiFecha
       CodigoCorresp = Ninguno
       MiMes = FechaMes(FechaSistema)
       MiFecha = FechaSistema
       Total = 0
       MBFechaI.Text = MiFecha
       Contador = 0
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          CodigoCorresp = Mid(Cod_Field, 1, 3)
          If CodigoCorresp <> "ZZZ" Then
             Codigo = Trim(Mid(Cod_Field, 79, 10))
             If Codigo = "" Then Codigo = "."
             Saldo = Val(Mid(Cod_Field, 32, 13) & "." & Mid(Cod_Field, 45, 2))
             Factura_No = Val(Mid(Cod_Field, 284, 7))
             If Len(Trim(Mid(Cod_Field, 284, 7))) < 7 Then
                Factura_No = Val("1" & Mid(Cod_Field, 7, 6))
             End If
             MiFecha = Mid(Cod_Field, 27, 2) & "/" & Mid(Cod_Field, 25, 2) & "/" & Mid(Cod_Field, 21, 4)
             TipoDoc = Trim(Mid(Cod_Field, 257, 8))
             CodigoCli = Ninguno
             If AdoQuery.Recordset.RecordCount > 0 Then
                AdoQuery.Recordset.MoveFirst
                AdoQuery.Recordset.Find ("CI_RUC Like '*" & Codigo & "*' ")
                If Not AdoQuery.Recordset.EOF Then
                   CodigoCli = AdoQuery.Recordset.Fields("Codigo")
                End If
             End If
             DiaV = Val(Mid(MiFecha, 1, 2))
             MesV = Val(Mid(MiFecha, 4, 2))
             AñoV = Val(Mid(MiFecha, 7, 4))
             If AñoV <= 1900 Then CodigoCli = Ninguno
             If Not IsNumeric(Mid(MiFecha, 1, 2)) Then CodigoCli = Ninguno
             If Not IsNumeric(Mid(MiFecha, 4, 2)) Then CodigoCli = Ninguno
             If Not IsNumeric(Mid(MiFecha, 7, 4)) Then CodigoCli = Ninguno
             FBancoPacifico.Caption = "FACTURACION: " & Format(Contador / ProgBarra.Max, "00%") & " - " & Codigo & " - " & MiFecha & " - " & Factura_No & " - Valor: " & Format(Saldo, "#,##0.00")
             If CodigoCli = Ninguno Then
                CodigoCli = Codigo
               'MsgBox CodigoCli
                SetAdoAddNew "Clientes"
                SetAdoFields "T", Normal
                SetAdoFields "FA", 1
                SetAdoFields "FactM", 1
                SetAdoFields "Codigo", CodigoCli
                SetAdoFields "CI_RUC", Codigo
                SetAdoFields "Fecha", MiFecha
                SetAdoFields "Grupo", NumEmpresa
                SetAdoFields "Cliente", "NO EXISTE ALUMNO: " & Codigo
                SetAdoFields "Direccion", "NO EXISTE ALUMNO: " & Codigo
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoUpdate
             
                sSQL = "SELECT Codigo,CI_RUC,Cliente " _
                     & "FROM Clientes " _
                     & "WHERE Codigo <> '.' " _
                     & "ORDER BY Codigo "
                SelectAdodc AdoQuery, sSQL
                TextoImprimio = TextoImprimio & Codigo & vbCrLf
             End If
            
             SetAdoAddNew "Facturas"
             SetAdoFields "T", Cancelado
             SetAdoFields "TC", "FM"
             SetAdoFields "CodigoC", CodigoCli
             SetAdoFields "Fecha", MiFecha
             SetAdoFields "Fecha_C", MiFecha
             SetAdoFields "Fecha_V", MiFecha
             SetAdoFields "Factura", Factura_No
             SetAdoFields "SubTotal", Saldo
             SetAdoFields "Total_MN", Saldo
             SetAdoFields "Saldo_MN", 0
             SetAdoFields "Cta_CxP", Cta_Cobrar
             SetAdoFields "Cod_CxC", NivelNo
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
          
             SetAdoAddNew "Detalle_Factura"
             SetAdoFields "T", Pendiente
             SetAdoFields "TC", "FM"
             SetAdoFields "CodigoC", CodigoCli
             SetAdoFields "Factura", Factura_No
             SetAdoFields "Fecha", MiFecha
             SetAdoFields "Codigo", CodigoInv
             SetAdoFields "Producto", Producto & " (" & MesesLetras(FechaMes(MiFecha)) & ")"
             SetAdoFields "Cantidad", 1
             SetAdoFields "Precio", Saldo
             SetAdoFields "Total", Saldo
             SetAdoFields "Cta_Venta", Cta_Ventas
             SetAdoFields "CodigoL", NivelNo
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
          
             SetAdoAddNew "Trans_Abonos"
             SetAdoFields "T", Cancelado
             SetAdoFields "TP", "FM"
             SetAdoFields "CodigoC", CodigoCli
             SetAdoFields "Fecha", MiFecha
             SetAdoFields "Comprobante", TipoDoc
             SetAdoFields "Factura", Factura_No
             SetAdoFields "Abono", Saldo
             SetAdoFields "Banco", "ABONO POR BANCO"
             SetAdoFields "Cheque", CStr(Saldo)
             SetAdoFields "Cta", Cta_CajaG
             SetAdoFields "Cta_CxP", Cta_Cobrar
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
             Total = Total + Saldo
          End If
          Contador = Contador + 1
          ProgBarra.Value = Contador
       Loop
  Close #NumFile
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  NumEmpresa = AuxNumEmp
  FBancoPacifico.Caption = "FACTURACION DE BANCO DEL PACIFICO"
  MsgBox "Total Factura USD " & Format(Total, "#,##0.00") & vbCrLf _
         & "Desde: " & FechaInicial & vbCrLf _
         & "Hasta: " & FechaFinal & vbCrLf _
         & "Fin del Proceso"
  If TextoImprimio <> "" Then
     TextoImprimio = "Warning:" & vbCrLf & TextoImprimio
     FInfoError.Show
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  SiguienteControl
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  NombreArchivo = File1.FileName
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.FileName
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.FileName
     File1.FileName = Dir1.Path & "\*.*"
  End If
End Sub

Private Sub File1_LostFocus()
  NombreArchivo = UCase(File1.FileName)
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  Drive1.Drive = Mid(RutaSysBases, 1, 2)
  RatonNormal
  RutaBackup = RutaSysBases & "\BANCO"
  TipoProcesos ""
  FBancoPacifico.Caption = "FACTURACION DE BANCOS"
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> 0 " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     AdoAux.Recordset.MoveLast
     Codigo = AdoAux.Recordset.Fields("CI_RUC")
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FBancoPacifico
  If CodigoUsuario = "ACCESO02" Then Command6.Enabled = True
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoQuery
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

