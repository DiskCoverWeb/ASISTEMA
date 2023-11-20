VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FTxtFacturas 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Espere un momento....     Estoy procesando las bases"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "FTxtFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Codigo Faltantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4305
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Codigo Faltantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3255
      Width           =   1065
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Recibir Abonos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4200
      TabIndex        =   14
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recibir &Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2205
      Width           =   1065
   End
   Begin VB.OptionButton OpcEnviar 
      BackColor       =   &H00FF8080&
      Caption         =   "&Enviar Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.OptionButton OpcRecibir 
      BackColor       =   &H00FF8080&
      Caption         =   "&Recibir Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2205
      TabIndex        =   1
      Top             =   105
      Width           =   1800
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   3990
      Left            =   2940
      TabIndex        =   12
      Top             =   1890
      Width           =   3165
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   5985
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Recibir Abonos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":156C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1155
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":1E12
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1065
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
      TabIndex        =   8
      Top             =   1890
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
      Height          =   960
      Left            =   6195
      Picture         =   "FTxtFact.frx":261C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5355
      Width           =   1065
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
      Height          =   3690
      Left            =   105
      TabIndex        =   9
      Top             =   2205
      Width           =   2850
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   3255
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
      Top             =   3570
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
      Top             =   3885
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
      TabIndex        =   3
      Top             =   420
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
      TabIndex        =   10
      Top             =   1680
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
      TabIndex        =   2
      Top             =   420
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2850
   End
End
Attribute VB_Name = "FTxtFacturas"
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

Public Sub TipoProcesos()
  NombreArchivo = ""
  If OpcEnviar.Value Then
     Dir1.Path = RutaBackup & "\Banco\Descuent\"
     File1.FileName = Dir1.Path & "\*.*"
  ElseIf OpcRecibir.Value Then
     Dir1.Path = RutaBackup & "\Banco\Facturas\"
     File1.FileName = Dir1.Path & "\*.*"
  Else
     Dir1.Path = RutaBackup & "\Banco\Abonos\"
     File1.FileName = Dir1.Path & "\*.*"
  End If
  Dir1.Refresh
End Sub

Private Sub Command1_Click()
Dim AuxNumEmp As String
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  'MsgBox RutaGeneraFile
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL = True " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Cta_Cobrar = AdoAux.Recordset.Fields("CxC")
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
       CodigoCorresp = Mid(Cod_Field, 1, 3)
       MiMes = Mid(Cod_Field, 4, 2)
       MiFecha = Mid(Cod_Field, 12, 2) & "/" & Mid(Cod_Field, 10, 2) & "/" & Mid(Cod_Field, 6, 4)
       Total = Val(CCur(Mid(Cod_Field, 14, 11) & "." & Mid(Cod_Field, 25, 2)))
       MBFechaI.Text = MiFecha
       ProgBarra.Max = Val(Mid(Cod_Field, 27, 5))
       Contador = 0
       sSQL = "DELETE * " _
            & "FROM Facturas " _
            & "WHERE Fecha = #" & BuscarFecha(MiFecha) & "# " _
            & "AND Item = '" & NumEmpresa & "' "
       ConectarAdoExecute sSQL

       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field

          Codigo = "FA" & Format(Val(Mid(Cod_Field, 1, 8)), "00000000")
          Saldo = Val(Mid(Cod_Field, 9, 9) & "." & Mid(Cod_Field, 18, 2))
          'MsgBox Saldo
          Factura = Val(Mid(Cod_Field, 26, 7))
          SetAdoAddNew "Facturas"
          SetAdoFields "T", Pendiente
          SetAdoFields "TC", "FM"
          SetAdoFields "CodigoC", Codigo
          SetAdoFields "Fecha", MiFecha
          SetAdoFields "Fecha_C", MiFecha
          SetAdoFields "Fecha_V", MiFecha
          SetAdoFields "Factura", Factura
          SetAdoFields "SubTotal", Saldo
          SetAdoFields "Total_MN", Saldo
          SetAdoFields "Saldo_MN", Saldo
          SetAdoFields "Cta_CxP", Cta_Cobrar
          SetAdoUpdate
          Contador = Contador + 1
          ProgBarra.Value = Contador
       Loop
  Close #NumFile
  ProgBarra.Value = 0
  Contador = 0
  sSQL = "SELECT F.CodigoC,F.Factura,F.SubTotal,CF.Codigo_Inv,CP.Producto,CP.Cta_Ventas " _
       & "FROM Facturas As F,Clientes_Facturacion As CF,Catalogo_Productos As CP " _
       & "WHERE F.Fecha = #" & BuscarFecha(MiFecha) & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.CodigoC = CF.Codigo " _
       & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
       & "AND F.Item = CF.Item " _
       & "AND F.Item = CP.Item " _
       & "AND CF.Item = CP.Item " _
       & "ORDER BY F.Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       ProgBarra.Max = .RecordCount
       sSQL = "DELETE * " _
            & "FROM Detalle_Factura " _
            & "WHERE Fecha = #" & BuscarFecha(MiFecha) & "# " _
            & "AND Item = '" & NumEmpresa & "' "
       ConectarAdoExecute sSQL
       Do While Not .EOF
          SetAdoAddNew "Detalle_Factura"
          SetAdoFields "T", Pendiente
          SetAdoFields "TC", "FM"
          SetAdoFields "CodigoC", .Fields("CodigoC")
          SetAdoFields "Factura", .Fields("Factura")
          SetAdoFields "Fecha", MiFecha
          SetAdoFields "Codigo", .Fields("Codigo_Inv")
          SetAdoFields "Producto", .Fields("Producto") & "(" & MesesLetras(FechaMes(MiFecha)) & ")"
          SetAdoFields "Cantidad", 1
          SetAdoFields "Precio", .Fields("SubTotal")
          SetAdoFields "Total", .Fields("SubTotal")
          SetAdoFields "Cta_Venta", .Fields("Cta_Ventas")
          SetAdoUpdate
          Contador = Contador + 1
          ProgBarra.Value = Contador
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  NumEmpresa = AuxNumEmp
  FTxtFacturas.Caption = "FACTURACION DE BANCOS"
  MsgBox "Total Factura USD " & Format(Total, "#,##0.00") & vbCrLf & "Fin del Proceso"
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload FTxtFacturas
End Sub

Private Sub Command3_Click()
  RatonReloj
  sSQL = "UPDATE Facturas " _
       & "SET Forma_Pago = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET Forma_Pago = 'F' " _
          & "FROM Facturas As F,Clientes As C "
  Else
     sSQL = "UPDATE Facturas As F,Clientes As C " _
          & "SET F.Forma_Pago = 'F' "
  End If
  sSQL = sSQL & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.CodigoC = C.Codigo "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT CodigoC,Factura,Total_MN " _
       & "FROM Facturas " _
       & "WHERE Forma_Pago = '.' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoAux, sSQL
  GenerarArchivoPlano FTxtFacturas, AdoAux, "FALTANTE.TXT", True
  RatonNormal
End Sub

Private Sub Command4_Click()
'MsgBox NombreFile
FechaValida MBFechaI
MiMes = Format(FechaMes(MBFechaI.Text), "00")
MiFecha = Format(MBFechaI.Text, "YYYYMMDD")
RutaGeneraFile = UCase(Dir1.Path & "\DESC" & CodigoDelBanco & ".TXT")
sSQL = "SELECT CF.*,CP.PVP " _
     & "FROM Clientes_Facturacion As CF, Catalogo_Productos As CP " _
     & "WHERE CF.Item = '" & NumEmpresa & "' " _
     & "AND CF.Valor > 0 " _
     & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
     & "AND CF.Item = CP.Item " _
     & "AND CF.Valor <> CP.PVP " _
     & "ORDER BY Codigo "
SelectAdodc AdoAux, sSQL
NumFile = FreeFile
Contador = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With AdoAux.Recordset
 If .RecordCount > 0 Then
     ProgBarra.Max = .RecordCount
     Do While Not .EOF
        Contador = Contador + 1
        ProgBarra.Value = Contador
        Codigo = Mid(.Fields("Codigo"), 3, 8)
        Print #NumFile, CodigoDelBanco;
        Print #NumFile, Codigo;
        Print #NumFile, "001";
        Print #NumFile, Format(.Fields("Valor"), "00000000.00");
        Print #NumFile, "V"
       .MoveNext
     Loop
 End If
End With
Close #NumFile
ProgBarra.Value = ProgBarra.Max
MsgBox "Fin del Proceso"
'Unload Me
''        Print #NumFile, MiMes;
''        Print #NumFile, MiFecha;
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
  RutaBackup = RutaSysBases
  TipoProcesos
  FTxtFacturas.Caption = "FACTURACION DE BANCOS"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FTxtFacturas
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

Private Sub OpcEnviar_Click()
  TipoProcesos
End Sub

Private Sub OpcRecibir_Click()
  TipoProcesos
End Sub

Private Sub Option1_Click()
  TipoProcesos
End Sub
