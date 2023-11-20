VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RAbonos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA                                                                      Cancelacion Factura"
   ClientHeight    =   5160
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1155
      TabIndex        =   7
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   10
      Format          =   "dd/mm/aaaa"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Pagos"
      Height          =   855
      Left            =   6930
      Picture         =   "RAbonos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1485
   End
   Begin MSDBGrid.DBGrid DBGIngCaja 
      Bindings        =   "RAbonos.frx":066A
      Height          =   4005
      Left            =   105
      OleObjectBlob   =   "RAbonos.frx":0683
      TabIndex        =   6
      Top             =   1050
      Width           =   9780
   End
   Begin VB.Data DataClien 
      Caption         =   "Clien"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4410
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2835
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   855
      Left            =   8505
      Picture         =   "RAbonos.frx":1044
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar Pagos"
      Height          =   855
      Left            =   5250
      Picture         =   "RAbonos.frx":12C6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1590
   End
   Begin VB.Data DataDiarioCaja 
      Caption         =   "DiarioCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2625
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSDBCtls.DBCombo DBCCliente 
      Bindings        =   "RAbonos.frx":1708
      DataSource      =   "DataClientes"
      Height          =   315
      Left            =   1155
      TabIndex        =   2
      Top             =   525
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "Cliente"
   End
   Begin VB.Data DataClientes 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1050
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   960
   End
End
Attribute VB_Name = "RAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoCliente As String
Dim NombreCliente As String

Private Sub Command1_Click()
  RatonReloj
  sSQL = "SELECT Factura,Monto_ME,Monto_MN,Caja_ME,Caja_MN,Caja_Vaucher,Abonos_ME,Abonos_MN,"
  sSQL = sSQL & "Saldo_ME,Saldo_MN,Fecha,Diario_No,Caja_No,CtaxCob,CtaxVent,Cotizacion "
  sSQL = sSQL & "FROM Diario_Caja "
  sSQL = sSQL & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
  sSQL = sSQL & "AND T = '" & Normal & "' "
  sSQL = sSQL & "AND Codigo_C = '" & CodigoCliente & "' "
  sSQL = sSQL & "AND TP = '" & CxC & "' "
  sSQL = sSQL & "ORDER BY Factura "
  SelectDBGrid DBGIngCaja, DataDiarioCaja, sSQL
  With DataDiarioCaja.Recordset
        Total = 0
        Do While Not .EOF
           Total = Total + .Fields("Abonos_MN")
          .MoveNext
        Loop
  End With
  RatonNormal
End Sub

Private Sub Command2_Click()
   Unload RAbonos
End Sub

Private Sub Command3_Click()
   ImprimirReciboCaja DataDiarioCaja, NombreCliente
   DBCCliente.SetFocus
End Sub


Private Sub DBCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Command1.SetFocus
End Sub

Private Sub DBCCliente_LostFocus()
   NombreCliente = DBCCliente.Text
   Codigo = Ninguno
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE Cliente = '" & DBCCliente.Text & "' "
   SelectData DataClien, sSQL, False
   If DataClien.Recordset.RecordCount > 0 Then
      Codigo = DataClien.Recordset.Fields("Codigo")
      CodigoCliente = Codigo
   Else
      MsgBox "Este Cliente no existe."
      DBCCliente.SetFocus
   End If
End Sub

Private Sub Form_Activate()
   MiFecha = BuscarFecha(FechaSistema)
   sSQL = "SELECT * FROM Clientes ORDER BY Cliente "
   SelectDBCombo DBCCliente, DataClientes, sSQL, "Cliente", False
   RatonNormal
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm RAbonos
   'MsgBox UCase(App.EXEName)
   Select Case UCase(App.EXEName)
     Case "FACTHOTEL"
          DataDiarioCaja.DatabaseName = RutaEmpresa & "\FACHOTEL.MDB"
          DataClientes.DatabaseName = RutaEmpresa & "\FACHOTEL.MDB"
          DataClien.DatabaseName = RutaEmpresa & "\FACHOTEL.MDB"
     Case "FACTTURS"
          DataDiarioCaja.DatabaseName = RutaEmpresa & "\FACTTURS.MDB"
          DataClientes.DatabaseName = RutaEmpresa & "\FACTTURS.MDB"
          DataClien.DatabaseName = RutaEmpresa & "\FACTTURS.MDB"
   End Select
End Sub

Private Sub MBoxFecha_GotFocus()
  MBoxFecha.Text = LimpiarFechas
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
End Sub


