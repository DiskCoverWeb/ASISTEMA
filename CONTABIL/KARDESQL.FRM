VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form KardexSQLs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTARIO DE KARDEX"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataProd 
      Caption         =   "Prod"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7140
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data DataProducto 
      Caption         =   "Producto"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSDBCtls.DBCombo DBCProducto 
      Bindings        =   "KardeSQL.frx":0000
      DataSource      =   "DataProducto"
      Height          =   360
      Left            =   2835
      TabIndex        =   10
      Top             =   945
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   635
      _Version        =   327680
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CONSULTAR"
      Height          =   435
      Left            =   6090
      TabIndex        =   11
      Top             =   735
      Width           =   1380
   End
   Begin VB.TextBox TextFechaDF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1470
      MaxLength       =   2
      TabIndex        =   6
      Top             =   945
      Width           =   435
   End
   Begin VB.TextBox TextFechaMF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   7
      Top             =   945
      Width           =   435
   End
   Begin VB.TextBox TextFechaAF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2310
      MaxLength       =   2
      TabIndex        =   8
      Top             =   945
      Width           =   435
   End
   Begin VB.TextBox TextFechaD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   105
      MaxLength       =   2
      TabIndex        =   2
      Top             =   945
      Width           =   435
   End
   Begin VB.TextBox TextFechaM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   525
      MaxLength       =   2
      TabIndex        =   3
      Top             =   945
      Width           =   435
   End
   Begin VB.TextBox TextFechaA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      MaxLength       =   2
      TabIndex        =   4
      Top             =   945
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir Consulta"
      Height          =   435
      Left            =   9345
      TabIndex        =   13
      Top             =   735
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir de la Consulta"
      Height          =   435
      Left            =   7560
      TabIndex        =   12
      Top             =   735
      Width           =   1695
   End
   Begin VB.Data DataDetKardex 
      Caption         =   "Query"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7140
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   315
      Visible         =   0   'False
      Width           =   2190
   End
   Begin MSDBGrid.DBGrid DBGQuery 
      Bindings        =   "KardeSQL.frx":0017
      Height          =   4965
      Left            =   105
      OleObjectBlob   =   "KardeSQL.frx":002F
      TabIndex        =   14
      Top             =   1365
      Width           =   9735
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Promedio Valor Unitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   18
      Top             =   6405
      Width           =   2640
   End
   Begin VB.Label LabelUnitario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6930
      TabIndex        =   17
      Top             =   6405
      Width           =   1800
   End
   Begin VB.Label LabelCant 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2310
      TabIndex        =   16
      Top             =   6405
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad Existente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   15
      Top             =   6405
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2835
      TabIndex        =   9
      Top             =   630
      Width           =   3165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   5
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4620
   End
End
Attribute VB_Name = "KardexSQLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  MousePointer = vbHourglass
  Call ImprimirData(0.5, 1, DataDetKardex, True, 2)
  MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  Unload KardexSQLs
End Sub

Private Sub Command3_Click()
  MousePointer = vbHourglass
  SSQL = "SELECT * FROM Materia_Prima "
  SSQL = SSQL & "WHERE Producto = '" & DBCProducto.Text & "' "
  DataProd.RecordSource = SSQL: DataProd.Refresh
  If DataProd.Recordset.RecordCount > 0 Then
     Codigo = DataProd.Recordset.Fields("Codigo")
  Else
     Codigo = Ninguno
  End If
  FechaIni = BuscarFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
  FechaFin = BuscarFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
  SSQL = "SELECT SUM(Valor_Unitario)/COUNT(*) AS Promedio "
  SSQL = SSQL & "FROM KARDEX "
  SSQL = SSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  SSQL = SSQL & "AND Codigo = '" & Codigo & "' "
  DataDetKardex.RecordSource = SSQL: DataDetKardex.Refresh
  If DataDetKardex.Recordset.RecordCount > 0 Then
     LabelUnitario.Caption = Format(DataDetKardex.Recordset.Fields("Promedio"), "#,##0.00")
  Else
     LabelUnitario.Caption = "0"
  End If
  SSQL = "SELECT Fecha,Referencia,Descripcion,Entrada,Salida,Valor_Unitario,Valor_Total,Cant_Existente,Saldo_Unitario,Saldo_Total "
  SSQL = SSQL & "FROM KARDEX "
  SSQL = SSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  SSQL = SSQL & "AND Codigo = '" & Codigo & "' "
  SSQL = SSQL & "ORDER BY Fecha,Referencia "
  DataDetKardex.RecordSource = SSQL: DataDetKardex.Refresh
  DBAnchoMallas KardexSQLs, DBGQuery, DataDetKardex
  If DataDetKardex.Recordset.RecordCount > 0 Then
     DataDetKardex.Recordset.MoveLast
     LabelCant.Caption = Format(DataDetKardex.Recordset.Fields("Cant_Existente"), "#,##0")
  Else
     LabelCant.Caption = "0"
  End If
  MousePointer = vbDefault
End Sub

Private Sub DataQuery_Validate(Action As Integer, Save As Integer)
End Sub

Private Sub Form_Activate()
  SSQL = "SELECT * FROM Materia_Prima "
  DataProducto.RecordSource = SSQL: DataProducto.Refresh
  DBCProducto.ListField = "Producto"
  DBCProducto.Text = DataProducto.Recordset.Fields("Producto")
End Sub

Private Sub Form_Load()
  AnchoForm KardexSQLs
  DBGQuery.Width = KardexSQLs.Width - 300
  DBGQuery.Left = ((KardexSQLs.Width - DBGQuery.Width) / 2) - 20
  Label1.Caption = " " & Empresa
  DataDetKardex.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
  DataProducto.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
  DataProd.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
End Sub

Private Sub TextFechaD_Change()
  If Len(TextFechaD.Text) >= TextFechaD.MaxLength Then TextFechaM.SetFocus
End Sub

Private Sub TextFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Key_Enter Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaD_LostFocus()
  If TextFechaD.Text = "" Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaM_Change()
  If Len(TextFechaM.Text) >= TextFechaM.MaxLength Then TextFechaA.SetFocus
End Sub

Private Sub TextFechaM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaM_LostFocus()
  If TextFechaM.Text = "" Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaA_Change()
  If Len(TextFechaA.Text) >= TextFechaA.MaxLength Then TextFechaD.SetFocus
End Sub

Private Sub TextFechaA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaA_LostFocus()
  If TextFechaA.Text = "" Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaDF_Change()
  If Len(TextFechaDF.Text) >= TextFechaDF.MaxLength Then TextFechaMF.SetFocus
End Sub

Private Sub TextFechaDF_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Key_Enter Then TextFechaDF.Text = Dia
End Sub

Private Sub TextFechaDF_LostFocus()
  If TextFechaDF.Text = "" Then TextFechaDF.Text = Dia
End Sub

Private Sub TextFechaMF_Change()
  If Len(TextFechaMF.Text) >= TextFechaMF.MaxLength Then TextFechaAF.SetFocus
End Sub

Private Sub TextFechaMF_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaMF.Text = Mes
End Sub

Private Sub TextFechaMF_LostFocus()
  If TextFechaMF.Text = "" Then TextFechaMF.Text = Mes
End Sub

Private Sub TextFechaAF_Change()
  If Len(TextFechaAF.Text) >= TextFechaAF.MaxLength Then Command3.SetFocus
End Sub

Private Sub TextFechaAF_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaAF.Text = Anio
End Sub

Private Sub TextFechaAF_LostFocus()
  If TextFechaAF.Text = "" Then TextFechaAF.Text = Anio
End Sub

