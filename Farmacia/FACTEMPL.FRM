VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FactEmpl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   6600
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar Factura"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Data DataAsiento 
      Caption         =   "Asiento"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBCtls.DBCombo DBCmbCliente 
      Bindings        =   "FACTEMPL.frx":0000
      DataSource      =   "DataCliente"
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
   End
   Begin MSDBCtls.DBCombo DBCmbVendedor 
      Bindings        =   "FACTEMPL.frx":001A
      DataSource      =   "DataVendedor"
      Height          =   315
      Left            =   1680
      TabIndex        =   11
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Vendedor"
   End
   Begin VB.Data DataProductos 
      Caption         =   "Productos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1812
   End
   Begin MSDBCtls.DBCombo DBCmbArticulo 
      Bindings        =   "FACTEMPL.frx":0035
      DataSource      =   "DataArticulo"
      Height          =   285
      Left            =   1080
      TabIndex        =   27
      Top             =   2520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483646
      Text            =   "Productos"
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   2040
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Data DataBancos 
      Caption         =   "Bancos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.TextBox TextVUnit 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5976
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Data DataDetalle 
      Caption         =   "Detalle"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Data DataFactura 
      Caption         =   "Factura"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data DataVendedor 
      Caption         =   "Vendedor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5952
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data DataCliente 
      Caption         =   "Cliente"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TextRep 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   288
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   29
      Text            =   "0"
      Top             =   2520
      Width           =   612
   End
   Begin VB.TextBox TextCant 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   288
      Left            =   4920
      MaxLength       =   6
      TabIndex        =   28
      Text            =   "0"
      Top             =   2520
      Width           =   732
   End
   Begin VB.TextBox TextBultos 
      Height          =   288
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "0"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox TextTransportador 
      Height          =   288
      Left            =   1680
      MaxLength       =   45
      TabIndex        =   9
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox TextPedidos 
      Height          =   288
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   10
      Text            =   "0"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox TextFechaA 
      Height          =   288
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   6
      Top             =   720
      Width           =   372
   End
   Begin VB.TextBox TextFechaM 
      Height          =   288
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   5
      Top             =   720
      Width           =   372
   End
   Begin VB.TextBox TextFechaD 
      Height          =   288
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   4
      Top             =   720
      Width           =   372
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   39
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label LabelCodigos 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label LabelTotalCredito 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Credito Pendiente S/."
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7320
      TabIndex        =   37
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   34
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label LabelIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   36
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label LabelSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   35
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total con IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total sin IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   32
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label LabelVTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7680
      TabIndex        =   31
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio Unitario"
      Height          =   255
      Left            =   6240
      TabIndex        =   25
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rep."
      Height          =   255
      Left            =   5640
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label17 
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
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   2280
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bultos No."
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Transportado por:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pedido No."
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vendedor:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label LabelTelefono 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7800
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telefono:"
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Pedido (DD/MM/AA):"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label LabelFacturaNo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "emp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
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
      Top             =   120
      Width           =   4185
   End
End
Attribute VB_Name = "FactEmpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload FactEmpl
End Sub

Private Sub Command2_Click()
   FormaPag1.Show
End Sub

Private Sub DBCmbArticulo_GotFocus()
   sSQL = "SELECT * FROM Articulo ORDER BY Articulo "
   DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
End Sub

Private Sub DBCmbArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Key_Esc Then Command2.SetFocus
End Sub

Private Sub DBCmbArticulo_LostFocus()
   LabelCodigos.Caption = ""
   sSQL = "SELECT * FROM Articulo "
   sSQL = sSQL & "WHERE Articulo = '" & DBCmbArticulo.Text & "' "
   DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
   If DataArticulo.Recordset.RecordCount > 0 Then
      TextVUnit.Text = Format(DataArticulo.Recordset.Fields("PVP"), "###0")
      LabelCodigos.Caption = DataArticulo.Recordset.Fields("Codigo")
      If CalcularIVA Then BanIVA = DataArticulo.Recordset.Fields("IVA") Else BanIVA = False
   Else
      Mensajes = "Este Producto no existe," & Chr(13)
      Mensajes = Mensajes & "Repita la operacion."
      MsgBox Mensajes
      DBCmbArticulo.SetFocus
   End If
End Sub

Private Sub DBCmbCliente_GotFocus()
   sSQL = "SELECT * FROM Clientes ORDER BY Cliente "
   DataCliente.RecordSource = sSQL: DataCliente.Refresh
End Sub

Private Sub DBCmbCliente_LostFocus()
   LabelCodigo.Caption = ""
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & " WHERE Cliente = '" & DBCmbCliente.Text & "' "
   DataCliente.RecordSource = sSQL: DataCliente.Refresh
   If DataCliente.Recordset.RecordCount > 0 Then
      LabelTelefono.Caption = DataCliente.Recordset.Fields("Telefono_1")
      LabelTotalCredito.Caption = Format(DataCliente.Recordset.Fields("Saldo_Pendiente"), "#,##0")
      LabelCodigo.Caption = DataCliente.Recordset.Fields("Codigo")
      CodigoCli = LabelCodigo.Caption
      DireccionCli = DataCliente.Recordset.Fields("Direccion")
   Else
      Mensajes = "Este Cliente No existe," & Chr(13)
      Mensajes = Mensajes & "Repita la operacion."
      MsgBox Mensajes
      DBCmbCliente.SetFocus
   End If
End Sub

Private Sub DBCmbVendedor_GotFocus()
  sSQL = "Vendedor"
  DataVendedor.RecordSource = sSQL: DataVendedor.Refresh
End Sub

Private Sub DBCmbVendedor_LostFocus()
   sSQL = "SELECT * FROM Vendedor "
   sSQL = sSQL & " WHERE Vendedor = '" & DBCmbVendedor.Text & "' "
   DataVendedor.RecordSource = sSQL: DataVendedor.Refresh
   If DataVendedor.Recordset.RecordCount > 0 Then
      CodigoVen = DataVendedor.Recordset.Fields("Codigo")
   Else
      Mensajes = "Este Vendedor no existe," & Chr(13)
      Mensajes = Mensajes & "repita la operacion."
      MsgBox Mensajes
      DBCmbVendedor.SetFocus
   End If
End Sub

Private Sub DBGDetalle_AfterDelete()
  CalculosTotales
End Sub

Private Sub DBGDetalle_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & Chr(13) & "("
  Mensajes = Mensajes & DataProductos.Recordset.Fields("CODIGO") & ") "
  Mensajes = Mensajes & DataProductos.Recordset.Fields("PRODUCTO") & "   TOTAL -> "
  Mensajes = Mensajes & DataProductos.Recordset.Fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Facturas' "
   DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
   NumComp = DataCodigos.Recordset.Fields("Numero")
   sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Asiento' "
   DataAsiento.RecordSource = sSQL: DataAsiento.Refresh
   Asiento = DataAsiento.Recordset.Fields("Numero")
   sSQL = "SELECT * FROM Clientes ORDER BY Cliente "
   DataCliente.RecordSource = sSQL: DataCliente.Refresh
   DBCmbCliente.ListField = "Cliente"
   sSQL = "SELECT * FROM Articulo ORDER BY Articulo "
   DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
   DBCmbArticulo.ListField = "Articulo"
   sSQL = "SELECT * FROM Cuentas ORDER BY Codigo "
   DataBancos.RecordSource = sSQL: DataBancos.Refresh
   DataTrans.RecordSource = "Transacciones": DataTrans.Refresh
   DataVendedor.RecordSource = "Vendedor": DataVendedor.Refresh
   DBCmbVendedor.ListField = "Vendedor"
   DataFactura.RecordSource = "Facturas": DataFactura.Refresh
   DataDetalle.RecordSource = "Detalle_Factura": DataDetalle.Refresh
   sSQL = "Detalle_Empleados"
   DataProductos.RecordSource = sSQL: DataProductos.Refresh
   DBGDetalle.Columns(0).NumberFormat = "Yes/No"
   DBGDetalle.Columns(DataProductos.Recordset.Fields.Count - 2).NumberFormat = "##,##0"
   DBGDetalle.Columns(DataProductos.Recordset.Fields.Count - 1).NumberFormat = "##,##0"
   LabelFacturaNo = Format(NumComp, "000000")
   DBCmbCliente.Text = DataCliente.Recordset.Fields("Cliente")
   DBCmbVendedor.Text = DataVendedor.Recordset.Fields("Vendedor")
   DBCmbArticulo.Text = DataArticulo.Recordset.Fields("Articulo")
   MDIFact.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   'Centrar el formulario
   Left = (Screen.width - width) / 2
   Top = (Screen.Height - Height) / 2 - 600
   Label1.Caption = " " & Empresa
   ' Abriendo bases relacionadas
   DataCodigos.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCliente.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataFactura.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataDetalle.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataVendedor.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataProductos.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataBancos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsiento.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   TextCant.Text = ""
   TextRep.Text = ""
   TextVUnit.Text = ""
   LabelVTotal.Caption = ""
   Modificar = False
   Bandera = True
   
End Sub

Private Sub TextBultos_LostFocus()
  If TextBultos.Text = "" Then
     TextBultos.Text = "0"
     TextBultos.Text = Format(TextBultos.Text, "0000000000")
  End If
End Sub

Private Sub TextCant_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "###0")
End Sub

Private Sub TextCant_LostFocus()
   If TextCant.Text = "" Then TextCant.Text = "0"
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0")
End Sub

Private Sub TextFechaA_Change()
  If Len(TextFechaA.Text) >= TextFechaA.MaxLength Then DBCmbCliente.SetFocus
End Sub

Private Sub TextFechaA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaA_LostFocus()
If TextFechaA.Text = "" Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaD_Change()
  If Len(TextFechaD.Text) >= TextFechaD.MaxLength Then
     TextFechaM.SetFocus
  End If
End Sub

Private Sub TextFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Key_Enter Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaD_LostFocus()
  If TextFechaD.Text = "" Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaM_Change()
  If Len(TextFechaM.Text) >= TextFechaM.MaxLength Then
     TextFechaA.SetFocus
  End If
End Sub

Private Sub TextFechaM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaM_LostFocus()
  If TextFechaM.Text = "" Then TextFechaM.Text = Mes
End Sub

Private Sub TextPedidos_LostFocus()
  If TextPedidos.Text = "" Then
     TextPedidos.Text = "0"
     TextPedidos.Text = Format(TextPedidos.Text, "0000000000")
  End If
End Sub

Private Sub TextRep_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "###0")
End Sub

Private Sub TextRep_LostFocus()
   If TextRep.Text = "" Then TextRep.Text = "0"
   TextRep.Text = Format(TextRep.Text, "###0")
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "###0")
End Sub

Private Sub TextTransportador_LostFocus()
 If TextTransportador.Text = "" Then TextTransportador.Text = Ninguno
End Sub

Private Sub TextVUnit_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "###0")
End Sub

Private Sub TextVUnit_LostFocus()
   If DataProductos.Recordset.RecordCount <= 15 Then
      Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
      LabelVTotal.Caption = Format(Real1, "###0")
      DataProductos.Recordset.AddNew
      DataProductos.Recordset.Fields("IVA") = BanIVA
      DataProductos.Recordset.Fields("CODIGO") = LabelCodigos.Caption
      DataProductos.Recordset.Fields("PRODUCTO") = DBCmbArticulo.Text
      DataProductos.Recordset.Fields("REP") = Val(TextRep.Text)
      DataProductos.Recordset.Fields("CANT") = Val(TextCant.Text)
      DataProductos.Recordset.Fields("PRECIO") = Val(TextVUnit.Text)
      DataProductos.Recordset.Fields("TOTAL") = Real1
      If BanIVA Then Real2 = Real1 * 0.1 Else Real2 = 0
      DataProductos.Recordset.Fields("Total_IVA") = Real2
      If ((Val(TextCant.Text) > 0) And (LabelCodigos.Caption <> "")) Then DataProductos.Recordset.Update
      CalculosTotales
      TextVUnit.Text = ""
      DBCmbArticulo.SetFocus
   Else
      MsgBox "Ya no se puede ingresar mas datos."
      Command1.SetFocus
   End If
End Sub

Public Sub CalculosTotales()
   Total = 0: Total_Con_IVA = 0: Total_Sin_IVA = 0: Total_IVA = 0
   If DataProductos.Recordset.RecordCount > 0 Then
      DataProductos.Recordset.MoveFirst
      Do While Not DataProductos.Recordset.EOF
         Total_IVA = Total_IVA + DataProductos.Recordset.Fields("Total_IVA")
         If DataProductos.Recordset.Fields("IVA") Then
            Total_Con_IVA = Total_Con_IVA + DataProductos.Recordset.Fields("TOTAL")
         Else
            Total_Sin_IVA = Total_Sin_IVA + DataProductos.Recordset.Fields("TOTAL")
         End If
         DataProductos.Recordset.MoveNext
      Loop
   End If
   Total_Con_IVA = Round(Total_Con_IVA)
   Total_Sin_IVA = Round(Total_Sin_IVA)
   Total_IVA = Round(Total_IVA)
   Total = Total_Sin_IVA + Total_Con_IVA + Total_IVA
   LabelSubTotal.Caption = Format(Total_Sin_IVA, "#,##0")
   LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0")
   LabelIVA.Caption = Format(Total_IVA, "#,##0")
   LabelTotal.Caption = Format(Total, "#,##0")
   TextCant.Text = ""
   TextRep.Text = ""
   LabelVTotal.Caption = ""
End Sub

Public Sub Imprimir()
  If DataProductos.Recordset.RecordCount > 0 Then
     Mensajes = "Desea imprimir la factura"
     Titulo = "Formulario de impresion."
     TipoDeCaja = 4 + 32: K = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If K = 6 Then Call ImprimirFactura(Val(LabelFacturaNo.Caption))
   Else
     MsgBox "No se puede imprimir la Factura," & Chr(13) & "falta datos."
   End If
End Sub

Public Sub ProcGrabar()
Dim IVAS As Double
Dim Factura_No As Long
Dim TipoCta As String
  If DataProductos.Recordset.RecordCount > 0 Then
     CalculosTotales
     Mensajes = "Esta Seguro que desea grabar: " & Chr(13) & " La Factura No. " & LabelFacturaNo
     Titulo = "Formulario de Grabacion"
     TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If J = 6 Then
        MousePointer = vbHourglass
        FechaTexto = FormatoFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
        FechaTexto1 = FormatoFecha(FormaPag.TextFechaDP.Text, FormaPag.TextFechaMP.Text, FormaPag.TextFechaAP.Text)
        Factura_No = NumComp
        DataProductos.Recordset.MoveFirst
        Do While Not DataProductos.Recordset.EOF
           DataDetalle.Recordset.AddNew
           DataDetalle.Recordset.Fields("Estado") = Normal
           DataDetalle.Recordset.Fields("Factura_No") = Factura_No
           DataDetalle.Recordset.Fields("Fecha") = FechaTexto
           DataDetalle.Recordset.Fields("IVA") = DataProductos.Recordset.Fields("IVA")
           DataDetalle.Recordset.Fields("Codigo") = DataProductos.Recordset.Fields("CODIGO")
           DataDetalle.Recordset.Fields("Cantidad") = DataProductos.Recordset.Fields("CANT")
           DataDetalle.Recordset.Fields("Reposicion") = DataProductos.Recordset.Fields("REP")
           DataDetalle.Recordset.Fields("Precio") = DataProductos.Recordset.Fields("PRECIO")
           DataDetalle.Recordset.Fields("Total") = DataProductos.Recordset.Fields("TOTAL")
           DataDetalle.Recordset.Fields("Total_IVA") = DataProductos.Recordset.Fields("Total_IVA")
           DataDetalle.Recordset.Update
           DataProductos.Recordset.MoveNext
        Loop
        Total_Factura = Total_Sin_IVA + Total_Con_IVA + Total_IVA
        DataFactura.Recordset.AddNew
        DataFactura.Recordset.Fields("Estado") = Pendiente
        DataFactura.Recordset.Fields("Factura") = Factura_No
        DataFactura.Recordset.Fields("Fecha") = FechaTexto
        DataFactura.Recordset.Fields("Fecha_Pago") = FechaTexto1
        DataFactura.Recordset.Fields("Codigo_C") = CodigoCli
        DataFactura.Recordset.Fields("Codigo_V") = CodigoVen
        DataFactura.Recordset.Fields("Pedido_No") = TextPedidos.Text
        DataFactura.Recordset.Fields("Bultos_No") = TextBultos.Text
        DataFactura.Recordset.Fields("Transportador") = TextTransportador.Text
        DataFactura.Recordset.Fields("Cheque") = FormaPag.TextCheque.Text
        DataFactura.Recordset.Fields("Banco") = FormaPag.TextBanco.Text
        DataFactura.Recordset.Fields("Condicion_Pago") = FormaPag.ComboPago.Text
        DataFactura.Recordset.Fields("Sin_IVA") = Total_Sin_IVA
        DataFactura.Recordset.Fields("Con_IVA") = Total_Con_IVA
        DataFactura.Recordset.Fields("IVA") = Total_IVA
        DataFactura.Recordset.Fields("Total") = Total_Factura
        DataFactura.Recordset.Fields("Saldo") = Total_Factura
        DataFactura.Recordset.Update
        'Grabamos cuentas por cobrar
        Cadena = DBCmbCliente.Text & ":Fact." & Factura_No
        DataTrans.Recordset.AddNew
        DataTrans.Recordset.Fields("Fecha") = FechaTexto
        DataTrans.Recordset.Fields("Asiento") = 0
        DataTrans.Recordset.Fields("TP") = CompFactura
        DataTrans.Recordset.Fields("Comp") = Factura_No
        DataTrans.Recordset.Fields("Concepto") = Cadena
        DataTrans.Recordset.Fields("Cheq_Dep") = "0"
        DataTrans.Recordset.Fields("Cta") = Cta_Cobrar
        TipoCta = "0"
        sSQL = "SELECT * FROM Cuentas "
        sSQL = sSQL & "WHERE Codigo = '" & Cta_Cobrar & "' "
        DataCtas.RecordSource = sSQL: DataCtas.Refresh
        If DataCtas.Recordset.RecordCount > 0 Then TipoCta = DataCtas.Recordset.Fields("TC")
        DataTrans.Recordset.Fields("TC") = TipoCta
        DataTrans.Recordset.Fields("Debe") = Total_Factura
        DataTrans.Recordset.Fields("Haber") = 0
        DataTrans.Recordset.Fields("Estado") = Normal
        DataTrans.Recordset.Fields("Asiento") = Asiento
        DataTrans.Recordset.Update
        'grabamos el iva si lo tubiera
        '=================================
        Cadena = "IVA de la Factura No. (" & Factura_No & ")"
        If Total_IVA > 0 Then
           DataTrans.Recordset.AddNew
           DataTrans.Recordset.Fields("Fecha") = FechaTexto
           DataTrans.Recordset.Fields("Asiento") = 0
           DataTrans.Recordset.Fields("TP") = CompFactura
           DataTrans.Recordset.Fields("Comp") = Factura_No
           DataTrans.Recordset.Fields("Concepto") = Cadena
           DataTrans.Recordset.Fields("Cheq_Dep") = "0"
           DataTrans.Recordset.Fields("Cta") = Cta_IVA
           TipoCta = "0"
           sSQL = "SELECT * FROM Cuentas "
           sSQL = sSQL & "WHERE Codigo = '" & Cta_IVA & "' "
           DataCtas.RecordSource = sSQL: DataCtas.Refresh
           If DataCtas.Recordset.RecordCount > 0 Then TipoCta = DataCtas.Recordset.Fields("TC")
           DataTrans.Recordset.Fields("TC") = TipoCta
           DataTrans.Recordset.Fields("Debe") = 0
           DataTrans.Recordset.Fields("Haber") = Total_IVA
           DataTrans.Recordset.Fields("Estado") = Normal
           DataTrans.Recordset.Fields("Asiento") = Asiento
           DataTrans.Recordset.Update
        End If
        'empezamos a llenar los detalle de la factura
        '=============================================
        sSQL = "SELECT * FROM Detalle_Factura "
        sSQL = sSQL & "WHERE Factura_No = " & Factura_No
        DataDetalle.RecordSource = sSQL: DataDetalle.Refresh
        If DataDetalle.Recordset.RecordCount > 0 Then
           DataDetalle.Recordset.MoveFirst
           Do While Not DataDetalle.Recordset.EOF
              Total_Detalle = DataDetalle.Recordset.Fields("Total")
              CodProd = DataDetalle.Recordset.Fields("Codigo")
              Cadena = "Detalle(" & CodProd & ")Factura No. " & Factura_No
              DataTrans.Recordset.AddNew
              DataTrans.Recordset.Fields("Fecha") = FechaTexto
              DataTrans.Recordset.Fields("Asiento") = 0
              DataTrans.Recordset.Fields("TP") = CompFactura
              DataTrans.Recordset.Fields("Comp") = Factura_No
              DataTrans.Recordset.Fields("Concepto") = Cadena
              DataTrans.Recordset.Fields("Debe") = 0
              DataTrans.Recordset.Fields("Haber") = Total_Detalle
              DataTrans.Recordset.Fields("Cheq_Dep") = "0"
              DataTrans.Recordset.Fields("Estado") = Normal
              DataTrans.Recordset.Fields("Asiento") = Asiento
              sSQL = "SELECT * FROM Articulo WHERE Codigo = '" & CodProd & "' "
              DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
              If DataArticulo.Recordset.RecordCount > 0 Then CodProd = DataArticulo.Recordset.Fields("Cta")
              TipoCta = "0"
              sSQL = "SELECT * FROM Cuentas "
              sSQL = sSQL & "WHERE Codigo = '" & CodProd & "' "
              DataCtas.RecordSource = sSQL: DataCtas.Refresh
              If DataCtas.Recordset.RecordCount > 0 Then TipoCta = DataCtas.Recordset.Fields("TC")
              DataTrans.Recordset.Fields("Cta") = CodProd
              DataTrans.Recordset.Fields("TC") = TipoCta
              DataTrans.Recordset.Update
              DataDetalle.Recordset.MoveNext
           Loop
        End If
     'Grabamos el saldo del cliente
     sSQL = "UPDATE Clientes SET Saldo_Pendiente = Saldo_Pendiente + " & Round(Total_Factura) & " "
     SSQ = sSQL & "WHERE Cliente = '" & DBCmbCliente.Text & "' "
     DataCliente.Database.Execute sSQL
     'Grabamos el numero de factura
     sSQL = "UPDATE Codigos SET Numero = Numero + 1 WHERE Concepto = 'Facturas' "
     DataCodigos.Database.Execute sSQL
     sSQL = "UPDATE Codigos SET Numero = Numero + 1 WHERE Concepto = 'Asiento' "
     DataAsiento.Database.Execute sSQL
     sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Facturas' "
     DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
     NumComp = DataCodigos.Recordset.Fields("Numero")
     sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Asiento' "
     DataAsiento.RecordSource = sSQL: DataAsiento.Refresh
     Asiento = DataAsiento.Recordset.Fields("Numero")
     sSQL = "DELETE * FROM Detalle_Empleados "
     DataProductos.Database.Execute sSQL
     sSQL = "Detalle_Empleados "
     DataProductos.RecordSource = sSQL: DataProductos.Refresh
     DBGDetalle.Columns(0).NumberFormat = "Yes/No"
     DBGDetalle.Columns(DataProductos.Recordset.Fields.Count - 2).NumberFormat = "##,##0"
     DBGDetalle.Columns(DataProductos.Recordset.Fields.Count - 1).NumberFormat = "##,##0"
     MousePointer = vbDefault
     Mensajes = "Desea imprimir la factura"
     Titulo = "Formulario de impresion."
     TipoDeCaja = 4 + 32: K = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If K = 6 Then Call ImprimirFactura(Val(LabelFacturaNo.Caption))
     LabelFacturaNo = Format(NumComp, "000000")
     End If
  Else
     MsgBox "No se puede grabar la Factura," & Chr(13) & "falta datos."
  End If
End Sub
