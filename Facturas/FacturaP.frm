VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturasP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   7065
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataContrato 
      Caption         =   "Contrato"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4410
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CheckBox CheckContrato 
      Caption         =   "Contrato No."
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
      Top             =   0
      Width           =   1485
   End
   Begin VB.TextBox TextGrupo_No 
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
      Left            =   6510
      MaxLength       =   15
      TabIndex        =   7
      Top             =   420
      Width           =   1065
   End
   Begin VB.CheckBox CheqFormato 
      Caption         =   "Formato Propio"
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
      Left            =   6825
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TextDesc 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "FacturaP.frx":0000
      Top             =   6615
      Width           =   1380
   End
   Begin VB.Data DataBuscarCli 
      Caption         =   "BuscarCli"
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
      Top             =   4410
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataAux 
      Caption         =   "Aux"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   5670
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox TextProducto 
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
      Height          =   330
      Left            =   105
      MaxLength       =   120
      TabIndex        =   28
      Text            =   "x"
      Top             =   2835
      Width           =   10725
   End
   Begin VB.TextBox TextVUnit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   7770
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "FacturaP.frx":0005
      Top             =   2520
      Width           =   1485
   End
   Begin VB.TextBox TextCant 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   6825
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "FacturaP.frx":000A
      Top             =   2520
      Width           =   960
   End
   Begin VB.TextBox TextObs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1575
      MaxLength       =   200
      TabIndex        =   22
      Top             =   1785
      Width           =   9255
   End
   Begin VB.TextBox TextNota 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1575
      MaxLength       =   200
      TabIndex        =   20
      Top             =   1470
      Width           =   9255
   End
   Begin MSDBCtls.DBCombo DBCmbCliente 
      Bindings        =   "FacturaP.frx":000C
      DataSource      =   "DataCliente"
      Height          =   315
      Left            =   1575
      TabIndex        =   15
      Top             =   1155
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextFacturaNo 
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
      Left            =   9240
      TabIndex        =   9
      Text            =   "0"
      Top             =   420
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
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
   Begin VB.Data DataLinea 
      Caption         =   "Linea"
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
      Top             =   4725
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataEjecutivo 
      Caption         =   "Ejecutivo"
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
      RecordSource    =   "Clientes"
      Top             =   5355
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataFact1 
      Caption         =   "Fact1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
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
      Height          =   645
      Left            =   1050
      TabIndex        =   42
      Top             =   6300
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar &Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   41
      Top             =   6300
      Width           =   855
   End
   Begin MSDBCtls.DBCombo DBCmbArticulo 
      Bindings        =   "FacturaP.frx":0026
      DataSource      =   "DataArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   27
      Top             =   2520
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483646
      Text            =   "Productos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data DataProductos 
      Caption         =   "Productos"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5670
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataDetalle 
      Caption         =   "Detalle"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4725
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataFactura 
      Caption         =   "Factura"
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
      Top             =   5670
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataCliente 
      Caption         =   "Cliente"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   5355
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataTrans 
      Caption         =   "DiarioCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4725
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSDBCtls.DBCombo DBCLinea 
      Bindings        =   "FacturaP.frx":0041
      DataSource      =   "DataLinea"
      Height          =   315
      Left            =   1575
      TabIndex        =   11
      Top             =   840
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data DataEjec 
      Caption         =   "Ejec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   5355
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Factura/ME"
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
      Left            =   9240
      TabIndex        =   13
      Top             =   840
      Width           =   1485
   End
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   4305
      TabIndex        =   5
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
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
   Begin VB.Data DataListFact 
      Caption         =   "ListFact"
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
      Top             =   4095
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSDBCtls.DBCombo DBCContrato 
      Bindings        =   "FacturaP.frx":0059
      DataSource      =   "DataContrato"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LabelTelefono 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9240
      TabIndex        =   18
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telefono:"
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
      Left            =   8085
      TabIndex        =   17
      Top             =   1155
      Width           =   1170
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   6720
      TabIndex        =   16
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label38 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &GRUPO:"
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
      Left            =   5670
      TabIndex        =   6
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Venc.:"
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
      Left            =   2940
      TabIndex        =   4
      Top             =   420
      Width           =   1380
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   9135
      TabIndex        =   38
      Top             =   6615
      Width           =   1695
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
      Height          =   330
      Left            =   9135
      TabIndex        =   35
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label LabelIVA 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7770
      TabIndex        =   37
      Top             =   6615
      Width           =   1380
   End
   Begin VB.Label LabelServ 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6405
      TabIndex        =   47
      Top             =   6615
      Width           =   1380
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
      Height          =   330
      Left            =   7770
      TabIndex        =   39
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio 10%"
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
      Height          =   330
      Left            =   6405
      TabIndex        =   32
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total &Desc."
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
      Left            =   5040
      TabIndex        =   43
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label LabelConIVA 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3570
      TabIndex        =   40
      Top             =   6615
      Width           =   1485
   End
   Begin VB.Label LabelVTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   9240
      TabIndex        =   31
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label LabelStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   5880
      TabIndex        =   46
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " OBSERVACION"
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
      TabIndex        =   21
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOTA"
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
      TabIndex        =   19
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
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
      TabIndex        =   14
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
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
      Height          =   330
      Left            =   7770
      TabIndex        =   8
      Top             =   420
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha &Emisión:"
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
      Width           =   1485
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LINEA:"
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
      TabIndex        =   10
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Left            =   9240
      TabIndex        =   26
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio Unitario"
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
      Left            =   7770
      TabIndex        =   25
      Top             =   2205
      Width           =   1485
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
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
      Left            =   6825
      TabIndex        =   24
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock"
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
      Left            =   5880
      TabIndex        =   45
      Top             =   2205
      Width           =   960
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
      Height          =   330
      Left            =   2100
      TabIndex        =   36
      Top             =   6615
      Width           =   1485
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
      Height          =   330
      Left            =   3570
      TabIndex        =   34
      Top             =   6300
      Width           =   1485
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
      Height          =   330
      Left            =   2100
      TabIndex        =   33
      Top             =   6300
      Width           =   1485
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   23
      Top             =   2205
      Width           =   5790
   End
End
Attribute VB_Name = "FacturasP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckContrato_Click()
  If CheckContrato.Value = 1 Then
     DBCContrato.Visible = True
  Else
     DBCContrato.Visible = False
  End If
End Sub

Private Sub CheqFormato_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Command1_Click()
  Unload FacturasP
End Sub

Private Sub Command2_Click()
   Mensajes = "Esta Seguro que desea grabar: " & Chr(13) & " La Factura No. " & TextFacturaNo.Text
   Titulo = "Formulario de Grabacion"
   TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
   If J = 6 Then
      If Check1.Value = 1 Then Moneda_US = True Else Moneda_US = False
      CalculosTotales FacturasP, DataProductos
      TextoFormaPago = PagoCont
      ProcGrabar
   End If
   NumComp = ReadSetDataNum("Facturas", True, False)
   TextFacturaNo.Text = NumComp
   TextFacturaNo.SetFocus
End Sub

Private Sub DBCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DBCLinea_LostFocus()
 Si_No = False
 CodigoL = SinEspaciosIzq(DBCLinea.Text)
 sSQL = "SELECT * FROM Linea_Producto "
 sSQL = sSQL & "WHERE Codigo = '" & CodigoL & "' "
 SelectData DataFact1, sSQL
 With DataFact1.Recordset
  If .RecordCount > 0 Then
      If .Fields("Cta_CxC") <> Ninguno Then Cta_Cobrar = .Fields("Cta_CxC")
      Si_No = .Fields("Todos")
  End If
 End With
 If Si_No Then
    sSQL = "SELECT Articulo & ' -> '& Codigo As Nom_Art " _
         & "FROM Articulo " _
         & "ORDER BY Articulo "
 Else
    sSQL = "SELECT Articulo & ' -> '& Codigo As Nom_Art " _
         & "FROM Articulo " _
         & "WHERE CodigoL = '" & CodigoL & "' " _
         & "ORDER BY Articulo "
 End If
 SelectDBCombo DBCmbArticulo, DataArticulo, sSQL, "Nom_Art", False
End Sub

Private Sub DBCmbArticulo_GotFocus()
   LabelStock.Caption = ""
End Sub

Private Sub DBCmbArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
     Case vbKeyEscape
          Empleados = False
          Command2.SetFocus
     Case vbKeyReturn
          TextProducto.SetFocus
   End Select
End Sub

Private Sub DBCmbArticulo_LostFocus()
   Producto = ""
   Codigos = SinEspaciosDer(DBCmbArticulo.Text)
   sSQL = "SELECT * FROM Articulo "
   sSQL = sSQL & "WHERE Codigo = '" & Codigos & "' "
   SelectData DataFact1, sSQL, False
   With DataFact1.Recordset
   If .RecordCount > 0 Then
      'DBCmbArticulo.Text = .Fields("Articulo")
      Producto = .Fields("Articulo")
      TextVUnit.Text = Format(.Fields("PVP"), "#,##0.00")
      LabelStock.Caption = .Fields("Stock")
      If Porc_IVA > 0 Then BanIVA = .Fields("IVA") Else BanIVA = False
      Cadena = ""
      Mensajes = "Escriba el nombre del Empleado ?"
      If Empleados Then
         Cadena = InputBox(Mensajes, , "")
         Cadena = " (" & Cadena & ")"
      End If
   Else
      Mensajes = "Este Producto no existe," & Chr(13)
      Mensajes = Mensajes & "Repita la operacion."
      MsgBox Mensajes
      DBCmbArticulo.SetFocus
   End If
   End With
End Sub

Private Sub DBCmbCliente_GotFocus()
   MarcarTexto DBCmbCliente
   'sSQL = "SELECT * FROM Clientes ORDER BY Cliente "
   'SelectData DataCliente, sSQL, False
End Sub

Private Sub DBCmbCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DBCmbCliente_LostFocus()
   Empleados = False
   LabelCodigo.Caption = ""
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & " WHERE Cliente = '" & DBCmbCliente.Text & "' "
   SelectData DataAux, sSQL, False
   With DataAux.Recordset
    If .RecordCount > 0 Then
        DBCmbCliente.Text = .Fields("Cliente")
        LabelTelefono.Caption = .Fields("Telefono")
        LabelCodigo.Caption = .Fields("Codigo")
        CodigoCli = LabelCodigo.Caption
        DireccionCli = .Fields("Direccion")
   Else
        DBCmbCliente.SetFocus
   End If
   End With
End Sub

Private Sub DBGDetalle_AfterDelete()
  CalculosTotales FacturasP, DataProductos
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
   CTDetalles_F
   TextGrupo_No.Text = "0"
   Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
   Label36.Caption = "Serv. " & Format(Porc_Serv * 100, "#0.00") & "%"
   TextCant.Text = "0"
   TextVUnit.Text = "0"
   LabelVTotal.Caption = "0"
   Modificar = False
   Bandera = True
   Mifecha = BuscarFecha(FechaSistema)
   
   sSQL = "SELECT Contrato_No & Space(4) & Fecha & ' , de: ' & Cliente "
   sSQL = sSQL & " & ' , Por MN: ' & format(Monto_MN,'#,###') "
   sSQL = sSQL & " & ' o ME: ' & format(Monto_ME,'#,###.00') As Contrato "
   sSQL = sSQL & "FROM Contratos_Meses,Clientes "
   sSQL = sSQL & "WHERE Contratos_Meses.Codigo_C = Clientes.Codigo "
   sSQL = sSQL & "AND Fecha <= #" & Mifecha & "# "
   sSQL = sSQL & "AND T = '" & Pendiente & "' "
   sSQL = sSQL & "ORDER BY Contrato_No,Fecha "
   SelectDBCombo DBCContrato, DataContrato, sSQL, "Contrato", False
   
   sSQL = "SELECT Codigo & '  ' & Linea As LineaP "
   sSQL = sSQL & "FROM Linea_Producto "
   sSQL = sSQL & "ORDER BY Linea "
   SelectDBCombo DBCLinea, DataLinea, sSQL, "LineaP", False
   CodigoL = SinEspaciosIzq(DBCLinea.Text)
   
   sSQL = "SELECT Articulo & ' -> '& Codigo As Nom_Art "
   sSQL = sSQL & "FROM Articulo "
   sSQL = sSQL & "WHERE CodigoL = '" & CodigoL & "' "
   sSQL = sSQL & "ORDER BY Articulo "
   SelectDBCombo DBCmbArticulo, DataArticulo, sSQL, "Nom_Art", False
   
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE E = 'C' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DBCmbCliente, DataCliente, sSQL, "Cliente", False
   
   sSQL = "SELECT * FROM Detalles_" & CodigoUsuario & " "
   SelectData DataProductos, sSQL
   sSQL = "DELETE * FROM Detalles_" & CodigoUsuario & " "
   DeleteData DataProductos, sSQL
   SelectDBGrid DBGDetalle, DataProductos, "Detalles_" & CodigoUsuario & " "
   
   NumComp = ReadSetDataNum("Facturas", True, False)
   TextFacturaNo.Text = NumComp
   RatonNormal
   CheckContrato.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FacturasP
  'Abriendo bases relacionadas
   DataAux.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataEjec.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataLinea.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCliente.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataEjecutivo.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataProductos.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataBuscarCli.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataFact1.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataListFact.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataFactura.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataDetalle.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataTrans.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataContrato.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
   FechaTexto1 = MBoxFecha.Text
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV
End Sub

Private Sub TextCant_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  If TextCant.Text = "" Then TextCant.Text = "0"
  Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
  LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
  TextoValido TextDesc, True
  CalculosTotales FacturasP, DataProductos
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
 If TextFacturaNo.Text = "" Then TextFacturaNo.Text = "0"
 Factura_No = Val(TextFacturaNo.Text)
 sSQL = "SELECT * FROM Facturas WHERE Factura = " & Factura_No & " "
 SelectData DataFactura, sSQL
 If DataFactura.Recordset.RecordCount > 0 Then
    MsgBox "Warning:" & Chr(13) & "Ya existe la Factura No. " & Format(Factura_No, "000000") & "."
    DBCLinea.SetFocus
 End If
End Sub

Private Sub TextGrupo_No_GotFocus()
  MarcarTexto TextGrupo_No
End Sub

Private Sub TextGrupo_No_LostFocus()
  TextoValido TextGrupo_No
  sSQL = "SELECT * FROM Clientes "
  sSQL = sSQL & "WHERE E = 'C' "
  sSQL = sSQL & "AND Grupo = '" & TextGrupo_No.Text & "' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBCombo DBCmbCliente, DataCliente, sSQL, "Cliente", False
  'DBCmbCliente.SetFocus
End Sub

Private Sub TextNota_GotFocus()
   MarcarTexto TextNota
End Sub

Private Sub TextNota_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNota_LostFocus()
  TextoValido TextNota
End Sub

Private Sub TextObs_GotFocus()
  MarcarTexto TextObs
End Sub

Private Sub TextObs_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextObs_LostFocus()
  TextoValido TextObs
End Sub

Private Sub TextProducto_GotFocus()
  TextProducto.Text = ""
End Sub

Private Sub TextProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextProducto_LostFocus()
   TextoValido TextProducto
End Sub

Private Sub TextVUnit_Change()
   'Real1 = CDbl(TextCant.Text) * CDbl(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
   TextoValido TextVUnit, True
   TextoValido TextCant, True
   If DataProductos.Recordset.RecordCount <= 30 Then
     'Producto = DBCmbArticulo.Text
      If TextProducto.Text <> Ninguno Then Producto = Producto & ", " & TextProducto.Text
      Real1 = CDbl(TextCant.Text) * CDbl(TextVUnit.Text)
      'Real2 = CDbl(TextRep.Text) * CDbl(TextVUnit.Text)
      LabelVTotal.Caption = Format(Real1, "#,##0.00")
      DataProductos.Recordset.AddNew
      DataProductos.Recordset.Fields("CODIGO") = Codigos
      DataProductos.Recordset.Fields("CODIGO_L") = CodigoL
      DataProductos.Recordset.Fields("PRODUCTO") = Mid(Producto, 1, 151)
      DataProductos.Recordset.Fields("REP") = 0
      DataProductos.Recordset.Fields("CANT") = CDbl(TextCant.Text)
      DataProductos.Recordset.Fields("PRECIO") = CDbl(TextVUnit.Text)
      DataProductos.Recordset.Fields("TOTAL") = Real1
      DataProductos.Recordset.Fields("Total_Desc") = Real2
      If BanIVA Then Real3 = (Real1 - Real2) * Porc_IVA Else Real3 = 0
      DataProductos.Recordset.Fields("Total_IVA") = Real3
      DataProductos.Recordset.Fields("Cod_Ejec") = Ninguno
      DataProductos.Recordset.Fields("Porc_C") = 0
      If ((Val(TextCant.Text) > 0) And (Codigos <> "")) Then DataProductos.Recordset.Update
      CalculosTotales FacturasP, DataProductos
      TextVUnit.Text = ""
      DBCmbArticulo.SetFocus
   Else
      MsgBox "Ya no se puede ingresar mas datos."
      Command1.SetFocus
   End If
End Sub

Public Sub ProcGrabar()
  'Seteamos los encabezados para las facturas
  If DataProductos.Recordset.RecordCount > 0 Then
     RatonReloj
     TextoValido TextNota
     TextoValido TextObs
     FechaValida MBoxFechaV, False
     If Check1.Value = 1 Then Moneda_US = True Else Moneda_US = False
     Asiento = ReadSetDataNum("Asiento", True, True)
     IngresoCaja = ReadSetDataNum("Ingreso Caja", True, True)
     SelectData DataFactura, "Facturas", False
     SelectData DataDetalle, "Detalle_Factura", False
    'SelectData DataTrans, "Transacciones ", False
     FechaTexto = MBoxFecha.Text
     Total_FacturaME = 0
     CalculosTotales FacturasP, DataProductos
     If Moneda_US Then
        Total_Factura = Round((Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio) * Dolar)
        Total_FacturaME = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
     Else
        Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
        Total_FacturaME = 0
     End If
     Saldo = Total_Factura
     Saldo_ME = Total_FacturaME
     DiarioCaja = 0
     NumComp = ReadSetDataNum("Facturas", True, False)
     If Val(TextFacturaNo.Text) <> NumComp Then
        Factura_No = Val(TextFacturaNo.Text)
     Else
        Factura_No = ReadSetDataNum("Facturas", True, True)
     End If
     sSQL = "DELETE * FROM Detalle_Factura "
     sSQL = sSQL & "WHERE Factura_No = " & Factura_No & " "
     DeleteData DataTrans, sSQL
     sSQL = "DELETE * FROM Facturas "
     sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
     DeleteData DataTrans, sSQL
     sSQL = "DELETE * FROM Diario_Caja "
     sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
     DeleteData DataTrans, sSQL
     SelectData DataTrans, "Diario_Caja ", False
     With DataTrans.Recordset
         .AddNew
          'If FacturaNueva Then
         .Fields("T") = Normal
          'Else
          '  .Fields("T") = Procesado
          'End If
         .Fields("TP") = ventas
         .Fields("Fecha") = FechaTexto
         .Fields("Diario_No") = DiarioCaja
         .Fields("Caja_No") = IngresoCaja
         .Fields("Factura") = Factura_No
         .Fields("Banco") = Ninguno
         .Fields("Cheque") = Ninguno
         .Fields("Monto_ME") = Total_FacturaME
         .Fields("Monto_MN") = Total_Factura
         .Fields("Caja_ME") = 0
         .Fields("Caja_MN") = 0
         .Fields("Caja_Vaucher") = 0
         .Fields("Abonos_ME") = 0
         .Fields("Abonos_MN") = 0
         .Fields("Saldo_MN") = 0
         .Fields("Saldo_ME") = 0
         .Fields("Codigo_C") = CodigoCli
         .Fields("CtaxCob") = Cta_Cobrar
         .Fields("CtaxVent") = Cta_Ventas
         .Fields("Saldo_ME") = Saldo_ME
         .Fields("Saldo_MN") = Saldo
         .Fields("Cotizacion") = 0
          If Moneda_US Then .Fields("Cotizacion") = Dolar
         .Update
     End With
     If Saldo < 0 Then Saldo = 0
     If Saldo > 0 Then
        TextoFormaPago = PagoCred
        T = Pendiente
     Else
        TextoFormaPago = PagoCont
        T = Cancelado
     End If
    'Grabamos el numero de factura
     TextoProc = Ninguno
     If CheckContrato.Value <> 0 Then
        Mifecha = ObtenerPalabra(DBCContrato.Text, 3)
        Mifecha = Mid(Mifecha, 1, Len(Mifecha))
        TextoProc = SinEspaciosIzq(DBCContrato.Text)
        sSQL = "UPDATE Contratos_Meses SET T = '" & Procesado & "' "
        sSQL = sSQL & "WHERE Contrato_No = '" & TextoProc & "' "
        sSQL = sSQL & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
        UpdateData DataContrato, sSQL
        
        sSQL = "SELECT Contrato_No & Space(4) & Fecha & ' , de: ' & Cliente "
        sSQL = sSQL & " & ' , Por MN: ' & format(Monto_MN,'#,###') "
        sSQL = sSQL & " & ' o ME: ' & format(Monto_ME,'#,###.00') As Contrato "
        sSQL = sSQL & "FROM Contratos_Meses,Clientes "
        sSQL = sSQL & "WHERE Contratos_Meses.Codigo_C = Clientes.Codigo "
        sSQL = sSQL & "AND Fecha <= #" & Mifecha & "# "
        sSQL = sSQL & "AND T = '" & Pendiente & "' "
        sSQL = sSQL & "ORDER BY Contrato_No,Fecha "
        SelectDBCombo DBCContrato, DataContrato, sSQL, "Contrato", False
     End If
     With DataFactura.Recordset
         .AddNew
         .Fields("T") = Pendiente
         .Fields("ME") = Moneda_US
         .Fields("Factura") = Factura_No
         .Fields("Fecha") = FechaTexto
         .Fields("Fecha_C") = FechaTexto
         .Fields("Fecha_V") = MBoxFechaV.Text
         .Fields("Codigo_C") = CodigoCli
         .Fields("Vendedor") = NombreUsuario
         .Fields("Pedido_No") = 0  'TextPedidos.Text
         .Fields("Bultos_No") = 0  'TextBultos.Text
         .Fields("Gavetas_No") = 0 'TextTransportador.Text
         .Fields("Forma_Pago") = TextoFormaPago
         .Fields("Sin_IVA") = Total_Sin_IVA
         .Fields("Con_IVA") = Total_Con_IVA
         .Fields("SubTotal") = Total_Sin_IVA + Total_Con_IVA
         .Fields("Descuento") = Total_Desc
         .Fields("IVA") = Total_IVA
         .Fields("Servicio") = Total_Servicio
         .Fields("Total_MN") = 0
         .Fields("Total_ME") = 0
         .Fields("Comision") = 0
          If Moneda_US Then
            .Fields("Total_ME") = Total_FacturaME
             Total = Total_FacturaME
          Else
            .Fields("Total_MN") = Total_Factura
             Total = Total_Factura
          End If
         .Fields("Saldo_MN") = Total_Factura
         .Fields("Saldo_ME") = Total_FacturaME
         .Fields("Cod_Ejec") = Ninguno
         .Fields("Porc_C") = 0
         .Fields("Cotizacion") = Dolar
         .Fields("Observacion") = TextObs.Text
         .Fields("Nota") = TextNota.Text
         .Fields("Cta_CxC") = Cta_Cobrar
         .Fields("Cta_Venta") = Cta_Ventas
         .Fields("Contrato_No") = TextoProc
         .Fields("Nivel") = Ninguno
         .Update
     End With
     DataProductos.Recordset.MoveFirst
     Do While Not DataProductos.Recordset.EOF
        With DataDetalle.Recordset
            .AddNew
            .Fields("T") = Pendiente
            .Fields("Factura_No") = Factura_No
            .Fields("Codigo_C") = CodigoCli
            .Fields("Fecha") = FechaTexto
            .Fields("Codigo") = DataProductos.Recordset.Fields("CODIGO")
            .Fields("Cantidad") = DataProductos.Recordset.Fields("CANT")
            .Fields("CodigoL") = DataProductos.Recordset.Fields("CODIGO_L")
            .Fields("Reposicion") = DataProductos.Recordset.Fields("REP")
            .Fields("Precio") = DataProductos.Recordset.Fields("PRECIO")
            .Fields("Total") = DataProductos.Recordset.Fields("TOTAL")
            .Fields("Total_Desc") = DataProductos.Recordset.Fields("Total_Desc")
            .Fields("Total_IVA") = DataProductos.Recordset.Fields("Total_IVA")
            .Fields("Producto") = DataProductos.Recordset.Fields("PRODUCTO")
            .Fields("Cod_Ejec") = DataProductos.Recordset.Fields("Cod_Ejec")
            .Fields("Porc_C") = DataProductos.Recordset.Fields("Porc_C")
            .Update
        End With
        DataProductos.Recordset.MoveNext
     Loop
    'Grabamos el numero de factura
     sSQL = "DELETE * FROM Detalles_" & CodigoUsuario & " "
     DeleteData DataProductos, sSQL
     sSQL = "SELECT * FROM Detalles_" & CodigoUsuario & " "
     SelectDBGrid DBGDetalle, DataProductos, sSQL
     RatonNormal
     Mensajes = "Pago al Contado"
     Titulo = "Formulario de Pago"
     TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If J = 6 Then
        Cadena = DBCmbCliente.Text
        FPagoContado.MBoxFecha.Text = MBoxFecha.Text
        FPagoContado.Show 1
     End If
     If CheqFormato.Value = 1 Then
        ImprimirFacturasCxC Facturas, DataFactura, DataDetalle, Factura_No, Factura_No
     Else
        ImprimirFacturas Factura_No, DataFactura, DataDetalle
     End If
  Else
     MsgBox "No se puede grabar la Factura," & Chr(13) & "falta datos."
  End If
End Sub

