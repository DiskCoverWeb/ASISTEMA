VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FIngresos 
   Caption         =   "COMPROBANTE DE INGRESO"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextConcepto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1890
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   3465
      Width           =   9570
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   6405
      TabIndex        =   8
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   735
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#########-#-###"
      Mask            =   "#########-#-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   735
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
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
   Begin VB.TextBox TextBenef 
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
      Left            =   1575
      MaxLength       =   35
      TabIndex        =   6
      Top             =   735
      Width           =   4740
   End
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4830
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   2010
   End
   Begin MSDBCtls.DBList DBLCuentas 
      Bindings        =   "Ingresos.frx":0000
      DataSource      =   "DataCuentas"
      Height          =   1815
      Left            =   1890
      TabIndex        =   28
      Top             =   4620
      Visible         =   0   'False
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   3201
      _Version        =   327680
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
   Begin VB.Frame FrameAsigna 
      Height          =   1170
      Left            =   9240
      TabIndex        =   32
      Top             =   3990
      Visible         =   0   'False
      Width           =   2115
      Begin VB.TextBox TextOpcTM 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "Ingresos.frx":0016
         Top             =   210
         Width           =   330
      End
      Begin VB.TextBox TextOpcDH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "Ingresos.frx":0018
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label20 
         Caption         =   " Haber    2"
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
         TabIndex        =   48
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   " 2.- M/E"
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
         Left            =   735
         TabIndex        =   47
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   " Debe     1"
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
         TabIndex        =   46
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Valores: 1.- M/N"
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
         TabIndex        =   45
         Top             =   210
         Width           =   1485
      End
   End
   Begin MSDBGrid.DBGrid DBGAsientos 
      Bindings        =   "Ingresos.frx":001A
      Height          =   2220
      Left            =   105
      OleObjectBlob   =   "Ingresos.frx":0031
      TabIndex        =   37
      Top             =   4725
      Width           =   11355
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2415
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Data DataCaja 
      Caption         =   "Caja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox TextCotiza 
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
      Height          =   330
      Left            =   8400
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   735
      Width           =   1485
   End
   Begin VB.Data DataSubCtaDet 
      Caption         =   "SubCtaDet"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4095
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataDetCli 
      Caption         =   "DetCli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataSQLDetCli 
      Caption         =   "SQLDetCli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataSubCta 
      Caption         =   "SubCta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6405
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6510
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox TextCuenta 
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
      Left            =   1890
      TabIndex        =   27
      Top             =   4305
      Width           =   7260
   End
   Begin VB.TextBox TextCodigo 
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
      TabIndex        =   26
      Text            =   "0"
      Top             =   4305
      Width           =   1695
   End
   Begin VB.TextBox TextValor 
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
      Height          =   330
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "Ingresos.frx":09E7
      Top             =   4305
      Width           =   2010
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   105
      TabIndex        =   36
      Top             =   1050
      Width           =   11355
      Begin MSDBGrid.DBGrid DBGBanco 
         Bindings        =   "Ingresos.frx":09EB
         Height          =   855
         Left            =   1365
         OleObjectBlob   =   "Ingresos.frx":0A05
         TabIndex        =   21
         Top             =   1365
         Visible         =   0   'False
         Width           =   9885
      End
      Begin VB.TextBox TextDeposito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "0"
         Top             =   945
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox TextMonto 
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
         Height          =   345
         Left            =   7875
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "Ingresos.frx":13B8
         Top             =   945
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSDBCtls.DBCombo DBCBanco 
         Bindings        =   "Ingresos.frx":13BA
         DataSource      =   "DataBanco"
         Height          =   315
         Left            =   1365
         TabIndex        =   18
         Top             =   945
         Visible         =   0   'False
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   556
         _Version        =   327680
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
      Begin VB.TextBox TextCantidad 
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
         Height          =   345
         Left            =   7875
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1485
      End
      Begin VB.CheckBox CheckBco 
         Caption         =   "&Bancos"
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
         TabIndex        =   14
         Top             =   1050
         Width           =   1065
      End
      Begin VB.CheckBox CheckEfect 
         Caption         =   "&Efectivo"
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
         TabIndex        =   11
         Top             =   735
         Width           =   1065
      End
      Begin MSDBCtls.DBCombo DBCCaja 
         Bindings        =   "Ingresos.frx":13CE
         DataSource      =   "DataCaja"
         Height          =   315
         Left            =   1365
         TabIndex        =   12
         Top             =   630
         Visible         =   0   'False
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   556
         _Version        =   327680
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEP."
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
         Left            =   9450
         TabIndex        =   17
         Top             =   945
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Left            =   7245
         TabIndex        =   15
         Top             =   945
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LabelMontoTotal 
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
         Height          =   330
         Left            =   9450
         TabIndex        =   42
         Top             =   210
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Valor Total Recibido"
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
         Left            =   7455
         TabIndex        =   41
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SM"
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
         Left            =   7245
         TabIndex        =   16
         Top             =   630
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataComprobantes 
      Caption         =   "Comprobantes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6510
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6510
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1575
      TabIndex        =   35
      Top             =   7035
      Width           =   1380
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
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
      TabIndex        =   34
      Top             =   7035
      Width           =   1380
   End
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6510
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataAsientBanco 
      Caption         =   "AsientBanco"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6405
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Data DataBancos 
      Caption         =   "Bancos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataBanco 
      Caption         =   "Banco"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " R.U.C./C.I."
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
      Left            =   6405
      TabIndex        =   7
      Top             =   525
      Width           =   1905
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia"
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
      Left            =   3045
      TabIndex        =   44
      Top             =   7035
      Width           =   1275
   End
   Begin VB.Label LabelDiferencia 
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
      Left            =   4410
      TabIndex        =   43
      Top             =   7035
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COTIZACION:"
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
      Left            =   8400
      TabIndex        =   9
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      TabIndex        =   24
      Top             =   4095
      Width           =   1695
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RECIBI DE:"
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
      Left            =   1575
      TabIndex        =   5
      Top             =   525
      Width           =   4740
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   4110
   End
   Begin VB.Label LabelHaber 
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
      Left            =   9345
      TabIndex        =   40
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label LabelDebe 
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
      Left            =   7455
      TabIndex        =   39
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      Left            =   6195
      TabIndex        =   38
      Top             =   7035
      Width           =   1170
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
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
      TabIndex        =   33
      Top             =   4095
      Width           =   2010
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIGITE CLAVE O BUSQUE LA CUENTA CONTABLE"
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
      Left            =   1890
      TabIndex        =   25
      Top             =   4095
      Width           =   7260
   End
   Begin VB.Label LabelComp 
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
      Height          =   330
      Left            =   9870
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label7 
      Caption         =   " Por concepto de:"
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
      TabIndex        =   22
      Top             =   3465
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   " Ingreso No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      TabIndex        =   1
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Width           =   1380
   End
End
Attribute VB_Name = "FIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBLCuentas_DblClick()
  SiguienteControl
End Sub

Private Sub DBLCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLCuentas_LostFocus()
  Cadena = ObtenerPalabra(DBLCuentas.Text, 4)
  DBLCuentas.Visible = False
  LeerCta DataCtas, Cadena
  TextCodigo.Text = Codigo
  TextCuenta.Text = Cuenta
  FrameAsigna.Visible = True
  If OpcCoop Then
     Label16.Visible = False
     Label18.Visible = False
     TextOpcTM.Visible = False
     TextOpcDH.SetFocus
  Else
     TextOpcTM.SetFocus
  End If
End Sub

Private Sub TextBenef_GotFocus()
   TextBenef.Text = ""
End Sub

Private Sub TextBenef_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextBenef_LostFocus()
   TextoValido TextBenef
End Sub

Private Sub TextOpcDH_Change()
  If 1 > Val(TextOpcDH.Text) Or Val(TextOpcDH.Text) > 2 Then
     TextOpcDH.Text = ""
  Else
     OpcDH = Val(TextOpcDH.Text)
     SiguienteControl
  End If
End Sub

Private Sub TextOpcDH_GotFocus()
  TextOpcDH.Text = ""
End Sub

Private Sub TextOpcDH_LostFocus()
  If OpcCoop Then
     If Moneda_US Then OpcTM = 2 Else OpcTM = 1
  End If
  OpcDH = Val(TextOpcDH.Text)
  If OpcTM >= 1 And OpcDH >= 1 Then
     FrameAsigna.Visible = False
     TextValor.Visible = True
     TextValor.SetFocus
  Else
     If OpcCoop = False Then TextOpcTM.SetFocus Else TextOpcDH.SetFocus
  End If
End Sub

Private Sub TextOpcTM_Change()
  If 1 > Val(TextOpcTM.Text) Or Val(TextOpcTM.Text) > 2 Then
     TextOpcTM.Text = ""
  Else
     TextOpcDH.SetFocus
  End If
End Sub

Private Sub TextOpcTM_GotFocus()
  TextOpcTM.Text = ""
  TextOpcDH.Text = ""
End Sub

Private Sub TextOpcTM_LostFocus()
  OpcTM = Val(TextOpcTM.Text)
End Sub

Private Sub CheckBco_Click()
   Label9.Caption = Moneda
   sSQL = "DELETE * FROM Asiento_Bancos_I_" & CodigoUsuario & " "
   DeleteData DataAsientBanco, sSQL
   If CheckBco.Value Then
      TextMonto.Visible = True
      DBCBanco.Visible = True
      DBGBanco.Visible = True
      TextDeposito.Visible = True
      Label6.Visible = True
      Label9.Visible = True
     'TextMonto.SetFocus
      DBCBanco.SetFocus
   Else
      DBCBanco.Visible = False
      DBGBanco.Visible = False
      TextMonto.Visible = False
      TextDeposito.Visible = False
      Label6.Visible = False
      Label9.Visible = False
      'LimpiarAsientosDe FIngresos, CompIngreso, DataAsientos, DBGAsientos
   End If
End Sub

Private Sub CheckEfect_Click()
   Label4.Caption = Moneda
   If CheckEfect.Value Then
     TextCantidad.Visible = True
     DBCCaja.Visible = True
     DBCCaja.SetFocus
     Label4.Visible = True
     'TextCantidad.SetFocus
   Else
     TextCantidad.Visible = False
     DBCCaja.Visible = False
     TextCantidad.Text = "0.00"
     Abono = 0
     Label4.Visible = False
   End If
End Sub

Private Sub CmdCancelar_Click()
   Unload FIngresos
End Sub

Private Sub CmdGrabar_Click()
  SumaBancos = CalculosTotalBancos(DataAsientBanco)
  Monto_Total = SumaBancos + Abono
  LabelMontoTotal.Caption = Format(Monto_Total, "#,##0.00")
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
  If Round_ME(SumaDebe) <> Round_ME(SumaHaber) Then
    Mensajes = "Las transacciones no cuadran correctamente" & Chr(13)
    Mensajes = Mensajes & "corrija los resultados de las cuentas"
    MsgBox Mensajes
    TextCodigo.SetFocus
  Else
    Mensajes = "Esta seguro de Grabar el Comprobante No. " & LabelComp.Caption & "]"
    Titulo = "Pregunta de grabación"
    TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
    If J = 6 Then
       If DataAsientos.Recordset.RecordCount > 0 Then
          RatonReloj
          If NumCompIng = 0 Then
             NumComp = ReadSetDataNum("Ingresos", True, True)
          Else
             NumComp = NumCompIng
          End If
          FechaTexto = MBoxFecha.Text
          sSQL = "SELECT * FROM Asiento_Bancos_I_" & CodigoUsuario & " "
          SelectData DataAsientBanco, sSQL, False
          sSQL = "SELECT * FROM Asientos_SC_I_" & CodigoUsuario & " "
          'sSQL = sSQL & "WHERE Valor <> 0 "
          SelectData DataSubCtaDet, sSQL, False
          Co.T = Normal
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.TP = CompIngreso
          Co.Concepto = TextConcepto.Text
          Co.Beneficiario = TextBenef.Text
          Co.RUC_CI = MBoxRUC.Text
          Co.Efectivo = Abono
          Co.Monto_Total = CDbl(LabelMontoTotal.Caption)
          GrabarComprobantes Co, DataAsientos, DataSubCtaDet, DataAsientBanco
          sSQL = "DELETE * FROM Asiento_Bancos_I_" & CodigoUsuario & " "
          UpdateData DataAsientBanco, sSQL
          sSQL = "DELETE * FROM Asientos_SC_I_" & CodigoUsuario & " "
          UpdateData DataSubCtaDet, sSQL
          IniciarAsientosDe CompIngreso, DataAsientos, DBGAsientos, DataAsientBanco, DBGBanco
          DBCBanco.Visible = False
          DBCCaja.Visible = False
          TextDeposito.Visible = False
          TextCantidad.Visible = False
          TextCantidad.Text = "0.00"
          Abono = 0
          CheckEfect.Value = 0
          CheckBco.Value = 0
          RatonNormal
          ImprimirComprobantesDe False, CompIngreso, NumComp, NumEmpresa, DataComprobantes, DataTransacciones, DataBancos
          NumComp = NumComp + 1
          LabelComp.Caption = Format(NumComp, "000000")
          If NumCompIng <> 0 Then
             Unload FIngresos
             Exit Sub
          Else
             MBoxFecha.SetFocus
          End If
       Else
          MsgBox "Warning: Falta de Ingresar datos."
          TextCodigo.SetFocus
       End If
       MBoxFecha.SetFocus
    Else
       MBoxFecha.SetFocus
    End If
  End If
End Sub

Private Sub DBCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
         If CheckBco.Value Then
            With DataAsientBanco.Recordset
             If .RecordCount > 0 Then
                .MoveFirst
                 Do While Not .EOF
                    Cadena = .Fields("CTA_BANCO")
                    Valor = .Fields("VALOR")
                    LeerCta DataCtas, Cadena
                    OpcDH = 1: ValorDH = Valor
                    InsertarAsientosC DataAsientos
                   .MoveNext
                 Loop
             End If
            End With
         End If
         SumaBancos = CalculosTotalBancos(DataAsientBanco)
         Monto_Total = SumaBancos + Abono
         LabelMontoTotal.Caption = Format(Monto_Total, "#,##0.00")
         TextConcepto.SetFocus
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub DBCBanco_LostFocus()
  LeerBanco DataCtas, DBCBanco.Text
  Label17.Caption = "VALOR M/N"
  If Moneda_US Then
     Label9.Caption = "M/E"
     Label17.Caption = "PARCIAL M/E"
  Else
     Label9.Caption = Moneda
  End If
End Sub

Private Sub DBCCaja_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBCCaja_LostFocus()
  TextCantidad.SetFocus
  LeerBanco DataCtas, DBCCaja.Text
  If Moneda_US Then Label4.Caption = "M/E" Else Label4.Caption = Moneda
End Sub

Private Sub DBGAsientos_BeforeDelete(Cancel As Integer)
  Codigo = DataAsientos.Recordset.Fields("CODIGO")
  Cancel = DeleteSiNo(DataAsientos)
  If Cancel = False Then EliminarSubCta DataSubCtaDet, Codigo
End Sub

Private Sub DBGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then DBCCtas.SetFocus
End Sub

Private Sub DBGAsientos_LostFocus()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub DBGBanco_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataAsientBanco)
End Sub

Private Sub Form_Deactivate()
 FIngresos.WindowState = 1
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_GotFocus()
  TextCotiza.Text = Dolar
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_LostFocus()
  TextoValido TextCotiza, True
  If Val(TextCotiza.Text) <= 0 Then TextCotiza.Text = Dolar
  Dolar = Val(TextCotiza.Text)
End Sub

Private Sub Form_Activate()
  FIngresos.WindowState = 2
  TipoDoc = CompIngreso
  CTAsiento_Bancos TipoDoc
  CTAsientos TipoDoc
  CTAsientos_SC TipoDoc
  IniciarAsientosDe TipoDoc, DataAsientos, DBGAsientos, DataAsientBanco, DBGBanco
  IniciarAsientoSC_De TipoDoc, DataSubCtaDet
  SelectCuentas DBLCuentas, DataCuentas
  TotalSubCta = 0: SubCtaGen = Ninguno
  Una_Vez = True
  CompCheque = True
  NumComp = ReadSetDataNum("Ingresos", True, False)
  If NumCompIng <> 0 Then
     NumComp = NumCompIng
     NumEmpresa = NumItem
  Else
     NumEmpresa = NumItemTemp
  End If
  LabelComp.Caption = Format(NumComp, "000000")
  Label12.Caption = Empresa
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCaja "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE TC = 'CJ' AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DBCCaja, DataCaja, sSQL, "NomCaja", False
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta AS NomCuenta "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE TC = 'BA' AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DBCBanco, DataBanco, sSQL, "NomCuenta", False
  TextMonto.Visible = False
  DBCBanco.Visible = False
  DBGBanco.Visible = False
  TextDeposito.Visible = False
  TextCantidad.Visible = False
  MBoxFecha.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  'Abriendo bases relacionadas
   DataSubCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataDetCli.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSQLDetCli.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSubCtaDet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataBancos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCaja.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
   MBoxFecha.Text = FechaSistema
   MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, True
End Sub

Private Sub TextCantidad_GotFocus()
   MarcarTexto TextCantidad
End Sub

Private Sub TextCantidad_LostFocus()
  TextoValido TextCantidad, True
  Abono = Val(TextCantidad.Text)
  TextCantidad.Text = Format(Abono, "#,##0.00")
  If CheckEfect.Value Then
     OpcDH = 1: ValorDH = Abono
     Cadena = SinEspaciosIzq(DBCCaja.Text)
     LeerCta DataCtas, Cadena
     If Moneda_US Then ValorDH = Round(ValorDH * Dolar)
     InsertarAsientosC DataAsientos
  End If
  CheckBco.SetFocus
End Sub

Private Sub TextCodigo_GotFocus()
  TextCodigo.Text = ""
  TextValor.Visible = False
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
  If Cta_General <> Ninguno And Una_Vez Then
     LeerCta DataCtas, Cta_General
     TextCodigo.Text = Cta_General
     Una_Vez = False
  End If
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         TextCodigo.Text = "-1"
         TextValor.Visible = True
         CmdGrabar.SetFocus
    Case vbKeyF2
         TextCodigo.Text = "-1"
         FIngCtas.Show
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, True
  LeerCodigoCta TextCodigo, TextCuenta, TextValor, DBLCuentas, DataCuentas, FrameAsigna, TextOpcTM, TextOpcDH
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
  TextCodigo.SetFocus
End Sub

Private Sub TextDeposito_GotFocus()
  TextDeposito.Text = ""
End Sub

Private Sub TextDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDeposito_LostFocus()
  TextoValido TextDeposito, False
  Valor = Val(TextMonto.Text)
  If Valor <> 0 Then InsertarAsientoBanco DataAsientBanco, TextDeposito.Text, Valor
  DBCBanco.SetFocus
End Sub

Private Sub TextMonto_GotFocus()
  MarcarTexto TextMonto
End Sub

Private Sub TextMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMonto_LostFocus()
  TextoValido TextMonto, True
End Sub

Private Sub TextValor_GotFocus()
   TextValor.Text = ""
   Label17.Caption = "VALOR M/N"
   If Moneda_US Or OpcTM = 2 Then Label17.Caption = "VALOR M/E"
End Sub

Private Sub TextValor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValor_LostFocus()
  TextoValido TextValor, True
  ValorDH = Val(TextValor.Text)
  InsertarAsiento DataAsientos
  TotalSubCta = ValorDH
  Select Case SubCta
    Case "C", "P", "G", "I"
         FechaTexto = MBoxFecha.Text
         SubCtaGen = Codigo
         FSubCtas.Show 1
  End Select
  RatonNormal
  TextCuenta.Text = ""
  TextCodigo.SetFocus
End Sub

