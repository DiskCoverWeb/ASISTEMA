VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FEgreso 
   Caption         =   "COMPROBANTE DE EGRESO"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   11580
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
      TabIndex        =   20
      Top             =   3150
      Width           =   9570
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   6300
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
      Left            =   1470
      MaxLength       =   35
      TabIndex        =   6
      Top             =   735
      Width           =   4740
   End
   Begin VB.Frame FrameAsigna 
      Height          =   1170
      Left            =   9240
      TabIndex        =   27
      Top             =   3780
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
         TabIndex        =   28
         Text            =   "FEgreso.frx":0000
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
         TabIndex        =   29
         Text            =   "FEgreso.frx":0002
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label18 
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
         TabIndex        =   46
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label11 
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   210
         Width           =   1485
      End
   End
   Begin MSDBCtls.DBList DBLCuentas 
      Bindings        =   "FEgreso.frx":0004
      DataSource      =   "DataCuentas"
      Height          =   2400
      Left            =   1890
      TabIndex        =   26
      Top             =   4305
      Visible         =   0   'False
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   4233
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
   Begin VB.CommandButton CmdCancelar 
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
      Height          =   330
      Left            =   1575
      Picture         =   "FEgreso.frx":001A
      TabIndex        =   33
      Top             =   7035
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid DBGAsientos 
      Bindings        =   "FEgreso.frx":045C
      Height          =   2535
      Left            =   105
      OleObjectBlob   =   "FEgreso.frx":0473
      TabIndex        =   31
      Top             =   4410
      Width           =   11355
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
      Top             =   6405
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8925
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   8295
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   735
      Width           =   1485
   End
   Begin VB.Data DataSQLDetCli 
      Caption         =   "SQLDetCli"
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
      Top             =   5460
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataDetCli 
      Caption         =   "DetCli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataSubCtaDet 
      Caption         =   "SubCtaDet"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataSubCta 
      Caption         =   "SubCta"
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
      Width           =   2430
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
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   3990
      Width           =   1905
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
      TabIndex        =   25
      Top             =   3990
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
      TabIndex        =   24
      Text            =   "0"
      Top             =   3990
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   735
      Width           =   1275
      _ExtentX        =   2249
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
      Height          =   2010
      Left            =   105
      TabIndex        =   11
      Top             =   1050
      Width           =   11355
      Begin VB.TextBox TextCheque 
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
         Left            =   9765
         MaxLength       =   8
         TabIndex        =   18
         Top             =   945
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSDBGrid.DBGrid DBGBanco 
         Bindings        =   "FEgreso.frx":0E29
         Height          =   540
         Left            =   1365
         OleObjectBlob   =   "FEgreso.frx":0E43
         TabIndex        =   39
         Top             =   1365
         Width           =   9885
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
         Height          =   330
         Left            =   9765
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   630
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSDBCtls.DBCombo DBCBanco 
         Bindings        =   "FEgreso.frx":17F6
         DataSource      =   "DataBanco"
         Height          =   315
         Left            =   1365
         TabIndex        =   16
         Top             =   945
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   556
         _Version        =   327680
         Text            =   "DBCombo1"
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
      Begin VB.CheckBox CheckEfect 
         Caption         =   "Efectivo"
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
         TabIndex        =   12
         Top             =   735
         Width           =   1065
      End
      Begin VB.CheckBox CheckBco 
         Caption         =   "Banco "
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
         TabIndex        =   15
         Top             =   1050
         Width           =   960
      End
      Begin MSDBCtls.DBCombo DBCCaja 
         Bindings        =   "FEgreso.frx":180A
         DataSource      =   "DataCaja"
         Height          =   315
         Left            =   1365
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   556
         _Version        =   327680
         Text            =   "DBCombo1"
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
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQ."
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
         Left            =   9030
         TabIndex        =   17
         Top             =   945
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQ."
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
         Left            =   9030
         TabIndex        =   40
         Top             =   630
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label LabelTotal 
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
         Left            =   9765
         TabIndex        =   38
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Total Pagado"
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
         Left            =   7875
         TabIndex        =   37
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
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
      Picture         =   "FEgreso.frx":181D
      TabIndex        =   32
      Top             =   7035
      Width           =   1380
   End
   Begin VB.Data DataComprobantes 
      Caption         =   "Comprobantes"
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
      Top             =   6090
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6090
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataBanco 
      Caption         =   "Banco"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
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
      Top             =   6090
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
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
      Width           =   2430
   End
   Begin VB.Data DataRet 
      Caption         =   "Ret"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6090
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
      Left            =   6510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataBancos 
      Caption         =   "Bancos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8925
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataDetRet 
      Caption         =   "DetRet"
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
      Top             =   5775
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label20 
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
      Left            =   6300
      TabIndex        =   7
      Top             =   525
      Width           =   1905
   End
   Begin VB.Label Label8 
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
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   7035
      Width           =   1695
   End
   Begin VB.Label Label9 
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
      Left            =   8295
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
      TabIndex        =   21
      Top             =   3780
      Width           =   1695
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
      TabIndex        =   36
      Top             =   7035
      Width           =   1170
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIGITE LA CLAVE O SELECCIONES LA CUENTA"
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
      TabIndex        =   22
      Top             =   3780
      Width           =   7260
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
      TabIndex        =   23
      Top             =   3780
      Width           =   1905
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
      TabIndex        =   35
      Top             =   7035
      Width           =   1800
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
      TabIndex        =   34
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
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
      Width           =   4215
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
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   " Egreso No."
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
      Left            =   8715
      TabIndex        =   1
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PAGADO A:"
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
      Left            =   1470
      TabIndex        =   5
      Top             =   525
      Width           =   4740
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
      TabIndex        =   19
      Top             =   3150
      Width           =   1590
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
      Left            =   10290
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "FEgreso"
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
  DBLCuentas.Visible = False
  Cadena = ObtenerPalabra(DBLCuentas.Text, 4)
  LeerCta DataCtas, Cadena
  TextCodigo.Text = Codigo
  TextCuenta.Text = Cuenta
  FrameAsigna.Visible = True
  If OpcCoop Then
     Label11.Visible = False
     Label16.Visible = False
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
     If SubCta = "R" Then
        FechaTexto = MBoxFecha.Text
        Cta_Ret_Egreso = Codigo
        Nombre_Cta_Ret = Cuenta
        Retencion1.Show 1
     End If
     TextValor.Visible = True
     TextValor.SetFocus
  Else
     TextOpcTM.SetFocus
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

Private Sub DBCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBCBanco_LostFocus()
  LeerBanco DataCtas, DBCBanco.Text
  DBCBanco.Text = Codigo & Space(5) & NombreBanco
End Sub

Private Sub DBCCaja_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBCCaja_LostFocus()
  LeerBanco DataCtas, DBCCaja.Text
  If Moneda_US Then Label6.Caption = "M/E" Else Label6.Caption = Moneda
End Sub

Private Sub DBGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Deactivate()
  FEgreso.WindowState = 1
End Sub

Private Sub MBoxFecha_GotFocus()
   MBoxFecha.Text = FechaSistema
   MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, True
End Sub

Private Sub TextCantidad_GotFocus()
   TextCantidad.Text = ""
   MarcarTexto TextCantidad
End Sub

Private Sub TextCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheque_GotFocus()
   TextCheque.Text = ""
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheque_LostFocus()
  TextoValido TextCheque, False
  IniciarAsientosBancoDe CompEgreso, DataAsientBanco, DBGBanco
  InsertarAsientoBanco DataAsientBanco, TextCheque.Text, 0
  TextConcepto.SetFocus
End Sub

Private Sub TextConcepto_Change()
  If TextoMaximo(TextConcepto) Then TextCodigo.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
  Monto_Total = SumaBancos + Abono
  LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_LostFocus()
  TextoValido TextCotiza, True
  If Val(TextCotiza.Text) <= 0 Then TextCotiza.Text = Dolar
  Dolar = Val(TextCotiza.Text)
End Sub

Private Sub TextValor_GotFocus()
  If SubCta = "R" And OpcDH = 2 Then
     TextValor.Text = Total_DetRet
     ValorDH = CDbl(TextValor.Text)
     'If Moneda_US Or OpcTM = 2 Then ValorDH = Round(ValorDH * Dolar)
     InsertarAsiento DataAsientos
     SiguienteControl
  Else
     TextValor.Text = ""
     Label17.Caption = "VALOR M/N"
     If Moneda_US Or OpcTM = 2 Then Label17.Caption = "VALOR M/E"
  End If
End Sub

Private Sub TextValor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValor_LostFocus()
  If SubCta = "R" And OpcDH = 2 Then
     'TextValor.Text = Total_DetRet
     'ValorDH = Val(TextValor.Text)
     'If Moneda_US Or OpcTM = 2 Then ValorDH = Round(ValorDH / Dolar)
     'InsertarAsiento DataAsientos
     'SiguienteControl
  Else
     TextoValido TextValor, True
     TotalSubCta = 0: SubCtaGen = Ninguno
     ValorDH = Val(TextValor.Text)
     InsertarAsiento DataAsientos
     TotalSubCta = ValorDH
     Select Case SubCta
       Case "C", "P", "G", "I"
            FechaTexto = MBoxFecha.Text
            SubCtaGen = Codigo
            FSubCtas.Show 1
     End Select
     TextCuenta.Text = ""
  End If
  TextCodigo.SetFocus
End Sub

Private Sub DBGAsientos_BeforeDelete(Cancel As Integer)
  Codigo = DataAsientos.Recordset.Fields("CODIGO")
  Cancel = DeleteSiNo(DataAsientos)
  If Cancel = False Then EliminarSubCta DataSubCtaDet, Codigo
End Sub

Private Sub DBGAsientos_GotFocus()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub DBGAsientos_LostFocus()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub CmdCancelar_Click()
  Unload FEgreso
End Sub

Private Sub CmdGrabar_Click()
  OpcSubCtaDH = 1
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
          If NumCompEgr = 0 Then
             NumComp = ReadSetDataNum("Egresos", True, True)
          Else
             NumComp = NumCompEgr
          End If
          FechaTexto = MBoxFecha.Text
          sSQL = "SELECT * FROM Asiento_Bancos_E_" & CodigoUsuario & " "
          SelectData DataAsientBanco, sSQL, False
          sSQL = "SELECT * FROM Asientos_R_" & CodigoUsuario & " "
          sSQL = sSQL & "ORDER BY CTA "
          SelectData DataDetRet, sSQL, False
          sSQL = "SELECT * FROM Asientos_SC_E_" & CodigoUsuario & " "
          'sSQL = sSQL & "WHERE Valor <> 0 "
          SelectData DataSubCtaDet, sSQL, False
         'Grabacion del Comp
          Co.T = Normal
          Co.TP = CompEgreso
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Monto_Total = Monto_Total
          Co.Concepto = TextConcepto.Text
          Co.Beneficiario = TextBenef.Text
          Co.RUC_CI = MBoxRUC.Text
          Co.Efectivo = Abono
          GrabarComprobantes Co, DataAsientos, DataSubCtaDet, DataAsientBanco, DataDetRet
         'Seteamos para el siguiente comprobante
          sSQL = "DELETE * FROM Asiento_Bancos_E_" & CodigoUsuario & " "
          UpdateData DataAsientBanco, sSQL
          sSQL = "DELETE * FROM Asientos_R_" & CodigoUsuario & " "
          UpdateData DataDetRet, sSQL
          sSQL = "DELETE * FROM Asientos_SC_E_" & CodigoUsuario & " "
          UpdateData DataSubCtaDet, sSQL
          IniciarAsientosDe CompEgreso, DataAsientos, DBGAsientos, DataAsientBanco, DBGBanco, DataDetRet
          DBGBanco.Visible = False
          RatonNormal
          ImprimirComprobantesDe False, CompEgreso, NumComp, NumEmpresa, DataComprobantes, DataTransacciones, DataBancos, DataRet
          NumComp = NumComp + 1
          LabelComp.Caption = Format(NumComp, "000000")
          If NumCompEgr <> 0 Then
             Unload FEgreso
             Exit Sub
          Else
             MBoxFecha.SetFocus
          End If
       Else
          MsgBox "Warning: Falta de Ingresar datos."
          TextCodigo.SetFocus
       End If
     Else
       TextCodigo.SetFocus
     End If
  End If
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
         TextValor.Visible = True
         If CheckEfect.Value Then
            OpcDH = 2: ValorDH = Abono
            Cadena = SinEspaciosIzq(DBCCaja.Text)
            LeerCta DataCtas, Cadena
            If Moneda_US Then ValorDH = Round(ValorDH * Dolar)
            InsertarAsientosC DataAsientos
         End If
         'CalculosTotalRetencion DataDetRet
         'If CheckBco.Value And Total_Ret <> 0 Then
         '   OpcDH = 2: ValorDH = Total_Ret
         '   LeerCta DataCtas, Cta_Ret_Egreso
         '   InsertarAsientosC DataAsientos
         'End If
         SumaBancos = 0
         CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
         If CheckBco.Value Then
            SumaBancos = SumaDebe - SumaHaber
            OpcDH = 2: ValorDH = SumaDebe - SumaHaber
            Cadena = SinEspaciosIzq(DBCBanco.Text)
            LeerCta DataCtas, Cadena
            With DataAsientBanco.Recordset
             If .RecordCount > 0 Then
                .MoveFirst
                .Edit
                .Fields("Valor") = SumaBancos
                .Update
             End If
            End With
            If OpcCoop And Moneda_US Then ValorDH = ValorDH * Dolar
            InsertarAsientosC DataAsientos
         End If
         Monto_Total = SumaBancos + Abono
         LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
         CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
         TextCodigo.Text = "-1"
         CmdGrabar.SetFocus
    Case vbKeyF2
         TextCodigo.Text = "-1"
         FIngCtas.Show
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub TextCodigo_LostFocus()
  If TextCodigo.Text = "" Then TextCodigo.Text = "0"
  LeerCodigoCta TextCodigo, TextCuenta, TextValor, DBLCuentas, DataCuentas, FrameAsigna, TextOpcTM, TextOpcDH
End Sub

Private Sub Form_Activate()
  FEgreso.WindowState = 2
  TipoDoc = CompEgreso
  CTAsientos_R
  CTAsiento_Bancos TipoDoc
  CTAsientos TipoDoc
  CTAsientos_SC TipoDoc
  IniciarAsientosDe TipoDoc, DataAsientos, DBGAsientos, DataAsientBanco, DBGBanco, DataDetRet
  IniciarAsientoSC_De TipoDoc, DataSubCtaDet
  SelectCuentas DBLCuentas, DataCuentas
  SubCtaCompCli = Cta_General
  SubCtaGen1 = Ninguno
  Una_Vez = True
  TextCotiza.Text = Dolar
  CtaCteNo = Ninguno
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE TC = 'BA' AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DBCBanco, DataBanco, sSQL, "NomCuenta", False
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCaja "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE TC = 'CJ' AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DBCCaja, DataCaja, sSQL, "NomCaja", False
  NumComp = ReadSetDataNum("Egresos", True, False)
  If NumCompEgr <> 0 Then
     NumComp = NumCompEgr
     NumEmpresa = NumItem
  Else
     NumEmpresa = NumItemTemp
  End If
  Label5.Caption = Empresa
  LabelComp.Caption = Format(NumComp, "000000")
  DBGBanco.Visible = False
  MBoxFecha.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  'Abriendo bases relacionadas
  DataSubCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSubCtaDet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataDetCli.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSQLDetCli.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataDetRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataBancos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataAsientBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataCaja.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextCantidad_LostFocus()
  TextoValido TextCantidad, True
  Abono = Val(TextCantidad.Text)
  TextCantidad.Text = Format(Abono, "#,##0.00")
  Monto_Total = Abono
  LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
  CheckBco.SetFocus
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
End Sub

Private Sub CheckBco_Click()
  If CheckBco.Value Then
     DBCBanco.Visible = True
     DBGBanco.Visible = True
     TextCheque.Visible = True
     Label24.Visible = True
     DBCBanco.SetFocus
  Else
     DBCBanco.Visible = False
     DBGBanco.Visible = False
     Label24.Visible = False
     TextCheque.Visible = False
  End If
End Sub

Private Sub CheckEfect_Click()
  Label6.Caption = Moneda
  If CheckEfect.Value Then
     TextCantidad.Visible = True
     DBCCaja.Visible = True
     DBCCaja.SetFocus
     Label6.Visible = True
  Else
     TextCantidad.Text = "0.00"
     Abono = 0
     TextCantidad.Visible = False
     DBCCaja.Visible = False
     Label6.Visible = True
  End If
End Sub

