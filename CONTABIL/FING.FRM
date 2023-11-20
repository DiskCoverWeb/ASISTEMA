VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMPROBANTE DE INGRESO"
   ClientHeight    =   7050
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
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
      Left            =   1785
      TabIndex        =   24
      Top             =   4200
      Width           =   5685
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
      TabIndex        =   23
      Text            =   "0"
      Top             =   4200
      Width           =   1590
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
      Left            =   9030
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "Fing.frx":0000
      Top             =   4200
      Width           =   1590
   End
   Begin VB.TextBox TextOpcDH 
      Alignment       =   2  'Center
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
      Left            =   8715
      MaxLength       =   1
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "Fing.frx":0004
      Top             =   4200
      Width           =   330
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   8820
      TabIndex        =   5
      Top             =   630
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   2220
      Left            =   105
      TabIndex        =   30
      Top             =   1050
      Width           =   10515
      Begin VB.TextBox TextCtaCte 
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
         Left            =   7980
         MaxLength       =   15
         TabIndex        =   16
         Top             =   735
         Width           =   1275
      End
      Begin VB.TextBox TextBanco 
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
         Left            =   4725
         MaxLength       =   20
         TabIndex        =   15
         Top             =   735
         Width           =   3165
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
         Left            =   1680
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   210
         Width           =   1380
      End
      Begin MSDBGrid.DBGrid DBGBanco 
         Bindings        =   "Fing.frx":0006
         Height          =   960
         Left            =   3150
         OleObjectBlob   =   "Fing.frx":0020
         TabIndex        =   18
         Top             =   1155
         Width           =   7260
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
         Left            =   3150
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Fing.frx":09D3
         Top             =   735
         Width           =   1485
      End
      Begin VB.CheckBox CheckBco 
         Caption         =   "&Cheques  S/."
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
         TabIndex        =   8
         Top             =   840
         Width           =   1485
      End
      Begin VB.CheckBox CheckEfect 
         Caption         =   "&Efectivo   S/."
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
         TabIndex        =   6
         Top             =   315
         Width           =   1485
      End
      Begin VB.TextBox TextDeposito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9345
         MaxLength       =   8
         TabIndex        =   17
         Text            =   "0"
         Top             =   735
         Width           =   1065
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
         Left            =   1680
         TabIndex        =   39
         Top             =   1785
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "V/Recibido S/."
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
         Left            =   210
         TabIndex        =   38
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Label Label3 
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
         Left            =   1680
         TabIndex        =   9
         Top             =   735
         Width           =   1380
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
         Left            =   3150
         TabIndex        =   10
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre del Banco"
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
         Left            =   4725
         TabIndex        =   11
         Top             =   420
         Width           =   3165
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cheque #"
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
         Left            =   9345
         TabIndex        =   13
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cta. Cte. #"
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
         Left            =   7980
         TabIndex        =   12
         Top             =   420
         Width           =   1275
      End
   End
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
      Left            =   1785
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3360
      Width           =   8835
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   630
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "0"
   End
   Begin MSDBGrid.DBGrid DBGAsientos 
      Bindings        =   "Fing.frx":09D5
      Height          =   1800
      Left            =   105
      OleObjectBlob   =   "Fing.frx":09EC
      TabIndex        =   31
      Top             =   4620
      Width           =   10515
   End
   Begin VB.TextBox TextRecibi 
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
      Left            =   3465
      MaxLength       =   30
      TabIndex        =   3
      Top             =   630
      Width           =   3900
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   6510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
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
      Height          =   330
      Left            =   3990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5355
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1785
      TabIndex        =   29
      Top             =   6510
      Width           =   1380
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "GRABAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   28
      Top             =   6510
      Width           =   1485
   End
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
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
      Top             =   5670
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
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5670
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
      Top             =   5670
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label15 
      Caption         =   " 2 = Haber"
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
      Left            =   7560
      TabIndex        =   41
      Top             =   4305
      Width           =   1065
   End
   Begin VB.Label Label11 
      Caption         =   " 1 = Debe"
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
      Left            =   7560
      TabIndex        =   40
      Top             =   3990
      Width           =   1065
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
      Top             =   3990
      Width           =   1590
   End
   Begin VB.Label Label10 
      Caption         =   " R.U.C. / C.I."
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
      TabIndex        =   4
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label 
      Caption         =   " Recibi de:"
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
      Left            =   2205
      TabIndex        =   2
      Top             =   735
      Width           =   1170
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
      Height          =   435
      Left            =   210
      TabIndex        =   37
      Top             =   105
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
      Left            =   8610
      TabIndex        =   34
      Top             =   6510
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
      Left            =   6615
      TabIndex        =   33
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SUMA DE TOTALES"
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
      TabIndex        =   32
      Top             =   6510
      Width           =   2325
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
      Left            =   8715
      TabIndex        =   26
      Top             =   3990
      Width           =   1905
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
      Left            =   1785
      TabIndex        =   22
      Top             =   3990
      Width           =   5685
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
      Left            =   9345
      TabIndex        =   36
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
      TabIndex        =   19
      Top             =   3360
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
      Left            =   7665
      TabIndex        =   35
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   " Fecha:"
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
      TabIndex        =   0
      Top             =   735
      Width           =   855
   End
End
Attribute VB_Name = "FIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBco_Click()
   If CheckBco.Value Then
      TextMonto.Visible = True
      TextBanco.Visible = True
      DBGBanco.Visible = True
      TextCtaCte.Visible = True
      TextDeposito.Visible = True
      sSQL = "SELECT * FROM Asiento_Bancos "
      SelectData DataAsientBanco, sSQL, False
      DBAnchoMallas FIngresos, DBGBanco, 0, DataAsientBanco
      TextMonto.SetFocus
   Else
      sSQL = "DELETE * FROM Asiento_Bancos "
      DeleteData DataAsientBanco, sSQL
      TextBanco.Visible = False
      DBGBanco.Visible = False
      TextMonto.Visible = False
      TextCtaCte.Visible = False
      TextDeposito.Visible = False
   End If
End Sub

Private Sub CheckEfect_Click()
   If CheckEfect.Value Then
     TextCantidad.Visible = True
     TextCantidad.SetFocus
   Else
     TextCantidad.Visible = False
     TextCantidad.Text = "0"
   End If
End Sub

Private Sub CmdCancelar_Click()
   Unload FIngresos
End Sub

Private Sub CmdGrabar_Click()
  SumaBancos = CalculosTotalBancos(DataAsientBanco, Label3)
  Monto_Total = SumaBancos + Abono
  LabelMontoTotal.Caption = Format(Monto_Total, "#,##0")
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
  If (SumaDebe <> SumaHaber) Then
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
          MousePointer = vbHourglass
          NumComp = ReadSetDataNum(DataCodigos, "Ingresos", True)
          FechaTexto = MBoxFecha.Text
          SelectData DataTransacciones, "Transacciones", False
          SelectData DataComprobantes, "Comprobantes", False
          SelectData DataBancos, "Bancos", False
          DataAsientos.Recordset.MoveFirst
          Do While Not DataAsientos.Recordset.EOF
             With DataTransacciones.Recordset
                 .AddNew
                 .Fields("T") = Normal
                 .Fields("Asiento") = 0
                 .Fields("Fecha") = FechaTexto
                 .Fields("TP") = CompIngreso
                 .Fields("Numero") = NumComp
                 .Fields("Cta") = DataAsientos.Recordset.Fields("CODIGO")
                 .Fields("Debe") = DataAsientos.Recordset.Fields("DEBE")
                 .Fields("Haber") = DataAsientos.Recordset.Fields("HABER")
                 .Update
             End With
             DataAsientos.Recordset.MoveNext
          Loop
          'Grabacion del Comp
          With DataComprobantes.Recordset
              .AddNew
              .Fields("T") = Normal
              .Fields("TP") = CompIngreso
              .Fields("Fecha") = FechaTexto
              .Fields("Beneficiario") = TextRecibi.Text
              .Fields("Numero") = NumComp
              .Fields("Monto_Total") = CDbl(LabelMontoTotal.Caption)
              .Fields("Concepto") = TextConcepto.Text
              .Fields("Efectivo") = Abono
              .Fields("RUC_CI") = MBoxRUC.Text
              .Update
          End With
          'llenar el desposito al banco
          If DataAsientBanco.Recordset.RecordCount > 0 Then
             DataAsientBanco.Recordset.MoveFirst
             Do While Not DataAsientBanco.Recordset.EOF
                With DataBancos.Recordset
                    .AddNew
                    .Fields("Fecha") = FechaTexto
                    .Fields("TP") = CompIngreso
                    .Fields("Numero") = NumComp
                    .Fields("Banco") = DataAsientBanco.Recordset.Fields("BANCO")
                    .Fields("Cta_Cte") = DataAsientBanco.Recordset.Fields("CTA_CTE")
                    .Fields("Cheq_Dep") = DataAsientBanco.Recordset.Fields("CHEQ_DEP")
                    .Fields("Valor") = DataAsientBanco.Recordset.Fields("VALOR")
                    .Update
                End With
                DataAsientBanco.Recordset.MoveNext
             Loop
          End If
          'Llenamos el encabezado del comprobante
          sSQL = "SELECT * FROM Comprobantes "
          sSQL = sSQL & "WHERE Numero = " & NumComp & " "
          sSQL = sSQL & "AND TP = '" & CompIngreso & "' "
          SelectData DataComprobantes, sSQL, False
          'Llenar cuentas
          sSQL = "SELECT Cta,Cuentas.Cuenta,Debe,Haber FROM Transacciones,Cuentas "
          sSQL = sSQL & "WHERE TP = '" & CompIngreso & "' "
          sSQL = sSQL & "AND Numero = " & NumComp & " "
          sSQL = sSQL & "AND Cuentas.Codigo = Cta "
          sSQL = sSQL & "ORDER BY Debe DESC "
          SelectData DataTransacciones, sSQL, False
          'Llenar Bancos
          sSQL = "SELECT * FROM Bancos "
          sSQL = sSQL & "WHERE TP = '" & CompIngreso & "' "
          sSQL = sSQL & "AND Numero = " & NumComp & " "
          SelectData DataBancos, sSQL, False
          MousePointer = vbDefault
          'Imprimir comprobante
          Mensajes = "Imprimir Comprobante No. " & Format(NumComp, "000000")
          Titulo = "Pregunta de Impresion"
          TipoDeCaja = 4 + 32: ResultBox = MsgBox(Mensajes, TipoDeCaja, Titulo)
          If ResultBox = 6 Then ImprimirCompIngreso MDIConta, DataComprobantes, DataBancos, DataTransacciones
          NumComp = ReadSetDataNum(DataCodigos, "Ingresos", False)
          LabelComp.Caption = Format(NumComp, "000000")
          sSQL = "DELETE * FROM Asientos "
          DeleteData DataAsientos, sSQL
          sSQL = "DELETE * FROM Asiento_Bancos "
          DeleteData DataAsientBanco, sSQL
          sSQL = "SELECT CODIGO,CUENTA,DEBE,HABER FROM Asientos "
          SelectData DataAsientos, sSQL, False
          sSQL = "SELECT * FROM Asiento_Bancos "
          SelectData DataAsientBanco, sSQL, False
          DBAnchoMallas FIngresos, DBGAsientos, 0, DataAsientos
          DBAnchoMallas FIngresos, DBGBanco, 0, DataAsientBanco
          DataComprobantes.Database.Close
          DataTransacciones.Database.Close
          DataBancos.Database.Close
       Else
          MsgBox "Warning: Falta de Ingresar datos."
          MBoxFecha.SetFocus
       End If
    End If
  End If
End Sub

Private Sub DBGAsientos_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataAsientos)
End Sub

Private Sub DBGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then DBCCtas.SetFocus
End Sub

Private Sub DBGAsientos_LostFocus()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
End Sub

Private Sub DBGBanco_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataAsientBanco)
End Sub

Private Sub MBoxRUC_GotFocus()
  MBoxRUC.Text = "000000000-0-000"
  MarcarTexto MBoxRUC
End Sub

Private Sub TextOpcDH_Change()
  If 1 > Val(TextOpcDH.Text) Or Val(TextOpcDH.Text) > 2 Then
     TextOpcDH.Text = ""
  Else
     TextValor.SetFocus
  End If
End Sub

Private Sub TextOpcDH_GotFocus()
  TextOpcDH.Text = ""
End Sub

Private Sub Form_Activate()
  sSQL = "DELETE * FROM Asientos "
  DeleteData DataAsientos, sSQL
  sSQL = "DELETE * FROM Asiento_Bancos "
  DeleteData DataAsientBanco, sSQL
  sSQL = "SELECT CODIGO,CUENTA,DEBE,HABER FROM Asientos "
  SelectData DataAsientos, sSQL, False
  sSQL = "SELECT * FROM Asiento_Bancos "
  SelectData DataAsientBanco, sSQL, False
  DBAnchoMallas FIngresos, DBGAsientos, 0, DataAsientos
  NumComp = ReadSetDataNum(DataCodigos, "Ingresos", False)
  LabelComp.Caption = Format(NumComp, "000000")
  Label12.Caption = Empresa
  TextMonto.Visible = False
  TextBanco.Visible = False
  DBGBanco.Visible = False
  TextCtaCte.Visible = False
  TextDeposito.Visible = False
  TextCantidad.Visible = False
  MDIConta.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   CentrarForm FIngresos
   'Abriendo bases relacionadas
   DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataBancos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
   MBoxFecha.Text = FechaSistema
   MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
   If MBoxFecha.Text = "00/00/00" Then MBoxFecha.Text = FechaSistema
End Sub

Private Sub TextBanco_GotFocus()
  TextBanco = ""
End Sub

Private Sub TextBanco_LostFocus()
  If TextBanco = "" Then TextBanco = Ninguno
End Sub

Private Sub TextCantidad_GotFocus()
   TextCantidad.Text = ""
End Sub

Private Sub TextCantidad_LostFocus()
  If TextCantidad.Text = "" Then TextCantidad.Text = "0"
  TextCantidad.Text = Format(CDbl(TextCantidad.Text), "#,##0")
  Abono = CDbl(TextCantidad.Text)
  SumaBancos = CalculosTotalBancos(DataAsientBanco, Label3)
  Monto_Total = SumaBancos + Abono
  LabelMontoTotal.Caption = Format(Monto_Total, "#,##0")
End Sub

Private Sub TextCodigo_GotFocus()
  TextCodigo.Text = ""
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then
     TextCodigo.Text = "-1"
     CmdGrabar.SetFocus
  End If
End Sub

Private Sub TextCodigo_LostFocus()
  If TextCodigo.Text = "" Then TextCodigo.Text = "0"
  If Val(TextCodigo.Text) = 0 Then
     ListCuentas.Show 1
     TextCodigo.Text = Cuenta
     TextCuenta.Text = Cadena
     TextOpcDH.SetFocus
  ElseIf Val(TextCodigo.Text) <> -1 Then
     sSQL = "SELECT * FROM Cuentas WHERE Codigo = '" & TextCodigo.Text & "' "
     sSQL = sSQL & "AND DG <> 'G' "
     SelectData DataCtas, sSQL, False
     If DataCtas.Recordset.RecordCount > 0 Then
        Cuenta = DataCtas.Recordset.Fields("Codigo")
        Cadena = DataCtas.Recordset.Fields("Cuenta")
        TipoCta = DataCtas.Recordset.Fields("TC")
        TextCodigo.Text = Cuenta
        TextCuenta.Text = Cadena
        TextOpcDH.SetFocus
     Else
        sSQL = "SELECT * FROM Cuentas "
        sSQL = sSQL & "WHERE Clave = " & Val(TextCodigo.Text) & " "
        SelectData DataCtas, sSQL, False
        If DataCtas.Recordset.RecordCount > 0 Then
           Cuenta = DataCtas.Recordset.Fields("Codigo")
           Cadena = DataCtas.Recordset.Fields("Cuenta")
           TipoCta = DataCtas.Recordset.Fields("TC")
           TextCodigo.Text = Cuenta
           TextCuenta.Text = Cadena
           TextOpcDH.SetFocus
        Else
           MsgBox "Warning:" & Chr(13) & " Esta Cuenta o Clave No existen, vuelva a ingresar."
           TextCodigo.SetFocus
        End If
     End If
  End If
End Sub

Private Sub TextConcepto_Change()
   If Len(TextConcepto.Text) >= TextConcepto.MaxLength Then TextCodigo.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
   TextConcepto.Text = ""
   SumaBancos = CalculosTotalBancos(DataAsientBanco, Label3)
   Monto_Total = Abono + SumaBancos
   LabelMontoTotal.Caption = Format(Monto_Total, "#,##0")
End Sub

Private Sub TextConcepto_LostFocus()
  If TextConcepto.Text = "" Then TextConcepto.Text = Ninguno
  DBAnchoMallas FIngresos, DBGAsientos, 0, DataAsientos
  SumaBancos = CalculosTotalBancos(DataAsientBanco, Label3)
  Monto_Total = SumaBancos + Abono
  LabelMontoTotal.Caption = Format(Monto_Total, "#,##0")
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
End Sub

Private Sub TextCtaCte_GotFocus()
   TextCtaCte.Text = ""
End Sub

Private Sub TextCtaCte_LostFocus()
  If TextCtaCte.Text = "" Then TextCtaCte.Text = Ninguno
End Sub

Private Sub TextDeposito_GotFocus()
   TextDeposito.Text = ""
End Sub

Private Sub TextDeposito_LostFocus()
  If TextDeposito.Text = "" Then TextDeposito.Text = Ninguno
  If CDbl(TextMonto.Text) <> 0 Then
     DataAsientBanco.Recordset.AddNew
     DataAsientBanco.Recordset.Fields("BANCO") = Mid(TextBanco.Text, 1, 20)
     DataAsientBanco.Recordset.Fields("CTA_CTE") = Mid(TextCtaCte.Text, 1, 15)
     DataAsientBanco.Recordset.Fields("CHEQ_DEP") = Mid(TextDeposito.Text, 1, 8)
     DataAsientBanco.Recordset.Fields("VALOR") = CDbl(TextMonto.Text)
     DataAsientBanco.Recordset.Update
  End If
  SumaBancos = CalculosTotalBancos(DataAsientBanco, Label3)
  Monto_Total = SumaBancos + Abono
  LabelMontoTotal.Caption = Format(Monto_Total, "#,##0")
  TextMonto.SetFocus
End Sub

Private Sub TextMonto_GotFocus()
  TextMonto.Text = ""
End Sub

Private Sub TextMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then TextConcepto.SetFocus
End Sub

Private Sub TextMonto_LostFocus()
  If TextMonto.Text = "" Then TextMonto.Text = "0"
End Sub

Private Sub TextRecibi_GotFocus()
  TextRecibi.Text = ""
End Sub

Private Sub TextValor_GotFocus()
   TextValor.Text = ""
End Sub

Private Sub TextValor_LostFocus()
  If TextValor.Text = "" Then TextValor.Text = "0"
  If Val(TextValor.Text) <> 0 Then
      Debe = 0: Haber = 0
      If Val(TextOpcDH.Text) = 1 Then
         Debe = Val(TextValor.Text)
      ElseIf Val(TextOpcDH.Text) = 2 Then
         Haber = Val(TextValor.Text)
      End If
      If Debe = 0 And Haber = 0 Then
         MsgBox "Warning: Opcion de Ingreso incorrecta."
      Else
         DataAsientos.Recordset.AddNew
         DataAsientos.Recordset.Fields("CODIGO") = Cuenta
         DataAsientos.Recordset.Fields("CUENTA") = Cadena
         DataAsientos.Recordset.Fields("DEBE") = Debe
         DataAsientos.Recordset.Fields("HABER") = Haber
         DataAsientos.Recordset.Update
      End If
  End If
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
  TextCuenta.Text = ""
  TextCodigo.SetFocus
End Sub

Private Sub TextRecibi_Change()
  If Len(TextRecibi.Text) >= TextRecibi.MaxLength Then MBoxRUC.SetFocus
End Sub

Private Sub TextRecibi_LostFocus()
  If TextRecibi.Text = "" Then TextRecibi.Text = Ninguno
End Sub
