VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9600
   Begin MSDBCtls.DBList DBLTP 
      Bindings        =   "FCredito.frx":0000
      DataSource      =   "DataTP"
      Height          =   2595
      Left            =   105
      TabIndex        =   24
      Top             =   4200
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   4577
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
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
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
      Top             =   6720
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
      Left            =   6090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataComp 
      Caption         =   "Comp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6405
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6090
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox TxtApellidos 
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
      Left            =   105
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1680
      Width           =   2430
   End
   Begin VB.TextBox TxtNombres 
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
      Left            =   105
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1050
      Width           =   2430
   End
   Begin VB.Data DataBanco 
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
      Top             =   6300
      Visible         =   0   'False
      Width           =   2010
   End
   Begin MSDBCtls.DBCombo DBCBanco 
      Bindings        =   "FCredito.frx":0011
      DataSource      =   "DataBanco"
      Height          =   360
      Left            =   2625
      TabIndex        =   29
      Top             =   5670
      Visible         =   0   'False
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   635
      _Version        =   327680
      Text            =   "Banco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
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
      Left            =   8400
      Picture         =   "FCredito.frx":0025
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Data DataListCtas 
      Caption         =   "ListCtas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6405
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
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
      Top             =   6090
      Visible         =   0   'False
      Width           =   2010
   End
   Begin MSDBCtls.DBList DBLCtas 
      Bindings        =   "FCredito.frx":068F
      DataSource      =   "DataListCtas"
      Height          =   1620
      Left            =   2625
      TabIndex        =   7
      Top             =   420
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   2858
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
   Begin VB.Data DataCtaNo 
      Caption         =   "CtaNo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7980
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataTP 
      Caption         =   "TP"
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
      Top             =   6405
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Caption         =   "TIPO DE CUENTA"
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
      Left            =   4935
      TabIndex        =   14
      Top             =   3885
      Width           =   2745
      Begin VB.OptionButton OpcME 
         Caption         =   "Extranjera"
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
         Left            =   1365
         TabIndex        =   16
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton OpcMN 
         Caption         =   "Nacional"
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
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.TextBox TextLinea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7245
      MaxLength       =   8
      TabIndex        =   33
      Top             =   3465
      Width           =   1065
   End
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
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
      Left            =   4410
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "FCredito.frx":06A6
      Top             =   4935
      Width           =   2325
   End
   Begin VB.TextBox TxtCheque 
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
      Left            =   2625
      MaxLength       =   8
      TabIndex        =   31
      Top             =   6405
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtBanco 
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
      Left            =   2625
      MaxLength       =   20
      TabIndex        =   28
      Top             =   5670
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.CommandButton Command1 
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
      Height          =   960
      Left            =   8400
      Picture         =   "FCredito.frx":06AA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8400
      Picture         =   "FCredito.frx":0AEC
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5250
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Height          =   435
      Left            =   2100
      TabIndex        =   12
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   2625
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   105
      TabIndex        =   11
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   2625
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   767
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Apellidos"
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
      TabIndex        =   4
      Top             =   1365
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombres"
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
      TabIndex        =   2
      Top             =   735
      Width           =   2430
   End
   Begin VB.Label LabelEncaje 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2625
      TabIndex        =   39
      Top             =   4200
      Width           =   2220
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Encaje"
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
      Left            =   2625
      TabIndex        =   40
      Top             =   3885
      Width           =   2220
   End
   Begin VB.Label LabelEnCheques 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2625
      TabIndex        =   22
      Top             =   3465
      Width           =   2220
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total por Efectivizar"
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
      Left            =   2625
      TabIndex        =   21
      Top             =   3150
      Width           =   2220
   End
   Begin VB.Label LabelEstado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2625
      TabIndex        =   37
      Top             =   4935
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      Left            =   2625
      TabIndex        =   38
      Top             =   4620
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuentas Abiertas"
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
      Left            =   2625
      TabIndex        =   6
      Top             =   105
      Width           =   6840
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
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
      Left            =   2100
      TabIndex        =   9
      Top             =   2310
      Width           =   1590
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   0
      X2              =   9555
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea No."
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
      TabIndex        =   32
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Cuenta No."
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
      TabIndex        =   8
      Top             =   2310
      Width           =   1800
   End
   Begin VB.Label LabelPorConf 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4935
      TabIndex        =   20
      Top             =   3465
      Width           =   2220
   End
   Begin VB.Label LabelDisponible 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   105
      TabIndex        =   18
      Top             =   3465
      Width           =   2430
   End
   Begin VB.Label LabelSocio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3885
      TabIndex        =   13
      Top             =   2625
      Width           =   5580
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
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
      Left            =   2625
      TabIndex        =   30
      Top             =   6090
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Banco"
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
      Left            =   2625
      TabIndex        =   27
      Top             =   5355
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto Transaccion"
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
      Left            =   4410
      TabIndex        =   25
      Top             =   4620
      Width           =   2325
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo de Transaccion"
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
      Top             =   3885
      Width           =   2430
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Total"
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
      Left            =   4935
      TabIndex        =   19
      Top             =   3150
      Width           =   2220
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo Disponible"
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
      TabIndex        =   17
      Top             =   3150
      Width           =   2430
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I. / R.U.C."
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
      TabIndex        =   0
      Top             =   105
      Width           =   2430
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre y Apellidos"
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
      Left            =   3885
      TabIndex        =   10
      Top             =   2310
      Width           =   5580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9555
      Y1              =   2100
      Y2              =   2100
   End
End
Attribute VB_Name = "FCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub InsertarMontos(MBFecha As MaskEdBox, DtaCta As Data, DtaBanc As Data, CuentaNo As MaskEdBox, TDebe As Double, THaber As Double, NoCheque As String, NomBanco As String)
  SaldoDisp = 0: SaldoCont = 0
  sSQL = "SELECT * FROM Transacciones " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
       & "ORDER BY Fecha,ID,Hora "
  SelectData DtaCta, sSQL, False
  With DtaCta.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
       End If
      .AddNew
      .Fields("CL") = True
      .Fields("Fecha") = MBFecha.Text
      .Fields("Cuenta_No") = CuentaNo.Text
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
       If TipoGrupo Then
          If THaber <> 0 Then
            .Fields("Saldo_Disp") = SaldoDisp
            .Fields("T") = Pendiente
          Else
            .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
            .Fields("T") = Normal
          End If
       Else
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
         .Fields("T") = Normal
       End If
      .Fields("Usuario") = NombreUsuario
      .Fields("ID") = NumeroLineas + 1
      .Fields("Hora") = Format(Time, FormatoTimes)
      .Fields("Item") = NumEmpresa
      .Fields("ME") = Moneda_US
      .Fields("Cheque") = NoCheque
      .Update
       Select Case TipoProc
         Case "APER", "DEP", "RET", "DEPS", "RETS", "BOVE"
             .AddNew
             .Fields("CL") = False
             .Fields("Fecha") = MBFecha.Text
             .Fields("Cuenta_No") = Ninguno
             .Fields("TP") = TipoProc
             .Fields("Debitos") = THaber
             .Fields("Creditos") = TDebe
             .Fields("Saldo_Cont") = 0
             .Fields("Saldo_Disp") = 0
             .Fields("T") = Normal
             .Fields("Usuario") = NombreUsuario
             .Fields("ID") = 0
             .Fields("Hora") = Format(Time, FormatoTimes)
             .Fields("Item") = NumEmpresa
             .Fields("ME") = Moneda_US
             .Fields("Cheque") = NoCheque
             .Update
       End Select
  End With
  If TipoGrupo Then
     sSQL = "SELECT * FROM Bancos " _
          & "WHERE Cuenta_No = '" & CuentaNo.Text & "' "
     SelectData DataCtaNo, sSQL, False
     With DtaBanc.Recordset
         .AddNew
         .Fields("T") = Pendiente
         .Fields("ME") = Moneda_US
         .Fields("Fecha") = MBFecha.Text
         .Fields("Cuenta_No") = CuentaNo.Text
         .Fields("TP") = TipoProc
         .Fields("Banco") = NomBanco
         .Fields("Cheque") = NoCheque
         .Fields("Valor") = THaber - TDebe
         .Fields("Dias") = 4
         .Fields("Item") = NumEmpresa
         .Update
     End With
  End If
End Sub

Public Sub ListarCuentas(CuentaNo As MaskEdBox)
  SaldoDisp = 0:  SaldoCont = 0
  TotalEncaje = 0
  NumeroLineas = 0
  Moneda_US = False
  LabelSocio.Caption = ""
  T = Normal
  TextLinea.Text = NumeroLineas
  LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
  LabelPorConf.Caption = Format(SaldoCont, "#,##0.00")
  LabelEnCheques.Caption = Format(SaldoCont - SaldoDisp, "#,##0.00")
  sSQL = "SELECT * FROM Bloqueos " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' "
  SelectData DataCuentas, sSQL, False
  With DataCuentas.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       TotalEncaje = .Fields("Valor")
   End If
  End With
  sSQL = "SELECT * FROM Cuentas " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' "
  SelectData DataCuentas, sSQL, False
  With DataCuentas.Recordset
   If .RecordCount > 0 Then
       ConLibreta = True
       T = .Fields("T")
       Moneda_US = .Fields("ME")
       NoCert = .Fields("No_Soc")
       LabelSocio.Caption = .Fields("Nombres") & "  " & .Fields("Apellidos")
       sSQL = "SELECT * FROM Transacciones " _
            & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
            & "ORDER BY Fecha,ID,Hora "
       SelectData DataCtaNo, sSQL, False
       If DataCtaNo.Recordset.RecordCount > 0 Then
          DataCtaNo.Recordset.MoveLast
          NumeroLineas = DataCtaNo.Recordset.Fields("ID")
          SaldoDisp = DataCtaNo.Recordset.Fields("Saldo_Disp")
          SaldoCont = DataCtaNo.Recordset.Fields("Saldo_Cont")
       End If
       If Moneda_US Then OpcME.Value = True Else OpcMN.Value = True
       
       TextLinea.Text = NumeroLineas
       LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
       LabelPorConf.Caption = Format(SaldoCont, "#,##0.00")
       LabelEnCheques.Caption = Format(SaldoCont - SaldoDisp, "#,##0.00")
       LabelEncaje.Caption = Format(TotalEncaje, "#,##0.00")
       'sSQL = "SELECT Proceso FROM Tipo_Proceso ORDER BY Proceso "
       'SelectDBList DBLTP, DataTP, sSQL, "Proceso"
       If T = Procesado Then
          LabelEstado.Caption = "Apertura"
          sSQL = "SELECT Proceso FROM Tipo_Proceso " _
               & "WHERE Mid(TP,1,3) = 'APE' " _
               & "AND Nivel = 1 " _
               & "ORDER BY Proceso "
       Else
          LabelEstado.Caption = "Normal"
          sSQL = "SELECT Proceso FROM Tipo_Proceso " _
               & "WHERE Mid(TP,1,3) <> 'APE' " _
               & "AND Nivel = 1 " _
               & "ORDER BY Proceso "
       End If
       SelectDBList DBLTP, DataTP, sSQL, "Proceso"
       MBoxFecha.SetFocus
   Else
       ConLibreta = False
       MsgBox "Esta cuenta no exite"
       LabelEstado.Caption = "NINGUNO"
       LabelSocio.Caption = "TRANSACCION SIN LIBRETA"
       sSQL = "SELECT Proceso FROM Tipo_Proceso " _
            & "WHERE Mid(TP,1,3) <> 'APE' " _
            & "AND Nivel = 1 " _
            & "ORDER BY Proceso "
       SelectDBList DBLTP, DataTP, sSQL, "Proceso"
       MBoxFecha.SetFocus
   End If
  End With
End Sub

Private Sub Command1_Click()
  Cadena = DBLTP.Text
  Monto_Total = 0
  If Cadena = "" Then Cadena = Ninguno
  sSQL = "SELECT * FROM Tipo_Proceso " _
       & "WHERE Proceso = '" & Cadena & "' "
  SelectData DataCtaNo, sSQL, False
  With DataCtaNo.Recordset
   If .RecordCount > 0 Then
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
   End If
  End With
  sSQL = "SELECT * FROM Monto_Apertura "
  sSQL = sSQL & "WHERE ME = " & Val(Moneda_US) & " "
  SelectData DataCtaNo, sSQL, False
  If DataCtaNo.Recordset.RecordCount > 0 Then
     Monto_Total = DataCtaNo.Recordset.Fields("Monto_Aper")
     Parcial = DataCtaNo.Recordset.Fields("Monto_Cert") * NoCert
  End If
  If NumCheque = "" Then NumCheque = Ninguno
  If NombreBanco = "" Then NombreBanco = Ninguno
  Select Case DC_Caja
    Case "C"
         InsertarMontos MBoxFecha, DataCuentas, DataCtaNo, MBoxCuenta, 0, Val(TxtMonto.Text), NumCheque, NombreBanco
         If T = Procesado Then
            NumeroLineas = NumeroLineas + 1
            TipoProc = "N/DG"
            InsertarMontos MBoxFecha, DataCuentas, DataCtaNo, MBoxCuenta, Monto_Total, 0, NumCheque, NombreBanco
            NumeroLineas = NumeroLineas + 1
            TipoProc = "N/DC"
            InsertarMontos MBoxFecha, DataCuentas, DataCtaNo, MBoxCuenta, Parcial, 0, NumCheque, NombreBanco
         End If
         sSQL = "UPDATE Cuentas SET T = '" & Normal & "' " _
              & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
         UpdateData DataCuentas, sSQL
    Case "D"
         InsertarMontos MBoxFecha, DataCuentas, DataCtaNo, MBoxCuenta, Val(TxtMonto.Text), 0, NumCheque, NombreBanco
         ValorDH = Val(TxtMonto.Text)
         If TipoGrupo Then
            'Procesar Diario
            NumComp = ReadSetDataNum("Diario", True, True)
            sSQL = "SELECT * FROM Asientos_D_" & CodigoUsuario & " "
            SelectData DataAsientos, sSQL, False
            sSQL = "DELETE * FROM Asientos_D_" & CodigoUsuario & " "
            DeleteData DataAsientos, sSQL
            sSQL = "SELECT * FROM Asientos_D_" & CodigoUsuario & " "
            SelectData DataAsientos, sSQL, False
            sSQL = "SELECT * FROM Asientos_SC_D_" & CodigoUsuario
            SelectData DataSubCta, sSQL, False
            InsertarAsientos DataAsientos, Cta_Libretas, 0, ValorDH, 0
            InsertarAsientos DataAsientos, SinEspaciosIzq(DBCBanco.Text), 0, 0, ValorDH
            Co.T = Normal
            Co.TP = CompDiario
            Co.Fecha = MBoxFecha.Text
            Co.Numero = NumComp
            Co.Concepto = "Retiro en Cheque No. " & NoCheque & ", del banco " & DBCBanco.Text & ", de la Cuenta No. " & MBoxCuenta.Text
            Co.Beneficiario = Ninguno
            Co.Efectivo = 0
            Co.Monto_Total = ValorDH
            GrabarComprobantes Co, DataAsientos, DataSubCta
            RatonNormal
            ImprimirComprobantesDe False, CompDiario, NumComp, NumEmpresa, DataComp, DataTrans
         End If
  End Select
  MBoxCuenta.SetFocus
End Sub

Private Sub Command2_Click()
   Unload FCaja
End Sub

Private Sub DBCBanco_LostFocus()
  NombreBanco = SinEspaciosIzq(DBCBanco.Text)
End Sub

Private Sub DBLTP_GotFocus()
  DBCBanco.Visible = False
  TxtBanco.Visible = False
End Sub

Private Sub DBLTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLTP_LostFocus()
  Cadena = DBLTP.Text
  If Cadena = "" Then Cadena = Ninguno
  sSQL = "SELECT * FROM Tipo_Proceso " _
       & "WHERE Proceso = '" & Cadena & "' "
  SelectData DataCtaNo, sSQL, False
  With DataCtaNo.Recordset
   If .RecordCount > 0 Then
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
       Label9.Visible = TipoGrupo
       Label10.Visible = TipoGrupo
       If DC_Caja = "D" Then
          DBCBanco.Visible = TipoGrupo
       Else
          TxtBanco.Visible = TipoGrupo
       End If
       TxtCheque.Visible = TipoGrupo
   End If
  End With
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Proceso FROM Tipo_Proceso " _
       & "WHERE Nivel = 1 " _
       & "ORDER BY Proceso "
  SelectDBList DBLTP, DataTP, sSQL, "Proceso"
  sSQL = "SELECT (Codigo & ' ' & Cuenta) As NomBanco FROM Catalogo " _
       & "WHERE TC = 'BA' " _
       & "ORDER BY Codigo "
  SelectDBCombo DBCBanco, DataBanco, sSQL, "NomBanco", False
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCaja
  DataTP.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataCtaNo.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataCuentas.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataListCtas.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataComp.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSubCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxCuenta_GotFocus()
   MBoxCuenta.Text = "00000000-0"
End Sub

Private Sub MBoxCuenta_LostFocus()
  NombreBanco = ""
  NumCheque = ""
  ListarCuentas MBoxCuenta
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
  DBLTP.SetFocus
End Sub

Private Sub MBoxRUC_LostFocus()
  sSQL = "SELECT (Cuenta_No  & ' => ' & Nombres & ' ' & Apellidos) As CuentasNo " _
       & "FROM Cuentas " _
       & "WHERE RUC_CI = '" & MBoxRUC.Text & "' " _
       & "ORDER BY Cuenta_No "
  SelectDBList DBLCtas, DataListCtas, sSQL, "CuentasNo"
End Sub

Private Sub TxtApellidos_LostFocus()
  I = Len(TxtApellidos.Text)
  If I > 0 Then
  sSQL = "SELECT (Cuenta_No  & ' => ' & Nombres & ' ' & Apellidos) As CuentasNo " _
       & "FROM Cuentas " _
       & "WHERE Mid(Apellidos,1," & I & ") = '" & UCase(TxtApellidos.Text) & "' " _
       & "ORDER BY Cuenta_No "
  SelectDBList DBLCtas, DataListCtas, sSQL, "CuentasNo"
  End If
End Sub

Private Sub TxtBanco_LostFocus()
  TextoValido TxtBanco, , True
  NombreBanco = TxtBanco.Text
End Sub

Private Sub TxtCheque_LostFocus()
  TextoValido TxtCheque, , True
  NumCheque = TxtCheque.Text
End Sub

Private Sub TxtMonto_GotFocus()
  TxtMonto.Text = ""
End Sub

Private Sub TxtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonto_LostFocus()
  TextoValido TxtMonto, True
  If SaldoDisp > 0 And ConLibreta Then
     If (DC_Caja = "D") Then
        If (Val(TxtMonto.Text) >= SaldoDisp - TotalEncaje) Then
            MsgBox "Excede el monto de encaje"
            TxtMonto.SetFocus
        End If
     End If
  End If
End Sub

Private Sub TxtNombres_LostFocus()
  I = Len(TxtNombres.Text)
  If I > 0 Then
  sSQL = "SELECT (Cuenta_No  & ' => ' & Nombres & ' ' & Apellidos) As CuentasNo " _
       & "FROM Cuentas " _
       & "WHERE Mid(Nombres,1," & I & ") = '" & UCase(TxtNombres.Text) & "' " _
       & "ORDER BY Cuenta_No "
  SelectDBList DBLCtas, DataListCtas, sSQL, "CuentasNo"
  End If
End Sub
