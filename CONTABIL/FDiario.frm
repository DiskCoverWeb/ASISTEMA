VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FDiario 
   Caption         =   "COMPROBANTE DE DIARIO"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin MSDBCtls.DBList DBLCuentas 
      Bindings        =   "FDiario.frx":0000
      DataSource      =   "DataCuentas"
      Height          =   4155
      Left            =   1890
      TabIndex        =   18
      Top             =   2310
      Visible         =   0   'False
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   7329
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
      TabIndex        =   19
      Top             =   1785
      Visible         =   0   'False
      Width           =   2115
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
         TabIndex        =   21
         Text            =   "FDiario.frx":0016
         Top             =   630
         Width           =   330
      End
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
         TabIndex        =   20
         Text            =   "FDiario.frx":0018
         Top             =   210
         Width           =   330
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
         TabIndex        =   34
         Top             =   210
         Width           =   1485
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
         TabIndex        =   33
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label9 
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
         TabIndex        =   32
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label6 
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
         TabIndex        =   31
         Top             =   840
         Width           =   1065
      End
   End
   Begin MSDBGrid.DBGrid DBGAsientos 
      Bindings        =   "FDiario.frx":001A
      Height          =   4425
      Left            =   105
      OleObjectBlob   =   "FDiario.frx":0031
      TabIndex        =   23
      Top             =   2415
      Width           =   11355
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
      Left            =   1890
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Concepto del Asiento de Diario"
      Top             =   1155
      Width           =   9570
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
      Width           =   4635
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
      TabIndex        =   22
      ToolTipText     =   "Monto del Asiento"
      Top             =   1995
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
      TabIndex        =   17
      Top             =   1995
      Width           =   7260
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
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
   Begin VB.Data DataSubCta 
      Caption         =   "SubCta"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataSubCtaDet 
      Caption         =   "SubCtaDet"
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
      Top             =   5880
      Visible         =   0   'False
      Width           =   1905
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
      TabIndex        =   16
      Text            =   "0"
      Top             =   1995
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
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
      Top             =   5880
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
      Left            =   7035
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   1905
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
      TabIndex        =   24
      Top             =   6930
      Width           =   1275
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
      Height          =   330
      Left            =   1470
      TabIndex        =   25
      Top             =   6930
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
      Left            =   4620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
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
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   2010
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
   Begin VB.Data DataRet 
      Caption         =   "Ret"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3150
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataDetRet 
      Caption         =   "DetRet"
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
      Top             =   2835
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label17 
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
   Begin VB.Label Label3 
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
      TabIndex        =   15
      Top             =   1785
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
      Left            =   2940
      TabIndex        =   30
      Top             =   6930
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
      Left            =   4305
      TabIndex        =   29
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO:"
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
      Width           =   4635
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
      TabIndex        =   13
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label LabelDiario 
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
      Width           =   1275
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
      TabIndex        =   28
      Top             =   6930
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
      TabIndex        =   27
      Top             =   6930
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
      Left            =   6090
      TabIndex        =   26
      Top             =   6930
      Width           =   1275
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIGITE LA CLAVE O BUSQUE LA CUENTA"
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
      TabIndex        =   14
      Top             =   1785
      Width           =   7260
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
      Width           =   4950
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
   Begin VB.Label Label2 
      Caption         =   " Diario No."
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
      Left            =   8505
      TabIndex        =   1
      Top             =   105
      Width           =   1275
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
      TabIndex        =   11
      Top             =   1155
      Width           =   1695
   End
End
Attribute VB_Name = "FDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub CmdGrabar_Click()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber, LabelDiferencia
  If Round_ME(SumaDebe) <> Round_ME(SumaHaber) Then
     Mensajes = "Las transacciones no cuadran correctamente" & Chr(13) _
              & "corrija los resultados de las cuentas"
     MsgBox Mensajes
     TextCodigo.SetFocus
  Else
     Mensajes = "Esta seguro de Grabar el Comprobante No. " & LabelDiario.Caption & "]"
     Titulo = "Pregunta de grabación"
     If BoxMensaje = 6 Then
       If DataAsientos.Recordset.RecordCount > 0 Then
          RatonReloj
          If NumCompDia = 0 Then
             NumComp = ReadSetDataNum("Diario", True, True)
          Else
             NumComp = NumCompDia
          End If
          sSQL = "SELECT * FROM Asientos_R_" & CodigoUsuario & " "
          sSQL = sSQL & "ORDER BY CTA "
          SelectData DataDetRet, sSQL, False
          sSQL = "SELECT * FROM Asientos_SC_D_" & CodigoUsuario & " "
          'sSQL = sSQL & "WHERE Valor <> 0 "
          SelectData DataSubCtaDet, sSQL, False
         'Grabacion del Comprobante
          FechaTexto = MBoxFecha.Text
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto.Text
          Co.Beneficiario = TextBenef.Text
          Co.RUC_CI = MBoxRUC.Text
          Co.Efectivo = 0
          Co.Monto_Total = SumaDebe
          GrabarComprobantes Co, DataAsientos, DataSubCtaDet, , DataDetRet
          sSQL = "SELECT * FROM Asientos_R_" & CodigoUsuario & " "
          SelectData DataDetRet, sSQL, False
          sSQL = "DELETE * FROM Asientos_SC_D_" & CodigoUsuario & " "
          UpdateData DataSubCtaDet, sSQL
          IniciarAsientosDe CompDiario, DataAsientos, DBGAsientos, , , DataDetRet
          RatonNormal
          If OpcCoop = False Then ImprimirComprobantesDe False, CompDiario, NumComp, NumEmpresa, DataComprobantes, DataTransacciones, , DataRet
          NumComp = NumComp + 1
          LabelDiario.Caption = Format(NumComp, "000000")
          If NumCompDia <> 0 Then
             Unload FDiario
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
     Label9.Visible = False
     Label16.Visible = False
     TextOpcTM.Visible = False
     TextOpcDH.SetFocus
  Else
     TextOpcTM.SetFocus
  End If
End Sub

Private Sub Form_Deactivate()
  FDiario.WindowState = 1
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
   If CDbl(TextCotiza.Text) <= 0 Then TextCotiza.Text = Dolar
   Dolar = CDbl(TextCotiza.Text)
   'MsgBox Cambio_Letras(Val(TextCotiza.Text))
End Sub

Private Sub TextValor_GotFocus()
  If SubCta = "R" And OpcDH = 2 Then
     TextValor.Text = Total_DetRet
     ValorDH = Val(TextValor.Text)
    'If Moneda_US Or OpcTM = 2 Then ValorDH = Round(ValorDH / Dolar)
     InsertarAsiento DataAsientos
     SiguienteControl
  Else
     TextValor.Text = ""
     Label3.Caption = "VALOR M/N"
     If Moneda_US Or OpcTM = 2 Then Label3.Caption = "VALOR M/E"
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
  ValorDH = Val(TextValor.Text)
  'MsgBox ValorDH
  InsertarAsiento DataAsientos
  TotalSubCta = ValorDH
  If ValorDH <> 0 Then
     Select Case SubCta
       Case "C", "P", "G", "I"
            FechaTexto = MBoxFecha.Text
            SubCtaGen = Codigo
            FSubCtas.Show 1
     End Select
  End If
  TextCuenta.Text = ""
  End If
  TextCodigo.SetFocus
End Sub

Private Sub DBGAsientos_BeforeDelete(Cancel As Integer)
  Codigo = DataAsientos.Recordset.Fields("CODIGO")
  Cancel = DeleteSiNo(DataAsientos)
  If Cancel = False Then EliminarSubCta DataSubCtaDet, Codigo
End Sub

Private Sub CmdCancelar_Click()
   Unload FDiario
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
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
  MarcarTexto TextCodigo
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
         SiguienteControl
    Case vbKeyEscape
         TextCodigo.Text = "-1"
         TextValor.Visible = True
         CmdGrabar.SetFocus
    Case vbKeyF2
         TextCodigo.Text = "-1"
         FIngCtas.Show
  End Select
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, True
  LeerCodigoCta TextCodigo, TextCuenta, TextValor, DBLCuentas, DataCuentas, FrameAsigna, TextOpcTM, TextOpcDH
End Sub

Private Sub Form_Activate()
  FDiario.WindowState = 2
  TipoDoc = CompDiario
  CTAsientos_R
  CTAsientos TipoDoc
  CTAsientos_SC TipoDoc
  IniciarAsientosDe TipoDoc, DataAsientos, DBGAsientos, , , DataDetRet
  IniciarAsientoSC_De TipoDoc, DataSubCtaDet
  SelectCuentas DBLCuentas, DataCuentas
  Una_Vez = True
  NumComp = ReadSetDataNum("Diario", True, False)
  If NumCompDia <> 0 Then
     NumComp = NumCompDia
     NumEmpresa = NumItem
  Else
     NumEmpresa = NumItemTemp
  End If
  Label5.Caption = Empresa
  LabelDiario.Caption = Format(NumComp, "000000")
  Codigo = ""
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataDetRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSubCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSubCtaDet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextConcepto_Change()
   If TextoMaximo(TextConcepto) Then TextCodigo.SetFocus
End Sub

Private Sub TextConcepto_LostFocus()
   TextoValido TextConcepto, False
End Sub

