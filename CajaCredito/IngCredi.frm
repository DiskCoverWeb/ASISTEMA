VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form IngPlanCreditos 
   Caption         =   "Ingreso/Modificacion de Productos de Inventario"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3780
      ScaleHeight     =   855
      ScaleWidth      =   3060
      TabIndex        =   45
      Top             =   7140
      Width           =   3060
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Crédito"
      Height          =   540
      Left            =   105
      TabIndex        =   42
      Top             =   7140
      Width           =   3585
      Begin VB.OptionButton Option1 
         Caption         =   "Por Días"
         Height          =   225
         Left            =   1365
         TabIndex        =   44
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton OpcM 
         Caption         =   "Mensual"
         Height          =   225
         Left            =   105
         TabIndex        =   43
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.TextBox TextSubCta 
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
      Left            =   5685
      MaxLength       =   80
      TabIndex        =   4
      Top             =   4200
      Width           =   5370
   End
   Begin VB.CommandButton Command6 
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
      Height          =   750
      Left            =   10290
      Picture         =   "IngCredi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   960
   End
   Begin ComctlLib.TreeView TVCatalogo 
      Height          =   4050
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "Un click en el dibujo de la Cta. y presionar la tecla <DEL> Borra la Cta."
      Top             =   105
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7144
      _Version        =   327682
      Indentation     =   794
      Sorted          =   -1  'True
      Style           =   5
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   420
      Top             =   1260
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "Inv"
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
      Height          =   750
      Left            =   10290
      Picture         =   "IngCredi.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1785
      Width           =   960
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
      Height          =   750
      Left            =   10290
      Picture         =   "IngCredi.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   945
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   420
      Top             =   1575
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoCodInv 
      Height          =   330
      Left            =   420
      Top             =   1890
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "CodInv"
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
   Begin MSMask.MaskEdBox MBoxCodigo 
      Height          =   330
      Left            =   1170
      TabIndex        =   5
      Top             =   4200
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
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
      Format          =   "CC.CC.CCC.CCCCC"
      Mask            =   "CC.CC.CCC.CCCCC"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MBCta_Vigente 
      Height          =   330
      Index           =   0
      Left            =   2100
      TabIndex        =   14
      Top             =   5040
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vigente 
      Height          =   330
      Index           =   1
      Left            =   2100
      TabIndex        =   15
      Top             =   5460
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vigente 
      Height          =   330
      Index           =   2
      Left            =   2100
      TabIndex        =   16
      Top             =   5880
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vigente 
      Height          =   330
      Index           =   3
      Left            =   2100
      TabIndex        =   17
      Top             =   6300
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vigente 
      Height          =   330
      Index           =   4
      Left            =   2100
      TabIndex        =   18
      Top             =   6720
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vencido 
      Height          =   330
      Index           =   0
      Left            =   5775
      TabIndex        =   25
      Top             =   5040
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vencido 
      Height          =   330
      Index           =   1
      Left            =   5775
      TabIndex        =   26
      Top             =   5460
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vencido 
      Height          =   330
      Index           =   2
      Left            =   5775
      TabIndex        =   27
      Top             =   5880
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vencido 
      Height          =   330
      Index           =   3
      Left            =   5775
      TabIndex        =   28
      Top             =   6300
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vencido 
      Height          =   330
      Index           =   4
      Left            =   5775
      TabIndex        =   29
      Top             =   6720
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Interes 
      Height          =   330
      Left            =   9450
      TabIndex        =   36
      Top             =   5040
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Interes_Mora 
      Height          =   330
      Left            =   9450
      TabIndex        =   37
      Top             =   5460
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Cobranza 
      Height          =   330
      Left            =   9450
      TabIndex        =   38
      Top             =   5880
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Seguro_D 
      Height          =   330
      Left            =   9450
      TabIndex        =   39
      Top             =   6300
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   330
      Left            =   9450
      TabIndex        =   40
      Top             =   6720
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De más de 360 días"
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
      Left            =   7455
      TabIndex        =   41
      Top             =   7140
      Width           =   2010
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X"
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
      Left            =   7455
      TabIndex        =   35
      Top             =   6720
      Width           =   2010
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Seguro de Desgrava."
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
      Left            =   7455
      TabIndex        =   34
      Top             =   6300
      Width           =   2010
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Servicios Operativos"
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
      Left            =   7455
      TabIndex        =   33
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VARIAS CUENTAS DE CREDITOS"
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
      Left            =   7455
      TabIndex        =   32
      Top             =   4620
      Width           =   3585
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interés por Mora"
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
      Left            =   7455
      TabIndex        =   31
      Top             =   5460
      Width           =   2010
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interés"
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
      Left            =   7455
      TabIndex        =   30
      Top             =   5040
      Width           =   2010
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De más de 360 días"
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
      Left            =   3780
      TabIndex        =   24
      Top             =   6720
      Width           =   2010
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 181 a 360 días"
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
      Left            =   3780
      TabIndex        =   23
      Top             =   6300
      Width           =   2010
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 91 a 180 días"
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
      Left            =   3780
      TabIndex        =   22
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTAS CREDITOS VENCIDO"
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
      Left            =   3780
      TabIndex        =   21
      Top             =   4620
      Width           =   3585
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 31 a 90 días"
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
      Left            =   3780
      TabIndex        =   20
      Top             =   5460
      Width           =   2010
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 1 a 30 días"
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
      Left            =   3780
      TabIndex        =   19
      Top             =   5040
      Width           =   2010
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De más de 360 días"
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
      TabIndex        =   13
      Top             =   6720
      Width           =   2010
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 181 a 360 días"
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
      TabIndex        =   12
      Top             =   6300
      Width           =   2010
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 91 a 180 días"
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
      TabIndex        =   11
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTAS CREDITOS VIGENTES"
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
      Top             =   4620
      Width           =   3585
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Codigo:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C&oncepto"
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
      Left            =   2955
      TabIndex        =   8
      Top             =   4200
      Width           =   2745
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 1 a 30 días"
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
      TabIndex        =   7
      Top             =   5040
      Width           =   2010
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " De 31 a 90 días"
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
      TabIndex        =   6
      Top             =   5460
      Width           =   2010
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngCredi.frx":114E
            Key             =   "Uno"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngCredi.frx":1468
            Key             =   "Dos"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngCredi.frx":1782
            Key             =   "Tres"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "IngPlanCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cta_Ini As String
Dim Cta_Fin As String
Dim nodX As Node

Public Sub AddNewCtaInv()
  If Len(Codigo) = 1 Then
     Set nodX = TVCatalogo.Nodes.Add(, , Codigo, Cuenta)
     nodX.Image = ImageList1.ListImages(1).key
     nodX.SelectedImage = ImageList1.ListImages(1).key
  Else
     Set nodX = TVCatalogo.Nodes.Add(Cta_Sup, tvwChild, Codigo, Cuenta)
     If Len(Codigo) = 1 Then IE = 1 Else IE = 3
     nodX.Image = ImageList1.ListImages(IE).key
     nodX.SelectedImage = ImageList1.ListImages(IE).key
     TVCatalogo.Tag = Codigo
  End If
End Sub

Private Sub Command1_Click()
  GrabarInv
End Sub

Private Sub Command2_Click()
  Unload IngPlanCreditos
End Sub

Private Sub Command6_Click()
  Imprimir_Codigos_Estanteria SinEspaciosIzq(TVCatalogo.SelectedItem)
End Sub

Private Sub Form_Activate()
  'FormatoMaskCta MBoxCta_Inv
  'FormatoMaskCta MBoxCta
  'FormatoMaskCta MBoxCta1
  'FormatoMaskCta MBoxCta_Ing
  'FormatoMaskCta MBoxCta_Ing0
  FormatoMaskCodC MBoxCodigo
  Si_No = False
  sSQL = "SELECT Item,Codigo,Periodo " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoTInv, sSQL
  If AdoTInv.Recordset.RecordCount > 0 Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Prestamo " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo "
     SelectAdodc AdoInv, sSQL
     If AdoInv.Recordset.RecordCount > 0 Then
        ReDim CodigoCtas(AdoInv.Recordset.RecordCount + 1) As String
        For I = 0 To AdoInv.Recordset.RecordCount
            CodigoCtas(I) = ""
        Next I
     End If
     Contador = 0
     Do While Not AdoTInv.Recordset.EOF
        Codigo = AdoTInv.Recordset.Fields("Codigo")
        Cta_Sup = CodigoCuentaSup(Codigo)
        'MsgBox Codigo
        With AdoInv.Recordset
         If .RecordCount > 0 Then
             Do While (Cta_Sup <> "0")
               .MoveFirst
               .Find ("Codigo Like '" & Cta_Sup & "' ")
                If Not .EOF And Cta_Sup <> "0" Then
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                Else
                   Si_No = True: Evaluar = True
                   For I = 0 To AdoInv.Recordset.RecordCount
                       If CodigoCtas(I) = Cta_Sup Then Evaluar = False
                   Next I
                   If Evaluar Then
                      SetAdoAddNew "Catalogo_Prestamo"
                      SetAdoFields "Item", NumEmpresa
                      SetAdoFields "Codigo", Cta_Sup
                      SetAdoFields "Descripcion", "NINGUN PRODUCTO"
                      SetAdoFields "Periodo", Periodo_Contable
                      SetAdoUpdate
                      CodigoCtas(Contador) = Cta_Sup
                      Contador = Contador + 1
                   End If
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                End If
             Loop
         End If
        End With
        AdoTInv.Recordset.MoveNext
     Loop
  End If
  RatonNormal
  If Si_No Then
     Cadena = vbCrLf
     For I = 0 To Contador
         Cadena = Cadena & CodigoCtas(I) & vbCrLf
     Next I
     MsgBox "Los siguientes Codigos no se han creado: " & vbCrLf _
            & Cadena & "ADVERTENCIA: REVIZAR."
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoInv, sSQL
  RatonReloj
  With AdoInv.Recordset
   If .RecordCount > 0 Then
'      Codigo = "C" & .Fields("Codigo_Inv")
'      Cta_Sup = "C" & CodigoCuentaSup(.Fields("Codigo_Inv"))
       .MoveFirst
        Do While Not .EOF
           If Len(.Fields("Codigo")) = 1 Then
              Codigo = .Fields("Codigo")
              Cta_Sup = .Fields("Codigo")
              Cuenta = .Fields("Codigo") & " - " & .Fields("Descripcion")
              AddNewCtaInv
           Else
              Codigo = .Fields("Codigo")
              Cta_Sup = CodigoCuentaSup(.Fields("Codigo"))
              Cuenta = .Fields("Codigo") & " - " & .Fields("Descripcion")
              AddNewCtaInv
           End If
          .MoveNext
        Loop
    End If
   End With
  'MBoxCodigo.SetFocus
  TVCatalogo.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  'CentrarForm IngProdInv
  ConectarAdodc AdoAux
  ConectarAdodc AdoInv
  ConectarAdodc AdoTInv
  ConectarAdodc AdoCodInv
End Sub

Private Sub MBoxCodigo_GotFocus()
  MarcarTexto MBoxCodigo
End Sub

Private Sub MBoxCodigo_LostFocus()
   MBoxCodigo.Text = UCase(MBoxCodigo.Text)
End Sub

'''Private Sub MBoxCta_Ing_GotFocus()
'''  MarcarTexto MBoxCta_Ing
'''End Sub
'''
'''Private Sub MBoxCta_Ing_KeyDown(KeyCode As Integer, Shift As Integer)
'''  PresionoEnter KeyCode
'''End Sub
'''
'''Private Sub MBoxCta_Ing0_GotFocus()
'''  MarcarTexto MBoxCta_Ing0
'''End Sub
'''
'''Private Sub MBoxCta_Ing0_KeyDown(KeyCode As Integer, Shift As Integer)
'''  PresionoEnter KeyCode
'''End Sub
'''
'''Private Sub MBoxCta_Inv_GotFocus()
'''  MarcarTexto MBoxCta_Inv
'''End Sub

'''Private Sub MBoxCta_Inv_KeyDown(KeyCode As Integer, Shift As Integer)
'''  PresionoEnter KeyCode
'''End Sub
'''
'''Private Sub MBoxCta1_GotFocus()
'''  MarcarTexto MBoxCta1
'''End Sub
'''
'''Private Sub MBoxCta1_KeyDown(KeyCode As Integer, Shift As Integer)
'''  PresionoEnter KeyCode
'''End Sub

Private Sub TextSubCta_GotFocus()
  MarcarTexto TextSubCta
End Sub

Private Sub TextSubCta_LostFocus()
  TextoValido TextSubCta
End Sub

Public Sub LlenarInv()
   'FormatoMaskCta MBoxCta_Inv
   'FormatoMaskCta MBoxCta
   'FormatoMaskCta MBoxCta1
'''   FormatoMaskCta MBoxCta_Ing
'''   FormatoMaskCta MBoxCta_Ing0
   FormatoMaskCodC MBoxCodigo
   TextSubCta.Text = ""
   With AdoInv.Recordset
    If .RecordCount > 0 Then
        Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
        'Codigo = SinEspaciosIzq(DCInv.Text)
        'MsgBox Codigo & vbCrLf & CodigosSinPuntos(Codigo)
       .MoveFirst
        TextoBusqueda = "Codigo Like '" & Codigo & "' "
       .Find (TextoBusqueda)
        If Not .EOF Then
           TextSubCta.Text = .Fields("Descripcion")
           MBoxCodigo.Text = FormatoCodigoCredito(.Fields("Codigo"))
'''           MBoxCta_Inv.Text = FormatoCodigoCta(.Fields("Cta_Inventario"))
'''           MBoxCta.Text = FormatoCodigoCta(.Fields("Cta_Proveedor"))
'''           MBoxCta1.Text = FormatoCodigoCta(.Fields("Cta_Costo_Venta"))
'''           MBoxCta_Ing.Text = FormatoCodigoCta(.Fields("Cta_Ventas"))
'''           MBoxCta_Ing0.Text = FormatoCodigoCta(.Fields("Cta_Ventas_0"))
        Else
            MsgBox "No existe"
        End If
    Else
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
End Sub

Public Sub GrabarInv()
  RatonReloj
  Nuevo = False
  'TextoValido TextPVP, True
  'CampoBusqueda = DGBusq.Columns(DGBusq.Col).Caption
  Codigo = UCase(CambioCodigoCta(MBoxCodigo))
  Codigo1 = Codigo
  Cta_Sup = CodigoCuentaSup(Codigo)
  Cuenta = Codigo & " - " & TextSubCta.Text
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoInv, sSQL
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       TextoBusqueda = "Codigo_Inv Like '" & Codigo & "' "
      .Find (TextoBusqueda)
       If .EOF Then
           SetAddNew AdoInv
           Nuevo = True
       End If
   Else
      SetAddNew AdoInv
      Nuevo = True
   End If
   'MsgBox Nuevo & vbCrLf & Codigo
   SetFields AdoInv, "Codigo", Codigo
   SetFields AdoInv, "Descripcion", TextSubCta.Text
   'SetFields AdoInv, "TC", Cadena
   SetFields AdoInv, "Item", NumEmpresa
   SetFields AdoInv, "Cta_Inventario", "0"
   SetFields AdoInv, "Cta_Costo_Venta", "0"
   SetFields AdoInv, "Cta_Ventas", "0"
   SetFields AdoInv, "Cta_Venta_Anticipada", "0"
   If Len(Codigo) > 1 Then
      SetFields AdoInv, "INV", Si_No
'''      SetFields AdoInv, "Cta_Inventario", CambioCodigoCta(MBoxCta_Inv.Text)
'''      SetFields AdoInv, "Cta_Costo_Venta", CambioCodigoCta(MBoxCta1.Text)
'''      SetFields AdoInv, "Cta_Ventas", CambioCodigoCta(MBoxCta_Ing.Text)
'''      SetFields AdoInv, "Cta_Ventas_0", CambioCodigoCta(MBoxCta_Ing0.Text)
   End If
   SetUpdate AdoInv
   If Nuevo Then
      Codigo2 = Codigo
      Codigo = Codigo1
      AddNewCtaInv
      Codigo = Codigo2
   Else
      IE = TVCatalogo.SelectedItem.Index
      TVCatalogo.Nodes(IE).Text = Codigo & " - " & TextSubCta.Text
      TVCatalogo.Refresh
   End If
  End With
  RatonNormal
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
     Cuenta = SinEspaciosDer(TVCatalogo.SelectedItem)
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo_Inv = '" & Codigo & "' "
     SelectAdodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        MsgBox "No se puede eliminar este codigo: " & Codigo & vbCrLf _
               & "Detalle: " & Cuenta & vbCrLf _
               & "existen datos procesados"
     Else
        Mensajes = "Seguro de Eliminar el Codigo:" & Codigo & vbCrLf _
                 & "?"
        Titulo = "ELIMINACION"
        If BoxMensaje = vbYes Then
           sSQL = "DELETE * " _
                & "FROM Catalogo_Productos " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Codigo_Inv = '" & Codigo & "' "
           ConectarAdoExecute sSQL
           TVCatalogo.Nodes.Remove TVCatalogo.SelectedItem.Index
        End If
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdodc AdoInv, True, 2, 8
  If CtrlDown And KeyCode = vbKeyU Then
  
     Unload IngPlanCreditos
  End If
End Sub

Private Sub TVCatalogo_LostFocus()
  LlenarInv
End Sub

