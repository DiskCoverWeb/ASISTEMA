VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FlujoDeLibretas 
   Caption         =   "Flujo de Libretas"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGFlujoCajaEfec 
      Bindings        =   "FlujoLib.frx":0000
      Height          =   4950
      Left            =   105
      TabIndex        =   27
      Top             =   1365
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   8731
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "FlujoLib.frx":001F
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   7980
      TabIndex        =   26
      Top             =   525
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "FlujoLib.frx":0033
      DataSource      =   "AdoUsuario"
      Height          =   315
      Left            =   4410
      TabIndex        =   25
      Top             =   525
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   315
      Top             =   1785
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "TP"
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
      Caption         =   "&Grabar Cierre"
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
      Left            =   10080
      Picture         =   "FlujoLib.frx":004C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   1575
      TabIndex        =   20
      Top             =   0
      Width           =   2745
      Begin VB.OptionButton OpcT 
         Caption         =   "Ordenar por Transaccion"
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
         Top             =   525
         Width           =   2535
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "Ordenar por Grupo"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.CheckBox CheckTP 
      Caption         =   "Tipo Proc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7980
      TabIndex        =   19
      Top             =   105
      Width           =   1275
   End
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por Cajero(a):"
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
      Left            =   4410
      TabIndex        =   18
      Top             =   105
      Width           =   3480
   End
   Begin VB.CommandButton Command5 
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
      Left            =   10080
      Picture         =   "FlujoLib.frx":048E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4095
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
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
      Left            =   10080
      Picture         =   "FlujoLib.frx":0710
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3045
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Caja Anterior"
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
      Left            =   10080
      Picture         =   "FlujoLib.frx":0D7A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSAdodcLib.Adodc AdoSaldoIni 
      Height          =   330
      Left            =   315
      Top             =   2100
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "SaldoIni"
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
   Begin MSAdodcLib.Adodc AdoFlujoCajaEfec 
      Height          =   330
      Left            =   315
      Top             =   2415
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "FlujoCajaEfec"
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   315
      Top             =   2730
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Usuario"
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   315
      Top             =   3045
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Caja"
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
   Begin VB.Label LabelSaldoIni 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8085
      TabIndex        =   23
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO ANTERIOR  "
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
      TabIndex        =   24
      Top             =   945
      Width           =   2220
   End
   Begin VB.Label LabelSaldoME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8190
      TabIndex        =   11
      Top             =   6720
      Width           =   1800
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO M/E"
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
      Left            =   6930
      TabIndex        =   12
      Top             =   6720
      Width           =   1275
   End
   Begin VB.Label LabelEgresosME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5040
      TabIndex        =   15
      Top             =   6720
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EGRESOS M/E"
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
      TabIndex        =   16
      Top             =   6720
      Width           =   1485
   End
   Begin VB.Label LabelIngresosME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1680
      TabIndex        =   13
      Top             =   6720
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS M/E"
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
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8190
      TabIndex        =   10
      Top             =   6405
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO M/N"
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
      Left            =   6930
      TabIndex        =   9
      Top             =   6405
      Width           =   1275
   End
   Begin VB.Label LabelEgresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5040
      TabIndex        =   6
      Top             =   6405
      Width           =   1905
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EGRESOS M/N"
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
      TabIndex        =   5
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label LabelIngresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1680
      TabIndex        =   8
      Top             =   6405
      Width           =   1905
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESOS M/N"
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
      Top             =   6405
      Width           =   1590
   End
   Begin VB.Label Label4 
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
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "FlujoDeLibretas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ListarFlujoDeLibretas()
  FechaValida MBoxFecha
  Mifecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT T.ME,T.Fecha,T.TP,T.Papeleta_No,T.Cuenta_No,T.Debitos,T.Creditos,T.CodigoU " _
       & "FROM Trans_Libretas As T,Clientes_Datos_Extras As C " _
       & "WHERE T.Fecha = #" & Mifecha & "# " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND T.T <> 'A' " _
       & "AND T.Cuenta_No = C.Cuenta_No " _
       & "AND T.Item = '" & NumEmpresa & "' "
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND T.CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND T.TP = '" & DCTP.Text & "' "
  If OpcG.value Then
     sSQL = sSQL & "ORDER BY T.ME,T.Fecha,T.TP,Hora,T.Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY T.Fecha,T.Hora "
  End If
  SelectDataGrid DGFlujoCajaEfec, AdoFlujoCajaEfec, sSQL
  sSQL = "SELECT TP " _
       & "FROM Trans_Libretas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  sSQL = "SELECT (Codigo & ' ' & Usuario) As NUsuario " _
       & "FROM Accesos " _
       & "WHERE Usuario <> '" & Ninguno & "' " _
       & "ORDER BY Usuario "
  SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NUsuario", False
  SaldoAnterior = 0: SaldoActual = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Saldo_Libretas " _
       & "WHERE CL = " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Fecha < #" & Mifecha & "# " _
       & "ORDER BY Fecha "
  SelectAdodc AdoSaldoIni, sSQL
  With AdoSaldoIni.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       SaldoAnterior = .Fields("Saldo_Anterior")
       SaldoActual = .Fields("Saldo_Actual")
   End If
  End With
  DGFlujoCajaEfec.Visible = False
  SumarIngEgrL AdoFlujoCajaEfec
  LabelSaldoIni.Caption = Format(SaldoActual, "#,##0.00")
  DGFlujoCajaEfec.Visible = True
  RatonNormal
End Sub

Public Sub SumarIngEgrL(DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("TP") <> "BOVE" Then
                 If .Fields("ME") Then
                     Debe_ME = Debe_ME + .Fields("Debitos")
                     Haber_ME = Haber_ME + .Fields("Creditos")
                 Else
                     Debe = Debe + .Fields("Debitos")
                     Haber = Haber + .Fields("Creditos")
                 End If
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  LabelSaldo.Caption = Format(SaldoActual + Haber - Debe, "#,##0.00")
  LabelIngresosME.Caption = Format(Debe_ME, "#,##0.00")
  LabelEgresosME.Caption = Format(Haber_ME, "#,##0.00")
  LabelSaldoME.Caption = Format(Haber_ME - Debe_ME, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command1_Click()
  ListarFlujoDeLibretas
End Sub

Private Sub Command2_Click()
  'If MBoxFecha.Text = FechaSistema Then
     sSQL = "DELETE * " _
          & "FROM Trans_Saldo_Libretas " _
          & "WHERE CL = " & Val(adFalse) & " " _
          & "AND Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
     ConectarAdoExecute sSQL
     SetAdoAddNew "Trans_Saldo_Libretas"
     SetAdoFields "CL", CBool(adFalse)
     SetAdoFields "Fecha", MBoxFecha.Text   'FechaSistema
     SetAdoFields "Saldo_Anterior", CCur(LabelSaldoIni.Caption)
     SetAdoFields "Saldo_Actual", CCur(LabelSaldo.Caption)
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoUpdate
     MsgBox "Proceso Terminado con exito"
  'Else
  '   MsgBox "No puede grabar de dias anterioeres"
 ' End If
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE DEPOSITOS DE AHORRO"
  Mensajes = "Imprimir Resumido"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = vbYes Then
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.value, False, CCur(LabelSaldoIni.Caption), True
  Else
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.value, False, CCur(LabelSaldoIni.Caption), False
  End If
  DGFlujoCajaEfec.Visible = True
End Sub

Private Sub Command5_Click()
  Unload FlujoDeLibretas
End Sub

Private Sub DCTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel(3) Or CNivel(4) Or CNivel(6) Then Command2.Enabled = False
  End If
  ListarFlujoDeLibretas
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
'CentrarForm FlujoDeLibretas
ConectarAdodc AdoTP
ConectarAdodc AdoCaja
ConectarAdodc AdoUsuario
ConectarAdodc AdoSaldoIni
ConectarAdodc AdoFlujoCajaEfec
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

