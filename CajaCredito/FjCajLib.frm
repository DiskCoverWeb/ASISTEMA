VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FlujoDeCajaLibreta 
   Caption         =   "Flujo de Caja"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11355
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   2415
      TabIndex        =   9
      Top             =   0
      Width           =   5790
      Begin VB.TextBox TextCant 
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
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   3675
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "FjCajLib.frx":0000
         Top             =   210
         Width           =   2010
      End
      Begin VB.OptionButton OpcE 
         Caption         =   "Egreso"
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
         Left            =   1365
         TabIndex        =   3
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "Ingreso"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
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
         Left            =   2520
         TabIndex        =   32
         Top             =   210
         Width           =   1170
      End
   End
   Begin MSDataListLib.DataCombo DCBancos 
      Bindings        =   "FjCajLib.frx":0007
      DataSource      =   "AdoBancos"
      Height          =   315
      Left            =   3150
      TabIndex        =   30
      Top             =   1155
      Width           =   6840
      _ExtentX        =   12065
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
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "FjCajLib.frx":001F
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   1680
      TabIndex        =   28
      Top             =   1155
      Width           =   1380
      _ExtentX        =   2434
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
      Bindings        =   "FjCajLib.frx":0033
      DataSource      =   "AdoUsuario"
      Height          =   315
      Left            =   1680
      TabIndex        =   29
      Top             =   735
      Width           =   4950
      _ExtentX        =   8731
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
   Begin MSDataGridLib.DataGrid DGFlujoCajaEfec 
      Bindings        =   "FjCajLib.frx":004C
      Height          =   3900
      Left            =   105
      TabIndex        =   27
      ToolTipText     =   "<F1> Depositar Cheques."
      Top             =   3150
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6879
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSDataGridLib.DataGrid DGFlujoCajaCheq 
      Bindings        =   "FjCajLib.frx":006B
      Height          =   1275
      Left            =   105
      TabIndex        =   26
      Top             =   1575
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   2249
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   420
      Top             =   4095
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
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":008A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1170
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
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":0394
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4725
      Width           =   1170
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
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":0C5E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3570
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar &Depósito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":10A0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2415
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "G&rabar Cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":14E2
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1260
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
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
      Height          =   1065
      Left            =   10080
      Picture         =   "FjCajLib.frx":1924
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1170
   End
   Begin VB.CheckBox CheckTP 
      Caption         =   "&Tipo Proc"
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
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar por "
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
      Left            =   8295
      TabIndex        =   20
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton OpcG 
         Caption         =   "Gr&upo"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OpcT 
         Caption         =   "&Transaccion"
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
         Top             =   420
         Width           =   1485
      End
   End
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "&Por Cajero(a):"
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
      Top             =   735
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   210
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
      Left            =   420
      Top             =   4410
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
   Begin MSAdodcLib.Adodc AdoFlujoCajaCheq 
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
      Caption         =   "FlujoCajaCheq"
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
      Left            =   420
      Top             =   4725
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   420
      Top             =   5040
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   420
      Top             =   5355
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
   Begin MSAdodcLib.Adodc AdoBancos 
      Height          =   330
      Left            =   420
      Top             =   5670
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
      Caption         =   "Bancos"
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   420
      Top             =   5985
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
      Caption         =   "Asientos"
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
      Left            =   8190
      TabIndex        =   8
      Top             =   2835
      Width           =   1800
   End
   Begin VB.Label Label8 
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
      Left            =   5355
      TabIndex        =   16
      Top             =   2835
      Width           =   2850
   End
   Begin VB.Label LabelSaldo 
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
      Left            =   8190
      TabIndex        =   15
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label LabelIngCheqMN 
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
      Left            =   3570
      TabIndex        =   17
      Top             =   2835
      Width           =   1800
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESOS CHEQUE M/N "
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
      TabIndex        =   18
      Top             =   2835
      Width           =   3480
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
      TabIndex        =   14
      Top             =   7035
      Width           =   1275
   End
   Begin VB.Label LabelEgresos 
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
      Left            =   5040
      TabIndex        =   11
      Top             =   7035
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
      TabIndex        =   10
      Top             =   7035
      Width           =   1485
   End
   Begin VB.Label LabelIngresos 
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
      Left            =   1680
      TabIndex        =   13
      Top             =   7035
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
      TabIndex        =   12
      Top             =   7035
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
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "FlujoDeCajaLibreta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SumarIngEgr(DtaCajaCheq As Adodc, DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total = Total + .Fields("Debitos")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngCheqMN.Caption = Format(Total, "#,##0.00")
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  LabelSaldo.Caption = Format(SaldoActual + Debe - Haber, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command1_Click()
  ListarFlujoCajas
End Sub

Private Sub Command2_Click()
Mensajes = "Desea Grabar"
Titulo = "Formulario de Grabacion"
If BoxMensaje = 6 Then
  Debe = 0: Haber = 0
  If OpcI.Value Then Debe = Round(Val(TextCant.Text), 2) Else Haber = Round(Val(TextCant.Text), 2)
  sSQL = "SELECT * FROM Trans_Cajas "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
      .AddNew
      .Fields("T") = Normal
      .Fields("ME") = False
      .Fields("Fecha") = MBoxFecha.Text
      .Fields("Cuenta_No") = Ninguno
      .Fields("TP") = "BOVE"
      .Fields("Debitos") = 0
      .Fields("Creditos") = 0
      .Fields("Debitos") = Debe
      .Fields("Creditos") = Haber
      .Fields("CodigoU") = CodigoUsuario
      .Fields("Hora") = Format(Time, FormatoTimes)
      .Fields("Item") = NumEmpresa
      .Fields("Cheque") = Ninguno
      .Fields("AC") = False
      .Fields("CHT") = False
      .Update
  End With
End If
  ListarFlujoCajas
End Sub

Private Sub Command3_Click()
  ValorDH = 0
  MiFecha = BuscarFecha(FechaSistema)
  With AdoFlujoCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          ValorDH = ValorDH + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ValorDH = Round(ValorDH, 2)
  If ValorDH <> 0 Then
     Trans_No = 23
     NumComp = ReadSetDataNum("Diario", True, True)
     IniciarAsientosAdo AdoAsientos
     InsertarAsientos AdoAsientos, SinEspaciosIzq(DCBancos.Text), 0, ValorDH, 0
     InsertarAsientos AdoAsientos, Cta_CajaG, 0, 0, ValorDH
     Co.T = Normal
     Co.TP = CompDiario
     Co.Fecha = MBoxFecha.Text
     Co.Numero = NumComp
     Co.Concepto = "Depósito en el Banco: " & DCBancos.Text & ", (" & NombreEmpresa & ")"
     Co.CodigoB = Ninguno
     Co.Efectivo = 0
     Co.Monto_Total = ValorDH
     Co.T_No = Trans_No
     Co.Usuario = CodigoUsuario
     Co.Item = NumEmpresa
     GrabarComprobante Co, AdoAsientos
     ImprimirComprobantesDe False, Co
     sSQL = "UPDATE Trans_Cajas SET AC = True " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# " _
          & "AND TP = 'DEPC' "
     ConectarAdoExecute sSQL
     sSQL = "UPDATE Trans_Cajas SET AC = True " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# " _
          & "AND TP = 'DDAC' "
     ConectarAdoExecute sSQL
     sSQL = "UPDATE Trans_Cajas SET AC = True " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# " _
          & "AND TP = 'APEC' "
     ConectarAdoExecute sSQL
     RatonNormal
  End If
  MiFecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
       & "FROM Trans_Cajas " _
       & "WHERE Fecha = #" & MiFecha & "# " _
       & "AND Debitos <> 0 " _
       & "AND Creditos = 0 " _
       & "AND CHT = True " _
       & "AND T = 'N' " _
       & "AND AC = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDataGrid DGFlujoCajaCheq, AdoFlujoCajaCheq, sSQL
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
       & "FROM Trans_Cajas " _
       & "WHERE Fecha = #" & MiFecha & "# AND T <> 'A' "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
  sSQL = sSQL & "AND CHT = False "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDataGrid DGFlujoCajaEfec, AdoFlujoCajaEfec, sSQL
  sSQL = "SELECT TP " _
       & "FROM Trans_Cajas " _
       & "GROUP BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  
  sSQL = "SELECT (Codigo & ' ' & Usuario) As NUsuario " _
       & "FROM Accesos " _
       & "ORDER BY Usuario "
  SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NUsuario"
  
  sSQL = "SELECT Codigo & ' => ' & Cuenta As NombreBanco " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBancos, AdoBancos, sSQL, "NombreBanco"
  SaldoAnterior = 0: SaldoActual = 0
  sSQL = "SELECT * FROM Saldo_Caja_Libreta " _
       & "WHERE CL = True " _
       & "AND Fecha < #" & MiFecha & "# " _
       & "ORDER BY Fecha "
  SelectAdodc AdoSaldoIni, sSQL
  With AdoSaldoIni.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       SaldoAnterior = .Fields("Saldo_Anterior")
       SaldoActual = .Fields("Saldo_Actual")
   End If
  End With
  SumarIngEgr AdoFlujoCajaCheq, AdoFlujoCajaEfec
  LabelSaldoIni.Caption = Format(SaldoActual, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE FLUJO DE CAJA"
  Mensajes = "Imprimir Resumido"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = 6 Then
     DGFlujoCajaCheq.Visible = False
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.Value, True, CCur(LabelSaldoIni.Caption), True
  Else
     DGFlujoCajaCheq.Visible = False
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.Value, True, CCur(LabelSaldoIni.Caption), False
  End If
  DGFlujoCajaCheq.Visible = True
  DGFlujoCajaEfec.Visible = True
End Sub

Private Sub Command5_Click()
  Unload FlujoDeCajaLibreta
End Sub

Private Sub Command6_Click()
  If MBoxFecha.Text = FechaSistema Then
  sSQL = "DELETE * FROM Saldo_Caja_Libreta " _
       & "WHERE CL = True " _
       & "AND Fecha = #" & BuscarFecha(FechaSistema) & "# "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * FROM Saldo_Caja_Libreta " _
       & "WHERE CL = True "
  SelectAdodc AdoSaldoIni, sSQL
  SaldoAnterior = Round(CDbl(LabelSaldoIni.Caption), 2)
  SaldoActual = Round(CDbl(LabelSaldo.Caption), 2)
  With AdoSaldoIni.Recordset
      .AddNew
      .Fields("CL") = True
      .Fields("Fecha") = FechaSistema
      .Fields("Saldo_Anterior") = SaldoAnterior
      .Fields("Saldo_Actual") = SaldoActual
      .Fields("CodigoU") = CodigoUsuario
      .Fields("Item") = NumEmpresa
      .Update
  End With
  Else
     MsgBox "No puede grabar de dias anterioeres"
  End If
End Sub

Private Sub DCBancos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGFlujoCajaEfec_KeyDown(KeyCode As Integer, Shift As Integer)
  MiFecha = BuscarFecha(MBoxFecha.Text)
  If KeyCode = vbKeyF1 Then
     Mensajes = "Seguro de Depositar los Cheques"
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = 6 Then
        sSQL = "UPDATE Trans_Libretas " _
             & "SET CHT = True " _
             & "WHERE Fecha = #" & MiFecha & "# " _
             & "AND TP = 'DEPC' " _
             & "OR TP = 'DDAC' " _
             & "OR TP = 'APEC' "
        ConectarAdoExecute sSQL
        sSQL = "UPDATE Trans_Cajas " _
             & "SET CHT = True " _
             & "WHERE Fecha = #" & MiFecha & "# " _
             & "AND TP = 'DEPC' OR TP = 'DDAC' OR TP = 'APEC' "
        ConectarAdoExecute sSQL
        sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
             & "FROM Trans_Cajas " _
             & "WHERE Fecha = #" & MiFecha & "# " _
             & "AND Debitos <> 0 " _
             & "AND Creditos = 0 " _
             & "AND CHT = True " _
             & "AND T = 'N' " _
             & "AND AC = False "
        If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DCUsuario.Text & "' "
        If OpcG.Value Then
           sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
        Else
           sSQL = sSQL & "ORDER BY Fecha,Hora "
        End If
        SelectDataGrid DGFlujoCajaCheq, AdoFlujoCajaCheq, sSQL
        sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
             & "FROM Trans_Cajas " _
             & "WHERE Fecha = #" & MiFecha & "# AND T <> 'A' "
        If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DCUsuario.Text & "' "
        If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
        sSQL = sSQL & "AND CHT = False "
        If OpcG.Value Then
           sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
        Else
           sSQL = sSQL & "ORDER BY Fecha,Hora "
        End If
        SelectDataGrid DGFlujoCajaEfec, AdoFlujoCajaEfec, sSQL
        
        sSQL = "SELECT TP FROM Trans_Cajas " _
             & "GROUP BY TP "
        SelectDBCombo DCTP, AdoTP, sSQL, "TP"
        
        sSQL = "SELECT (Codigo & ' ' & Usuario) As NUsuario " _
             & "FROM Accesos " _
             & "ORDER BY Usuario "
        SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NUsuario"
        
        sSQL = "SELECT Codigo & ' => ' & Cuenta As NombreBanco " _
             & "FROM Catalogo_Cuentas " _
             & "WHERE TC = 'BA' " _
             & "ORDER BY Codigo "
       SelectDBCombo DCBancos, AdoBancos, sSQL, "NombreBanco"
       SaldoAnterior = 0: SaldoActual = 0
       sSQL = "SELECT * FROM Saldo_Caja_Libreta " _
            & "WHERE CL = True " _
            & "AND Fecha < #" & MiFecha & "# " _
            & "ORDER BY Fecha "
       SelectAdodc AdoSaldoIni, sSQL
       With AdoSaldoIni.Recordset
        If .RecordCount > 0 Then
           .MoveLast
            SaldoAnterior = .Fields("Saldo_Anterior")
            SaldoActual = .Fields("Saldo_Actual")
        End If
       End With
       SumarIngEgr AdoFlujoCajaCheq, AdoFlujoCajaEfec
       LabelSaldoIni.Caption = Format(SaldoActual, "#,##0.00")
       RatonNormal
     End If
  End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel_3 Or CNivel_4 Or CNivel_6 Then
        Command2.Enabled = False
        Command3.Enabled = False
        Command6.Enabled = False
     End If
  End If
  sSQL = "SELECT TP " _
       & "FROM Tipo_Proceso " _
       & "ORDER BY Nivel,TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  sSQL = "SELECT (Codigo & ' ' & Usuario) As NUsuario " _
       & "FROM Accesos " _
       & "ORDER BY Usuario "
  SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NUsuario"
  sSQL = "SELECT Codigo & ' => ' & Cuenta As NombreBanco " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBancos, AdoBancos, sSQL, "NombreBanco"
  ListarFlujoCajas
  FlujoDeCajaLibreta.WindowState = vbMaximized
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
ConectarAdodc AdoTP
ConectarAdodc AdoCaja
ConectarAdodc AdoBancos
ConectarAdodc AdoUsuario
ConectarAdodc AdoAsientos
ConectarAdodc AdoSaldoIni
ConectarAdodc AdoFlujoCajaCheq
ConectarAdodc AdoFlujoCajaEfec
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
End Sub

Private Sub OpcE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_GotFocus()
  TextCant.Text = ""
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub

Public Sub ListarFlujoCajasLibretas(EsFlujoCaja As Boolean)
  FechaValida MBoxFecha
  MiFecha = BuscarFecha(MBoxFecha.Text)
  If EsFlujoCaja Then
     sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
          & "FROM Trans_Cajas " _
          & "WHERE Fecha = #" & MiFecha & "# " _
          & "AND Debitos <> 0 " _
          & "AND Creditos = 0 " _
          & "AND CHT = True " _
          & "AND T = 'N' " _
          & "AND AC = False "
     If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
     If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
     If OpcG.Value Then
        sSQL = sSQL & "ORDER BY ME,Fecha,TP,Hora,Cuenta_No "
     Else
        sSQL = sSQL & "ORDER BY Fecha,Hora "
     End If
     SelectDataGrid DGFlujoCajaCheq, AdoFlujoCajaCheq, sSQL
  End If
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU "
  If EsFlujoCaja Then
     sSQL = sSQL & "FROM Trans_Cajas "
  Else
     sSQL = sSQL & "FROM Trans_Libretas "
  End If
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# " _
       & "AND T <> 'A' "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
  If EsFlujoCaja Then sSQL = sSQL & "AND CHT = False "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Hora,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDataGrid DGFlujoCajaEfec, AdoFlujoCajaEfec, sSQL
  SaldoAnterior = 0: SaldoActual = 0
  sSQL = "SELECT * FROM Saldo_Caja_Libreta " _
       & "WHERE CL = " & EsFlujoCaja & " " _
       & "AND Fecha < #" & MiFecha & "# " _
       & "ORDER BY Fecha "
  SelectAdodc AdoSaldoIni, sSQL
  With AdoSaldoIni.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       SaldoAnterior = .Fields("Saldo_Anterior")
       SaldoActual = .Fields("Saldo_Actual")
   End If
  End With
  DGFlujoCajaCheq.Visible = False
  DGFlujoCajaEfec.Visible = False
  SumarIngEgr AdoFlujoCajaCheq, AdoFlujoCajaEfec
  LabelSaldoIni.Caption = Format(SaldoActual, "#,##0.00")
  DGFlujoCajaCheq.Visible = True
  DGFlujoCajaEfec.Visible = True
  RatonNormal
End Sub
