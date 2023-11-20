VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form EstadoResult12Meses 
   Caption         =   "ESTADO DE RESULTADOS POR MES"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "EstRes12.frx":0000
      DataSource      =   "AdoAgencias"
      Height          =   360
      Left            =   4515
      TabIndex        =   15
      Top             =   420
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqMeses 
      Caption         =   "&Por Meses"
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
      TabIndex        =   20
      Top             =   105
      Width           =   2115
   End
   Begin VB.OptionButton OpcER 
      Caption         =   "Estado de Resultado"
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
      Left            =   4560
      TabIndex        =   19
      Top             =   840
      Value           =   -1  'True
      Width           =   2220
   End
   Begin VB.OptionButton OpcES 
      Caption         =   "Estado de Situación"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   840
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
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
      Left            =   210
      Picture         =   "EstRes12.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2205
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5580
      Left            =   105
      TabIndex        =   10
      Top             =   1785
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   9843
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.- RESULTADOS NUMERICOS"
      TabPicture(0)   =   "EstRes12.frx":08E4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "AdoBalanceG"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DGBalanceG"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&2.- RESULTADOS PORCENTUALES"
      TabPicture(1)   =   "EstRes12.frx":0900
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command3"
      Tab(1).Control(1)=   "AdoBalanceP"
      Tab(1).Control(2)=   "DGBalanceP"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DGBalanceG 
         Bindings        =   "EstRes12.frx":091C
         Height          =   4740
         Left            =   1470
         TabIndex        =   16
         Top             =   420
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   8361
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
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir &Reporte"
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
         Left            =   -74895
         Picture         =   "EstRes12.frx":0936
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir Sin Presupuesto"
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
         Left            =   105
         Picture         =   "EstRes12.frx":0C40
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "I&mprimir Con Presupuesto"
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
         Left            =   105
         Picture         =   "EstRes12.frx":14C2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2730
         Width           =   1275
      End
      Begin MSAdodcLib.Adodc AdoBalanceG 
         Height          =   330
         Left            =   1470
         Top             =   5145
         Width           =   9885
         _ExtentX        =   17436
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
         Caption         =   "BalanceG"
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
      Begin MSAdodcLib.Adodc AdoBalanceP 
         Height          =   330
         Left            =   -73530
         Top             =   5145
         Width           =   9885
         _ExtentX        =   17436
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
         Caption         =   "BalanceP"
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
      Begin MSDataGridLib.DataGrid DGBalanceP 
         Bindings        =   "EstRes12.frx":1D8C
         Height          =   4740
         Left            =   -73530
         TabIndex        =   17
         Top             =   420
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   8361
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
   End
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   105
      Top             =   1785
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Agencias"
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
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "No&mbre de Agencia:"
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
      Left            =   4515
      TabIndex        =   11
      Top             =   105
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   4320
      Begin VB.OptionButton OpcSM 
         Caption         =   "P&rocesar sin SubCuentas de Bloque"
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
         TabIndex        =   9
         Top             =   525
         Width           =   3375
      End
      Begin VB.OptionButton OpcT 
         Caption         =   "Pr&ocesar Mayores y Subcuentas de Bloque"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   4005
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Procesar Util./Perd."
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
      Left            =   10185
      Picture         =   "EstRes12.frx":1DA6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2940
      TabIndex        =   4
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   105
      Top             =   2100
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   105
      Top             =   2415
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Trans"
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
   Begin MSAdodcLib.Adodc AdoResultado 
      Height          =   330
      Left            =   105
      Top             =   2730
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Resultado"
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
   Begin MSAdodcLib.Adodc AdoPres 
      Height          =   330
      Left            =   105
      Top             =   3045
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Pres"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "......"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   105
      TabIndex        =   0
      Top             =   1365
      Width           =   11385
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hasta"
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
      Left            =   2205
      TabIndex        =   3
      Top             =   945
      Width           =   750
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Desde"
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
      TabIndex        =   1
      Top             =   945
      Width           =   750
   End
End
Attribute VB_Name = "EstadoResult12Meses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TB As String

Private Sub CheckAgencia_Click()
  If CheckAgencia.value = 1 Then
     DCAgencia.Visible = True
  Else
     DCAgencia.Visible = False
  End If
End Sub

Private Sub Command1_Click()
  DGBalanceG.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS ANALITICOS MENSUALES"
  SQLMsg2 = "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  ImprimirEstResultPresupuesto AdoBalanceG, MBFechaF.Text
  DGBalanceG.Visible = True
End Sub

Private Sub Command2_Click()
  DGBalanceG.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS ANALITICOS MENSUALES"
  SQLMsg2 = "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  ImprimirEstResultAnalitico AdoBalanceG, MBFechaF.Text
  DGBalanceG.Visible = True
End Sub

Private Sub Command3_Click()
  DGBalanceP.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS ANALITICOS MENSUALES Y PORCENTUALES"
  SQLMsg2 = "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  Imprimir12MesesPorc AdoBalanceP, MBFechaI, MBFechaF
  DGBalanceP.Visible = True
End Sub

Private Sub Command4_Click()
  Unload EstadoResult12Meses
End Sub

Private Sub Command5_Click()
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
  Control_Procesos Normal, "Est. Anal. Mens. del " & MBFechaI.Text & "  al " & MBFechaF.Text
  DGBalanceG.Visible = False
  RatonReloj
  NumItemTemp = 0
  If CheckAgencia.value = 1 Then
     NumItemTemp = SinEspaciosIzq(DCAgencia.Text)
  Else
     NumItemTemp = NumEmpresa
  End If
  FechaValida MBFechaI
  FechaValida MBFechaF
 
 'Insertamos las Cuentas y SubModulos del Balance Analitico Mensual
  Generar_Cuentas_Reportes
  
 'Empezamos a sacar los saldos
  Select Case TB
    Case "ER": Label5.Caption = "ESTADO DE RESULTADOS ANALITICOS MENSUALES "
    Case "ES": Label5.Caption = "ESTADO DE SITUACIOB ANALITICOS MENSUALES "
  End Select
  Label5.Caption = Label5.Caption & "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  ProcesarBalance12Meses EstadoResult12Meses, MBFechaI, MBFechaF, OpcT.value, NumItemTemp, TB
  Total_Presupuesto = 0
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumItemTemp & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'E12M' " _
       & "ORDER BY Ln,Cta "
  SelectAdodc AdoBalanceG, sSQL
  With AdoBalanceG.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total_Presupuesto = Total_Presupuesto + .Fields("Presupuesto")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  
  DGBalanceG.Visible = True
  sSQL = "SELECT DG,Cta As Codigo,Comprobante As Cuenta,"
  For NoMeses = Month(MBFechaI.Text) To Month(MBFechaF.Text)
      sSQL = sSQL & MesesLetras(NoMeses) & ","
  Next NoMeses
  sSQL = sSQL & "Total "
  If Total_Presupuesto <> 0 Then sSQL = sSQL & ",Presupuesto,Diferencia "
  sSQL = sSQL & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumItemTemp & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'E12M' "
  If OpcER.value Then
     sSQL = sSQL & "AND TB = 'ER' "
  Else
     sSQL = sSQL & "AND TB = 'ES' "
  End If
  sSQL = sSQL & "ORDER BY Codigo_Aux "
  SelectDataGrid DGBalanceG, AdoBalanceG, sSQL
  Select Case TB
    Case "ER": Label5.Caption = "ESTADO DE RESULTADOS ANALITICOS MENSUALES "
    Case "ES": Label5.Caption = "ESTADO DE SITUACION ANALITICOS MENSUALES "
  End Select
  Label5.Caption = Label5.Caption & "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  
  RatonNormal
End Sub

Private Sub DGBalanceG_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then GenerarDataTexto EstadoResult12Meses, AdoBalanceG
End Sub

Private Sub DGBalanceP_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then GenerarDataTexto EstadoResult12Meses, AdoBalanceP
End Sub

Private Sub Form_Activate()
  If CNivel(7) Then
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload EstadoResult12Meses
     RatonNormal
  Else

  Label5.Caption = "ESTADO DE RESULTADOS ANALITICOS MENSUALES"
  sSQL = "SELECT (Item & '  ' & Empresa) As NombreEmpresa " _
       & "FROM Empresas " _
       & "WHERE Empresa <> '" & Ninguno & "' " _
       & "ORDER BY Item "
  SelectDBCombo DCAgencia, AdoAgencias, sSQL, "NombreEmpresa", False
  RatonNormal
  End If
End Sub

Private Sub Form_Load()
 'CentrarForm EstadoResult12Meses
  ConectarAdodc AdoPres
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoResultado
  ConectarAdodc AdoAgencias
  ConectarAdodc AdoBalanceG
  ConectarAdodc AdoBalanceP
  
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 100
  SSTab1.width = MDI_X_Max - SSTab1.Left
  
  DGBalanceP.width = SSTab1.width - Command4.width - 400
  DGBalanceP.Height = SSTab1.Height - DGBalanceP.Top - 400
  AdoBalanceP.Top = DGBalanceP.Top + DGBalanceP.Height + 5
  AdoBalanceP.width = SSTab1.width - Command4.width - 400
  
  DGBalanceG.width = SSTab1.width - Command4.width - 400
  DGBalanceG.Height = SSTab1.Height - DGBalanceG.Top - 400
  AdoBalanceG.Top = DGBalanceG.Top + DGBalanceG.Height + 5
  AdoBalanceG.width = SSTab1.width - Command4.width - 400
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF, False
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI, False
End Sub

Public Sub Generar_Cuentas_Reportes()
    If OpcES.value Then TB = "ES" Else TB = "ER"
    
    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE TP = 'E12M' " _
         & "AND Item = '" & NumItemTemp & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    ConectarAdoExecute sSQL
    
    sSQL = "INSERT INTO Saldo_Diarios " _
         & "(TB,TP,Cta,Codigo_Aux,CodigoC,Comprobante,CodigoU,TC,DG,Presupuesto,Item,Cta_Aux) " _
         & "SELECT '" & TB & "' As TB1,'E12M' As TP1,Codigo,Codigo As C_A,'.' As C_C,Mid$(Cuenta,1,65)," _
         & "'" & CodigoUsuario & "' As CU,TC,DG,Presupuesto,'" & NumItemTemp & "' As Items,Codigo As Cta1 " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumItemTemp & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    If OpcES.value Then
       sSQL = sSQL & "AND Mid$(Codigo,1,1) BETWEEN '1' AND '3' "
    Else
       sSQL = sSQL & "AND Mid$(Codigo,1,1) BETWEEN '4' AND '6' "
    End If
    ConectarAdoExecute sSQL
    
    sSQL = "INSERT INTO Saldo_Diarios (TB,TP,Cta,Codigo_Aux,CodigoC,Comprobante,CodigoU,TC,DG,Item,Cta_Aux) " _
         & "SELECT '" & TB & "' As TB1,'E12M' As TP1,Cta,Cta + '.' + Codigo As C_A,Codigo,' *' As Cuenta1," _
         & "'" & CodigoUsuario & "' As CU,TC,'D' As DG1,'" & NumItemTemp & "' As Items,Cta As Cta1 " _
         & "FROM Trans_SubCtas " _
         & "WHERE Item = '" & NumItemTemp & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    If OpcES.value Then
       sSQL = sSQL & "AND Mid$(Cta,1,1) BETWEEN '1' AND '3' "
    Else
       sSQL = sSQL & "AND Mid$(Cta,1,1) BETWEEN '4' AND '6' "
    End If
    sSQL = sSQL & "GROUP BY Cta,Codigo,TC "
    ConectarAdoExecute sSQL
    
    
'''  If SQL_Server Then
'''     sSQL = "UPDATE Transacciones " _
'''          & "SET Cheq_Dep = TB.Cheq_Dep " _
'''          & "FROM Transacciones As T,Trans_Bancos As TB "
'''  Else
'''     sSQL = "UPDATE Transacceiones As T,Trans_Bancos As TB " _
'''          & "SET T.Cheq_Dep = TB.Cheq_Dep "
'''  End If
    
    If SQL_Server Then
       sSQL = "UPDATE Saldo_Diarios " _
            & "SET Comprobante = '  * ' + RTRIM(LTRIM(SUBSTRING(C.Cliente,1,60))) " _
            & "FROM Saldo_Diarios As SD, Clientes As C "
    Else
       sSQL = "UPDATE Saldo_Diarios As SD, Clientes As C " _
            & "SET Comprobante = '  * ' + Trim(Mid$(C.Cliente,1,60)) "
    End If
    sSQL = sSQL _
         & "WHERE SD.CodigoC = C.Codigo " _
         & "AND SD.TP = 'E12M' " _
         & "AND SD.CodigoC <> '.' "
    ConectarAdoExecute sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Saldo_Diarios " _
            & "SET Comprobante = '  * ' + RTRIM(LTRIM(SUBSTRING(CS.Detalle ,1,60))) " _
            & "FROM Saldo_Diarios As SD, Catalogo_SubCtas As CS "
    Else
       sSQL = "UPDATE Saldo_Diarios As SD, Catalogo_SubCtas As CS " _
            & "SET Comprobante = '  * ' + Trim(Mid$(CS.Detalle ,1,60)) "
    End If
    sSQL = sSQL _
         & "WHERE SD.CodigoC = CS.Codigo " _
         & "AND SD.TP = 'E12M' " _
         & "AND SD.CodigoC <> '.' "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Saldo_Diarios " _
         & "SET Cta = '  ' " _
         & "WHERE CodigoC <> '.' " _
         & "AND TP = 'E12M' " _
         & "AND Item = '" & NumItemTemp & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    ConectarAdoExecute sSQL
    
                      
    sSQL = "SELECT * " _
        & "FROM Trans_Presupuestos " _
        & "WHERE Item = '" & NumItemTemp & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta = '" & TextoBusqueda & "' " _
        & "AND Codigo = 'Codigo' "

  NoMeses = Round((CFechaLong(MBFechaF) - CFechaLong(MBFechaI)) / 30)
  If NoMeses = 0 Then NoMeses = 1

''  sSQL = "DELETE * " _
''       & "FROM Saldo_Diarios " _
''       & "WHERE Item = '" & NumItemTemp & "' " _
''       & "AND CodigoU = '" & CodigoUsuario & "' " _
''       & "AND TP = 'E12M' "
''  ConectarAdoExecute sSQL
''  sSQL = "SELECT * " _
''       & "FROM Saldo_Diarios " _
''       & "WHERE Item = '" & NumItemTemp & "' " _
''       & "AND CodigoU = '" & CodigoUsuario & "' " _
''       & "AND TP = 'E12M' "
''  SelectAdodc AdoBalanceG, sSQL
''  sSQL = "SELECT TC,DG,Codigo,Cuenta,Presupuesto " _
''       & "FROM Catalogo_Cuentas " _
''       & "WHERE Item = '" & NumItemTemp & "' " _
''       & "AND Periodo = '" & Periodo_Contable & "' "
''  If OpcES.value Then
''     sSQL = sSQL & "AND Mid$(Codigo,1,1) BETWEEN '1' AND '3' "
''  Else
''     sSQL = sSQL & "AND Mid$(Codigo,1,1) BETWEEN '4' AND '6' "
''  End If
''  sSQL = sSQL & "ORDER BY Codigo "
''  SelectAdodc AdoCtas, sSQL
''
''  sSQL = "SELECT TC,Cta,Codigo " _
''       & "FROM Trans_SubCtas " _
''       & "WHERE Item = '" & NumItemTemp & "' " _
''       & "AND Periodo = '" & Periodo_Contable & "' " _
''       & "AND LEN(Codigo) > 1 " _
''       & "GROUP BY TC,Cta,Codigo "
''  SelectAdodc AdoTrans, sSQL
  Ln_No = 0: Contador = 0
''
''  With AdoCtas.Recordset
''   If .RecordCount > 0 Then
''       Do While Not .EOF
''          Ln_No = Ln_No + 1
''          Contador = Contador + 1
''          TextoBusqueda = .Fields("Codigo")
''          EstadoResult12Meses.Caption = "Actualizando Cuentas: (" & Round((Contador / .RecordCount) * 100) & "%) " & TextoBusqueda
''          SetAddNew AdoBalanceG
''          SetFields AdoBalanceG, "TC", .Fields("TC")
''          SetFields AdoBalanceG, "DG", .Fields("DG")
''          SetFields AdoBalanceG, "Cta", TextoBusqueda
''          SetFields AdoBalanceG, "Comprobante", Mid$(.Fields("Cuenta"), 1, 60)
''          SetFields AdoBalanceG, "CodigoC", Ninguno
''          SaldoCont = .Fields("Presupuesto")
''          If NoMeses >= 1 And CheqMeses.value = 1 Then SaldoCont = CCur((SaldoCont / 12) * NoMeses)
''          SetFields AdoBalanceG, "Presupuesto", SaldoCont
''          SetFields AdoBalanceG, "Item", NumItemTemp
''          SetFields AdoBalanceG, "CodigoU", CodigoUsuario
''          SetFields AdoBalanceG, "TP", "E12M"
''          SetFields AdoBalanceG, "Ln", Ln_No
''          Select Case Mid$(TextoBusqueda, 1, 1)
''            Case "1", "2", "3": SetFields AdoBalanceG, "TB", "ES"
''            Case "4", "5", "6": SetFields AdoBalanceG, "TB", "ER"
''            Case Else:          SetFields AdoBalanceG, "TB", "CO"
''          End Select
''          SetUpdate AdoBalanceG
''          If OpcT.value Then
''          If AdoTrans.Recordset.RecordCount > 0 Then
''             AdoTrans.Recordset.MoveFirst
''             AdoTrans.Recordset.Find ("Cta = '" & TextoBusqueda & "'")
''             If Not AdoTrans.Recordset.EOF Then
''                Codigo = AdoTrans.Recordset.Fields("Cta")
''                Do While Not AdoTrans.Recordset.EOF
''                   If Codigo = AdoTrans.Recordset.Fields("Cta") Then
''                      SaldoCont = 0
''                      sSQL = "SELECT * " _
''                           & "FROM Trans_Presupuestos " _
''                           & "WHERE Item = '" & NumItemTemp & "' " _
''                           & "AND Periodo = '" & Periodo_Contable & "' " _
''                           & "AND Cta = '" & TextoBusqueda & "' " _
''                           & "AND Codigo = '" & AdoTrans.Recordset.Fields("Codigo") & "' "
''                      SelectAdodc AdoPres, sSQL
''                      If AdoPres.Recordset.RecordCount > 0 Then
''                         SaldoCont = AdoPres.Recordset.Fields("Presupuesto")
''                         If NoMeses >= 1 And CheqMeses.value = 1 Then SaldoCont = CCur((SaldoCont / 12) * NoMeses)
''                      End If
''                      Ln_No = Ln_No + 1
''                      SetAddNew AdoBalanceG
''                      SetFields AdoBalanceG, "TC", .Fields("TC")
''                      SetFields AdoBalanceG, "DG", .Fields("DG")
''                      SetFields AdoBalanceG, "Cta", TextoBusqueda
''                      SetFields AdoBalanceG, "Comprobante", Mid$(.Fields("Cuenta"), 1, 60)
''                      SetFields AdoBalanceG, "CodigoC", AdoTrans.Recordset.Fields("Codigo")
''                      SetFields AdoBalanceG, "Presupuesto", SaldoCont
''                      SetFields AdoBalanceG, "Item", NumItemTemp
''                      SetFields AdoBalanceG, "CodigoU", CodigoUsuario
''                      SetFields AdoBalanceG, "TP", "E12M"
''                      SetFields AdoBalanceG, "Ln", Ln_No
''                      Select Case Mid$(TextoBusqueda, 1, 1)
''                        Case "1", "2", "3": SetFields AdoBalanceG, "TB", "ES"
''                        Case "4", "5", "6": SetFields AdoBalanceG, "TB", "ER"
''                        Case Else:          SetFields AdoBalanceG, "TB", "CO"
''                      End Select
''                      SetUpdate AdoBalanceG
''                   End If
''                   AdoTrans.Recordset.MoveNext
''                Loop
''             End If
''          End If
''          End If
''         .MoveNext
''       Loop
''   End If
''  End With
    
    SetAdoAddNew "Saldo_Diarios"
    SetAdoFields "TC", "N"
    SetAdoFields "DG", "G"
    SetAdoFields "Cta", "(+/-)"
    SetAdoFields "Codigo_Aux", "9.9.99.99.99.999"
    SetAdoFields "Comprobante", "UTILIDAD O PERDIDA"
    SetAdoFields "CodigoC", Ninguno
    SetAdoFields "Item", NumItemTemp
    SetAdoFields "CodigoU", CodigoUsuario
    SetAdoFields "TP", "E12M"
    If OpcES.value Then
       SetAdoFields "TB", "ES"
    Else
       SetAdoFields "TB", "ER"
    End If
    SetAdoUpdate
    
    Eliminar_Nullos "Saldo_Diarios"
End Sub
