VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ListarSuscripciones 
   Caption         =   "LISTADO DE SUSCRIPCIONES"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5985
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Entre&gas"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":09F6
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5250
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Et&iquetas"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":12C0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4515
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&mails"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3780
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "S&uscrip."
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":240C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3045
      Width           =   960
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&2.-Hasta"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":2CD6
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2310
      Width           =   960
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&1.- Desde"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":336C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1575
      Width           =   960
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Detallado"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":3A02
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   840
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "LstSuscr.frx":3E44
      Height          =   3375
      Left            =   105
      TabIndex        =   31
      Top             =   1785
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.PictureBox PictBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   105
      ScaleHeight     =   750
      ScaleWidth      =   1170
      TabIndex        =   38
      Top             =   1785
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Top             =   525
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
      Begin VB.OptionButton OpcListarT 
         Caption         =   "Lista Todas"
         Height          =   210
         Left            =   5775
         TabIndex        =   11
         Top             =   315
         Width           =   1200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Procesar"
         Height          =   855
         Left            =   9555
         Picture         =   "LstSuscr.frx":3E5B
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   210
         Width           =   750
      End
      Begin VB.OptionButton OpcTodas 
         Caption         =   "Ac&tivas"
         Height          =   210
         Left            =   3360
         TabIndex        =   10
         Top             =   630
         Width           =   990
      End
      Begin VB.OptionButton OpcRenov 
         Caption         =   "&Renovaciones"
         Height          =   210
         Left            =   4410
         TabIndex        =   8
         Top             =   630
         Width           =   1395
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "&Terminadas"
         Height          =   210
         Left            =   4410
         TabIndex        =   9
         Top             =   315
         Width           =   1185
      End
      Begin VB.OptionButton OpcSusp 
         Caption         =   "&Suspensas"
         Height          =   210
         Left            =   2205
         TabIndex        =   6
         Top             =   630
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Caption         =   "Por:"
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
         Left            =   7035
         TabIndex        =   19
         Top             =   105
         Width           =   2430
         Begin VB.OptionButton OpcSemestral 
            Caption         =   "Sem&estral"
            Height          =   210
            Left            =   1260
            TabIndex        =   37
            Top             =   630
            Width           =   1080
         End
         Begin VB.OptionButton OpcTrimestral 
            Caption         =   "&Trimestral"
            Height          =   210
            Left            =   1260
            TabIndex        =   36
            Top             =   420
            Width           =   1080
         End
         Begin VB.OptionButton OpcAnual 
            Caption         =   "&Anual"
            Height          =   210
            Left            =   1260
            TabIndex        =   35
            Top             =   210
            Width           =   975
         End
         Begin VB.OptionButton OpcMensual 
            Caption         =   "&Mensual"
            Height          =   210
            Left            =   105
            TabIndex        =   20
            Top             =   210
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OpcQuincenal 
            Caption         =   "&Quincenal"
            Height          =   210
            Left            =   105
            TabIndex        =   21
            Top             =   420
            Width           =   1080
         End
         Begin VB.OptionButton OpcSemanal 
            Caption         =   "Se&manal"
            Height          =   210
            Left            =   105
            TabIndex        =   22
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
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
         Left            =   105
         TabIndex        =   12
         Top             =   840
         Width           =   10200
         Begin MSDataListLib.DataCombo DCCliente 
            Bindings        =   "LstSuscr.frx":4165
            DataSource      =   "AdoCliente"
            Height          =   315
            Left            =   2940
            TabIndex        =   18
            Top             =   420
            Visible         =   0   'False
            Width           =   6420
            _ExtentX        =   11324
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
         Begin VB.CheckBox CheqCliente 
            Caption         =   "Listar Por Cliente"
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
            TabIndex        =   17
            Top             =   105
            Width           =   2640
         End
         Begin VB.OptionButton OpcSec 
            Caption         =   "S&ector"
            Height          =   315
            Left            =   2100
            TabIndex        =   15
            Top             =   105
            Width           =   885
         End
         Begin VB.OptionButton OpcCiud 
            Caption         =   "&Ciudad"
            Height          =   315
            Left            =   1155
            TabIndex        =   14
            Top             =   105
            Width           =   885
         End
         Begin VB.TextBox TxtPatron 
            Height          =   330
            Left            =   105
            TabIndex        =   16
            Top             =   420
            Width           =   2745
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Ninguno"
            Height          =   315
            Left            =   105
            TabIndex        =   13
            Top             =   105
            Value           =   -1  'True
            Width           =   1080
         End
      End
      Begin VB.OptionButton OpcAnul 
         Caption         =   "&Anuladas"
         Height          =   210
         Left            =   2205
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OpcNueva 
         Caption         =   "&Nuevas"
         Height          =   210
         Left            =   3360
         TabIndex        =   7
         Top             =   315
         Width           =   975
      End
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   840
         TabIndex        =   2
         Top             =   210
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
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Hasta:"
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
         TabIndex        =   3
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Desde:"
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
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&onsultar"
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
      Left            =   10605
      Picture         =   "LstSuscr.frx":417E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   105
      Top             =   5250
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
      Caption         =   "Listado de Facturas"
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
   Begin MSAdodcLib.Adodc AdoCiudad 
      Height          =   330
      Left            =   525
      Top             =   2835
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Ciudad"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   525
      Top             =   2520
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoNiveles 
      Height          =   330
      Left            =   525
      Top             =   2205
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Niveles"
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
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   525
      Top             =   3150
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Query1"
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
   Begin MSAdodcLib.Adodc AdoParte 
      Height          =   330
      Left            =   525
      Top             =   3465
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Parte"
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
   Begin MSDataGridLib.DataGrid DGParte 
      Bindings        =   "LstSuscr.frx":45C0
      Height          =   1590
      Left            =   105
      TabIndex        =   33
      Top             =   5670
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   2805
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      Caption         =   "LISTADO DE LLAMADAS"
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Ctrl+F2> Activar|<Ctrl+F3> Anular|<Ctrl+F4> Suspender|<Ctrl+B> Buscar|<Ctrl+P> Ing. Parte"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   2520
      TabIndex        =   32
      Top             =   5250
      Width           =   7995
   End
End
Attribute VB_Name = "ListarSuscripciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqCliente_Click()
  If CheqCliente.value = 1 Then DCCliente.Visible = True Else DCCliente.Visible = False
End Sub

Private Sub Command1_Click()
  Opcion = 0
  ListarSuscriptores Opcion
End Sub

Private Sub Command10_Click()
RatonReloj
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
Codigo = Ninguno
If CheqCliente.value = 1 Then
   With AdoCliente.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente Like '" & DCCliente.Text & "' ")
        If Not .EOF Then
           Codigo = .fields("Codigo")
        Else
           MsgBox "Cliente no Asignado"
        End If
   Else
        MsgBox "No existen datos"
    End If
   End With
End If
TextoValido TxtPatron
If OpcMensual.value Then TipoComp = "MENS"
If OpcQuincenal.value Then TipoComp = "QUNC"
If OpcSemanal.value Then TipoComp = "SEMA"
If OpcAnual.value Then TipoComp = "ANUA"
If OpcTrimestral.value Then TipoComp = "TRIM"
If OpcSemestral.value Then TipoComp = "SEME"
If OpcListarT.value <> True Then
   If OpcNueva.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES NUEVAS"
      TipoDoc = "N"
   End If
   If OpcRenov.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES RENOVACIONES"
      TipoDoc = "R"
   End If
   If OpcAnul.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ANULADAS"
      TipoDoc = "A"
   End If
   If OpcSusp.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES SUSPENSAS"
      TipoDoc = "S"
   End If
   If OpcCanc.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES TERMINADAS"
      TipoDoc = "T"
   End If
   If OpcTodas.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ACTIVAS"
      TipoDoc = "X"
   End If
Else
   SQLMsg1 = "LISTADO DE TODAS LAS SUSCRIPCIONES"
   TipoDoc = "Y"
End If
sSQL = "SELECT F.Factura,F.Total_MN,(F.Total_MN-F.Saldo_MN) AS Abono,F.Saldo_MN,F.Descuento,P.T,P.Fecha As Fecha_I,P.Fecha_C As Fecha_F,P.Credito_No As Contrato_No," _
     & "C.Cliente,P.Atencion,C.Sexo,C.Telefono,C.TelefonoT,C.FAX,C.Pais,C.Prov,C.Ciudad," _
     & "P.Sector,C.Direccion,C.DirNumero,C.CI_RUC,C.Email " _
     & "FROM Clientes As C,Prestamos As P,Facturas As F " _
     & "WHERE P.Item = '" & NumEmpresa & "' " _
     & "AND TC = 'FA' " _
     & "AND P.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND P.TP = '" & TipoComp & "' " _
     & "AND C.Codigo = P.Cuenta_No " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND P.Cuenta_No = F.CodigoC " _
     & "AND P.Numero = F.Factura " _
     & "AND P.Item = F.Item "
If OpcCiud.value Then sSQL = sSQL & "AND C.Ciudad = '" & TxtPatron.Text & "' "
If OpcSec.value Then sSQL = sSQL & "AND P.Sector = '" & TxtPatron.Text & "' "
If CheqCliente.value = 1 Then sSQL = sSQL & "AND C.Codigo = '" & Codigo & "' "
If OpcListarT.value <> True Then
   If TipoDoc <> "X" Then
      sSQL = sSQL & "AND P.T = '" & TipoDoc & "' "
   Else
      sSQL = sSQL & "AND P.T <> 'A' " _
           & "AND P.T <> 'S' " _
           & "AND P.T <> 'T' "
   End If
End If
sSQL = sSQL & "ORDER BY C.Ciudad,P.Sector,C.Cliente "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
DGQuery.Caption = SQLMsg1
DGQuery.Visible = True
RatonNormal
End Sub

Private Sub Command2_Click()
   DGQuery.Visible = False
   ImprimirEtiquetas AdoQuery, PictBar
   If AdoQuery.Recordset.RecordCount > 0 Then AdoQuery.Recordset.MoveFirst
   DGQuery.Visible = True
   MsgBox "Proceso Terminado"
End Sub

Private Sub Command3_Click()
  Unload ListarSuscripciones
End Sub

Private Sub Command4_Click()
   DGQuery.Visible = False
   Mifecha = MBFechaF.Text
   ImprimirSuscriptores AdoQuery
   DGQuery.Visible = True
End Sub

Private Sub Command5_Click()
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CaptionOld As String
Dim NombreFile As String
Dim CadFileReg As String
Dim ContadorReg As Long
Dim TotalCampo As Integer
Dim ValorBool As String
RatonReloj
DGQuery.Visible = False
ContadorReg = 0
If FileResp <= 0 Then FileResp = 1
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     TotalReg = .RecordCount
     TotalCampo = .fields.Count - 1
    .MoveFirst
     NombreFile = "Email " & Format$(Day(FechaSistema), "00") & "_" & Format$(Month(FechaSistema), "00") _
                & "_" & Format$(Year(FechaSistema), "0000") & ".TXT"
     RutaGeneraFile = Left(CurDir$, 2) & "\SYSBASES\EMAILS\" & NombreFile
     NumFile = FreeFile
     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
    .MoveFirst
     Do While Not .EOF
        ContadorReg = ContadorReg + 1
        ListarSuscripciones.Caption = "Creando Archivo de Mails: " _
                            & "(" & Format$(ContadorReg / TotalReg, "##0.00%") _
                            & ") " & String(ContadorReg Mod 40, "|")
        If EsUnEmail(.fields("Email")) Then Print #NumFile, TrimStrg(.fields("Email")) & ";";
       .MoveNext
     Loop
    .MoveFirst
 End If
End With
Close #NumFile
RatonNormal
DGQuery.Visible = True
MsgBox "Se ha procesado un archivo en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Private Sub Command6_Click()
sSQL = "SELECT Item,Contrato_No,Fecha,E,Ent_No " _
     & "FROM Trans_Suscripciones " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND E = " & Val(adFalse) & " " _
     & "AND T <> 'A' " _
     & "AND T <> 'T' " _
     & "AND T <> 'S' " _
     & "ORDER BY Item,Contrato_No,Fecha "
Select_Adodc AdoQuery1, sSQL
RatonReloj
With AdoQuery1.Recordset
 If .RecordCount > 0 Then
     Si_No = True
     Credito_No = .fields("Contrato_No")
     Do While Not .EOF
        If Credito_No <> .fields("Contrato_No") Then
           Credito_No = .fields("Contrato_No")
           Si_No = True
        End If
        If Si_No Then
           Si_No = False
          .fields("E") = adTrue
          .Update
        End If
       .MoveNext
     Loop
 End If
End With
sSQL = "SELECT Contrato_No,COUNT(Contrato_No) As X,0 as Y " _
     & "FROM Trans_Suscripciones As S " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND E = " & Val(adTrue) & " " _
     & "AND T <> 'T' " _
     & "AND T <> 'A' " _
     & "GROUP BY Contrato_No " _
     & "UNION " _
     & "SELECT Contrato_No,0 As X,COUNT(Contrato_No) As Y " _
     & "FROM Trans_Suscripciones As S " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND T <> 'T' " _
     & "AND T <> 'A' " _
     & "GROUP BY Contrato_No " _
     & "ORDER BY Contrato_No,X DESC "
Select_Adodc AdoQuery1, sSQL
Contador = 0
With AdoQuery1.Recordset
 If .RecordCount > 0 Then
     Contrato_No = .fields("Contrato_No")
     I = .fields("X")
     Do While Not .EOF
        If Contrato_No <> .fields("Contrato_No") Then
           If I >= J Then
              sSQL = "UPDATE Prestamos " _
                   & "SET T = 'T',Fecha_C = #" & BuscarFecha(FechaSistema) & "# " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Credito_No = '" & Contrato_No & "' "
              Ejecutar_SQL_SP sSQL
              
              sSQL = "UPDATE Trans_Suscripciones " _
                   & "SET T = 'T' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Contrato_No = '" & Contrato_No & "' "
              Ejecutar_SQL_SP sSQL
              Contador = Contador + 1
           End If
           Contrato_No = .fields("Contrato_No")
           I = .fields("X")
        End If
        J = .fields("Y")
       .MoveNext
     Loop
 End If
End With
RatonNormal
If Contador > 0 Then MsgBox "Hoy se terminaron: " & Contador & " Contratos."
Unload ListarSuscripciones
End Sub

Private Sub Command7_Click()
  TipoConsultaCxC
End Sub

Private Sub Command8_Click()
  Opcion = 1
  ListarSuscriptores Opcion
End Sub

Private Sub Command9_Click()
  Opcion = 2
  ListarSuscriptores Opcion
End Sub

Private Sub DGParte_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto ListarSuscripciones, AdoParte
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  DGQuery.Visible = True
  Contrato_No = Ninguno
  If AdoQuery.Recordset.RecordCount > 0 Then
  Contrato_No = DGQuery.Columns(3)
  If OpcNueva.value Then TipoDoc = "N"
  If OpcRenov.value Then TipoDoc = "R"
  If OpcCanc.value Then TipoDoc = "T"
  If OpcAnul.value Then TipoDoc = "A"
  If OpcTodas.value Then TipoDoc = "X"
  If CtrlDown And KeyCode = vbKeyF2 Then
     Mensajes = "Activar Contrato No. " & Contrato_No
     Titulo = "ACTIVACION"
     If BoxMensaje = vbYes Then
        ActualizaSuscriptores Contrato_No, "R"
        ListarSuscriptores Opcion
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF3 Then
     Mensajes = "Anular Contrato No. " & Contrato_No
     Titulo = "ANULACION"
     If BoxMensaje = vbYes Then
        ActualizaSuscriptores Contrato_No, "A"
        ListarSuscriptores Opcion
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF4 Then
     Mensajes = "Suspender Contrato No. " & Contrato_No
     Titulo = "SUSPENDER"
     If BoxMensaje = vbYes Then
        ActualizaSuscriptores Contrato_No, "S"
        ListarSuscriptores Opcion
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     Cadena = InputBox("Contrato No. " & Contrato_No & ": ", "INGRESO DE PARTE")
     If Cadena <> "" Then
        SetAddNew AdoParte
        SetFields AdoParte, "Contrato_No", Contrato_No
        SetFields AdoParte, "Fecha", FechaSistema
        SetFields AdoParte, "Parte", Cadena
        SetFields AdoParte, "Item", NumEmpresa
        SetUpdate AdoParte
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF11 Then
     Mensajes = "Contrato No. " & Contrato_No & vbCrLf & "Contador:"
     Titulo = "ACTUALIZACION"
     Contador = Val(InputBox(Mensajes, Titulo, 0))
     If Contador <> 0 Then
        sSQL = "UPDATE Trans_Suscripciones " _
             & "SET E = " & Val(adTrue) & " " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Ent_No <= " & CInt(Contador) & " " _
             & "AND Contrato_No = '" & Contrato_No & "' "
        Ejecutar_SQL_SP sSQL
        sSQL = "UPDATE Prestamos " _
             & "SET Dia = " & Contador & " " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        Ejecutar_SQL_SP sSQL
        RatonNormal
        MsgBox "Actualizacion exitosa"
        ListarSuscriptores Opcion
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF10 Then
     Mensajes = "Contrato No. " & Contrato_No & vbCrLf & "DD/MM/AAAA:"
     Titulo = "ACTUALIZACION"
     Mifecha = InputBox(Mensajes, Titulo, FechaSistema)
     If Len(Mifecha) = 10 Then
        sSQL = "UPDATE Trans_Suscripciones " _
             & "SET E = " & Val(adTrue) & " " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Fecha <= #" & BuscarFecha(Mifecha) & "# " _
             & "AND Contrato_No = '" & Contrato_No & "' "
        Ejecutar_SQL_SP sSQL
        RatonNormal
        MsgBox "Actualizacion exitosa"
        ListarSuscriptores Opcion
     End If
  End If
  'If CtrlDown And KeyCode = vbKeyB Then Buscar_Datos DGQuery, AdoQuery
  If KeyCode = vbKeyF1 Then GenerarDataTexto ListarSuscripciones, AdoQuery
  End If
  sSQL = "SELECT Contrato_No,Fecha,Parte,Item " _
       & "FROM Trans_Llamadas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Contrato_No = '" & Contrato_No & "' " _
       & "ORDER BY Fecha DESC "
  Select_Adodc_Grid DGParte, AdoParte, sSQL
  DGParte.Caption = "PARTE CONTRATO No. " & Contrato_No
End Sub

Private Sub Form_Activate()
   FechaValida MBFechaI
   FechaValida MBFechaF
   sSQL = "SELECT Codigo,Cliente " _
        & "FROM Clientes " _
        & "WHERE Cliente <> '.' " _
        & "AND FA = " & Val(adTrue) & " " _
        & "ORDER BY Cliente "
   SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
   Opcion = 0
   ListarSuscriptores Opcion
   RatonNormal
End Sub

Private Sub Form_Load()
   'CentrarForm ListarSuscripciones
   ConectarAdodc AdoParte
   ConectarAdodc AdoQuery
   ConectarAdodc AdoQuery1
   ConectarAdodc AdoCliente
   ConectarAdodc AdoCiudad
   ConectarAdodc AdoNiveles
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Public Sub ActualizaSuscriptores(CreditoNo As String, Tipo As String)
  RatonReloj
  sSQL = "UPDATE Trans_Suscripciones " _
       & "SET T = '" & Tipo & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TP = '" & TipoComp & "' " _
       & "AND Contrato_No = '" & CreditoNo & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Prestamos " _
       & "SET T = '" & Tipo & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Credito_No = '" & CreditoNo & "' "
  Ejecutar_SQL_SP sSQL
  RatonNormal
  MsgBox Mensajes & vbCrLf & "Actualizacion exitosa"
End Sub

Public Sub ListarSuscriptores(OpcReport As Integer)
RatonReloj
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
Codigo = Ninguno
If CheqCliente.value = 1 Then
   With AdoCliente.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente Like '" & DCCliente.Text & "' ")
        If Not .EOF Then
           Codigo = .fields("Codigo")
        Else
           MsgBox "Cliente no Asignado"
        End If
   Else
        MsgBox "No existen datos"
    End If
   End With
End If
TextoValido TxtPatron
If OpcMensual.value Then TipoComp = "MENS"
If OpcQuincenal.value Then TipoComp = "QUNC"
If OpcSemanal.value Then TipoComp = "SEMA"
If OpcAnual.value Then TipoComp = "ANUA"
If OpcTrimestral.value Then TipoComp = "TRIM"
If OpcSemestral.value Then TipoComp = "SEME"
If OpcListarT.value <> True Then
   If OpcNueva.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES NUEVAS"
      TipoDoc = "N"
   End If
   If OpcRenov.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES RENOVACIONES"
      TipoDoc = "R"
   End If
   If OpcAnul.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ANULADAS"
      TipoDoc = "A"
   End If
   If OpcSusp.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES SUSPENSAS"
      TipoDoc = "S"
   End If
   If OpcCanc.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES TERMINADAS"
      TipoDoc = "T"
   End If
   If OpcTodas.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ACTIVAS"
      TipoDoc = "X"
   End If
Else
   SQLMsg1 = "LISTADO DE TODAS LAS SUSCRIPCIONES"
   TipoDoc = "Y"
End If
sSQL = "SELECT P.T,P.Fecha As Fecha_I,P.Fecha_C As Fecha_F,P.Credito_No As Contrato_No,C.Cliente,P.Atencion,C.Telefono,C.Ciudad," _
     & "P.Sector,C.Direccion,C.DirNumero,(CSTR(P.Dia) & '/' & CSTR(P.Meses)) As Contador,C.Email," _
     & "' ' As Firma " _
     & "FROM Clientes As C,Prestamos As P " _
     & "WHERE P.Item = '" & NumEmpresa & "' "
Select Case OpcReport
  Case 0: sSQL = sSQL & "AND P.Fecha >= #" & FechaIni & "# AND P.Fecha_C <= #" & FechaFin & "# "
  Case 1: sSQL = sSQL & "AND P.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  Case 2: sSQL = sSQL & "AND P.Fecha_C BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
End Select
sSQL = sSQL & "AND P.TP = '" & TipoComp & "' " _
     & "AND C.Codigo = P.Cuenta_No "
If OpcCiud.value Then sSQL = sSQL & "AND C.Ciudad = '" & TxtPatron.Text & "' "
If OpcSec.value Then sSQL = sSQL & "AND P.Sector = '" & TxtPatron.Text & "' "
If CheqCliente.value = 1 Then sSQL = sSQL & "AND C.Codigo = '" & Codigo & "' "
If OpcListarT.value <> True Then
   If TipoDoc <> "X" Then
      sSQL = sSQL & "AND P.T = '" & TipoDoc & "' "
   Else
      sSQL = sSQL & "AND P.T <> 'A' " _
           & "AND P.T <> 'S' " _
           & "AND P.T <> 'T' "
   End If
End If
sSQL = sSQL & "ORDER BY C.Ciudad,P.Sector,C.Cliente "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
DGQuery.Caption = SQLMsg1
DGQuery.Visible = True
RatonNormal
End Sub

Public Sub TipoConsultaCxC()
  RatonReloj
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
DGQuery.Caption = "LISTADO DE FACTURAS"
If OpcMensual.value Then TipoComp = "MENS"
If OpcQuincenal.value Then TipoComp = "QUNC"
If OpcSemanal.value Then TipoComp = "SEMA"

sSQL = "SELECT S.Contrato_No,COUNT(Contrato_No) As X,0 As Y " _
     & "FROM Trans_Suscripciones As S " _
     & "WHERE S.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND S.E = " & Val(adTrue) & " " _
     & "AND S.TP = '" & TipoComp & "' " _
     & "AND S.Item = '" & NumEmpresa & "' " _
     & "AND S.T <> 'A' " _
     & "AND S.T <> 'T' " _
     & "GROUP BY S.Contrato_No " _
     & "UNION " _
     & "SELECT S.Contrato_No,0 As X,COUNT(Contrato_No) As Y " _
     & "FROM Trans_Suscripciones As S " _
     & "WHERE S.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND S.E = " & Val(adFalse) & " " _
     & "AND S.TP = '" & TipoComp & "' " _
     & "AND S.Item = '" & NumEmpresa & "' " _
     & "AND S.T <> 'A' " _
     & "AND S.T <> 'T' " _
     & "GROUP BY S.Contrato_No " _
     & "ORDER BY S.Contrato_No "
Select_Adodc AdoCliente, sSQL
With AdoCliente.Recordset
 If .RecordCount > 0 Then
     RatonReloj
     Credito_No = .fields("Contrato_No")
     Contador = 0
     IE = 0: JE = 0
     Do While Not .EOF
        If Credito_No <> .fields("Contrato_No") Then
           ListarSuscripciones.Caption = "(" & Format$(Contador / .RecordCount, "00%") & ")" & Credito_No
           KE = IE + JE
           IE = IE + 1
           sSQL = "UPDATE Prestamos " _
                & "SET Meses = " & CInt(KE) & ", " _
                & "Dia = " & CInt(IE) & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Credito_No = '" & Credito_No & "' " _
                & "AND TP = '" & TipoComp & "' "
           Ejecutar_SQL_SP sSQL
           Credito_No = .fields("Contrato_No")
           IE = 0: JE = 0
        End If
        IE = IE + .fields("X")
        JE = JE + .fields("Y")
        Contador = Contador + 1
       .MoveNext
     Loop
     KE = IE + JE
     IE = IE + 1
     sSQL = "UPDATE Prestamos " _
          & "SET Meses = " & CInt(KE) & ", " _
          & "Dia = " & CInt(IE) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Credito_No = '" & Credito_No & "' " _
          & "AND TP = '" & TipoComp & "' "
     Ejecutar_SQL_SP sSQL
 End If
End With
'Actualizamos suscripciones terminadas
ListarSuscripciones.Caption = "Actualizando Suscripciones..."
If SQL_Server Then
   sSQL = "UPDATE Prestamos " _
        & "SET T = TP.T " _
        & "FROM Prestamos As P, Trans_Suscripciones As TP "
Else
   sSQL = "UPDATE Prestamos As P,Trans_Suscripciones As TP " _
        & "SET P.T = TP.T "
End If
sSQL = sSQL & "WHERE TP.Item = '" & NumEmpresa & "' " _
     & "AND TP.Contrato_No = P.Credito_No " _
     & "AND TP.TP = P.TP " _
     & "AND TP.Item = P.Item " _
     & "AND TP.T <> P.T "
Ejecutar_SQL_SP sSQL
ListarSuscriptores Opcion
ListarSuscripciones.Caption = "LISTADO DE SUSCRIPCIONES"
RatonNormal
MsgBox "Proceso Terminado, Puede Consulte"
End Sub

Private Sub TxtPatron_GotFocus()
   MarcarTexto TxtPatron
End Sub

Private Sub TxtPatron_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
