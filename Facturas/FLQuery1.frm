VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCarteraCli 
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FLQuery1.frx":0000
      Height          =   2220
      Left            =   105
      TabIndex        =   18
      Top             =   2940
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   3916
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
   Begin VB.CheckBox CheqSoloPensiones 
      Caption         =   "Solo imprimir los cobros por pensiones"
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
      TabIndex        =   29
      Top             =   5145
      Width           =   11460
   End
   Begin MSDataGridLib.DataGrid DGCobros 
      Bindings        =   "FLQuery1.frx":0017
      Height          =   1065
      Left            =   105
      TabIndex        =   28
      Top             =   5565
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   1879
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FLQuery1.frx":002F
      DataSource      =   "AdoListCtas"
      Height          =   1155
      Left            =   5145
      TabIndex        =   4
      Top             =   315
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   2037
      _Version        =   393216
      Style           =   1
      ForeColor       =   8388608
      Text            =   "Cliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtCIRUC 
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
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1155
      Width           =   4950
   End
   Begin VB.CheckBox CheqCom 
      Caption         =   "Numeros"
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
      Left            =   10290
      TabIndex        =   25
      Top             =   0
      Width           =   1170
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por &Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5145
      TabIndex        =   1
      Top             =   0
      Width           =   5070
   End
   Begin VB.OptionButton OpcBusq 
      Caption         =   "&Por Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Value           =   -1  'True
      Width           =   4965
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
      Height          =   1485
      Left            =   105
      TabIndex        =   5
      Top             =   1470
      Width           =   11460
      Begin VB.CommandButton Command6 
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
         Left            =   10290
         Picture         =   "FLQuery1.frx":0049
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir &Resumido"
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
         Left            =   9240
         Picture         =   "FLQuery1.frx":0A3F
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
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
         Height          =   1065
         Left            =   8190
         Picture         =   "FLQuery1.frx":0D49
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Por Pensiones"
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
         Left            =   7140
         Picture         =   "FLQuery1.frx":1613
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Por &Facturas"
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
         Left            =   6090
         Picture         =   "FLQuery1.frx":191D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Por Clientes"
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
         Left            =   5040
         Picture         =   "FLQuery1.frx":1FB3
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton OpcPend 
         Caption         =   "&Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "&Canceladas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   11
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton OpcAnul 
         Caption         =   "&Anuladas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   12
         Top             =   525
         Width           =   1185
      End
      Begin VB.OptionButton OpcTodas 
         Caption         =   "&Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   13
         Top             =   1155
         Width           =   885
      End
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   840
         TabIndex        =   9
         Top             =   630
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
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   840
         TabIndex        =   7
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
      Begin MSDataListLib.DataCombo DCTipo 
         Bindings        =   "FLQuery1.frx":23F5
         DataSource      =   "AdoTipo"
         Height          =   315
         Left            =   840
         TabIndex        =   26
         Top             =   1050
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Tipo:"
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
         TabIndex        =   27
         Top             =   1050
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
         TabIndex        =   6
         Top             =   210
         Width           =   750
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
         TabIndex        =   8
         Top             =   630
         Width           =   750
      End
   End
   Begin VB.ListBox LstCampos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   105
      TabIndex        =   2
      Top             =   315
      Width           =   4950
   End
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   7245
      Top             =   525
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
      Caption         =   "Tarjetas"
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   7245
      Top             =   840
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
      Caption         =   "ListCtas"
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   5355
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   9240
      Top             =   840
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
      Caption         =   "Cuentas"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   9240
      Top             =   525
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
      BackColor       =   -2147483644
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
      Caption         =   "Creditos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   5355
      Top             =   525
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   105
      Top             =   6720
      Width           =   3900
      _ExtentX        =   6879
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   5355
      Top             =   1155
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Tipo"
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
   Begin MSAdodcLib.Adodc AdoCobros 
      Height          =   330
      Left            =   7245
      Top             =   1155
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
      Caption         =   "Cobros"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado"
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
      TabIndex        =   22
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label LabelFacturado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   6300
      TabIndex        =   21
      Top             =   6720
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total x Cobrar"
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
      Left            =   8085
      TabIndex        =   20
      Top             =   6720
      Width           =   1485
   End
   Begin VB.Label LabelAbonado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   9555
      TabIndex        =   19
      Top             =   6720
      Width           =   1800
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":240B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":2725
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":2A3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":2D59
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":3073
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":338D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":36A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":39C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":3CDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":3FF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":430F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":4629
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":133BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":136D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":139EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":13D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLQuery1.frx":13EBB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FCarteraCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Consultar_Cobros(CodigoC As String)
   If CodigoC = "" Then CodigoC = Ninguno
   sSQL = "SELECT C.Cliente,DF.Fecha,DF.Mes,DF.Total,DF.Producto,DF.Ticket As Año " _
        & "FROM Detalle_Factura As DF,Clientes As C " _
        & "WHERE DF.Fecha <= #" & FechaFin & "# " _
        & "AND DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.CodigoC = '" & CodigoC & "' "
   If CheqSoloPensiones.value <> 0 Then sSQL = sSQL & "AND DF.Mes <> '.' "
   sSQL = sSQL _
        & "AND DF.CodigoC = C.Codigo " _
        & "ORDER BY DF.Fecha,DF.Mes "
   SelectDataGrid DGCobros, AdoCobros, sSQL
End Sub

Public Sub TipoConsultaCxC(Tipo As String)
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)
DGQuery.Caption = "LISTADO DE FACTURAS"
Actualizar_Ejecutivo_Facturas FechaIni, FechaFin, True
RatonReloj
Contador = 0
DGQuery.Visible = False

Contador = 0
If Tipo = "C" Then
   sSQL = "SELECT F.T,C.Cliente,F.Fecha,F.Factura,F.Total_MN,F.Total_ME,F.Saldo_MN,F.Saldo_ME," _
        & "C.CI_RUC,C.Telefono,C.Celular,C.FAX,C.Ciudad,C.Direccion,C.Email,C.Prov,C.Codigo, "
ElseIf Tipo = "F" Then
   sSQL = "SELECT F.T,F.Fecha,F.Factura,C.Cliente,F.Total_MN," _
        & "(F.Total_MN-F.Saldo_MN) As Abono_MN,F.Saldo_MN,C.Telefono,C.Codigo, "
End If
If SQL_Server Then
   sSQL = sSQL & "F.Fecha_V,DATEDIFF(day,F.Fecha,'" & BuscarFecha(MBFechaF) & "') As Dias_De_Mora,"
Else
   sSQL = sSQL & "F.Fecha_V,DATEDIFF('d',F.Fecha,#" & BuscarFecha(MBFechaF) & "#) As Dias_De_Mora,"
End If
sSQL = sSQL & "A.Nombre_Completo As Ejecutivo " _
     & "FROM Facturas As F,Clientes As C,Accesos As A " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND A.Codigo = F.Cod_Ejec " _
     & "AND F.TC <> 'C' " _
     & "AND F.TC <> 'P' "
If OpcBusq.value Then
TextoBusqueda = TxtCIRUC.Text
If TextoBusqueda <> Ninguno Then
   If LstCampos.Text <> "Ninguno" Then
      TipoDatoBusq = AdoListCtas.Recordset.Fields(LstCampos.Text).Type
      If TipoDatoBusq <> 0 Then
         Select Case TipoDatoBusq
           Case TadByte, TadInteger, TadLong, TadSingle, TadDouble, TadBoolean
                TextoBusqueda = "C." & LstCampos.Text & " = " & Val(TextoBusqueda) & " "
           Case TadText, TadMemo
                TextoBusqueda = "Mid(C." & LstCampos.Text & ", 1," & Len(TextoBusqueda) & ") = '" & TextoBusqueda & "' "
           Case Else
                TextoBusqueda = "Mid(C." & LstCampos.Text & ", 1," & Len(TextoBusqueda) & ") = '" & TextoBusqueda & "' "
         End Select
         sSQL = sSQL & "AND " & TextoBusqueda
      End If
   End If
End If
Else
   sSQL = sSQL & "AND F.CodigoC = '" & CodigoCli & "' "
End If
If OpcPend.value Then sSQL = sSQL & "AND F.T = '" & Pendiente & "' "
If OpcAnul.value Then sSQL = sSQL & "AND F.T = '" & Anulado & "' "
If OpcCanc.value Then sSQL = sSQL & "AND F.T = '" & Cancelado & "' "
If OpcTodas.value Then sSQL = sSQL & "AND F.T <> '" & Anulado & "' "
If Tipo = "C" Then sSQL = sSQL & "ORDER BY C.Cliente,F.Fecha,F.Factura "
If Tipo = "F" Then sSQL = sSQL & "ORDER BY F.Factura,F.Fecha "
SelectDataGrid DGQuery, AdoQuery, sSQL
Total = 0: Saldo = 0
DGQuery.Visible = False
FCarteraCli.Caption = "Totalizando"
With AdoQuery.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
  If Tipo = "C" Then Codigo4 = .Fields("Direccion") Else Codigo4 = TxtCIRUC.Text
  Do While Not .EOF
     ' Contador = Contador + 1
     Total = Total + .Fields("Total_MN")
     Saldo = Saldo + .Fields("Saldo_MN")
    .MoveNext
  Loop
  .MoveFirst
  End If
End With
DGQuery.Visible = True
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Saldo, "#,##0.00")
FCarteraCli.Caption = "CARTERA POR CLIENTES"
RatonNormal
End Sub

Public Sub ListarClientes(Optional LlenarCliente As Boolean)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  If LlenarCliente Then sSQL = sSQL & "AND FA <> " & Val(adFalse) & " "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBCombo DCCliente, AdoListCtas, sSQL, "Cliente"
  'Label2.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format(AdoListCtas.Recordset.RecordCount, "000000")
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "'")
       If Not .EOF Then
          CodigoCli = .Fields("Codigo")
          Mifecha = PrimerDiaMes(FechaSistema)
          Dia = Day(Mifecha)
          Mes = Month(Mifecha)
          Anio = Year(Mifecha)
          FechaIni = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anio, "0000")
          FechaFin = FechaSistema
          Total = 0: Saldo = 0: Contador = 1
       Else
          MsgBox "No Existe"
       End If
   Else
     MsgBox "No Existe"
   End If
  End With
End Sub

Private Sub Command1_Click()
  TipoDoc = "F"
  TipoConsultaCxC TipoDoc
  Opcion = 2
End Sub

Private Sub Command2_Click()
  TipoDoc = "C"
  TipoConsultaCxC TipoDoc
  Opcion = 1
End Sub

Private Sub Command3_Click()
Dim ValorMes(20) As Currency
Dim ValorAnio(1990 To 2020) As Currency
Dim ContFacturas As Long
Dim FechaOld As String
RatonReloj
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)
FechaOld = "01/01/" & Format(Year(MBFechaF), "0000")
FechaIniN = Year(MBFechaI) - 7
FechaFinN = Year(MBFechaF)
sSQL = "DELETE * " _
     & "FROM Saldo_Diarios " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND TP = 'PENS' " _
     & "AND CodigoU = '" & CodigoUsuario & "' "
ConectarAdoExecute sSQL
       
For IE = 0 To 19
    ValorMes(IE) = 0
Next IE
For IE = FechaIniN To FechaFinN
    ValorAnio(IE) = 0
Next IE

DGQuery.Caption = "LISTADO DE PENDIENTES"
RatonReloj
Contador = 0
sSQL = "SELECT F.T,F.Fecha,F.Saldo_MN,F.Factura,C.Codigo,C.Grupo,F.Serie,F.Autorizacion " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Fecha <= #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND F.CodigoC = C.Codigo " _
     & "AND NOT F.TC IN ('C','P') "
If Option2.value Then
   sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
Else
   If LstCampos.Text = "Grupo" Then sSQL = sSQL & "AND C.Grupo = '" & TxtCIRUC & "' "
End If
If OpcPend.value Then sSQL = sSQL & "AND F.T = '" & Pendiente & "' "
If OpcAnul.value Then sSQL = sSQL & "AND F.T = '" & Anulado & "' "
If OpcCanc.value Then sSQL = sSQL & "AND F.T = '" & Cancelado & "' "
sSQL = sSQL & "ORDER BY C.Codigo,F.Fecha,F.Factura "
'MsgBox sSQL
SelectAdodc AdoCta, sSQL
Total = 0: Saldo = 0
With AdoCta.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     CodigoCli = .Fields("Codigo")
     TipoDoc = .Fields("T")
     FechaOld = .Fields("Fecha")
     Do While Not .EOF
        If CFechaLong(FechaOld) > CFechaLong(.Fields("Fecha")) Then FechaOld = .Fields("Fecha")
        'MsgBox TipoDoc & vbCrLf & .Fields("Saldo_MN") & vbCrLf & .Fields("Factura") & vbCrLf & .Fields("Fecha")
        If CodigoCli <> .Fields("Codigo") Then
           Total = 0: Saldo = 0
           SetAdoAddNew "Saldo_Diarios"
           SetAdoFields "TP", "PENS"
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoFields "CodigoC", CodigoCli
           SetAdoFields "T", TipoDoc
           For IE = 1 To 12
               If CheqCom.value = 1 Then
                  SetAdoFields UCase(Mid(MesesLetras(IE), 1, 3)) & "_No", ValorMes(IE)
               Else
                  SetAdoFields UCase(Mid(MesesLetras(IE), 1, 3)), ValorMes(IE)
               End If
               Total = Total + ValorMes(IE)
           Next IE
           For IE = FechaIniN To FechaFinN
               SetAdoFields "P_" & Format(IE, "0000"), ValorAnio(IE)
               Total = Total + ValorAnio(IE)
           Next IE
           SetAdoFields "Total", Total
           If Total > 0 Then SetAdoUpdate
           For IE = 0 To 19
               ValorMes(IE) = 0
           Next IE
           For IE = FechaIniN To FechaFinN
               ValorAnio(IE) = 0
           Next IE
           
           CodigoCli = .Fields("Codigo")
           TipoDoc = .Fields("T")
        End If
        IE = Month(.Fields("Fecha"))
        JE = Year(.Fields("Fecha"))
        If CheqCom.value = 1 Then
           ValorMes(IE) = .Fields("Factura")
        Else
           If (Year(MBFechaI) <= JE) And (JE <= Year(MBFechaI)) Then
              ValorMes(IE) = ValorMes(IE) + .Fields("Saldo_MN")
           Else
              ValorAnio(JE) = ValorAnio(JE) + .Fields("Saldo_MN")
           End If
        End If
        FCarteraCli.Caption = Format(Contador / .RecordCount, "00%")
        Contador = Contador + 1
       .MoveNext
  Loop
  End If
End With
Total = 0: Saldo = 0
SetAdoAddNew "Saldo_Diarios"
SetAdoFields "TP", "PENS"
SetAdoFields "Item", NumEmpresa
SetAdoFields "CodigoU", CodigoUsuario
SetAdoFields "CodigoC", CodigoCli
SetAdoFields "T", TipoDoc
For IE = 1 To 12
    SetAdoFields UCase(Mid(MesesLetras(IE), 1, 3)), ValorMes(IE)
    Total = Total + ValorMes(IE)
Next IE
For IE = FechaIniN To FechaFinN
    SetAdoFields "P_" & Format(IE, "0000"), ValorAnio(IE)
    Total = Total + ValorAnio(IE)
Next IE
SetAdoFields "Total", Total
If Total > 0 Then SetAdoUpdate

'MsgBox FechaOld
 sSQL = "SELECT Item,TP"
 For KE = Year(FechaOld) To Year(MBFechaF)
     sSQL = sSQL & ",SUM(P_" & KE & ") As PP_" & KE
 Next KE
 sSQL = sSQL & " FROM Saldo_Diarios " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND TP = 'PENS' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "GROUP BY Item,TP "
 SelectAdodc AdoQuery, sSQL
'Presentamos las columnas
 sSQL = "SELECT C.Cliente,"
 JE = FechaMes(MBFechaI)
 If (FechaAnio(MBFechaF) - FechaAnio(MBFechaI)) > 0 Then
     KE = 13 - FechaMes(MBFechaF)
     KE = KE + JE + 1
 Else
     KE = FechaMes(MBFechaF) - JE
     KE = KE + JE
 End If
'Campos de Años que deben
 With AdoQuery.Recordset
  If .RecordCount > 0 Then
      For IE = 0 To .Fields.Count - 1
          If Mid(.Fields(IE).Name, 1, 3) = "PP_" Then
             If .Fields(IE) > 0 Then sSQL = sSQL & "SD.P_" & Mid(.Fields(IE).Name, 4, 4) & ","
          End If
      Next IE
  End If
 End With
'MsgBox sSQL
 If KE > 12 Then KE = 12
 For IE = 1 To KE
     If CheqCom.value = 1 Then
        sSQL = sSQL & "SD." & UCase(Mid(MesesLetras(JE), 1, 3)) & "_No,"
     Else
        sSQL = sSQL & "SD." & UCase(Mid(MesesLetras(JE), 1, 3)) & ","
     End If
     JE = JE + 1
     If JE > 12 Then JE = 1
 Next IE
 If CheqCom.value = 1 Then
    sSQL = sSQL & "C.CI_RUC As Codigos,C.Direccion,C.Grupo "
 Else
    sSQL = sSQL & "SD.Total,C.CI_RUC As Codigos,C.Direccion,C.Grupo "
 End If
 sSQL = sSQL & "FROM Saldo_Diarios As SD,Clientes As C " _
      & "WHERE SD.Item = '" & NumEmpresa & "' " _
      & "AND SD.TP = 'PENS' " _
      & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
      & "AND SD.CodigoC = C.Codigo " _
      & "ORDER BY C.Grupo,C.Cliente "
'MsgBox sSQL
 Total = 0: Saldo = 0
 If CheqCom.value = 1 Then
    SelectDataGrid DGQuery, AdoQuery, sSQL
 Else
    SelectDataGrid DGQuery, AdoQuery, sSQL, , True
    RatonReloj
    DGQuery.Visible = False
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Saldo = Saldo + .Fields("Total")
           .MoveNext
         Loop
     End If
    End With
    RatonNormal
    DGQuery.Visible = True
 End If
FCarteraCli.Caption = "LISTADO DE CxC Pensiones"
DGQuery.Visible = True
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Saldo, "#,##0.00")
Opcion = 3
RatonNormal
End Sub

Private Sub Command4_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   SQLMsg2 = "Periodo del " & MBFechaI & " al " & MBFechaF
   Mifecha = MBFechaF.Text
   If Opcion = 1 Then
      ImprimirCtasCob AdoQuery, sSQL, True
   Else
      If CheqCom.value = 1 Then
         Imprimir_Pendientes_Facturacion AdoQuery, Opcion
      Else
         Imprimir_Pendientes_Facturacion AdoQuery, Opcion, True
      End If
   End If
   DGQuery.Visible = True
End Sub

Private Sub Command5_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBFechaF.Text
   If Opcion = 1 Then ImprimirResumenCartera AdoQuery, Codigo4
   DGQuery.Visible = True
End Sub

Private Sub Command6_Click()
  Unload FCarteraCli
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyF12 Then
     sSQL = "SELECT * FROM Clientes " _
          & "WHERE Grupo <> '.' " _
          & "ORDER BY Cliente "
     SelectAdodc AdoListCtas, sSQL
     RatonReloj
     With AdoListCtas.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
            .Fields("Cliente") = UCase(CompilarString(.Fields("Cliente")))
            .Update
            .MoveNext
          Loop
      End If
     End With
     RatonNormal
     Unload FClientes
  End If
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente.Text
  TipoDoc = "M"
End Sub

Private Sub DGCobros_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGCobros.Visible = False
     GenerarDataTexto FCarteraCli, AdoCobros
     DGCobros.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     DGCobros.Visible = False
     MensajeEncabData = "COBROS FACTURADOS"
     SQLMsg2 = ""
     SQLMsg3 = ""
     ImprimirAdodc AdoCobros, 1, 8
     DGCobros.Visible = True
  End If
End Sub

Private Sub DGQuery_DblClick()
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
       SQLMsg1 = "CURSO: " & DGQuery.Columns(DGQuery.Columns.Count - 7)
       'MsgBox SQLMsg1
       Consultar_Cobros DGQuery.Columns(DGQuery.Columns.Count - 4)
   End If
  End With
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FCarteraCli, AdoQuery
     DGQuery.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  ListarClientes CliFact
  LstCampos.Clear
  LstCampos.AddItem "Ninguno"
  LstCampos.AddItem "Cliente"
  LstCampos.AddItem "Direccion"
  LstCampos.AddItem "Ciudad"
  LstCampos.AddItem "CI_RUC"
  LstCampos.AddItem "Codigo"
  LstCampos.AddItem "Telefono"
  LstCampos.AddItem "FAX"
  LstCampos.AddItem "Celular"
  LstCampos.AddItem "Prov"
  LstCampos.AddItem "Grupo"
  LstCampos.AddItem "Email"
  LstCampos.Text = "Cliente"
  RatonReloj
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TP = 'PENS' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectAdodc AdoAux, sSQL

  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC DESC "
  SelectDBCombo DCTipo, AdoTipo, sSQL, "TC"
  RatonNormal
  FCarteraCli.WindowState = vbMaximized
  DCCliente.SetFocus
  If Nuevo Then
     TxtApellidosS.Text = NombreCliente
     LblCodigo.Caption = "Ninguno"
     TxtGrupo.Text = NumEmpresa
     TxtApellidosS.SetFocus
  Else
     ListarCuenta DCCliente.Text
     DCCliente.SetFocus
  End If
End Sub

Private Sub Form_Deactivate()
  FCarteraCli.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   CentrarForm FCarteraCli
   FCarteraCli.Caption = "CREACION DEL CLIENTE"
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoTipo
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCobros
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
      
   DGQuery.Height = MDI_Y_Max / 3
   DGQuery.width = MDI_X_Max - 200
   DGCobros.width = MDI_X_Max - 200
  
   CheqSoloPensiones.Top = DGQuery.Top + DGQuery.Height + 3
   DGCobros.Top = DGQuery.Top + DGQuery.Height + 350
   DGCobros.Height = MDI_Y_Max - CheqSoloPensiones.Top - 700
   
   Label2.Top = DGCobros.Top + DGCobros.Height
   Label4.Top = DGCobros.Top + DGCobros.Height
   Label1.Top = DGCobros.Top + DGCobros.Height
   LabelAbonado.Top = DGCobros.Top + DGCobros.Height
   LabelFacturado.Top = DGCobros.Top + DGCobros.Height
   AdoQuery.Top = DGCobros.Top + DGCobros.Height
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

Private Sub TxtCIRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCIRUC_LostFocus()
  RatonReloj
  TextoValido TxtCIRUC, , True
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
       TextoBusqueda = TxtCIRUC.Text
       If TextoBusqueda <> Ninguno Then
          RatonReloj
          If LstCampos.Text <> "Ninguno" Then
          TipoDatoBusq = .Fields(LstCampos.Text).Type
          If TipoDatoBusq <> 0 Then
             Select Case TipoDatoBusq
               Case TadDate, TadDate1
                    TextoBusqueda = LstCampos.Text & " like '" & TextoBusqueda & "' "
               Case TadByte, TadInteger, TadLong, TadSingle, TadDouble, TadBoolean
                    TextoBusqueda = LstCampos.Text & " like " & Val(TextoBusqueda) & " "
               Case TadText, TadMemo
                    TextoBusqueda = LstCampos.Text & " like '" & TextoBusqueda & "*' "
               Case Else
                    TextoBusqueda = LstCampos.Text & " like '" & TextoBusqueda & "*' "
             End Select
            .MoveFirst
            .Find (TextoBusqueda)
             RatonNormal
             If Not .EOF Then
                DCCliente.Text = .Fields("Cliente")
                CodigoCli = .Fields("Codigo")
                SiguienteControl
             Else
                MsgBox "No existe Datos que buscar"
             End If
          End If
          End If
       Else
          MsgBox "No existe Datos que buscar"
       End If
   Else
       MsgBox "No existe Datos que buscar"
   End If
  End With
  RatonNormal
End Sub

