VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FlujoDeCaja 
   Caption         =   "Flujo de Caja"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7035
      MaxLength       =   25
      TabIndex        =   9
      Top             =   420
      Width           =   3165
   End
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "Flujcaja.frx":0000
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   1365
      TabIndex        =   43
      Top             =   1260
      Visible         =   0   'False
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
      Bindings        =   "Flujcaja.frx":0014
      DataSource      =   "AdoUsuario"
      Height          =   315
      Left            =   4305
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
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
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   1575
      TabIndex        =   15
      Top             =   0
      Width           =   2220
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
         Height          =   225
         Left            =   1155
         TabIndex        =   3
         Top             =   315
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
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin MSDataListLib.DataCombo DCBancos 
      Bindings        =   "Flujcaja.frx":002D
      DataSource      =   "AdoBancos"
      Height          =   315
      Left            =   5040
      TabIndex        =   45
      Top             =   1260
      Width           =   5160
      _ExtentX        =   9102
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
      Bindings        =   "Flujcaja.frx":0045
      Height          =   2115
      Left            =   105
      TabIndex        =   42
      ToolTipText     =   "<F1> Depositar Cheques."
      Top             =   3990
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3731
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
      Bindings        =   "Flujcaja.frx":0064
      Height          =   1695
      Left            =   105
      TabIndex        =   41
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2990
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
      Top             =   4515
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
      Left            =   10290
      Picture         =   "Flujcaja.frx":0083
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Left            =   10290
      Picture         =   "Flujcaja.frx":038D
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Left            =   10290
      Picture         =   "Flujcaja.frx":0C57
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Left            =   10290
      Picture         =   "Flujcaja.frx":1099
      Style           =   1  'Graphical
      TabIndex        =   39
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
      Left            =   10290
      Picture         =   "Flujcaja.frx":14DB
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1260
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Dep/Ret de Caja"
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
      Picture         =   "Flujcaja.frx":191D
      Style           =   1  'Graphical
      TabIndex        =   10
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
      TabIndex        =   38
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenar por"
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
      Left            =   105
      TabIndex        =   35
      Top             =   735
      Width           =   2535
      Begin VB.OptionButton OpcG 
         Caption         =   "Grup&o"
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
         TabIndex        =   37
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton OpcT 
         Caption         =   "T&ransaccion"
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
         Left            =   1050
         TabIndex        =   36
         Top             =   210
         Width           =   1380
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
      Left            =   2730
      TabIndex        =   34
      Top             =   840
      Width           =   1590
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   3885
      TabIndex        =   22
      Top             =   0
      Width           =   1590
      Begin VB.OptionButton OpcME 
         Caption         =   "ME"
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
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton OpcMN 
         Caption         =   "MN"
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
         Left            =   840
         TabIndex        =   5
         Top             =   315
         Width           =   645
      End
   End
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5565
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Flujcaja.frx":1D5F
      Top             =   420
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc AdoSaldoIni 
      Height          =   330
      Left            =   420
      Top             =   4830
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
      Top             =   2625
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
      Top             =   5145
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
      Top             =   5460
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
      Left            =   2835
      Top             =   4515
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
      Left            =   2835
      Top             =   4830
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
      Left            =   2835
      Top             =   5145
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Detalle de la Transacción"
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
      Left            =   7035
      TabIndex        =   8
      Top             =   105
      Width           =   3165
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " APERTURAS"
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
      TabIndex        =   50
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label LabelAper 
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
      Left            =   1785
      TabIndex        =   49
      Top             =   6720
      Width           =   1905
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CERTIFICADOS"
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
      Left            =   3675
      TabIndex        =   48
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label LabelCert 
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
      Left            =   5250
      TabIndex        =   47
      Top             =   6720
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta a Depositar"
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
      Left            =   2835
      TabIndex        =   46
      Top             =   1260
      Width           =   2220
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
      Left            =   8400
      TabIndex        =   14
      Top             =   3675
      Width           =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<F1> Enviar Depósitos de Cheques al Banco                                             SALDO ANTERIOR  "
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
      Top             =   3675
      Width           =   8310
   End
   Begin VB.Label LabelIngCheqME 
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
      Left            =   8400
      TabIndex        =   32
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESOS CHEQUE M/E "
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
      TabIndex        =   33
      Top             =   3360
      Width           =   3060
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
      Left            =   8400
      TabIndex        =   21
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Label LabelIngCheqMN 
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
      Left            =   3570
      TabIndex        =   30
      Top             =   3360
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
      TabIndex        =   31
      Top             =   3360
      Width           =   3480
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
      Left            =   8400
      TabIndex        =   23
      Top             =   6405
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
      Left            =   7140
      TabIndex        =   24
      Top             =   6405
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
      Left            =   5250
      TabIndex        =   27
      Top             =   6405
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
      Left            =   3675
      TabIndex        =   28
      Top             =   6405
      Width           =   1590
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
      Left            =   1785
      TabIndex        =   25
      Top             =   6405
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
      TabIndex        =   26
      Top             =   6405
      Width           =   1695
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
      Left            =   7140
      TabIndex        =   20
      Top             =   6090
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
      Left            =   5250
      TabIndex        =   17
      Top             =   6090
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
      Left            =   3675
      TabIndex        =   16
      Top             =   6090
      Width           =   1590
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
      Left            =   1785
      TabIndex        =   19
      Top             =   6090
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
      TabIndex        =   18
      Top             =   6090
      Width           =   1695
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
      Left            =   5565
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "FlujoDeCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SumarIngEgr(DtaCajaCheq As Adodc, DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  MontoAper = 0: MontoCert = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total = Total + .Fields("Retiros")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Select Case .Fields("TP")
            Case "BOVE": Debe = Debe + .Fields("Depositos")
                         Haber = Haber + .Fields("Retiros")
            Case "APER", "DEP", "N/CE", "DEPP", "DEFR": Debe = Debe + .Fields("Depositos")
            Case "RET", "CIER": Haber = Haber + .Fields("Retiros")
            Case "N/DC": MontoCert = MontoCert + .Fields("Retiros")
            Case "N/DG": MontoAper = MontoAper + .Fields("Retiros")
          End Select
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngCheqMN.Caption = Format(Total, "#,##0.00")
  LabelIngCheqME.Caption = Format(Saldo, "#,##0.00")
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  LabelAper.Caption = Format(MontoAper, "#,##0.00")
  LabelCert.Caption = Format(MontoCert, "#,##0.00")
  LabelSaldo.Caption = Format(SaldoActual + Debe - Haber, "#,##0.00")
  LabelIngresosME.Caption = Format(Debe_ME, "#,##0.00")
  LabelEgresosME.Caption = Format(Haber_ME, "#,##0.00")
  LabelSaldoME.Caption = Format(Debe_ME - Haber_ME, "#,##0.00")
  RatonNormal
End Sub

Private Sub CheckTP_Click()
   If CheckTP.value = 0 Then DCTP.Visible = False Else DCTP.Visible = True
End Sub

Private Sub CheckUsuario_Click()
  If CheckUsuario.value = 0 Then DCUsuario.Visible = False Else DCUsuario.Visible = True
End Sub

Private Sub Command1_Click()
  ListarFlujoCajas
End Sub

Private Sub Command2_Click()
Mensajes = "Desea Grabar"
Titulo = "Formulario de Grabacion"
If BoxMensaje = vbYes Then
   TextoValido TxtDetalle, , True
  Debe = 0: Haber = 0
  If OpcI.value Then
     Debe = Round(CCur(TextCant.Text), 2)
  Else
     Haber = Round(CCur(TextCant.Text), 2)
  End If
  SetAdoAddNew "Trans_Libretas"
  SetAdoFields "T", Normal
  SetAdoFields "ME", False
  SetAdoFields "Fecha", MBoxFecha
  SetAdoFields "Cuenta_No", "BOVEDA"
  SetAdoFields "TP", "BOVE"
  SetAdoFields "Debitos", Haber
  SetAdoFields "Creditos", Debe
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoFields "Hora", Format(Time, FormatoTimes)
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "Cheque", Ninguno
  SetAdoFields "Banco", TxtDetalle
  SetAdoFields "ACC", CBool(adFalse)
  SetAdoFields "CHT", CBool(adFalse)
  SetAdoUpdate
  Imprimir_Boveda TxtDetalle, Debe, Haber
End If
 ListarFlujoCajas
End Sub

Private Sub Command3_Click()
  Trans_No = 53
  IniciarAsientosAdo AdoAsientos
  ValorDH = 0
  NombreEmpresa = ""
  Mifecha = BuscarFecha(FechaSistema)
  With AdoFlujoCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          DetalleComp = "Cuenta de Ahorro No. " & .Fields("Cuenta_No")
          InsertarAsientos AdoAsientos, SinEspaciosIzq(DCBancos.Text), 0, .Fields("Depositos"), 0
          ValorDH = ValorDH + .Fields("Depositos")
         'NombreEmpresa = NombreEmpresa & .Fields("Cuenta_No")
         .MoveNext
       Loop
   End If
  End With
  ValorDH = Round(ValorDH, 2)
  If ValorDH <> 0 Then
     NumComp = ReadSetDataNum("Diario", True, True)
     DetalleComp = ""
     InsertarAsientos AdoAsientos, Cta_Libretas, 0, 0, ValorDH  'Cta_CajaG
     Co.T = Normal
     Co.TP = CompDiario
     Co.Fecha = MBoxFecha.Text
     Co.Numero = NumComp
     Co.Concepto = "(" & NumEmpresa & ") Depósito en: " & DCBancos.Text
     Co.CodigoB = Ninguno
     Co.Efectivo = 0
     Co.Monto_Total = ValorDH
     Co.T_No = Trans_No
     Co.Usuario = CodigoUsuario
     Co.Item = NumEmpresa
     GrabarComprobante Co
     ImprimirComprobantesDe False, Co
     
     sSQL = "UPDATE Trans_Libretas " _
          & "SET ACC = " & Val(adTrue) & ", AC = " & Val(adTrue) & " " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha) & "# " _
          & "AND TP = 'DEPC' "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Libretas " _
          & "SET ACC = " & Val(adTrue) & ", AC = " & Val(adTrue) & " " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha) & "# " _
          & "AND TP = 'DDAC' "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Libretas " _
          & "SET ACC = " & Val(adTrue) & ", AC = " & Val(adTrue) & " " _
          & "WHERE Fecha = #" & BuscarFecha(MBoxFecha) & "# " _
          & "AND TP = 'APEC' "
     ConectarAdoExecute sSQL
     RatonNormal
     '& "AND ACL = " & Val(adTrue) & " "
  End If
  ListarFlujoCajas
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE FLUJO DE CAJA"
  Mensajes = "Imprimir Resumido"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = vbYes Then
     DGFlujoCajaCheq.Visible = False
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.value, True, CCur(LabelSaldoIni.Caption), True, True
  Else
     DGFlujoCajaCheq.Visible = False
     DGFlujoCajaEfec.Visible = False
     ImprimirFlujoCajaCoop AdoFlujoCajaEfec, True, 1, 9, OpcG.value, True, CCur(LabelSaldoIni.Caption), False, True
  End If
  DGFlujoCajaCheq.Visible = True
  DGFlujoCajaEfec.Visible = True
End Sub

Private Sub Command5_Click()
  Unload FlujoDeCaja
End Sub

Private Sub Command6_Click()
' Si CL es true es movimientos de Caja
  'If MBoxFecha.Text = FechaSistema Then
     sSQL = "DELETE * " _
          & "FROM Trans_Saldo_Libretas " _
          & "WHERE CL <> " & Val(adFalse) & " " _
          & "AND Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
     ConectarAdoExecute sSQL
     SetAdoAddNew "Trans_Saldo_Libretas"
     SetAdoFields "CL", CBool(adTrue)
     SetAdoFields "Fecha", MBoxFecha   'FechaSistema
     SetAdoFields "Saldo_Anterior", CCur(LabelSaldoIni.Caption)
     SetAdoFields "Saldo_Actual", CCur(LabelSaldo.Caption)
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoUpdate
     MsgBox "Proceso Terminado con exito"
  'Else
     'MsgBox "No puede grabar de dias anteriores"
  'End If
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
  Keys_Especiales Shift
  Mifecha = BuscarFecha(MBoxFecha.Text)
  If CtrlDown And KeyCode = vbKeyP Then
     Debe = Val(DGFlujoCajaEfec.Columns(4).Text)
     Haber = Val(DGFlujoCajaEfec.Columns(5).Text)
     'MsgBox Val(DGFlujoCajaEfec.Columns(4).Text)
     Imprimir_Boveda DGFlujoCajaEfec.Columns(6), Debe, Haber
  End If
  If KeyCode = vbKeyF1 Then
     Mensajes = "Seguro de Depositar los Cheques"
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = vbYes Then
        sSQL = "UPDATE Trans_Libretas " _
             & "SET CHT = " & Val(adTrue) & " " _
             & "WHERE Fecha = #" & Mifecha & "# " _
             & "AND TP = 'DEPC' " _
             & "OR TP = 'DDAC' " _
             & "OR TP = 'APEC' "
        ConectarAdoExecute sSQL
        
        ListarFlujoCajas
        
        RatonNormal
     End If
  End If
End Sub

Private Sub Form_Activate()
  Trans_No = 53
  sSQL = "SELECT (Codigo & ' ' & Usuario) As NUsuario " _
       & "FROM Accesos " _
       & "ORDER BY Usuario "
  SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NUsuario"
  
  sSQL = "SELECT Codigo & ' => ' & Cuenta As NombreBanco " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBancos, AdoBancos, sSQL, "NombreBanco", False

  If Supervisor = False Then
     If CNivel(3) Or CNivel(4) Or CNivel(6) Then
        Command2.Enabled = False
        Command3.Enabled = False
        Command6.Enabled = False
     End If
  End If
  ListarFlujoCajas
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
'CentrarForm FlujoDeCaja
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
  FechaValida MBoxFecha
End Sub

Private Sub OpcE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcME_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcMN_KeyDown(KeyCode As Integer, Shift As Integer)
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

Public Sub ListarFlujoCajas()
  FechaValida MBoxFecha
  Mifecha = BuscarFecha(MBoxFecha)
' Depositos de Cheques al Banco
  sSQL = "SELECT ME,Fecha,TP,Papeleta_No,Cuenta_No,Creditos As Depositos,Debitos As Retiros,CodigoU " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Debitos = 0 " _
       & "AND Creditos <> 0 " _
       & "AND CHT <> " & Val(adFalse) & " " _
       & "AND T = 'N' " _
       & "AND ACC = " & Val(adFalse) & " "
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario) & "' "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
  If OpcG.value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Hora,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDataGrid DGFlujoCajaCheq, AdoFlujoCajaCheq, sSQL
 'MsgBox sSQL & vbCrLf & AdoFlujoCajaCheq.Recordset.RecordCount
  
' Flujo de Cajas
  sSQL = "SELECT ME,Fecha,TP,Papeleta_No,Cuenta_No,Creditos As Depositos,Debitos As Retiros,Banco As Detalle,CodigoU " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND T <> 'A' "
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND TP = '" & DCTP.Text & "' "
  sSQL = sSQL & "AND CHT = " & Val(adFalse) & " "
  If OpcG.value Then
     sSQL = sSQL & "ORDER BY TP,Cuenta_No DESC,Fecha,Hora "
  Else
     sSQL = sSQL & "ORDER BY TP,Fecha,Hora "
  End If
  SelectDataGrid DGFlujoCajaEfec, AdoFlujoCajaEfec, sSQL
  
  If CheckTP.value = 1 Then Cta_Aux = DCTP.Text
  sSQL = "SELECT TP " _
       & "FROM Trans_Libretas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  If CheckTP.value = 1 Then DCTP.Text = Cta_Aux
  
  SaldoAnterior = 0: SaldoActual = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Saldo_Libretas " _
       & "WHERE CL = True " _
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
  DGFlujoCajaCheq.Visible = False
  DGFlujoCajaEfec.Visible = False
  SumarIngEgr AdoFlujoCajaCheq, AdoFlujoCajaEfec
  LabelSaldoIni.Caption = Format(SaldoActual, "#,##0.00")
  DGFlujoCajaCheq.Visible = True
  DGFlujoCajaEfec.Visible = True
  RatonNormal
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTexto TxtDetalle
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Imprimir_Boveda(Detalle As String, Ingresos As Currency, Egresos As Currency)
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Comprobante de Pago"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 0.5: InicioY = 0
   Escala_Centimetro Orientacion_Pagina, TipoTimes, 8, True
   Pagina = 1
   EncabezadoEmpresa 0.1
   PrinterPaint LogoTipo, 2, 0.1, 3, 1.5
   Printer.FontName = TipoTimes
   Printer.FontSize = 12: Printer.FontBold = True
   If Ingresos > 0 Then
      PrinterTexto 2, 1.7, "COMPROBANTE DE INGRESO A BOVEDA"
   Else
      PrinterTexto 2, 1.7, "COMPROBANTE DE EGRESO DE BOVEDA"
   End If
   Printer.FontSize = 11
   PrinterTexto 2, 2.5, NombreCiudad & ", " & FechaStrg(FechaSistema)
   PrinterTexto 2, 3.5, "Detalle: "
   PrinterTexto 2, 4.5, "La cantidad de:"
   PrinterTexto 5, 4.5, Moneda
   Printer.FontBold = False
   PrinterTexto 3.5, 3.5, Detalle
   If Ingresos > 0 Then
      PrinterVariables 6, 4.5, Format(Ingresos, "#,##0.00")
   Else
      PrinterVariables 6, 4.5, Format(Egresos, "#,##0.00")
   End If
   PrinterTexto 2.1, 6.5, "Cajero: " & CodigoUsuario
   PrinterTexto 8.3, 6.5, "Aprovación"
   RatonNormal
   Printer.EndDoc
   MensajeEncabData = ""
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub
