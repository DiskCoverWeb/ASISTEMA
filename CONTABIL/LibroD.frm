VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form LibroDiario 
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "&Excel"
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
      Left            =   210
      Picture         =   "LibroD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3570
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Eliminar Comproban. Incompletos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10500
      Picture         =   "LibroD.frx":0C42
      TabIndex        =   38
      Top             =   7980
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoDiario 
      Height          =   330
      Left            =   105
      Top             =   7560
      Width           =   4845
      _ExtentX        =   8546
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
      Caption         =   "Diario"
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
      Caption         =   "I&mprimir"
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
      Left            =   210
      Picture         =   "LibroD.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4410
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
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
      Left            =   210
      Picture         =   "LibroD.frx":1906
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1890
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Autorizar"
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
      Left            =   210
      Picture         =   "LibroD.frx":1D48
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5355
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Left            =   210
      Picture         =   "LibroD.frx":218A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2730
      Width           =   960
   End
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
      Left            =   210
      Picture         =   "LibroD.frx":2A54
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6195
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   105
      TabIndex        =   36
      Top             =   1470
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   10583
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DIARIO GENERAL"
      TabPicture(0)   =   "LibroD.frx":331E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGDiario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SUB MODULOS"
      TabPicture(1)   =   "LibroD.frx":333A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGSubCtas"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGDiario 
         Bindings        =   "LibroD.frx":3356
         Height          =   4320
         Left            =   1155
         TabIndex        =   39
         Top             =   420
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7620
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
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
            Name            =   "Arial"
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
      Begin MSDataGridLib.DataGrid DGSubCtas 
         Bindings        =   "LibroD.frx":336E
         Height          =   4320
         Left            =   -73845
         TabIndex        =   37
         Top             =   420
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7620
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
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
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
   End
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "LibroD.frx":3387
      DataSource      =   "AdoAgencias"
      Height          =   345
      Left            =   7455
      TabIndex        =   18
      Top             =   1050
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "LibroD.frx":33A1
      DataSource      =   "AdoUsuario"
      Height          =   345
      Left            =   1575
      TabIndex        =   16
      Top             =   1050
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   315
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
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por &Usuario:"
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
      TabIndex        =   15
      Top             =   1050
      Width           =   1380
   End
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "Agencia:"
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
      Left            =   6300
      TabIndex        =   17
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "COMPROBANTES DE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2205
      TabIndex        =   4
      Top             =   105
      Width           =   9255
      Begin VB.TextBox TextNumNo1 
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
         TabIndex        =   14
         Text            =   "0"
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.OptionButton OpcNC 
         Caption         =   "Notas de &Crédito"
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
         Left            =   2100
         TabIndex        =   9
         Top             =   525
         Width           =   1800
      End
      Begin VB.OptionButton OpcND 
         Caption         =   "Notas de &Débito"
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
         Top             =   525
         Width           =   1800
      End
      Begin VB.TextBox TextNumNo 
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
         TabIndex        =   13
         Text            =   "0"
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CheckBox CheckNum 
         Caption         =   "Desde el &No."
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
         Left            =   5355
         TabIndex        =   12
         Top             =   525
         Width           =   1485
      End
      Begin VB.OptionButton OpcA 
         Caption         =   "&Anulados"
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
         Left            =   5355
         TabIndex        =   11
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcT 
         Caption         =   "&Todos"
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
         Left            =   4095
         TabIndex        =   10
         Top             =   525
         Width           =   960
      End
      Begin VB.OptionButton OpcCD 
         Caption         =   "&Diario"
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
         Left            =   4095
         TabIndex        =   7
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcCE 
         Caption         =   "&Egreso"
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
         Left            =   2100
         TabIndex        =   6
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton OpcCI 
         Caption         =   "&Ingreso"
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
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   630
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   210
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   315
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
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
      Left            =   315
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   315
      Top             =   3780
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
   Begin MSAdodcLib.Adodc AdoSubCtas 
      Height          =   330
      Left            =   315
      Top             =   4095
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
      Caption         =   "SubCtas"
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
   Begin MSAdodcLib.Adodc AdoConceptos 
      Height          =   330
      Left            =   315
      Top             =   4410
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
      Caption         =   "Conceptos"
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
   Begin VB.Label LabelTotHaberME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   4935
      TabIndex        =   33
      Top             =   8295
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Haber ME"
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
      TabIndex        =   32
      Top             =   8295
      Width           =   1485
   End
   Begin VB.Label LabelTotHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   4935
      TabIndex        =   26
      Top             =   7980
      Width           =   1905
   End
   Begin VB.Label LabelTotDebeME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   1575
      TabIndex        =   31
      Top             =   8295
      Width           =   1905
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Debe ME"
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
      TabIndex        =   30
      Top             =   8295
      Width           =   1485
   End
   Begin VB.Label LabelTotDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   1575
      TabIndex        =   28
      Top             =   7980
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Debe"
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
      Top             =   7980
      Width           =   1485
   End
   Begin VB.Label LabelTotSaldoME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   8505
      TabIndex        =   35
      Top             =   8295
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe - Haber ME"
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
      TabIndex        =   34
      Top             =   8295
      Width           =   1695
   End
   Begin VB.Label LabelTotSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   8505
      TabIndex        =   24
      Top             =   7980
      Width           =   1905
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe - Haber:"
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
      TabIndex        =   25
      Top             =   7980
      Width           =   1695
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Haber:"
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
      TabIndex        =   27
      Top             =   7980
      Width           =   1485
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
      Left            =   105
      TabIndex        =   2
      Top             =   630
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
      TabIndex        =   0
      Top             =   210
      Width           =   750
   End
End
Attribute VB_Name = "LibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 5
  Progreso_Barra.Mensaje_Box = "Consultando Diario General"
  Progreso_Esperar
  RatonReloj
  NumItem = 0
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)
  sSQL = "SELECT T.Fecha,T.TP,T.Numero,CL.Cliente As Beneficiario,Co.Concepto,T.Cta,C.Cuenta," _
       & "T.Parcial_ME,T.Debe,T.Haber,T.Detalle,Ac.Nombre_Completo,Co.CodigoU,Co.Autorizado,T.Item,T.ID " _
       & "FROM Transacciones As T,Catalogo_Cuentas As C,Comprobantes As Co,Clientes As CL,Accesos As Ac " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  If OpcCI.value Then
     sSQL = sSQL & "AND T.TP = '" & CompIngreso & "' "
  ElseIf OpcCE.value Then
     sSQL = sSQL & "AND T.TP = '" & CompEgreso & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND T.TP = '" & CompDiario & "' "
  ElseIf OpcND.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaDebito & "' "
  ElseIf OpcNC.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaCredito & "' "
  End If
  If OpcA.value Then
     sSQL = sSQL & "AND T.T = '" & Anulado & "' "
  Else
     sSQL = sSQL & "AND T.T = '" & Normal & "' "
  End If
  If CheckAgencia.value = 1 Then
     sSQL = sSQL & "AND Co.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     If Not ConSucursal Then sSQL = sSQL & "AND Co.Item = '" & NumEmpresa & "' "
  End If
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND Co.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
  If CheckNum.value = 1 Then sSQL = sSQL & "AND Co.Numero BETWEEN " & CLng(TextNumNo.Text) & " and " & CLng(TextNumNo1.Text) & " "
  sSQL = sSQL _
       & "AND T.Item = Co.Item " _
       & "AND T.Item = C.Item " _
       & "AND C.Item = Co.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Periodo = Co.Periodo " _
       & "AND C.Periodo = Co.Periodo " _
       & "AND T.Cta = C.Codigo " _
       & "AND T.TP = Co.TP " _
       & "AND T.Numero = Co.Numero " _
       & "AND T.Fecha = Co.Fecha " _
       & "AND Co.Codigo_B = CL.Codigo " _
       & "AND Co.CodigoU = Ac.Codigo " _
       & "ORDER BY T.Fecha,T.TP,T.Numero,T.ID "
 'MsgBox sSQL
  Select_Adodc_Grid DGDiario, AdoDiario, sSQL, , , , "Diario_General"
  
  sSQLTotales = "SELECT T.Fecha,SUM(T.Parcial_ME) As TParcial_ME, SUM(T.Debe) As TDebe, SUM(T.Haber) As THaber " _
              & "FROM Transacciones As T,Catalogo_Cuentas As C,Comprobantes As Co,Clientes As CL,Accesos As Ac " _
              & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
              & "AND T.Periodo = '" & Periodo_Contable & "' "
  If OpcCI.value Then
     sSQLTotales = sSQLTotales & "AND T.TP = '" & CompIngreso & "' "
  ElseIf OpcCE.value Then
     sSQLTotales = sSQLTotales & "AND T.TP = '" & CompEgreso & "' "
  ElseIf OpcCD.value Then
     sSQLTotales = sSQLTotales & "AND T.TP = '" & CompDiario & "' "
  ElseIf OpcND.value Then
     sSQLTotales = sSQLTotales & "AND T.TP = '" & CompNotaDebito & "' "
  ElseIf OpcNC.value Then
     sSQLTotales = sSQLTotales & "AND T.TP = '" & CompNotaCredito & "' "
  End If
  If OpcA.value Then
     sSQLTotales = sSQLTotales & "AND T.T = '" & Anulado & "' "
  Else
     sSQLTotales = sSQLTotales & "AND T.T = '" & Normal & "' "
  End If
  If CheckAgencia.value = 1 Then
     sSQLTotales = sSQLTotales & "AND Co.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     If Not ConSucursal Then sSQLTotales = sSQLTotales & "AND Co.Item = '" & NumEmpresa & "' "
  End If
  If CheckUsuario.value = 1 Then sSQLTotales = sSQLTotales & "AND Co.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
  If CheckNum.value = 1 Then sSQLTotales = sSQLTotales & "AND Co.Numero BETWEEN " & CLng(TextNumNo.Text) & " and " & CLng(TextNumNo1.Text) & " "
  sSQLTotales = sSQLTotales _
              & "AND T.Item = Co.Item " _
              & "AND T.Item = C.Item " _
              & "AND C.Item = Co.Item " _
              & "AND T.Periodo = C.Periodo " _
              & "AND T.Periodo = Co.Periodo " _
              & "AND C.Periodo = Co.Periodo " _
              & "AND T.Cta = C.Codigo " _
              & "AND T.TP = Co.TP " _
              & "AND T.Numero = Co.Numero " _
              & "AND T.Fecha = Co.Fecha " _
              & "AND Co.Codigo_B = CL.Codigo " _
              & "AND Co.CodigoU = Ac.Codigo " _
              & "GROUP BY T.Fecha "
  Select_Adodc AdoTrans, sSQLTotales
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  DGDiario.Visible = False
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      'MsgBox .RecordCount
       RatonReloj
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          Debe = Debe + .fields("TDebe")
          Haber = Haber + .fields("THaber")
          If .fields("TParcial_ME") > 0 Then
              Debe_ME = Debe_ME + .fields("TParcial_ME")
          Else
              Haber_ME = Haber_ME + (-.fields("TParcial_ME"))
          End If
          Progreso_Barra.Mensaje_Box = "Consultando Diario General " & .fields("Fecha")
          Progreso_Esperar
          'MsgBox "..."
         .MoveNext
       Loop
       RatonNormal
      .MoveFirst
   End If
  End With
  
' SubModulos
  Progreso_Barra.Mensaje_Box = "Consultando SubModulos del Diario General"
  Progreso_Esperar
  sSQL = "SELECT T.Fecha,T.TP,T.Numero,C.Cliente,T.Cta,T.TC,T.Factura,T.Debitos,T.Creditos,T.Prima,T.Codigo " _
       & "FROM Trans_SubCtas As T,Clientes As C " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  If CheckAgencia.value = 1 Then
     sSQL = sSQL & "AND T.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  End If
  If OpcCI.value Then
     sSQL = sSQL & "AND T.TP = '" & CompIngreso & "' "
  ElseIf OpcCE.value Then
     sSQL = sSQL & "AND T.TP = '" & CompEgreso & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND T.TP = '" & CompDiario & "' "
  ElseIf OpcND.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaDebito & "' "
  ElseIf OpcNC.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaCredito & "' "
  End If
  If OpcA.value Then
     sSQL = sSQL & "AND T.T = '" & Anulado & "' "
  Else
     sSQL = sSQL & "AND T.T = '" & Normal & "' "
  End If
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND T.CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckNum.value = 1 Then sSQL = sSQL & "AND T.Numero BETWEEN " & CLng(TextNumNo.Text) & " and " & CLng(TextNumNo1.Text) & " "
  sSQL = sSQL _
       & "AND T.TC IN ('C','P') " _
       & "AND T.Codigo = C.Codigo " _
       & "UNION " _
       & "SELECT T.Fecha,T.TP,T.Numero,Detalle As Cliente,T.Cta,T.TC,T.Factura,T.Debitos,T.Creditos,T.Prima,T.Codigo " _
       & "FROM Trans_SubCtas As T,Catalogo_SubCtas As C " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  If CheckAgencia.value = 1 Then
     sSQL = sSQL & "AND T.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  End If
  If OpcCI.value Then
     sSQL = sSQL & "AND T.TP = '" & CompIngreso & "' "
  ElseIf OpcCE.value Then
     sSQL = sSQL & "AND T.TP = '" & CompEgreso & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND T.TP = '" & CompDiario & "' "
  ElseIf OpcND.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaDebito & "' "
  ElseIf OpcNC.value Then
     sSQL = sSQL & "AND T.TP = '" & CompNotaCredito & "' "
  End If
  If OpcA.value Then
     sSQL = sSQL & "AND T.T = '" & Anulado & "' "
  Else
     sSQL = sSQL & "AND T.T = '" & Normal & "' "
  End If
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND T.CodigoU = '" & SinEspaciosIzq(DCUsuario.Text) & "' "
  If CheckNum.value = 1 Then sSQL = sSQL & "AND T.Numero BETWEEN " & CLng(TextNumNo.Text) & " and " & CLng(TextNumNo1.Text) & " "
  sSQL = sSQL _
       & "AND T.TC NOT IN ('C','P') " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Codigo = C.Codigo " _
       & "ORDER BY T.Fecha,T.TP,T.Numero,T.Cta,T.Factura "
  Select_Adodc_Grid DGSubCtas, AdoSubCtas, sSQL
  Progreso_Esperar
  LabelTotDebe.Caption = Format(Debe, "#,###.00")
  LabelTotHaber.Caption = Format(Haber, "#,###.00")
  LabelTotSaldo.Caption = Format(Debe - Haber, "#,###.00")
  LabelTotDebeME.Caption = Format(Debe_ME, "#,###.00")
  LabelTotHaberME.Caption = Format(Haber_ME, "#,###.00")
  LabelTotSaldoME.Caption = Format(Debe_ME - Haber_ME, "#,###.00")
  LibroDiario.Caption = "DIARIO GENERAL"
  DGDiario.Visible = True
  RatonNormal
  Progreso_Final
End Sub

Private Sub Command2_Click()
  DGDiario.Visible = False
  DGSubCtas.Visible = False
  RatonReloj
  If OpcCoop Then
     Imprimir_Diario_General_Coop AdoDiario
  Else
     Imprimir_Diario_General AdoDiario, AdoSubCtas, AdoConceptos
  End If
  RatonNormal
  DGDiario.Visible = True
  DGSubCtas.Visible = True
End Sub

Private Sub Command3_Click()
  Unload LibroDiario
End Sub

Private Sub Command4_Click()
  RatonReloj
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  sSQL = "UPDATE Comprobantes " _
       & "SET Autorizado = '" & CodigoUsuario & "' " _
       & "WHERE Fecha " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Autorizado = '" & Ninguno & "' "
  Ejecutar_SQL_SP sSQL
  DGDiario.Visible = False
  sSQL = "SELECT T.Fecha,T.TP,T.Numero,Cta,Cuenta,Co.Concepto," _
       & "Parcial_ME,Debe,Haber,CodigoU,Autorizado " _
       & "FROM Transacciones As T,Catalogo_Cuentas As C,Comprobantes As Co " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Periodo = Co.Periodo " _
       & "AND T.T = '" & Normal & "' " _
       & "AND T.Cta = C.Codigo " _
       & "AND T.TP = Co.TP " _
       & "AND T.Numero = Co.Numero " _
       & "AND T.Fecha = Co.Fecha " _
       & "AND T.Item = Co.Item "
  If OpcCI.value Then
     sSQL = sSQL & "AND T.TP = '" & CompIngreso & "' "
  ElseIf OpcCE.value Then
     sSQL = sSQL & "AND T.TP = '" & CompEgreso & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND T.TP = '" & CompDiario & "' "
  End If
  sSQL = sSQL & "ORDER BY T.Fecha,T.TP,T.Numero,T.ID "
  Select_Adodc_Grid DGDiario, AdoTrans, sSQL
  DGDiario.Visible = True
  RatonNormal
End Sub

Private Sub CheckNum_Click()
   SiguienteControl
End Sub

Private Sub CheckNum_LostFocus()
   If CheckNum.value = 1 Then
      TextNumNo.Text = "0"
      TextNumNo.Visible = True
      TextNumNo1.Text = "0"
      TextNumNo1.Visible = True
   Else
      TextNumNo.Visible = False
      TextNumNo1.Visible = False
   End If
End Sub

Private Sub Command5_Click()
  DGDiario.Visible = False
  RatonReloj
  SQLMsg1 = "D I A R I O    G E N E R A L"
  SQLMsg2 = "Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
  ImprimirDiarioGeneralSimple AdoDiario
  RatonNormal
  DGDiario.Visible = True
End Sub

Private Sub Command6_Click()
  RatonReloj
  Actualiza_Comprobantes_Incompletos "Trans_Kardex"
  Actualiza_Comprobantes_Incompletos "Trans_SubCtas"
  Actualiza_Comprobantes_Incompletos "Transacciones"
  'Actualiza_Comprobantes_Incompletos ""
  RatonNormal
  MsgBox "Fin del Proceso"
End Sub

Private Sub DGDiario_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NombreArchivo As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGDiario.Visible = False
     GenerarDataTexto LibroDiario, AdoDiario
     DGDiario.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyA Then
     RatonReloj
     DGDiario.Visible = False
     sSQL = "SELECT T.Numero AS NUMDOCUMENTO, T.Numero AS NUMEROASIENTO, T.Fecha AS FECHAINGRESO, '00:00:00' AS HORACREACION, " _
          & "Ac.Nombre_Completo AS USUARIOREGISTRA, Ac.Nombre_Completo AS USUARIOMODIFICA, T.Fecha AS FECHAMODIFICACION, " _
          & "'00:00:00' AS HORAMODIFICACION, Ac.Nombre_Completo AS USUARIOAPROBO, T.Fecha AS FECHAAPROBACION, '00:00:00' AS HORAAPROBACION, " _
          & "T.Fecha AS FECHACONTABLE, C.Codigo_Ext AS CODIGOCUENTA, T.Debe AS VALORDEBITO, T.Haber AS VALORCREDITO, T.TP AS TIPOASIENTO, " _
          & "Co.T AS Estado, Ac.Nombre_Completo AS CARGOUSUARIO, T.TP AS TIPOTRANSACCION " _
          & "FROM Transacciones As T,Catalogo_Cuentas As C,Comprobantes As Co,Clientes As CL,Accesos As Ac " _
          & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND T.Periodo = '" & Periodo_Contable & "' " _
          & "AND T.Item = '" & NumEmpresa & "' " _
          & "AND T.Item = Co.Item " _
          & "AND T.Item = C.Item " _
          & "AND C.Item = Co.Item " _
          & "AND T.Periodo = C.Periodo " _
          & "AND T.Periodo = Co.Periodo " _
          & "AND C.Periodo = Co.Periodo " _
          & "AND T.Cta = C.Codigo " _
          & "AND T.TP = Co.TP " _
          & "AND T.Numero = Co.Numero " _
          & "AND T.Fecha = Co.Fecha " _
          & "AND Co.Codigo_B = CL.Codigo " _
          & "AND Co.CodigoU = Ac.Codigo " _
          & "ORDER BY T.Fecha,T.TP,T.Numero,T.ID "
     Select_Adodc_Grid DGDiario, AdoDiario, sSQL
     
     NombreArchivo = "DIARIO GENERAL " & Format(MBoxFechaI, "YYYYMMDD") & " - " & Format(MBoxFechaF, "YYYYMMDD") & ".txt"
     GenerarArchivoPlano LibroDiario, AdoDiario, NombreArchivo, True
     DGDiario.Visible = True
     
     RatonNormal
     MsgBox "BUSQUE EL ARCHIVO EN : " & RutaSysBases & "\TEMP\" & NombreArchivo
  End If
  If KeyCode = vbKeyF10 Then
     If ClaveAuxiliar Then
        FechaComp = DGDiario.Columns(0).Text
        TipoComp = DGDiario.Columns(1).Text
        NumComp = DGDiario.Columns(2).Text
        NumItem = NumEmpresa
        NumeroComp = NumComp
        Mensajes = "Seguro que quiere Modificar Comprobante " & TipoComp & " No. " & NumeroComp
        Titulo = "Pregunta de Eliminacion"
        If BoxMensaje = vbYes Then
           CopiarComp = False
           NuevoComp = False
           Trans_No = 1
           IniciarAsientosAdo AdoAsientos
           Unload LibroDiario
           FComprobantes.Show
        End If
     End If
  End If
End Sub

Private Sub DGSubCtas_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyF1 Then
       DGSubCtas.Visible = False
       GenerarDataTexto LibroDiario, AdoSubCtas
       DGSubCtas.Visible = True
    End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     Command2.Enabled = CNivel(1) Or CNivel(2)
     Command4.Enabled = CNivel(1) Or CNivel(2)
     Command5.Enabled = CNivel(1) Or CNivel(2)
  End If
  If NombreUsuario = "Administrador de Red" Then Command6.Visible = True
  Command4.Visible = OpcCoop
  sSQL = "SELECT (Nombre_Completo & '  ' & Codigo) As CodUsuario " _
       & "FROM Accesos " _
       & "WHERE Codigo <> '*' " _
       & "ORDER BY Nombre_Completo "
  SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "CodUsuario", False
  If ConSucursal Then
     sSQL = "SELECT (Item & '  ' & Empresa) As NomEmpresa " _
          & "FROM Empresas " _
          & "WHERE Item IN (" & ListSucursales & ") " _
          & "ORDER BY Item,Empresa "
     SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NomEmpresa"
     CheckAgencia.value = 0
     DCAgencia.Visible = True
     CheckAgencia.Visible = True
  Else
     DCAgencia.Visible = False
     CheckAgencia.Visible = False
  End If
  LibroDiario.Caption = "DIARIO GENERAL"
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoDiario
  ConectarAdodc AdoSubCtas
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoAgencias
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoConceptos
  
   SSTab1.Height = MDI_Y_Max - SSTab1.Top - 1050
   SSTab1.width = MDI_X_Max - SSTab1.Left
   
   DGSubCtas.Left = SSTab1.Left + Command1.width + 100
   DGSubCtas.Height = SSTab1.Height - DGSubCtas.Top - 100
   DGSubCtas.width = SSTab1.width - DGSubCtas.Left - 100
   
   DGDiario.Height = SSTab1.Height - DGDiario.Top - 100
   DGDiario.width = SSTab1.width - DGDiario.Left - 100
   
   AdoDiario.Top = SSTab1.Top + SSTab1.Height + 50
   AdoDiario.width = SSTab1.width - 50
   Command6.Top = AdoDiario.Top + AdoDiario.Height + 50
   Label6.Top = AdoDiario.Top + AdoDiario.Height + 50
   Label9.Top = AdoDiario.Top + AdoDiario.Height + 50
   Label11.Top = AdoDiario.Top + AdoDiario.Height + 50
   
   LabelTotDebe.Top = AdoDiario.Top + AdoDiario.Height + 50
   LabelTotHaber.Top = AdoDiario.Top + AdoDiario.Height + 50
   LabelTotSaldo.Top = AdoDiario.Top + AdoDiario.Height + 50
   
   Label34.Top = Label6.Top + Label6.Height + 20
   Label5.Top = Label6.Top + Label6.Height + 20
   Label7.Top = Label6.Top + Label6.Height + 20

   LabelTotDebeME.Top = Label6.Top + Label6.Height + 50
   LabelTotHaberME.Top = Label6.Top + Label6.Height + 50
   LabelTotSaldoME.Top = Label6.Top + Label6.Height + 50
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub TextNumNo_GotFocus()
  TextNumNo.Text = ""
End Sub

Private Sub TextNumNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextNumNo_LostFocus()
  TextoValido TextNumNo, True
End Sub

Private Sub TextNumNo1_GotFocus()
  TextNumNo1.Text = ""
End Sub

Private Sub TextNumNo1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextNumNo1_LostFocus()
  TextoValido TextNumNo1, True
End Sub


Public Sub Actualiza_Comprobantes_Incompletos(Nombre_Tabla As String)
 'Enceramos Bandera de Verificacion
  sSQL = "UPDATE " & Nombre_Tabla & " " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL
 'Actualizamosd si esta completo el Comprobante
  If SQL_Server Then
     sSQL = "UPDATE " & Nombre_Tabla & " " _
          & "SET X = 'X' " _
          & "FROM " & Nombre_Tabla & " As X,Comprobantes As C "
  Else
     sSQL = "UPDATE " & Nombre_Tabla & " As X,Comprobantes As C " _
          & "SET X.X = 'X' "
  End If
  sSQL = sSQL & "WHERE C.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND X.Item = C.Item " _
       & "AND X.Periodo = C.Periodo " _
       & "AND X.TP = C.TP " _
       & "AND X.Fecha = C.Fecha " _
       & "AND X.Numero = C.Numero "
  Ejecutar_SQL_SP sSQL
 'Eliminacion de los comprobantes Incompletos
  sSQL = "DELETE * " _
       & "FROM " & Nombre_Tabla & " " _
       & "WHERE X = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
End Sub
