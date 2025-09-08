VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FMatriculados 
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.CheckBox OpcTodasM 
      Caption         =   "Todas las Materias del Curso"
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
      Left            =   1155
      TabIndex        =   23
      Top             =   840
      Width           =   2850
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3060
      Left            =   1155
      TabIndex        =   13
      Top             =   1260
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   5398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CURSOS Y PARALELOS"
      TabPicture(0)   =   "FLCxC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DLGrupo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DCMaterias"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "RESUMEN TOTALES DE CURSOS"
      TabPicture(1)   =   "FLCxC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGTotalCursos"
      Tab(1).ControlCount=   1
      Begin MSDataListLib.DataCombo DCMaterias 
         Bindings        =   "FLCxC.frx":0038
         DataSource      =   "AdoMaterias"
         Height          =   315
         Left            =   105
         TabIndex        =   21
         Top             =   2625
         Visible         =   0   'False
         Width           =   9465
         _ExtentX        =   16695
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
      Begin MSDataListLib.DataList DLGrupo 
         Bindings        =   "FLCxC.frx":0052
         DataSource      =   "AdoCuentas"
         Height          =   2220
         Left            =   105
         TabIndex        =   14
         Top             =   420
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   3916
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGTotalCursos 
         Bindings        =   "FLCxC.frx":006B
         Height          =   2220
         Left            =   -74895
         TabIndex        =   20
         Top             =   420
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   3916
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
   End
   Begin VB.CheckBox CheqDescuento 
      Caption         =   "Solo con Descuentos"
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
      Left            =   4095
      TabIndex        =   11
      Top             =   840
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   420
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
   Begin VB.CheckBox CheqTodos 
      Caption         =   "Todos"
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
      Top             =   840
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   2205
      TabIndex        =   4
      Top             =   0
      Width           =   8625
      Begin VB.OptionButton OpcRepre 
         Caption         =   "Nomina con Repre."
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
         Left            =   5670
         TabIndex        =   22
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcNotas 
         Caption         =   "Notas de Profesores"
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
         Left            =   4305
         TabIndex        =   8
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcListasProf 
         Caption         =   "Nomina Profesores"
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
         Left            =   2940
         TabIndex        =   7
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcListas 
         Caption         =   "Lista de Alumnos"
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
         Left            =   1785
         TabIndex        =   6
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton OpcFact 
         Caption         =   "Por Facturación"
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
         Left            =   7140
         TabIndex        =   9
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton OpcSinFact 
         Caption         =   "Por Listas Con Facturas"
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
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1590
      End
   End
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
      Height          =   855
      Left            =   105
      Picture         =   "FLCxC.frx":0088
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4095
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Direccion"
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
      Left            =   105
      Picture         =   "FLCxC.frx":0A7E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3150
      Width           =   960
   End
   Begin VB.CommandButton Command5 
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
      Height          =   855
      Left            =   105
      Picture         =   "FLCxC.frx":1300
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2205
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCProductos 
      Bindings        =   "FLCxC.frx":160A
      DataSource      =   "AdoProductos"
      Height          =   315
      Left            =   5565
      TabIndex        =   12
      Top             =   840
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FLCxC.frx":1625
      Height          =   2850
      Left            =   1155
      TabIndex        =   18
      Top             =   4410
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   5027
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin VB.CommandButton Command3 
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
      Height          =   855
      Left            =   105
      Picture         =   "FLCxC.frx":163C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1260
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   1575
      Top             =   2940
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   1575
      Top             =   2625
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   1575
      Top             =   2310
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
      Left            =   1155
      Top             =   7350
      Width           =   9675
      _ExtentX        =   17066
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   1575
      Top             =   1995
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
      Caption         =   "Productos"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoTotalCursos 
      Height          =   330
      Left            =   1575
      Top             =   1680
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
      Caption         =   "Productos"
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
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   1575
      Top             =   3255
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
      Caption         =   "Materias"
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
   Begin MSAdodcLib.Adodc AdoRepresentante 
      Height          =   330
      Left            =   1575
      Top             =   3570
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
      Caption         =   "Representante"
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
      TabIndex        =   2
      Top             =   420
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
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLCxC.frx":1946
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FMatriculados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ListarCuenta(TextoBusqueda As String)
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "' ")
       If Not .EOF Then
          CodigoCli = .fields("Codigo")
          Mifecha = PrimerDiaMes(FechaSistema)
          Dia = Day(Mifecha)
          Mes = Month(Mifecha)
          Anio = Year(Mifecha)
          FechaIni = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
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
  DGQuery.Visible = False
  MarcarTexto MBFechaI
  MarcarTexto MBFechaF
  Total = 0
  MensajeEncabData = "L I S T A D O    D E    M A T R I C U L A D O S"
  SQLMsg1 = " "
  SQLMsg2 = "" 'AdoQuery.Recordset.Fields("Grupo") & " - " & AdoQuery.Recordset.Fields("Direccion")
  SQLMsg3 = ""
  Si_No = False
  If Month(MBFechaI) <> Month(MBFechaF) Then Si_No = True
  Imprimir_Lista_Alumnos_Dir AdoQuery, DCProductos
  DGQuery.Visible = True
End Sub

Private Sub Command3_Click()
Dim ValorMes(20) As Currency
Dim CodigoMat As String
Dim PorCodMat As Boolean
RatonReloj
CodigoP = SinEspaciosDer(DCProductos)
If CodigoP = "" Then CodigoP = Ninguno
For IE = 0 To 19
    ValorMes(IE) = 0
Next IE
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)

PorCodMat = False
sSQL = "DELETE * " _
     & "FROM Saldo_Diarios " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND CodigoU = '" & CodigoUsuario & "' " _
     & "AND TP = 'NOMI' "
Ejecutar_SQL_SP sSQL
DGQuery.Caption = "LISTADO DE MOROSIDAD"
RatonReloj
Contador = 0
Cadena1 = Tipo_Acceso_Educativo("C.", "Grupo")
If OpcFact.value Then
   sSQL = "SELECT C.Grupo As Grupo_No,C.Sexo,C.Cliente,F.CodigoC,C.Casilla,C.CI_RUC As Codigos " _
        & "FROM Detalle_Factura As F,Clientes As C " _
        & "WHERE F.Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC NOT IN ('C','P') " _
        & "AND F.Codigo = '" & CodigoP & "' " _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.T <> '" & Anulado & "' "
   If CheqTodos.value <> 1 Then
      sSQL = sSQL & "AND C.Grupo = '" & SinEspaciosIzq(DLGrupo) & "' "
   Else
      sSQL = sSQL & Cadena1 & " "
   End If
   If CheqDescuento.value Then sSQL = sSQL & "AND F.Total_Desc > 0 "
   sSQL = sSQL & "GROUP BY C.Grupo,C.Sexo,C.Cliente,F.CodigoC,C.Casilla,CI_RUC,F.Codigo " _
        & "ORDER BY C.Grupo,C.Cliente "
ElseIf OpcSinFact.value Then
   sSQL = "SELECT C.Grupo As Grupo_No,C.Sexo,C.Cliente,F.Codigo,C.Casilla,C.CI_RUC As Codigos,SUM(F.Valor-F.Descuento) As Totales " _
        & "FROM Clientes_Facturacion As F,Clientes As C " _
        & "WHERE F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Year(MBFechaI) & "' " _
        & "AND F.Codigo = C.Codigo "
   If CheqTodos.value <> 1 Then
      sSQL = sSQL & "AND C.Grupo = '" & SinEspaciosIzq(DLGrupo) & "' "
   Else
      sSQL = sSQL & Cadena1 & " "
   End If
   If CheqDescuento.value Then sSQL = sSQL & "AND F.Descuento > 0 "
   sSQL = sSQL & "GROUP BY C.Grupo,C.Sexo,C.Cliente,F.Codigo,C.Casilla,CI_RUC " _
        & "ORDER BY C.Grupo,C.Cliente "
Else
  'Educativo:
   PorCodMat = True
   CodigoMat = Ninguno
   If AdoMaterias.Recordset.RecordCount > 0 Then
      AdoMaterias.Recordset.MoveFirst
      AdoMaterias.Recordset.Find ("Materia = '" & DCMaterias & "' ")
      If Not AdoMaterias.Recordset.EOF Then
         CodigoMat = AdoMaterias.Recordset.fields("CodMat")
      End If
   End If
   If OpcListas.value <> 0 Then
        sSQL = "SELECT C.Cliente,C.Direccion,C.Casilla,C.CI_RUC As Codigos,CM.* " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE C.FA <> 0 " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' "
        If CheqTodos.value <> 1 Then sSQL = sSQL & "AND CM.Grupo_No = '" & SinEspaciosIzq(DLGrupo) & "' "
        sSQL = sSQL & "AND C.Codigo = CM.Codigo " _
             & "ORDER BY CM.Grupo_No,C.Cliente "
        PorCodMat = False
   Else
        sSQL = "SELECT C.Cliente,C.Direccion,C.Casilla,C.CI_RUC As Codigos,TN.CodMat,CM.* " _
             & "FROM Clientes As C,Clientes_Matriculas As CM,Trans_Notas As TN " _
             & "WHERE CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' "
        If CheqTodos.value <> 1 Then sSQL = sSQL & "AND CM.Grupo_No = '" & SinEspaciosIzq(DLGrupo) & "' "
        If OpcTodasM.value <> 1 Then sSQL = sSQL & "AND TN.CodMat = '" & CodigoMat & "' "
        sSQL = sSQL _
             & "AND C.Codigo = CM.Codigo " _
             & "AND C.Codigo = TN.Codigo " _
             & "AND CM.Item = TN.Item " _
             & "AND CM.Periodo = TN.Periodo " _
             & "AND CM.Grupo_No = TN.CodE " _
             & "ORDER BY CM.Grupo_No,TN.CodMat,C.Cliente "
    End If
End If
'MsgBox sSQL
Select_Adodc AdoAux, sSQL
Total = 0: Saldo = 0: Contador = 0
With AdoAux.Recordset
 If .RecordCount > 0 Then
    'MsgBox sSQL & vbCrLf & .RecordCount
    .MoveFirst
     Codigo = .fields("Grupo_No")
     Cadena = Leer_Datos_del_Curso(.fields("Grupo_No"), 1)
     If PorCodMat Then CodigoMat = .fields("CodMat")
    'MsgBox Codigo
     Do While Not .EOF
        If Codigo <> .fields("Grupo_No") Then
           Codigo = .fields("Grupo_No")
           Contador = 0
        End If
        Contador = Contador + 1
        'MsgBox Codigo & vbCrLf & Contador
        SetAdoAddNew "Saldo_Diarios"
        SetAdoFields "TC", Format$(Contador, "00")
        SetAdoFields "TP", "NOMI"
        SetAdoFields "Item", NumEmpresa
        SetAdoFields "CodigoU", CodigoUsuario
        If OpcFact.value Then
           SetAdoFields "CodigoC", .fields("CodigoC")
        Else
           SetAdoFields "CodigoC", .fields("Codigo")
        End If
        Cta = MidStrg(.fields("Casilla"), 1, 1)
        Select Case Cta
          Case "D", "V": Cta = Cta
          Case Else: Cta = Ninguno
        End Select
        If PorCodMat Then CodigoMat = .fields("CodMat")
        SetAdoFields "T", Cta
        If AdoMaterias.Recordset.RecordCount > 0 Then
           AdoMaterias.Recordset.MoveFirst
           AdoMaterias.Recordset.Find ("CodMat = '" & CodigoMat & "' ")
           If Not AdoMaterias.Recordset.EOF Then
              SetAdoFields "Dato_Aux1", TrimStrg(UCaseStrg(MidStrg(AdoMaterias.Recordset.fields("Materia"), 1, 50)))
              SetAdoFields "Dato_Aux2", TrimStrg(UCaseStrg(MidStrg(AdoMaterias.Recordset.fields("Profesores"), 1, 50)))
           End If
        End If
        If OpcListas.value Or OpcListasProf.value Or OpcNotas.value Then
           SetAdoFields "Comprobante", TrimStrg(MidStrg(Dato_Curso.Bachiller, 1, 50))
           SetAdoFields "Dato_Aux3", TrimStrg(MidStrg(Dato_Curso.Especialidad, 1, 50))
        End If
        SetAdoFields "Dato_Aux4", TrimStrg(MidStrg(Dato_Curso.Descripcion, 1, 50))
        SetAdoFields "Cta", .fields("Grupo_No")
        SetAdoUpdate
        FMatriculados.Caption = Format$(Contador / .RecordCount, "00%")
       .MoveNext
     Loop
 End If
End With
Contador = 0
Total = 0: Saldo = 0
sSQL = "SELECT SD.T,SD.TC As No_, C.Cliente,"
 
If OpcListas.value = False Then
   JE = FechaMes(MBFechaI.Text)
   If (FechaAnio(MBFechaF.Text) - FechaAnio(MBFechaI.Text)) > 0 Then
      KE = 13 - FechaMes(MBFechaF)
      KE = KE + JE + 1
   Else
      KE = FechaMes(MBFechaF) - JE
      KE = KE + JE
   End If
  'MsgBox KE
   If KE < 12 Then KE = 17
   If OpcNotas.value Then KE = 5
'''   For IE = 1 To KE
'''       sSQL = sSQL & "SD." & UCaseStrg(MidStrg(MesesLetras(JE), 1, 3)) & ","
'''       JE = JE + 1
'''       If JE > 12 Then JE = 1
'''   Next IE
   For IE = 1 To KE
       sSQL = sSQL & "D_" & Format$(IE, "00") & ","
       JE = JE + 1
       If JE > 12 Then JE = 1
   Next IE
End If
Opcion = 3
'MsgBox OpcRepre.Value
If OpcRepre.value Then
   sSQL = "SELECT CM.Grupo_No,C.Cliente,CM.Representante,CM.Cedula_R,CM.Lugar_Trabajo_R,CM.Telefono_R " _
        & "FROM Clientes As C,Clientes_Matriculas As CM " _
        & "WHERE C.FA <> " & Val(adFalse) & " " _
        & "AND CM.Item = '" & NumEmpresa & "' " _
        & "AND CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Grupo_No = '" & SinEspaciosIzq(DLGrupo) & "' " _
        & "AND C.Codigo = CM.Codigo " _
        & "ORDER BY CM.Grupo_No,C.Cliente "
   Opcion = 99
Else
   sSQL = sSQL & "C.CI_RUC As Codigos,Dato_Aux4 As Curso,SD.Cta As Grupo,C.DireccionT As Domicilio,C.Telefono,C.Fecha_N, " _
        & "SD.Comprobante As Bachiller,SD.Dato_Aux3 As Especialidad,Dato_Aux1 As Materia,Dato_Aux2 As Profesor " _
        & "FROM Saldo_Diarios As SD,Clientes As C " _
        & "WHERE SD.Item = '" & NumEmpresa & "' " _
        & "AND SD.TP = 'NOMI' " _
        & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
        & "AND SD.CodigoC = C.Codigo " _
        & "ORDER BY C.Grupo,C.Cliente,C.Sexo,Dato_Aux1,Dato_Aux2 "
End If
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
Saldo = AdoQuery.Recordset.RecordCount
FMatriculados.Caption = "LISTADO DE ALUMNOS MATRICULADOS"
DGQuery.Visible = True
RatonNormal
End Sub

Private Sub Command5_Click()
  DGQuery.Visible = False
  MarcarTexto MBFechaI
  MarcarTexto MBFechaF
  'MsgBox Opcion
  Select Case Opcion
    Case 3
         Total = 0
         If CheqDescuento.value Then
            MensajeEncabData = "L I S T A D O    D E    B E C A D O S"
         Else
            MensajeEncabData = "L I S T A D O    D E    M A T R I C U L A D O S"
         End If
         SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
         SQLMsg2 = "" 'AdoQuery.Recordset.Fields("Grupo") & " - " & AdoQuery.Recordset.Fields("Direccion")
         SQLMsg3 = ""
         Si_No = False
         If Month(MBFechaI) <> Month(MBFechaF) Then Si_No = True
         If CheqTodos.value Then Encontro = True Else Encontro = False
         'MsgBox Encontro
         If OpcListas.value Then
            MensajeEncabData = ""
            Imprimir_Nomina_Alumnos AdoQuery, MBFechaI, MBFechaF, True, Encontro, 1
         ElseIf OpcListasProf.value Then
            MensajeEncabData = ""
            Imprimir_Nomina_Alumnos AdoQuery, MBFechaI, MBFechaF, True, Encontro, 2
         ElseIf OpcNotas.value Then
            MensajeEncabData = ""
            Imprimir_Nomina_Alumnos AdoQuery, MBFechaI, MBFechaF, True, Encontro, 3
         Else
            Imprimir_Lista_Alumnos AdoQuery, MBFechaI, MBFechaF, True, Encontro
         End If
    Case 4
         MensajeEncabData = "L I S T A D O    D E    B E C A D O S"
         SQLMsg1 = UCaseStrg(DCProductos)
         SQLMsg2 = "DE: " & UCaseStrg(DLGrupo)
         SQLMsg3 = ""
         ImprimirAdodc AdoQuery, 1, 9
    Case Else
         MensajeEncabData = "L I S T A D O    D E    A L U M N O S"
         SQLMsg1 = ""
         SQLMsg2 = "DE: " & UCaseStrg(DLGrupo)
         SQLMsg3 = ""
         ImprimirAdodc AdoQuery, 1, 9
  End Select
  DGQuery.Visible = True
End Sub

Private Sub Command6_Click()
  Unload FMatriculados
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FMatriculados, AdoQuery
     DGQuery.Visible = True
  End If
End Sub

Private Sub DGTotalCursos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FMatriculados, AdoQuery
     DGQuery.Visible = True
  End If
End Sub

Private Sub DLGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLGrupo_LostFocus()
  Codigo = SinEspaciosIzq(DLGrupo)
  Lista_Materias_Curso Codigo
  If DCMaterias.Visible Then DCMaterias.SetFocus
End Sub

Private Sub Form_Activate()
  FMatriculados.Caption = "CREACION DEL CLIENTE"
  Cadena1 = Tipo_Acceso_Educativo("", "Grupo")
  sSQL = "SELECT Producto & ' ' & Codigo_Inv As NombProducto " _
       & "FROM Catalogo_Productos " _
       & "WHERE TC = 'P' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDB_Combo DCProductos, AdoProductos, sSQL, "NombProducto"
  If UCaseStrg(Modulo) = "EDUCATIVO" Then
     sSQL = "SELECT (Curso & ' - ' & Descripcion) As Curso,COUNT(CM.Grupo_No) As CantAlum " _
          & "FROM Catalogo_Cursos As CC,Clientes_Matriculas As CM " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND CC.Item = CM.Item " _
          & "AND CC.Periodo = CM.Periodo " _
          & "AND CC.Curso = CM.Grupo_No " _
          & "GROUP BY CC.Curso,CC.Descripcion " _
          & "ORDER BY CC.Curso,CC.Descripcion "
  Else
     sSQL = "SELECT (Grupo & ' - ' & Direccion) As Curso,COUNT(DF.Codigo) As CantAlum " _
          & "FROM Clientes As C,Detalle_Factura As DF " _
          & "WHERE C.FA <> " & Val(adFalse) & " " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = DF.CodigoC " _
          & "GROUP BY C.Grupo,C.Direccion " _
          & "ORDER BY C.Grupo "
  End If
 'MsgBox sSQL
  SelectDB_List DLGrupo, AdoCuentas, sSQL, "Curso"
  SQL1 = "SELECT Grupo,Direccion As Curso,COUNT(CM.Codigo) As Alumnos " _
       & "FROM Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE C.FA <> " & Val(adFalse) & " " _
       & "AND CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = CM.Codigo "
  If UCaseStrg(Modulo) = "EDUCATIVO" Then SQL1 = SQL1 & Cadena1 & " "
  SQL1 = SQL1 & "GROUP BY C.Grupo,C.Direccion " _
       & "ORDER BY C.Grupo "
  Select_Adodc_Grid DGTotalCursos, AdoTotalCursos, SQL1
  I = 0
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If I < .fields("CantAlum") Then I = .fields("CantAlum")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  Lista_Materias_Curso Ninguno
  Frame1.Caption = " |Tope máxino Estudiantes: " & I & "| "
  DCMaterias.Text = "SELECCIONE UN GRADO"
  FMatriculados.WindowState = vbMaximized
  RatonNormal
'  MBFechaI.SetFocus
End Sub

Private Sub Form_Deactivate()
  FMatriculados.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoMaterias
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoProductos
   ConectarAdodc AdoTotalCursos
   ConectarAdodc AdoRepresentante
   If UCaseStrg(Modulo) = "EDUCATIVO" Then
      DCProductos.Visible = False
      CheqDescuento.Visible = False
      OpcFact.Visible = False
   Else
      DCProductos.Visible = True
      CheqDescuento.Visible = True
      OpcFact.Visible = True
   End If
   DGQuery.Height = MDI_Y_Max - DGQuery.Top - 350
   DGQuery.width = MDI_X_Max - DGQuery.Left - 100
   AdoQuery.width = MDI_X_Max - DGQuery.Left - 100
   AdoQuery.Top = DGQuery.Top + DGQuery.Height
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

Private Sub OpcFact_Click()
  DCMaterias.Visible = False
End Sub

Private Sub OpcListas_Click()
  DCMaterias.Visible = False
End Sub

Private Sub OpcListasProf_Click()
  DCMaterias.Visible = False
End Sub

Private Sub OpcNotas_Click()
  DCMaterias.Visible = True
End Sub

Private Sub OpcRepre_Click()
  DCMaterias.Visible = False
End Sub

Private Sub OpcSinFact_Click()
  DCMaterias.Visible = False
End Sub

Public Sub Lista_Materias_Curso(CodigoCurso As String)
  sSQL = "SELECT CE.TC,CE.CodigoE, CM.Materia, C.Cliente As Profesores, CE.CodMat " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias As CM,Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.CodMat <> '.' " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Profesor = C.Codigo " _
       & "AND MidStrg(CE.CodigoE,1," & Len(CodigoCurso) & ") = '" & CodigoCurso & "' " _
       & "ORDER BY CE.CodigoE "
  SelectDB_Combo DCMaterias, AdoMaterias, sSQL, "Materia"
  If OpcListasProf.value = True Then DCMaterias.Visible = True
End Sub
