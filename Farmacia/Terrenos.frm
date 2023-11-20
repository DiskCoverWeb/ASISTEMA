VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FTerrenos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "&Abonos"
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
      Left            =   10815
      Picture         =   "Terrenos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1050
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Productos"
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
      Left            =   10815
      Picture         =   "Terrenos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   105
      Width           =   1065
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
      Height          =   855
      Left            =   10815
      Picture         =   "Terrenos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1995
      Width           =   1065
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
      Height          =   855
      Left            =   10815
      Picture         =   "Terrenos.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2940
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   105
      TabIndex        =   13
      Top             =   0
      Width           =   10620
      Begin VB.OptionButton OpcLote 
         Caption         =   "Por &Lote"
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
         Left            =   2940
         TabIndex        =   28
         Top             =   735
         Width           =   2850
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "Terrenos.frx":1A18
         DataSource      =   "AdoCliente"
         Height          =   315
         Left            =   6300
         TabIndex        =   6
         Top             =   1050
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   105
         TabIndex        =   21
         Top             =   1050
         Width           =   2640
         Begin VB.OptionButton OpcPend 
            Caption         =   "&Normales"
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
            Left            =   105
            TabIndex        =   11
            Top             =   0
            Value           =   -1  'True
            Width           =   1185
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
            Left            =   1365
            TabIndex        =   12
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.TextBox TextPatron 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6300
         TabIndex        =   7
         Top             =   1050
         Width           =   4215
      End
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   1470
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   1470
         TabIndex        =   1
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
      Begin MSDataListLib.DataCombo DCFactura 
         Bindings        =   "Terrenos.frx":1A31
         DataSource      =   "AdoFacturas"
         Height          =   315
         Left            =   3255
         TabIndex        =   9
         Top             =   420
         Width           =   1710
         _ExtentX        =   3016
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
      Begin VB.ListBox ListCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6300
         TabIndex        =   5
         Top             =   420
         Width           =   4215
      End
      Begin VB.CheckBox CheckTodos 
         Caption         =   "LISTAR &TODOS LOS CLIENTES"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   1470
         Width           =   2715
      End
      Begin VB.OptionButton OpcFactura 
         Caption         =   "&Por Contrato"
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
         Left            =   2940
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   3060
      End
      Begin VB.OptionButton OpcCliente 
         Caption         =   "Procesar por Patron de &Busqueda:"
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
         TabIndex        =   4
         Top             =   210
         Width           =   3270
      End
      Begin MSDataListLib.DataCombo DCObs 
         Bindings        =   "Terrenos.frx":1A4B
         DataSource      =   "AdoObs"
         Height          =   315
         Left            =   2940
         TabIndex        =   26
         Top             =   1050
         Width           =   3270
         _ExtentX        =   5768
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
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   330
         Left            =   2940
         TabIndex        =   27
         Top             =   1470
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   582
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Fecha Final:"
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
         Width           =   1380
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Fecha Inicial:"
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
         Width           =   1380
      End
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Terrenos.frx":1A60
      Height          =   4740
      Left            =   105
      TabIndex        =   20
      Top             =   1890
      Width           =   10620
      _ExtentX        =   18733
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   360
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Facturas"
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
      Left            =   360
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc AdoHistoria 
      Height          =   330
      Left            =   360
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Historia"
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
   Begin MSAdodcLib.Adodc AdoObs 
      Height          =   330
      Left            =   315
      Top             =   3675
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
      Caption         =   "Obs"
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
   Begin VB.Label LabelSaldo 
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
      Left            =   8925
      TabIndex        =   18
      Top             =   6615
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
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
      TabIndex        =   19
      Top             =   6615
      Width           =   645
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
      Left            =   6510
      TabIndex        =   14
      Top             =   6615
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cobrado"
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
      Left            =   5565
      TabIndex        =   15
      Top             =   6615
      Width           =   960
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
      Left            =   3780
      TabIndex        =   16
      Top             =   6615
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Facturado"
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
      TabIndex        =   17
      Top             =   6615
      Width           =   1065
   End
End
Attribute VB_Name = "FTerrenos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
   DGQuery.Visible = False
   If Opcion = 1 Then
      MensajeEncabData = "ESTADO DE CUENTA DE CLIENTES"
      SQLMsg1 = DGQuery.Caption
      SQLMsg2 = "CORTE DEL " & MBoxFechaI.Text & " AL " & MBoxFechaF.Text
      MiFecha = MBoxFechaF.Text
      Imprimir_Estado_Cuenta_Clientes AdoQuery, 1, 8, True
   Else
      MensajeEncabData = "ESTADO DE ABONOS DE CLIENTES"
      SQLMsg1 = "CORTE DEL " & MBoxFechaI.Text & " AL " & MBoxFechaF.Text
      SQLMsg2 = ""
      MiFecha = MBoxFechaF.Text
      ImprimirAbonosDeCaja AdoQuery, MBoxFechaI.Text, MBoxFechaF.Text
   End If
   DGQuery.Visible = True
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command4_Click()
Opcion = 1
DGQuery.Visible = False
FechaValida MBoxFechaI
FechaValida MBoxFechaF
FechaIni = BuscarFecha(MBoxFechaI.Text)
FechaFin = BuscarFecha(MBoxFechaF.Text)
Factura_No = Val(DCFactura.Text)
DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
RatonReloj
sSQL = "DELETE * " _
     & "FROM Saldo_Diarios " _
     & "WHERE CodigoU = '" & CodigoUsuario & "' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND TP = 'TERR' "
ConectarAdoExecute sSQL

sSQL = "SELECT F.Codigo,C.Cliente,F.TC,F.T,F.Fecha,F.CodigoC,F.Factura,F.Producto,F.Total,F.Total_IVA,F.Cantidad,F.Ticket " _
     & "FROM Clientes As C,Detalle_Factura As F " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
If OpcPend.Value Then sSQL = sSQL & "AND F.T <> '" & Anulado & "' "
If OpcAnul.Value Then sSQL = sSQL & "AND F.T = '" & Anulado & "' "
sSQL = sSQL & "AND C.Codigo = F.CodigoC " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.TC <> 'C' " _
     & "AND F.TC <> 'P' "
If CheckTodos.Value = 0 Then
   If OpcCliente.Value Then
      With AdoCliente.Recordset
       Select Case .Fields(ListCliente.Text).Type
         Case dbDate
              sSQL = sSQL & "AND " & ListCliente.Text & " = #" & BuscarFecha(TextPatron.Text) & "# "
         Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbBoolean
              sSQL = sSQL & "AND " & ListCliente.Text & " = " & Val(TextPatron.Text) & " "
         Case dbText, dbMemo
              LongStrg = Len(TextPatron.Text)
              If ListCliente.Text = "Codigo" Then
                 sSQL = sSQL & "AND Mid(Ucase(CodigoC),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              Else
                 sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              End If
         Case Else
              LongStrg = Len(TextPatron.Text)
              If ListCliente.Text = "Codigo" Then
                 sSQL = sSQL & "AND Mid(Ucase(CodigoC),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              Else
                 sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              End If
       End Select
      End With
   End If
   If OpcFactura.Value Then sSQL = sSQL & "AND F.Factura = " & Factura_No & " "
   If OpcLote.Value Then sSQL = sSQL & "AND Ucase(F.Ruta) = '" & UCase(DCObs.Text) & "' "
   
End If
sSQL = sSQL & "ORDER BY C.Cliente,F.Codigo,F.Fecha "
SelectData AdoHistoria, sSQL
RatonReloj
ProgBar.Min = 0
ProgBar.Value = 0
With AdoHistoria.Recordset
 If .RecordCount > 0 Then
     Saldo = 0
     Contador = 0
     ProgBar.Max = .RecordCount + 1
     CodigoCliente = .Fields("CodigoC")
     Factura_No = .Fields("Factura")
     FechaStr = .Fields("Fecha")
     TipoDoc = .Fields("T")
     TipoProc = .Fields("TC")
     NivelNo = .Fields("Codigo")
     MiFecha = FechaStr
     FechaTexto = MiFecha
     DGQuery.Caption = "ESTADO DE CUENTA DE CONJUNTOS HABITACIONALES"
     FTerrenos.Caption = "ESTADO DE CUENTA DE CONJUNTOS HABITACIONALES"
     sSQL = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND CodigoC = '" & CodigoCliente & "' "
     SelectAdodc AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then
        DGQuery.Caption = "PROYECTO: " & UCase(AdoQuery.Recordset.Fields("Fecha_Tours"))
        CodigoL = UCase(AdoQuery.Recordset.Fields("Definitivo"))
     End If
     Do While Not .EOF
        If (CodigoCliente <> .Fields("CodigoC")) Or (NivelNo <> .Fields("Codigo")) Then
          'Verificar si existen facturas
           sSQL = "SELECT * " _
                & "FROM Trans_Abonos " _
                & "WHERE Fecha <= #" & FechaFin & "# " _
                & "AND Factura = " & Factura_No & " " _
                & "AND CodigoC = '" & CodigoCliente & "' " _
                & "AND Codigo_Inv = '" & NivelNo & "' " _
                & "ORDER BY Fecha,Codigo_Inv,Abono "
           SelectAdodc AdoQuery, sSQL
           RatonReloj
           If AdoQuery.Recordset.RecordCount > 0 Then
              Do While Not AdoQuery.Recordset.EOF
                 Contador = Contador + 1
                 MiFecha = AdoQuery.Recordset.Fields("Fecha")
                 NombreCliente = AdoQuery.Recordset.Fields("Banco")
                 NoCheque = AdoQuery.Recordset.Fields("Cheque")
                 NivelNo = AdoQuery.Recordset.Fields("Codigo_Inv")
                 If NoCheque = Ninguno And AdoQuery.Recordset.Fields("Recibo_No") > 0 Then NoCheque = CStr(AdoQuery.Recordset.Fields("Recibo_No"))
                 Total = 0
                 Abono = AdoQuery.Recordset.Fields("Abono")
                 Saldo = Saldo - Abono
                 ProcesarProducto
                 AdoQuery.Recordset.MoveNext
              Loop
           End If
           CodigoCliente = .Fields("CodigoC")
           NombreCliente = .Fields("Producto")
           NivelNo = .Fields("Codigo")
           Factura_No = .Fields("Factura")
           FechaStr = .Fields("Fecha")
           TipoDoc = .Fields("T")
           TipoProc = .Fields("TC")
           MiFecha = FechaStr
           FechaTexto = MiFecha
           Factura_No = Numero
           Saldo = 0
           Contador = 0
           Total = 0
           Total_IVA = 0
        End If
        Contador = Contador + 1
        MiFecha = .Fields("Fecha")
        FechaTexto = MiFecha
        Factura_No = .Fields("Factura")
        Abono = 0
        Total = .Fields("Total")
        Saldo = Saldo + Total
        NoCheque = Ninguno
        ProcesarProducto
        ProgBar.Value = ProgBar.Value + 1
       .MoveNext
     Loop
    'Verificar si existen facturas
     sSQL = "SELECT * " _
          & "FROM Trans_Abonos " _
          & "WHERE Fecha <= #" & FechaFin & "# " _
          & "AND Factura = " & Factura_No & " " _
          & "AND CodigoC = '" & CodigoCliente & "' " _
          & "AND Codigo_Inv = '" & NivelNo & "' " _
          & "ORDER BY Fecha,Abono "
     SelectAdodc AdoQuery, sSQL
     RatonReloj
     If AdoQuery.Recordset.RecordCount > 0 Then
        Do While Not AdoQuery.Recordset.EOF
           Contador = Contador + 1
           MiFecha = AdoQuery.Recordset.Fields("Fecha")
           NombreCliente = AdoQuery.Recordset.Fields("Banco")
           NoCheque = AdoQuery.Recordset.Fields("Cheque")
           NivelNo = AdoQuery.Recordset.Fields("Codigo_Inv")
           If NoCheque = Ninguno And AdoQuery.Recordset.Fields("Recibo_No") > 0 Then NoCheque = CStr(AdoQuery.Recordset.Fields("Recibo_No"))
           Total = 0
           Abono = AdoQuery.Recordset.Fields("Abono")
           Saldo = Saldo - Abono
           ProcesarProducto
           AdoQuery.Recordset.MoveNext
        Loop
     End If
     ProgBar.Value = ProgBar.Max
 End If
End With
DGQuery.Visible = False
'Codigo = DCCliente.Text

'ANDRADE MACIAS EDUARDO
sSQL = "SELECT SD.Recibo,C.Cliente,SD.Numero As Factura,SD.Fecha,SD.Comprobante As Detalle,SD.Cta As Recibo_No," _
     & "Total As Ventas,Egresos As Abono,Saldo_Actual " _
     & "FROM Saldo_Diarios As SD,Clientes As C " _
     & "WHERE SD.Item = '" & NumEmpresa & "' " _
     & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
     & "AND SD.CodigoC = C.Codigo " _
     & "AND TP = 'TERR' " _
     & "ORDER BY C.Cliente,SD.Recibo,SD.ME "
SelectDataGrid DGQuery, AdoQuery, sSQL, , True
Total = 0: Abono = 0
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     Do While Not .EOF
        Total = Total + .Fields("Ventas")
        Abono = Abono + .Fields("Abono")
       .MoveNext
     Loop
    .MoveFirst
 End If
End With
Saldo = Total - Abono
'SelectDataGrid DGQuery, AdoQuery, SQLMsg1
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Total - Abono, "#,##0.00")
DGQuery.Visible = True
RatonNormal
End Sub

Private Sub Command5_Click()
Opcion = 2
DGQuery.Visible = False
FechaValida MBoxFechaI
FechaValida MBoxFechaF
FechaIni = BuscarFecha(MBoxFechaI.Text)
FechaFin = BuscarFecha(MBoxFechaF.Text)
'Codigo = DCCliente.Text
   Factura_No = Val(DCFactura.Text)
   DGQuery.Caption = "ABONOS DE FACTURAS"
   RatonReloj
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT TA.TP,TA.Fecha,C.Cliente,TA.Factura,TA.Banco,TA.Cheque,TA.Abono,TA.Comprobante,TA.Cta,TA.Cta_CxP " _
        & "FROM Trans_Abonos As TA,Clientes C " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.TP <> 'C' " _
        & "AND TA.TP <> 'P' " _
        & "AND TA.T <> 'A' " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.CodigoC = C.Codigo "
   If CheckTodos.Value = 0 Then
   If OpcCliente.Value Then
      'sSQL = sSQL & "AND C.Cliente = '" & Codigo & "' "
      With AdoCliente.Recordset
       Select Case .Fields(ListCliente.Text).Type
         Case dbDate
              sSQL = sSQL & "AND " & ListCliente.Text & " = #" & BuscarFecha(TextPatron.Text) & "# "
         Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbBoolean
              sSQL = sSQL & "AND " & ListCliente.Text & " = " & Val(TextPatron.Text) & " "
         Case dbText, dbMemo
              LongStrg = Len(TextPatron.Text)
              If ListCliente.Text = "Codigo" Then
                 sSQL = sSQL & "AND Mid(Ucase(CodigoC),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              Else
                 sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              End If
         Case Else
              LongStrg = Len(TextPatron.Text)
              If ListCliente.Text = "Codigo" Then
                 sSQL = sSQL & "AND Mid(Ucase(CodigoC),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              Else
                 sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
              End If
       End Select
      End With
   End If
   If OpcFactura.Value Then sSQL = sSQL & "AND TA.Factura = " & Factura_No & " "
   End If
   sSQL = sSQL & "ORDER BY C.Cliente,TA.Factura,TA.Fecha,TA.Banco "
   SelectDataGrid DGQuery, AdoQuery, sSQL
   With AdoQuery.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Total = Total + .Fields("Abono")
          .MoveNext
        Loop
    End If
   End With
   LabelAbonado.Caption = Format(Total, "#,##0.00")
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
End Sub

Private Sub DCCliente_LostFocus()
  TextPatron.Text = DCCliente.Text
  DCCliente.Visible = False
  TextPatron.Visible = True
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FTerrenos, AdoQuery
     DGQuery.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyG Then
     With AdoQuery.Recordset
          Numero = .Fields("Factura")
          CodigoCli = .Fields("CodigoC")
          TipoDoc = .Fields("Tipo")
          Abono = .Fields("Abono")
          Total = .Fields("Total")
     End With
     Codigo = InputBox("CAMBIO DE GRUPO", "NUEVO GRUPO: ", NumEmpresa)
     If Codigo <> "" Then
        sSQL = "UPDATE Facturas " _
             & "SET Definitivo = '" & Codigo & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND TC = '" & TipoDoc & "' " _
             & "AND Factura = " & Numero & " " _
             & "AND CodigoC = '" & CodigoCli & "' "
        ConectarAdoExecute sSQL
        MsgBox "Proceso realizado con exito"
     End If
  End If
End Sub

Private Sub Form_Activate()
   ListCliente.Visible = False
   TextPatron.Visible = False
   CheckTodos.Value = 0
   sSQL = "SELECT Definitivo " _
        & "FROM Facturas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "GROUP BY Definitivo " _
        & "ORDER BY Definitivo "
   SelectDBCombo DCObs, AdoObs, sSQL, "Definitivo"
   sSQL = "SELECT Codigo,Cliente,Telefono,Ciudad,Direccion,Grupo,DirNumero,FactM " _
        & "FROM Clientes " _
        & "WHERE Grupo <> '.' " _
        & "ORDER BY Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
   ListCliente.Clear
   With AdoCliente.Recordset
    For I = 0 To .Fields.Count - 1
        ListCliente.AddItem .Fields(I).Name
    Next I
   End With
   ListCliente.Text = ListCliente.List(0)
   sSQL = "SELECT Factura " _
        & "FROM Facturas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T <> 'A' " _
        & "AND TC <> 'C' " _
        & "AND TC <> 'P' " _
        & "GROUP BY Factura " _
        & "ORDER BY Factura "
   SelectDBCombo DCFactura, AdoFacturas, sSQL, "Factura"
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FTerrenos
   ConectarAdodc AdoObs
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCliente
   ConectarAdodc AdoHistoria
   ConectarAdodc AdoFacturas
End Sub

Private Sub ListCliente_LostFocus()
  If ListCliente.Text = "Cliente" Then
     DCCliente.Visible = True
     TextPatron.Visible = False
     DCCliente.SetFocus
  End If
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub OpcCliente_Click()
  If OpcCliente.Value Then
     CheckTodos.Value = 0
     ListCliente.Visible = True
     TextPatron.Visible = True
  End If
End Sub

Private Sub OpcFactura_Click()
  ListCliente.Visible = False
  TextPatron.Visible = False
  CheckTodos.Value = 0
End Sub

Public Sub ProcesarProducto(Optional EsSaldo As Boolean)
   SetAdoAddNew "Saldo_Diarios"
   SetAdoFields "ME", CByte(Contador)
   SetAdoFields "TP", "TERR"
   SetAdoFields "CodigoU", CodigoUsuario
   SetAdoFields "Comprobante", NombreCliente
   SetAdoFields "Item", NumEmpresa
   SetAdoFields "T", TipoDoc
   SetAdoFields "CodigoC", CodigoCliente
   SetAdoFields "Numero", Factura_No
   SetAdoFields "Fecha", MiFecha
   SetAdoFields "Fecha_Venc", FechaTexto
   SetAdoFields "Cta", NoCheque
   SetAdoFields "Recibo", NivelNo
   SetAdoFields "Total", Total
   SetAdoFields "Ingresos", Total_IVA
   SetAdoFields "Egresos", Abono
   SetAdoFields "Saldo_Actual", Saldo
   SetAdoUpdate
End Sub

Private Sub TextPatron_GotFocus()
  MarcarTexto TextPatron
End Sub

Public Sub Sumatoria_Abonos()
Dim ContAbonos As Long
    ContAbonos = 0
     'Verificar si existen facturas
      sSQL = "SELECT * " _
           & "FROM Trans_Abonos " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Factura = " & Factura_No & " " _
           & "AND CodigoC = '" & CodigoCliente & "' " _
           & "ORDER BY Fecha,Abono "
      SelectAdodc AdoQuery, sSQL
      RatonReloj
      If AdoQuery.Recordset.RecordCount > 0 Then
         Do While Not AdoQuery.Recordset.EOF
            ContAbonos = ContAbonos + 1
            Contador = Contador + 1
            FechaTexto = AdoQuery.Recordset.Fields("Fecha")
            FechaTexto1 = AdoQuery.Recordset.Fields("Fecha")
            Abono = AdoQuery.Recordset.Fields("Abono")
            NombreCliente = Format(ContAbonos, "00") & ".- " & AdoQuery.Recordset.Fields("Banco")
            CodigoP = AdoQuery.Recordset.Fields("Cheque")
            Saldo = Saldo - Abono
            'ProcesarFactura
            AdoQuery.Recordset.MoveNext
         Loop
      Else
         Contador = Contador + 1
         'ProcesarFactura True
      End If
      TipoDoc = "-"
      TipoProc = "--"
      MiFecha = FechaSistema
      FechaTexto = FechaSistema
      Total = 0: Abono = 0: Saldo = 0
      Contador = Contador + 1
End Sub
