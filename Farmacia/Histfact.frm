VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form HistorialFacturas 
   Caption         =   "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   10815
      TabIndex        =   1
      Top             =   0
      Width           =   5370
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
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   210
         Width           =   1410
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
         Height          =   315
         Left            =   4410
         TabIndex        =   5
         Top             =   210
         Width           =   870
      End
      Begin VB.OptionButton OpcPend 
         Caption         =   "&Pendiente"
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
         TabIndex        =   3
         Top             =   210
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
         Height          =   315
         Left            =   3045
         TabIndex        =   4
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.ListBox ListCliente 
      DataSource      =   "AdoQuery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7665
      TabIndex        =   20
      Top             =   1050
      Width           =   2010
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Histfact.frx":0000
      DataSource      =   "AdoCliente"
      Height          =   2100
      Left            =   9765
      TabIndex        =   21
      Top             =   735
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3704
      _Version        =   393216
      Style           =   1
      Text            =   ""
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
   Begin VB.CheckBox CheqMarca 
      Caption         =   "Tipo de Marca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   16
      Top             =   2310
      Width           =   1665
   End
   Begin VB.CheckBox CheqDescItem 
      Caption         =   "Descripcion Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   17
      Top             =   2625
      Width           =   2295
   End
   Begin VB.CheckBox CheqCxC 
      Caption         =   "Cuenta por Cobrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   10
      Top             =   735
      Width           =   2610
   End
   Begin VB.CheckBox CheqProd 
      Caption         =   "Listar por Productos y/o Nivel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   13
      Top             =   1575
      Width           =   3030
   End
   Begin VB.TextBox TxtNivel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6930
      TabIndex        =   14
      Text            =   "0"
      Top             =   1575
      Width           =   330
   End
   Begin VB.CheckBox CheqAbonos 
      Caption         =   "Tipo de Cuenta de Abono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4620
      TabIndex        =   11
      Top             =   735
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   105
      TabIndex        =   29
      Top             =   7770
      Width           =   540
   End
   Begin MSComctlLib.ImageList ImageListFacturas 
      Left            =   5775
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":0019
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":08F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":11CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":14E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1801
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":252D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":2847
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3145
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":345F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3779
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3A93
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":436D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":4C47
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":73F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":7553
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":786D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":7CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":7E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":7F73
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Histfact.frx":828D
      Height          =   4530
      Left            =   105
      TabIndex        =   28
      Top             =   2940
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   7990
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   630
      Top             =   7770
      Width           =   3795
      _ExtentX        =   6694
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   315
      Top             =   5040
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
      Left            =   315
      Top             =   4410
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
      Left            =   315
      Top             =   4725
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   315
      Top             =   5355
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
   Begin MSAdodcLib.Adodc AdoCxC 
      Height          =   330
      Left            =   315
      Top             =   5670
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
      Caption         =   "CxC"
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   315
      Top             =   5985
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
      Caption         =   "Producto"
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   330
      Left            =   315
      Top             =   6300
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
      Caption         =   "Marca"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   1785
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
      Left            =   105
      TabIndex        =   7
      Top             =   1050
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
   Begin MSComctlLib.Toolbar ToolbarMenu 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageListFacturas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir los resultados"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Facturas"
            Object.ToolTipText     =   "Presenta el Resumen de Facturas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Productos"
            Object.ToolTipText     =   "Presenta facturas con detalles"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resumen"
            Object.ToolTipText     =   "Resumen de Productos"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clientes"
            Object.ToolTipText     =   "Ventas por Clientes"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Costos"
            Object.ToolTipText     =   "Resumen de Ventas/Costos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VentasxProductos"
            Object.ToolTipText     =   "Ventas por Productos"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abonos"
            Object.ToolTipText     =   "Detalle de Abonos de Facturas/Notas de Ventas"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Protestado"
            Object.ToolTipText     =   "Cheques protestados"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Por_Clientes"
            Object.ToolTipText     =   "Listado de Facturas ordenadas por Clientes"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Por_Facturas"
            Object.ToolTipText     =   "Listado de Clientes ordenados por Facturas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Retenciones_NC"
            Object.ToolTipText     =   "Presentar Retenciones y Notas de Credito"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abonos_Ant"
            Object.ToolTipText     =   "Abonos Anticipados"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abonos_Erroneos"
            Object.ToolTipText     =   "Presenta Abonos mal procesados"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CxC_Clientes"
            Object.ToolTipText     =   "Listado de Cartera por Meses"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Listar_Por_Meses"
            Object.ToolTipText     =   "Listar Clientes por Rubro de Meses "
            ImageIndex      =   19
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Estado_Cuenta_Cliente"
            Object.ToolTipText     =   "Lista el Estado de cuenta de Clientes"
            ImageIndex      =   20
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   16275
         TabIndex        =   30
         Top             =   0
         Width           =   2430
         Begin VB.CheckBox CheqPreFa 
            Caption         =   "Incluir Prefacturación"
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
            Top             =   210
            Width           =   2220
         End
      End
   End
   Begin MSDataListLib.DataCombo DCCxC 
      Bindings        =   "Histfact.frx":82A4
      DataSource      =   "AdoCxC"
      Height          =   315
      Left            =   1470
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
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
   Begin MSDataListLib.DataCombo DCProducto 
      Bindings        =   "Histfact.frx":82B9
      DataSource      =   "AdoProducto"
      Height          =   315
      Left            =   1470
      TabIndex        =   15
      Top             =   1890
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
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
   Begin MSDataListLib.DataCombo DCMarca 
      Bindings        =   "Histfact.frx":82D3
      DataSource      =   "AdoMarca"
      Height          =   315
      Left            =   2415
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
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
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Patron de &Busqueda:"
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
      Left            =   7665
      TabIndex        =   19
      Top             =   735
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Inicial"
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
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Final"
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
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label LabelSaldo 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   10395
      TabIndex        =   26
      Top             =   7770
      Width           =   1590
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
      Left            =   9555
      TabIndex        =   27
      Top             =   7770
      Width           =   855
   End
   Begin VB.Label LabelAbonado 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7980
      TabIndex        =   22
      Top             =   7770
      Width           =   1590
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
      Left            =   7035
      TabIndex        =   23
      Top             =   7770
      Width           =   960
   End
   Begin VB.Label LabelFacturado 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5460
      TabIndex        =   24
      Top             =   7770
      Width           =   1590
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
      Left            =   4410
      TabIndex        =   25
      Top             =   7770
      Width           =   1065
   End
End
Attribute VB_Name = "HistorialFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PorCxC As Boolean
Dim PorProd As Boolean
Dim PorMarca As Boolean

Private Sub CheqAbonos_Click()
  If CheqAbonos.value = 1 Then
     sSQL = "SELECT (TA.Cta & ' - ' & CC.Cuenta) As NomCxC " _
          & "FROM Trans_Abonos As TA,Catalogo_Cuentas As CC " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND CC.Periodo = TA.Periodo " _
          & "AND CC.Item = TA.Item " _
          & "AND CC.Codigo = TA.Cta " _
          & "GROUP BY Cta,Cuenta " _
          & "ORDER BY Cta "
     SelectDBCombo DCCxC, AdoCxC, sSQL, "NomCxC"
     CheqCxC.value = 0
     DCCxC.Visible = True
  Else
     DCCxC.Visible = False
  End If
End Sub

Private Sub CheqCxC_Click()
  If CheqCxC.value = 1 Then
     sSQL = "SELECT (Codigo & ' - ' & Concepto) As NomCxC " _
          & "FROM Catalogo_Lineas " _
          & "WHERE TL <> " & Val(adFalse) & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo "
          
     sSQL = "SELECT (F.Cta_CxP & ' - ' & CC.Cuenta) As NomCxC " _
          & "FROM Facturas As F,Catalogo_Cuentas As CC " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Item = F.Item " _
          & "AND CC.Codigo = F.Cta_CxP " _
          & "GROUP BY F.Cta_CxP,Cuenta " _
          & "ORDER BY F.Cta_CxP "
     SelectDBCombo DCCxC, AdoCxC, sSQL, "NomCxC"
     CheqAbonos.value = 0
     DCCxC.Visible = True
  Else
     DCCxC.Visible = False
  End If
End Sub

Private Sub CheqDescItem_Click()
  If CheqDescItem.value = 1 Then
     sSQL = "SELECT Desc_Item " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Desc_Item <> '" & Ninguno & "' " _
          & "GROUP BY Desc_Item " _
          & "ORDER BY Desc_Item "
     SelectDBCombo DCMarca, AdoMarca, sSQL, "Desc_Item"
     DCMarca.Visible = True
     CheqMarca.value = 0
  Else
     DCMarca.Visible = False
  End If
End Sub

Private Sub CheqMarca_Click()
  If CheqMarca.value = 1 Then
     sSQL = "SELECT (CodMar & ' - ' & Marca) As NomMarca " _
          & "FROM Catalogo_Marcas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodMar <> '" & Ninguno & "' " _
          & "ORDER BY Marca "
     SelectDBCombo DCMarca, AdoMarca, sSQL, "NomMarca"
     DCMarca.Visible = True
     CheqDescItem.value = 0
  Else
     DCMarca.Visible = False
  End If
End Sub

Private Sub CheqProd_Click()
  If CheqProd.value = 1 Then
     DCProducto.Visible = True
  Else
     DCProducto.Visible = False
  End If
End Sub

Public Sub Historico_Facturas()
'Facturas
 Opcion = 1
 DGQuery.Caption = "HISTORIAL DE FACTURAS"
 Label2.Caption = " Facturado"
 Label4.Caption = " Cobrado"
 Label3.Caption = " Saldo"
 PorCxC = False
 PorProd = False
 If CheqCxC.value = 1 Then PorCxC = True
 If CheqProd.value = 1 Then PorProd = True
 Actualizar_Saldos_Facturas MBFechaF
 RatonReloj
 Total = 0
 Abono = 0
 Saldo = 0
 sSQL = "SELECT C.Cliente,F.T,F.Factura,F.Fecha As Fecha_Emis,F.Fecha_C As Abonado_El," _
      & "F.Total_MN As Total,(F.Total_MN - F.Saldo_Actual) As Abono,F.Saldo_Actual," _
      & "F.CodigoC,C.CI_RUC,F.TC,F.Serie,F.Autorizacion,C.Grupo,A.Nombre_Completo As Ejecutivo," _
      & "C.Plan_Afiliado As Sectorizacion,F.Cta_CxP " _
      & "FROM Facturas As F,Clientes As C,Accesos As A " _
      & "WHERE F.Item = '" & NumEmpresa & "' " _
      & "AND F.Fecha <= #" & FechaFin & "# " _
      & Tipo_De_Consulta() _
      & "AND F.TC NOT IN ('C','P') " _
      & "AND F.CodigoC = C.Codigo " _
      & "AND A.Codigo = F.Cod_Ejec " _
      & "ORDER BY C.Cliente,F.Factura,F.Fecha "
 SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
 RatonNormal
 Totales_CxC_Abonos
End Sub

Public Sub Ventas_Productos()
'Resumen Ventas por Producto
Opcion = 8
RatonReloj
DGQuery.Visible = False
  Si_No = False
  Con_Costeo = " "
  Mensajes = "Reporte Con Costeo "
  Titulo = "Formulario de Confirmación"
  If BoxMensaje = vbYes Then
     If ClaveAdministrador Then Si_No = True
  End If
Label2.Caption = " Ventas"
Label4.Caption = "  "
Label3.Caption = "  "
DGQuery.Visible = False
PorCxC = False
PorProd = False
PorMarca = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True
If CheqMarca.value = 1 Then PorMarca = True

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT * " _
     & "FROM Catalogo_Productos " _
     & "WHERE TC = 'P' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv "
SelectAdodc AdoHistoria, sSQL

sSQL = "SELECT * " _
     & "FROM Trans_Kardex " _
     & "WHERE Fecha <= #" & FechaFin & "# " _
     & "AND T <> 'A' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv,Fecha,Kardex "
SelectAdodc AdoQuery, sSQL
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     FEsperar.Show
     Codigo = .Fields("Codigo_Inv")
     TotalIngreso = 0
     Do While Not .EOF
        If Codigo <> .Fields("Codigo_Inv") Then
           If AdoHistoria.Recordset.RecordCount > 0 Then
              AdoHistoria.Recordset.MoveFirst
              AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
              If Not AdoHistoria.Recordset.EOF Then
                 If AdoHistoria.Recordset.Fields("Valor_Compra") <> TotalIngreso Then
                    AdoHistoria.Recordset.Fields("Valor_Compra") = TotalIngreso
                    AdoHistoria.Recordset.Update
                 End If
              End If
           End If
           Codigo = .Fields("Codigo_Inv")
           TotalIngreso = 0
        End If
        TotalIngreso = .Fields("Valor_Unitario")
       .MoveNext
     Loop
     If AdoHistoria.Recordset.RecordCount > 0 Then
        AdoHistoria.Recordset.MoveFirst
        AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
        If Not AdoHistoria.Recordset.EOF Then
           If AdoHistoria.Recordset.Fields("Valor_Compra") <> TotalIngreso Then
              AdoHistoria.Recordset.Fields("Valor_Compra") = TotalIngreso
              AdoHistoria.Recordset.Update
           End If
        End If
     End If
     Unload FEsperar
 End If
End With
sSQL = "UPDATE Detalle_Factura " _
     & "SET Producto_Aux = Mid(Producto,1,50) " _
     & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' "
ConectarAdoExecute sSQL

sSQL = "SELECT F.T,CL.Cliente,F.Factura,F.Fecha,F.Codigo,F.Producto_Aux,F.Cantidad,F.Total " & Con_Costeo
If Si_No Then sSQL = sSQL & ",F.Precio,Valor_Compra As Costos "
sSQL = sSQL & "FROM Detalle_Factura As F,Catalogo_Productos As C,Clientes As CL " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND C.INV <> " & Val(adFalse) & " " _
     & Tipo_De_Consulta() _
     & "AND F.Item = C.Item " _
     & "AND F.Periodo = C.Periodo " _
     & "AND F.Codigo = C.Codigo_Inv " _
     & "AND F.CodigoC = CL.Codigo " _
     & "ORDER BY F.Factura,F.Fecha,F.Codigo,C.Producto "
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
'MsgBox sSQL
RatonReloj
DGQuery.Visible = False
Total = 0: Abono = 0: Saldo = 0
Total = T_Fields("Total")
If Si_No Then
   Abono = T_Fields("Precio")
   Saldo = T_Fields("Costos")
End If
'3301738
Label2.Caption = "Facturado"
Label4.Caption = "PVP"
Label3.Caption = "Costo"

LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Saldo, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Productos_Abonos()
'Productos y Abonos
Opcion = 2
RatonReloj
Label2.Caption = " Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
'Actualizar_Saldos_Facturas MBFechaF, TipoFactura
PorCxC = False
PorProd = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
Total = 0: Abono = 0
DGQuery.Visible = False

  sSQL = "UPDATE Facturas " _
       & "SET T = 'P' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' "
  ConectarAdoExecute sSQL

  sSQL = "UPDATE Facturas " _
       & "SET T = 'C' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Saldo_MN <= 0 " _
       & "AND T <> 'A' "
  ConectarAdoExecute sSQL

  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET T = F.T " _
          & "FROM Detalle_Factura As DF,Facturas As F "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Facturas As F " _
          & "SET DF.T = F.T "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.TC = F.TC " _
       & "AND DF.Serie = F.Serie " _
       & "AND DF.Factura = F.Factura " _
       & "AND DF.Autorizacion = F.Autorizacion " _
       & "AND DF.CodigoC = F.CodigoC " _
       & "AND DF.Periodo = F.Periodo " _
       & "AND DF.Item = F.Item " _
       & "AND DF.T <> F.T "
  ConectarAdoExecute sSQL

  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET T = F.T " _
          & "FROM Trans_Abonos As DF,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As DF,Facturas As F " _
          & "SET DF.T = F.T "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.TP = F.TC " _
       & "AND DF.Serie = F.Serie " _
       & "AND DF.Factura = F.Factura " _
       & "AND DF.Autorizacion = F.Autorizacion " _
       & "AND DF.CodigoC = F.CodigoC " _
       & "AND DF.Periodo = F.Periodo " _
       & "AND DF.Item = F.Item " _
       & "AND DF.T <> F.T "
  ConectarAdoExecute sSQL

''  If SQL_Server Then
''     sSQL = "UPDATE Detalle_Factura " _
''          & "SET Producto = TB.Producto " _
''          & "FROM Detalle_Factura As T,Catalogo_Productos As TB "
''  Else
''     sSQL = "UPDATE Detalle_Factura As T,Catalogo_Productos As TB " _
''          & "SET T.Producto = TB.Producto "
''  End If
''  sSQL = sSQL _
''       & "WHERE T.Item = '" & NumEmpresa & "' " _
''       & "AND T.Periodo = '" & Periodo_Contable & "' " _
''       & "AND T.Producto = '.' " _
''       & "AND T.Codigo = TB.Codigo_Inv " _
''       & "AND T.Periodo = TB.Periodo " _
''       & "AND T.Item = TB.Item " _
''       & "AND T.Producto <> TB.Producto "
''   ConectarAdoExecute sSQL

sSQL = "SELECT F.TC As Tipo,C.Cliente,Factura,F.Fecha,Cantidad,Producto As Producto_,Total As Ventas,0 As Abonos," _
     & "F.Serie,F.Autorizacion " _
     & "FROM Detalle_Factura As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & Tipo_De_Consulta() _
     & "AND F.CodigoC = C.Codigo "
SQL1 = "UNION " _
     & "SELECT F.TC As Tipo,C.Cliente,Factura,F.Fecha,0 As Cantidad,'* TOTAL I.V.A. ' As Producto_,IVA As Ventas,0 As Abonos," _
     & "F.Serie,F.Autorizacion " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.IVA > 0 " _
     & Tipo_De_Consulta() _
     & "AND F.TC NOT IN ('C','P') " _
     & "AND F.CodigoC = C.Codigo "
SQL1 = SQL1 & "UNION " _
     & "SELECT F.TC As Tipo,C.Cliente,Factura,F.Fecha,0 As Cantidad,'* TOTAL SERVICIO ' As Producto_,Servicio As Ventas,0 As Abonos," _
     & "F.Serie,F.Autorizacion " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Servicio > 0 " _
     & Tipo_De_Consulta() _
     & "AND F.TC NOT IN ('C','P') " _
     & "AND F.CodigoC = C.Codigo "
SQL2 = "UNION " _
     & "SELECT F.TC As Tipo,C.Cliente,Factura,F.Fecha,0 As Cantidad,'* TOTAL DESCUENTO ' As Producto_,-F.Descuento As Ventas,0 As Abonos," _
     & "F.Serie,F.Autorizacion " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Descuento > 0 " _
     & Tipo_De_Consulta() _
     & "AND F.TC NOT IN ('C','P') " _
     & "AND F.CodigoC = C.Codigo "
SQL2 = SQL2 & "UNION " _
     & "SELECT F.TP As Tipo,C.Cliente,F.Factura,F.Fecha,0 As Cantidad,('- ' & Banco & ' No. ' & Cheque) As Producto_,0 As Ventas,F.Abono As Abonos," _
     & "F.Serie,F.Autorizacion " _
     & "FROM Trans_Abonos As F,Clientes As C,Facturas As F1 " _
     & "WHERE F.Fecha <= #" & FechaFin & "# " _
     & "AND F1.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Abono > 0 " _
     & Tipo_De_Consulta(True) _
     & "AND F1.TC NOT IN ('C','P') " _
     & "AND F.CodigoC = C.Codigo " _
     & "AND F1.CodigoC = F.CodigoC " _
     & "AND F1.TC = F.TP " _
     & "AND F1.Serie = F.Serie " _
     & "AND F1.Autorizacion = F.Autorizacion " _
     & "AND F1.Factura = F.Factura " _
     & "AND F1.T = F.T " _
     & "ORDER BY C.Cliente,Tipo,F.Factura,F.Fecha,F.Cantidad DESC,Producto_ DESC "
sSQL = sSQL & SQL1 & SQL2
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
Total = T_Fields("Ventas")
Abono = T_Fields("Abonos")
'With AdoQuery.Recordset
' If .RecordCount > 0 Then
'     ProgBar.Max = .RecordCount + 1
'     Do While Not .EOF
'        Total = Total + .Fields("Ventas")
'        Abono = Abono + .Fields("Abonos")
'        ProgBar.Value = ProgBar.Value + 1
'       .MoveNext
'     Loop
'    .MoveFirst
' End If
'End With
Saldo = Total - Abono
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Total - Abono, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Abonos_Facturas(Optional Ret_NC As Boolean)
RatonReloj
Opcion = 6
Label2.Caption = " Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
PorProd = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True
   DGQuery.Caption = "ABONOS DE FACTURAS"
   If SQL_Server Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Mes = DF.Mes, Mes_No = DF.Mes_No " _
           & "FROM Trans_Abonos As TA,Detalle_Factura As DF "
   Else
      sSQL = "UPDATE Trans_Abonos As TA,Detalle_Factura As DF " _
           & "SET TA.Mes = DF.Mes, TA.Mes_No = DF.Mes_No "
   End If
   sSQL = sSQL _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.T <> 'A' " _
        & "AND TA.Item = DF.Item " _
        & "AND TA.Periodo = DF.Periodo " _
        & "AND TA.Factura = DF.Factura " _
        & "AND TA.Serie = DF.Serie " _
        & "AND TA.Autorizacion = DF.Autorizacion " _
        & "AND TA.CodigoC = DF.CodigoC "
   ConectarAdoExecute sSQL
       
   Total = 0
  'Asientos de CxC Cheque
   If Ret_NC Then
      sSQL = "SELECT F.TP,F.Fecha,C.Cliente,F.Factura,F.Banco,F.Cheque,F.Abono,F.Mes,F.Comprobante,F.Serie,F.Autorizacion,F.Base_Imponible,F.Porc,C.Representante As Razon_Social,F.Cta,F.Cta_CxP " _
           & "FROM Trans_Abonos As F,Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & Tipo_De_Consulta(True) _
           & "AND Mid(F.Banco,1,9) = 'RETENCION' " _
           & "AND F.CodigoC = C.Codigo " _
           & "UNION " _
           & "SELECT F.TP,F.Fecha,C.Cliente,F.Factura,F.Banco,F.Cheque,F.Abono,F.Mes,F.Comprobante,F.Serie,F.Autorizacion,F.Base_Imponible,F.Porc,C.Representante As Razon_Social,F.Cta,F.Cta_CxP " _
           & "FROM Trans_Abonos As F,Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & Tipo_De_Consulta(True) _
           & "AND F.Banco = 'NOTA DE CREDITO' " _
           & "AND F.CodigoC = C.Codigo " _
           & "ORDER BY F.Banco,F.Cheque,C.Cliente,F.Factura,F.Fecha "
   Else
      sSQL = "SELECT F.TP,F.Fecha,C.Cliente,F.Factura,F.Banco,F.Cheque,F.Abono,F.Mes,F.Comprobante,F.Serie,F.Autorizacion,C.Representante As Razon_Social,F.Cta,F.Cta_CxP " _
           & "FROM Trans_Abonos As F,Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & Tipo_De_Consulta(True) _
           & "AND F.CodigoC = C.Codigo " _
           & "ORDER BY C.Cliente,F.Factura,F.Fecha,F.Banco "
   End If
   SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
   Total = T_Fields("Abono")
   LabelAbonado.Caption = Format(Total, "#,##0.00")
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
   RatonNormal
End Sub

Public Sub Abonos_Anticipados()
RatonReloj
Opcion = 6
Label2.Caption = " Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
PorProd = False
   DGQuery.Caption = "ABONOS DE ANTICIPADOS"
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT TA.TP,F.Serie,F.Autorizacion,F.Fecha,F.Factura," _
        & "TA.Fecha As Fecha_Abono,TA.Abono " _
        & "FROM Facturas As F, Trans_Abonos As TA " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC = TA.TP " _
        & "AND F.Serie = TA.Serie " _
        & "AND F.Autorizacion = TA.Autorizacion " _
        & "AND F.Factura = TA.Factura " _
        & "AND F.Fecha > TA.Fecha " _
        & "ORDER BY TA.Fecha,F.Factura "
   SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
   LabelAbonado.Caption = "0.00"
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
   RatonNormal
End Sub

Public Sub Abonos_Erroneos()
RatonReloj
Opcion = 6
Label2.Caption = " Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
PorProd = False
   DGQuery.Caption = "ABONOS MAL PROCESADOS"
   Total = 0
   sSQL = "UPDATE Trans_Abonos " _
        & "SET X = 'E' " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
   ConectarAdoExecute sSQL
   If SQL_Server Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET X = '.' " _
           & "FROM Trans_Abonos As TA,Facturas As F "
   Else
      sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
           & "SET TA.X = '.' "
   End If
   sSQL = sSQL & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC = TA.TP " _
        & "AND F.Serie = TA.Serie " _
        & "AND F.Autorizacion = TA.Autorizacion " _
        & "AND F.Factura = TA.Factura " _
        & "AND F.CodigoC = TA.CodigoC "
   ConectarAdoExecute sSQL
  'Asientos de CxC Cheque
   sSQL = "SELECT TP,Serie,Autorizacion,Fecha,Factura,Abono " _
        & "FROM Trans_Abonos " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND X = 'E' " _
        & "AND TP <> 'CB' " _
        & "ORDER BY Fecha,Factura "
   SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
   LabelAbonado.Caption = "0.00"
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
   RatonNormal
End Sub

Public Sub Resumen_Productos()
'Resumen Productos
Opcion = 3
RatonReloj
Label2.Caption = " I.V.A."
Label4.Caption = " VENTAS"
Label3.Caption = " TOTAL"
DGQuery.Visible = False
PorCxC = False
PorProd = False
PorMarca = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True
If CheqMarca.value = 1 Then PorMarca = True


DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT C.Cliente,SUM(F.Cantidad) As Cant_Prod,CP.Producto,F.Codigo,SUM(F.Total_IVA) As IVA," _
     & "SUM(F.Total) As Ventas,SUM(F.Cantidad*CP.Gramaje/1000) As Kilos,CP.Gramaje " _
     & "FROM Clientes As C,Detalle_Factura As F,Catalogo_Productos AS CP " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & Tipo_De_Consulta() _
     & "AND F.CodigoC = C.Codigo " _
     & "AND F.Item = CP.Item " _
     & "AND F.Periodo = CP.Periodo " _
     & "AND F.Codigo = CP.Codigo_Inv " _
     & "GROUP BY C.Cliente,F.Codigo,F.CodigoC,CP.Producto,CP.Gramaje " _
     & "ORDER BY C.Cliente,F.Codigo,F.CodigoC,CP.Producto,CP.Gramaje "
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
DGQuery.Visible = False
Total = 0: Abono = 0
Total = T_Fields("IVA")
Abono = T_Fields("Ventas")
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Total + Abono, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Ventas_Cliente()
'Resumen Ventas por Cliente
Opcion = 4
RatonReloj
Label2.Caption = " Ventas"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
PorProd = False
PorMarca = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True
If CheqMarca.value = 1 Then PorMarca = True

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT C.Cliente,F.TC,COUNT(F.CodigoC) As Cant_Fact,SUM(F.Total) As Ventas," _
     & "SUM(F.Total_IVA) As I_V_A,SUM(F.Total + F.Total_IVA) As Total_Facturado " _
     & "FROM Detalle_Factura As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & Tipo_De_Consulta() _
     & "AND F.CodigoC = C.Codigo " _
     & "GROUP BY C.Cliente,F.TC " _
     & "ORDER BY SUM(F.Total + F.Total_IVA) DESC,C.Cliente "
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
RatonReloj
DGQuery.Visible = False
Total = 0: Abono = 0
Total = T_Fields("Ventas")
Abono = T_Fields("I_V_A")
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Total - Abono, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Resumen_Ventas_Costos()
'Resumen Ventas por Producto
Opcion = 5
RatonReloj
DGQuery.Visible = False
  Si_No = False
  Con_Costeo = " "
  Mensajes = "Reporte Con Costeo "
  Titulo = "Formulario de Confirmación"
  If BoxMensaje = vbYes Then
     If ClaveAdministrador Then Si_No = True
  End If
Label2.Caption = " Ventas"
Label4.Caption = "  "
Label3.Caption = "  "
DGQuery.Visible = False
PorCxC = False
PorProd = False
PorMarca = False
If CheqCxC.value = 1 Then PorCxC = True
If CheqProd.value = 1 Then PorProd = True
If CheqMarca.value = 1 Then PorMarca = True

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT * " _
     & "FROM Catalogo_Productos " _
     & "WHERE TC = 'P' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv "
SelectAdodc AdoHistoria, sSQL

sSQL = "SELECT * " _
     & "FROM Trans_Kardex " _
     & "WHERE Fecha <= #" & FechaFin & "# " _
     & "AND T <> 'A' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv,Fecha,Kardex "
SelectAdodc AdoQuery, sSQL
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     FEsperar.Show
     Codigo = .Fields("Codigo_Inv")
     TotalIngreso = 0
     Do While Not .EOF
        If Codigo <> .Fields("Codigo_Inv") Then
           If AdoHistoria.Recordset.RecordCount > 0 Then
              AdoHistoria.Recordset.MoveFirst
              AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
              If Not AdoHistoria.Recordset.EOF Then
                 If AdoHistoria.Recordset.Fields("Valor_Compra") <> TotalIngreso Then
                    AdoHistoria.Recordset.Fields("Valor_Compra") = TotalIngreso
                    AdoHistoria.Recordset.Update
                 End If
              End If
           End If
           Codigo = .Fields("Codigo_Inv")
           TotalIngreso = 0
        End If
        TotalIngreso = .Fields("Valor_Unitario")
       .MoveNext
     Loop
     If AdoHistoria.Recordset.RecordCount > 0 Then
        AdoHistoria.Recordset.MoveFirst
        AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
        If Not AdoHistoria.Recordset.EOF Then
           If AdoHistoria.Recordset.Fields("Valor_Compra") <> TotalIngreso Then
              AdoHistoria.Recordset.Fields("Valor_Compra") = TotalIngreso
              AdoHistoria.Recordset.Update
           End If
        End If
     End If
     Unload FEsperar
 End If
End With
sSQL = "SELECT F.Codigo,CP.Producto,SUM(F.Cantidad) As Cant_Prod,SUM(F.Total) As Ventas," _
     & "SUM(F.Cantidad*CP.Gramaje/1000) As Kilos,CP.Desc_Item " & Con_Costeo
If Si_No Then sSQL = sSQL & ",AVG(F.Precio) As PVP,Valor_Compra As Costos "
sSQL = sSQL & "FROM Detalle_Factura As F,Catalogo_Productos As CP,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND CP.INV <> " & Val(adFalse) & " " _
     & Tipo_De_Consulta()
If CheqDescItem.value <> 0 Then sSQL = sSQL & "AND CP.Desc_Item = '" & DCMarca & "' "
sSQL = sSQL & "AND F.Item = CP.Item " _
     & "AND F.Periodo = CP.Periodo " _
     & "AND F.Codigo = CP.Codigo_Inv " _
     & "AND F.CodigoC = C.Codigo "
If Si_No Then
   If CheqDescItem.value <> 0 Then
      sSQL = sSQL & "GROUP BY CP.Desc_Item,F.Codigo,CP.Valor_Compra "
   Else
      sSQL = sSQL & "GROUP BY F.Codigo,CP.Valor_Compra,CP.Producto,CP.Desc_Item "
   End If
Else
   If CheqDescItem.value <> 0 Then
      sSQL = sSQL & "GROUP BY CP.Desc_Item,F.Codigo,CP.Producto "
   Else
      sSQL = sSQL & "GROUP BY F.Codigo,CP.Producto,CP.Desc_Item "
   End If
End If
If CheqDescItem.value <> 0 Then
   sSQL = sSQL & "ORDER BY CP.Desc_Item,F.Codigo,SUM(F.Total) DESC "
Else
   sSQL = sSQL & "ORDER BY F.Codigo,SUM(F.Total) DESC,CP.Producto "
End If
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
'MsgBox sSQL
RatonReloj
DGQuery.Visible = False
Total = 0: Abono = 0: Saldo = 0
Total = T_Fields("Ventas")
If Si_No Then
   Abono = T_Fields("PVP")
   Saldo = T_Fields("Costos")
End If
'3301738
Label2.Caption = "Facturado"
Label4.Caption = "PVP"
Label3.Caption = "Costo"

LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Abono, "#,##0.00")
LabelSaldo.Caption = Format(Saldo, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Cheques_Protestados()
'Cheques protestado
RatonReloj
Label2.Caption = " Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"

   DGQuery.Caption = "ABONOS DE FACTURAS"
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT F.TP,F.Fecha,C.Cliente,F.Factura,F.Banco,F.Cheque,F.Abono,F.Comprobante,F.Cta,F.Cta_CxP " _
        & "FROM Trans_Abonos As F,Clientes C " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & Tipo_De_Consulta() _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.Protestado <> " & Val(adFalse) & " " _
        & "ORDER BY C.Cliente,F.Factura,F.Fecha,F.Banco "
   SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
   Opcion = 7
   Totales_CxC_Abonos
   RatonNormal
End Sub

Private Sub Command1_Click()
  Unload HistorialFacturas
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto HistorialFacturas, AdoQuery
     DGQuery.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyB Then
     'CampoBusqueda
     BuscarDatos DGQuery, AdoQuery
  End If
End Sub

Private Sub Form_Activate()
Dim OpcBusqueda() As String
Dim TempOpcBusqueda As String
Dim CantOpcBusqueda As Byte
   HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
   sSQL = "SELECT (Codigo_Inv & ' - ' & Producto) As Codigos " _
        & "FROM Catalogo_Productos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Codigo_Inv "
   SelectDBCombo DCProducto, AdoProducto, sSQL, "Codigos"
   
   sSQL = "SELECT C.Cliente,C.CI_RUC,C.Ciudad,C.Codigo,C.Grupo,C.Plan_Afiliado,Count(F.Factura) As Fact_Proc " _
        & "FROM Clientes As C,Facturas As F " _
        & "WHERE Cliente = '.' " _
        & "AND C.Codigo = F.CodigoC " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "GROUP BY C.Codigo,C.CI_RUC,C.Cliente,C.Ciudad,C.Grupo,C.Plan_Afiliado " _
        & "ORDER BY C.Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
   
   CantOpcBusqueda = AdoCliente.Recordset.Fields.Count + 2
   ReDim OpcBusqueda(CantOpcBusqueda) As String
   With AdoCliente.Recordset
    For I = 0 To .Fields.Count - 2
        OpcBusqueda(I) = .Fields(I).Name
    Next I
   End With
   OpcBusqueda(I) = "Tipo Documento"
   OpcBusqueda(I + 1) = "Factura"
   OpcBusqueda(I + 2) = "Forma_Pago"
   For I = 0 To CantOpcBusqueda - 1
       For J = I + 1 To CantOpcBusqueda - 1
           If OpcBusqueda(I) > OpcBusqueda(J) Then
              TempOpcBusqueda = OpcBusqueda(I)
              OpcBusqueda(I) = OpcBusqueda(J)
              OpcBusqueda(J) = TempOpcBusqueda
           End If
       Next J
   Next I
   ListCliente.Clear
   For I = 0 To CantOpcBusqueda - 1
       ListCliente.AddItem OpcBusqueda(I)
   Next I
   ListCliente.AddItem "Todos"
   ListCliente.Text = "Todos"
   
   If TipoFactura = "" Then TipoFactura = Ninguno
   
   sSQL = "SELECT (Codigo & ' - ' & Concepto) As NomCxC " _
        & "FROM Catalogo_Lineas " _
        & "WHERE TL <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Codigo "
   SelectDBCombo DCCxC, AdoCxC, sSQL, "NomCxC"
   
   sSQL = "SELECT (CodMar & ' - ' & Marca) As NomMarca " _
        & "FROM Catalogo_Marcas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND CodMar <> '" & Ninguno & "' " _
        & "ORDER BY Marca "
   SelectDBCombo DCMarca, AdoMarca, sSQL, "NomMarca"
   
   DCCliente.Visible = False
   
   DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
   DGQuery.width = MDI_X_Max - 100
   AdoQuery.Top = DGQuery.Top + DGQuery.Height
   Command1.Top = DGQuery.Top + DGQuery.Height
   Label2.Top = DGQuery.Top + DGQuery.Height
   Label3.Top = DGQuery.Top + DGQuery.Height
   Label4.Top = DGQuery.Top + DGQuery.Height
   LabelFacturado.Top = DGQuery.Top + DGQuery.Height
   LabelAbonado.Top = DGQuery.Top + DGQuery.Height
   LabelSaldo.Top = DGQuery.Top + DGQuery.Height
   RatonNormal
   MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
   'CentrarForm HistorialFacturas
   ConectarAdodc AdoCxC
   ConectarAdodc AdoTipo
   ConectarAdodc AdoMarca
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCliente
   ConectarAdodc AdoHistoria
   ConectarAdodc AdoFacturas
   ConectarAdodc AdoProducto
End Sub

Private Sub ListCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub ListCliente_LostFocus()
Dim Nombre_Campo As String
  Nombre_Campo = ListCliente.Text
  DCCliente.Visible = True
  Select Case ListCliente.Text
    Case "Codigo"
         sSQL = "SELECT Codigo,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Codigo " _
              & "ORDER BY Codigo "
         Nombre_Campo = "Codigo"
    Case "CI_RUC"
         sSQL = "SELECT CI_RUC,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY CI_RUC " _
              & "ORDER BY CI_RUC "
         Nombre_Campo = "CI_RUC"
    Case "Cliente"
         sSQL = "SELECT Cliente,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Cliente " _
              & "ORDER BY Cliente "
         Nombre_Campo = "Cliente"
    Case "Ciudad"
         sSQL = "SELECT Ciudad,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Ciudad " _
              & "ORDER BY Ciudad "
         Nombre_Campo = "Ciudad"
    Case "Grupo"
         sSQL = "SELECT Grupo,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Grupo " _
              & "ORDER BY Grupo "
         Nombre_Campo = "Grupo"
    Case "Factura"
         sSQL = "SELECT Factura " _
              & "FROM Facturas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND TC NOT IN ('C','P') " _
              & "GROUP BY Factura " _
              & "ORDER BY Factura DESC "
         Nombre_Campo = "Factura"
    Case "Forma_Pago"
         sSQL = "SELECT Forma_Pago " _
              & "FROM Facturas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND TC NOT IN ('C','P') " _
              & "GROUP BY Forma_Pago " _
              & "ORDER BY Forma_Pago "
         Nombre_Campo = "Forma_Pago"
    Case "Tipo Documento"
         sSQL = "SELECT TC,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "AND TC NOT IN ('C','P') " _
              & "GROUP BY TC " _
              & "ORDER BY TC "
         Nombre_Campo = "TC"
    Case "Plan_Afiliado"
         sSQL = "SELECT Plan_Afiliado,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "AND LEN(C.Plan_Afiliado) > 3 " _
              & "AND TC NOT IN ('C','P') " _
              & "AND C.Codigo = F.CodigoC " _
              & "GROUP BY Plan_Afiliado " _
              & "ORDER BY Plan_Afiliado "
         Nombre_Campo = "Plan_Afiliado"
    Case Else
         DCCliente.Visible = False
         sSQL = "SELECT Cliente,Count(Factura) As Fact_Proc " _
              & "FROM Clientes As C,Facturas As F " _
              & "WHERE C.Codigo = F.CodigoC " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Cliente " _
              & "ORDER BY Cliente "
         Nombre_Campo = "Cliente"
  End Select
  SelectDBCombo DCCliente, AdoCliente, sSQL, Nombre_Campo
  If DCCliente.Visible Then DCCliente.SetFocus
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
  FechaValida MBFechaI
End Sub

Public Sub ProcesarFactura(Optional EsSaldo As Boolean)
   SetAdoAddNew "Saldo_Diarios"
   SetAdoFields "ME", CByte(Contador)
   SetAdoFields "TP", "HIST"
   SetAdoFields "CodigoU", CodigoUsuario
   SetAdoFields "Comprobante", NombreCliente
   SetAdoFields "Cta", CICliente
   SetAdoFields "Item", NumEmpresa
   SetAdoFields "T", TipoDoc
   SetAdoFields "TC", TipoProc
   SetAdoFields "CodigoC", CodigoCliente
   SetAdoFields "Numero", Factura_No
   SetAdoFields "Fecha", Mifecha
   SetAdoFields "Fecha_Venc", FechaTexto
   SetAdoFields "Recibo", TipoProc & ": " & Format(Factura_No, "0000000")
   If EsSaldo Then SetAdoFields "Total", Total
   SetAdoFields "Egresos", Abono
   SetAdoFields "Saldo_Actual", Saldo
   SetAdoUpdate
End Sub

Public Sub ProcesarProducto(Optional EsSaldo As Boolean)
   SetAdoAddNew "Saldo_Diarios"
   SetAdoFields "ME", CInt(Contador)
   SetAdoFields "TP", "PROD"
   SetAdoFields "CodigoU", CodigoUsuario
   SetAdoFields "Comprobante", NombreCliente
   SetAdoFields "Item", NumEmpresa
   SetAdoFields "T", TipoDoc
   SetAdoFields "CodigoC", CodigoCliente
   SetAdoFields "Numero", Factura_No
   SetAdoFields "Fecha", Mifecha
   SetAdoFields "Fecha_Venc", FechaTexto
   SetAdoFields "Cta", NoCheque
   SetAdoFields "Total", Total
   SetAdoFields "Ingresos", Total_IVA
   SetAdoFields "Egresos", Abono
   SetAdoFields "Saldo_Actual", Saldo
   SetAdoUpdate
End Sub

Public Sub ProcesarProductoVentas(Optional EsSaldo As Boolean)
   SetAdoAddNew "Saldo_Diarios"
   SetAdoFields "ME", CInt(Contador)
   SetAdoFields "TP", "PROD"
   SetAdoFields "CodigoU", CodigoUsuario
   SetAdoFields "Comprobante", NombreCliente
   SetAdoFields "Item", NumEmpresa
   SetAdoFields "T", TipoDoc
   SetAdoFields "CodigoC", CodigoCliente
   SetAdoFields "Numero", Factura_No
   SetAdoFields "Fecha", Mifecha
   SetAdoFields "Fecha_Venc", FechaTexto
   SetAdoFields "Cta", NoCheque
   'If EsSaldo Then
   SetAdoFields "Saldo_Anterior", Cantidad
   SetAdoFields "Total", Total
   SetAdoFields "Ingresos", Total_IVA
   SetAdoFields "Egresos", Abono
   SetAdoFields "Saldo_Actual", Saldo
   SetAdoUpdate
End Sub

Public Sub Tipo_Consulta_CxC(Tipo As String)
Dim Actualiza_Buses As Boolean
Actualiza_Buses = Leer_Campo_Empresa("Actualizar_Buses")
DGQuery.Caption = ""
Actualizar_Saldos_Facturas MBFechaF
RatonReloj
If Actualiza_Buses Then
   If SQL_Server Then
      sSQL = "UPDATE Facturas " _
           & "SET Forma_Pago = MID(DF.Producto,1,10) " _
           & "FROM Facturas AS F,Detalle_Factura AS DF "
   Else
      sSQL = "UPDATE Facturas AS F,Detalle_Factura AS DF " _
           & "SET F.Forma_Pago = MID(DF.Producto,1,10) "
   End If
   sSQL = sSQL _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC = DF.TC " _
        & "AND F.Serie = DF.Serie " _
        & "AND F.Autorizacion = DF.Autorizacion " _
        & "AND F.CodigoC = DF.CodigoC " _
        & "AND F.Fecha = DF.Fecha " _
        & "AND F.Factura = DF.Factura " _
        & "AND F.Item = DF.Item " _
        & "AND F.Periodo = DF.Periodo "
   ConectarAdoExecute sSQL
End If
If TipoFactura = "" Then TipoFactura = Ninguno
If Tipo = "C" Then
   sSQL = "SELECT F.T,F.Razon_Social,C.Cliente,F.Fecha,F.Factura,"
   Opcion = 9
ElseIf Tipo = "F" Then
   sSQL = "SELECT F.T,F.Fecha,F.Factura,F.Razon_Social,C.Cliente,"
   Opcion = 10
End If
sSQL = sSQL _
     & "F.Total_MN,(F.Total_MN-F.Saldo_MN) As Abono_MN,F.Saldo_MN," _
     & "F.Total_ME,F.Saldo_ME,C.CI_RUC,F.Forma_Pago,F.RUC_CI As RUC_CI_SRI,C.Telefono,C.Celular,C.Ciudad,C.Direccion,C.DireccionT As DireccionT,C.Email,"
If SQL_Server Then
   sSQL = sSQL & "F.Fecha_V,DATEDIFF(day,F.Fecha,'" & BuscarFecha(MBFechaF) & "') As Dias_De_Mora,"
Else
   sSQL = sSQL & "F.Fecha_V,DATEDIFF('d',F.Fecha,#" & BuscarFecha(MBFechaF) & "#) As Dias_De_Mora,"
End If
sSQL = sSQL & "F.TC,F.Serie,F.Autorizacion,A.Nombre_Completo As Ejecutivo,C.Plan_Afiliado As Sectorizacion " _
     & "FROM Facturas As F,Clientes As C,Accesos As A " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND A.Codigo = F.Cod_Ejec " _
     & "AND F.TC NOT IN ('C','P') " _
     & Tipo_De_Consulta()
sSQL = sSQL & "AND F.Periodo = '" & Periodo_Contable & "' "
If Tipo = "C" Then sSQL = sSQL & "ORDER BY C.Cliente,F.Fecha,F.Factura "
If Tipo = "F" Then sSQL = sSQL & "ORDER BY F.Factura,F.Fecha "
SelectDataGrid DGQuery, AdoQuery, sSQL, , , True
Total = 0: Saldo = 0
If Tipo = "C" Then
   Total = T_Fields("Total_MN")
   Saldo = T_Fields("Saldo_MN")
Else
   Total = T_Fields("Total_MN")
   Saldo = T_Fields("Saldo_MN")
End If
DGQuery.Visible = True
TipoDoc = Tipo
'LabelFacturado.Caption = Format(Total, "#,##0.00")
'LabelAbonado.Caption = Format(Saldo, "#,##0.00")
Totales_CxC_Abonos
RatonNormal
End Sub

Public Function Tipo_De_Consulta(Optional Opcion_TP As Boolean) As String
Dim SQL3X As String
Dim Patron_Busqueda As String
  'MsgBox Opcion
   Patron_Busqueda = DCCliente.Text
   If Patron_Busqueda = "" Then Patron_Busqueda = Ninguno
   Cta_Cobrar = Trim(SinEspaciosIzq(DCCxC.Text))
   Cod_Marca = Trim(SinEspaciosIzq(DCMarca.Text))
   Codigo_Inv = Trim(SinEspaciosIzq(DCProducto))
   Nivel_No = Val(TxtNivel)
  'Encabezado de Factura
   SQL3X = ""
   If OpcPend.value Then
      If Opcion = 1 Then
         Select Case Opcion
           Case 1: SQL3X = SQL3X & "AND Saldo_Actual <> 0 "
           Case 9, 10: SQL3X = SQL3X & "AND Saldo_MN <> 0 "
         End Select
      Else
         SQL3X = SQL3X & "AND F.T = '" & Pendiente & "' "
      End If
   ElseIf OpcCanc.value Then
      SQL3X = SQL3X & "AND F.T = '" & Cancelado & "' "
   ElseIf OpcAnul.value Then
      SQL3X = SQL3X & "AND F.T = '" & Anulado & "' "
   Else
      SQL3X = SQL3X & "AND F.T <> '" & Anulado & "' "
   End If
   Select Case ListCliente.Text
     Case "Codigo"
          SQL3X = SQL3X & "AND C.Codigo = '" & Patron_Busqueda & "' "
     Case "CI_RUC"
          SQL3X = SQL3X & "AND C.CI_RUC = '" & Patron_Busqueda & "' "
     Case "Cliente"
          LongStrg = Len(Patron_Busqueda)
          SQL3X = SQL3X & "AND Ucase(Mid(C.Cliente,1," & LongStrg & ")) = '" & Patron_Busqueda & "' "
     Case "Ciudad"
          SQL3X = SQL3X & "AND C.Ciudad = '" & Patron_Busqueda & "' "
     Case "Grupo"
          SQL3X = SQL3X & "AND C.Grupo = '" & Patron_Busqueda & "' "
     Case "Factura"
          SQL3X = SQL3X & "AND F.Factura = " & Val(Patron_Busqueda) & " "
     Case "Forma_Pago"
          SQL3X = SQL3X & "AND F.Forma_Pago = '" & Patron_Busqueda & "' "
     Case "Plan_Afiliado"
          SQL3X = SQL3X & "AND C.Plan_Afiliado = '" & Patron_Busqueda & "' "
     Case "Tipo Documento"
          If Opcion_TP Then
             SQL3X = SQL3X & "AND F.TP = '" & Patron_Busqueda & "' "
          Else
             SQL3X = SQL3X & "AND F.TC = '" & Patron_Busqueda & "' "
          End If
          TipoFactura = Patron_Busqueda
   End Select
   SQL3X = SQL3X & "AND F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' "
        
  'Detalle Factura:
   If CheqDescItem.value = 1 Then
      If Nivel_No > 0 Then
         I = 1
         J = 1
         Contador = 0
         CNivel_1 = True
         For K = 1 To Len(Codigo_Inv)
             If Mid(Codigo, K, 1) = "." Then
                Contador = Contador + 1
                If CNivel_1 Then
                   CNivel_1 = Not (CNivel_1)
                   J = K - 1
                Else
                   I = J + 2
                   J = K - 1
                End If
                If Contador = Val(TxtNivel) Then
                   SQL3X = SQL3X & "AND Mid(F.Codigo," & I & "," & (J - I + 1) & ") = '" & Mid(Codigo_Inv, I, J - I + 1) & "' "
                End If
             End If
         Next K
         I = J + 2
         J = K - 1
         Contador = Contador + 1
         If Contador = Val(TxtNivel) Then
            SQL3X = SQL3X & "AND Mid(F.Codigo," & I & "," & (J - I + 1) & ") = '" & Mid(Codigo, I, J - I + 1) & "' "
         End If
      Else
         SQL3X = SQL3X & "AND Mid(F.Codigo,1," & Len(Codigo) & ") = '" & Codigo & "' "
      End If
   End If
   If CheqCxC.value = 1 Then SQL3X = SQL3X & "AND F.Cta_CxP = '" & Cta_Cobrar & "' "
   If CheqAbonos.value = 1 Then SQL3X = SQL3X & "AND F.Cta = '" & Cta_Cobrar & "' "
   If CheqMarca.value = 1 Then SQL3X = SQL3X & "AND F.CodMarca = '" & Cod_Marca & "' "
  'MsgBox SQL3X
   Tipo_De_Consulta = SQL3X
End Function

Public Sub Totales_CxC_Abonos()
  RatonReloj
  DGQuery.Visible = False
  Total = 0
  Abono = 0
  Saldo = 0
  Select Case Opcion
    Case 1
         Total = T_Fields("Total")
         Abono = T_Fields("Abono")
         Saldo = Total - Abono
         LabelFacturado.Caption = Format(Total, "#,##0.00")
         LabelAbonado.Caption = Format(Abono, "#,##0.00")
         LabelSaldo.Caption = Format(Saldo, "#,##0.00")
    Case 7
         Total = T_Fields("Abono")
         LabelFacturado.Caption = Format(Total, "#,##0.00")
         LabelAbonado.Caption = Format(Abono, "#,##0.00")
         LabelSaldo.Caption = Format(Saldo, "#,##0.00")
    Case 9, 10
         Total = T_Fields("Total_MN")
         Saldo = T_Fields("Saldo_MN")
         Abono = Total - Saldo
         LabelFacturado.Caption = Format(Total, "#,##0.00")
         LabelAbonado.Caption = Format(Abono, "#,##0.00")
         LabelSaldo.Caption = Format(Saldo, "#,##0.00")
  End Select
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub ToolbarMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
 'MsgBox Button.key
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Mifecha = FechaIni
  FechaTexto = FechaFin
  Select Case Button.key
    Case "Salir"
         Unload HistorialFacturas
    Case "Imprimir"
         Impresiones
    Case "Facturas"
         Historico_Facturas
    Case "Productos"
         Productos_Abonos
    Case "Resumen"
         Resumen_Productos
    Case "Clientes"
         Ventas_Cliente
    Case "Costos"
         Resumen_Ventas_Costos
    Case "VentasxProductos"
         Ventas_Productos
    Case "Abonos"
         Abonos_Facturas
    Case "Protestado"
         Cheques_Protestados
    Case "Por_Clientes"
         Tipo_Consulta_CxC "C"
    Case "Por_Facturas"
         Tipo_Consulta_CxC "F"
    Case "Retenciones_NC"
         Abonos_Facturas True
    Case "Abonos_Ant"
         Abonos_Anticipados
    Case "Abonos_Erroneos"
         Abonos_Erroneos
    Case "CxC_Clientes"
         Listado_Facturas_Por_Meses True
    Case "Listar_Por_Meses"
         Listado_Facturas_Por_Meses False
    Case "Estado_Cuenta_Cliente"
         Estado_Cuenta_Cliente
  End Select
  If Button.key <> "Salir" Then Command1.Caption = "&S(" & Opcion & ")"
End Sub

Private Sub TxtNivel_GotFocus()
  MarcarTexto TxtNivel
End Sub

Private Sub TxtNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNivel_LostFocus()
  TextoValido TxtNivel
  If Not IsNumeric(TxtNivel) Then TxtNivel = "0"
End Sub

Public Sub Impresiones()
    DGQuery.Visible = False
   'MsgBox Opcion
    Select Case Opcion
      Case 1
          MensajeEncabData = "ESTADO DE CUENTA DE CLIENTES"
          SQLMsg1 = "Corte al " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          Imprimir_Saldo_Factura AdoQuery, 9
      Case 2
          MensajeEncabData = "ESTADO DE CUENTA DE CLIENTES"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          Imprimir_Saldo_Clientes AdoQuery, 8
      Case 3
          MensajeEncabData = "ESTADO DE PRODUCTOS POR CLIENTES"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          Imprimir_Resumen_Productos AdoQuery, 8
      Case 4
          MensajeEncabData = "RESUMEN DE VENTAS POR CLIENTES"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          ImprimirAdo AdoQuery, True, 2, 9
      Case 5
          MensajeEncabData = "RESUMEN DE VENTAS POR PRODUCTOS"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          ImprimirVentasCosto AdoQuery, True, 2, 9
      Case 6
          MensajeEncabData = "ESTADO DE ABONOS DE CLIENTES"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          Imprimir_Abonos_De_Caja AdoQuery, MBFechaI.Text, MBFechaF.Text
      Case 7
          MensajeEncabData = "ESTADO DE CHEQUES PROTESTADOS"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          Imprimir_Abonos_De_Caja AdoQuery, MBFechaI.Text, MBFechaF.Text
      Case 8
          MensajeEncabData = "VENTAS POR PRODUCTOS"
          SQLMsg1 = "CORTE DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
          Mifecha = MBFechaF.Text
          ImprimirAdo AdoQuery, True, 1, 8
      Case 9, 10
          DGQuery.Visible = False
          If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
          If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
          If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
          If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
          Mifecha = MBFechaF.Text
          If TipoDoc = "C" Then ImprimirResumenCartera AdoQuery, Codigo4
          If TipoDoc = "F" Then ImprimirCtasCob AdoQuery, sSQL, True
          DGQuery.Visible = True
      Case 11
          DGQuery.Visible = False
          If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
          If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
          If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
          If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
          Mifecha = MBFechaF.Text
          Imprimir_Pendientes_Facturacion AdoQuery, Opcion, True
          DGQuery.Visible = True
    End Select
    HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
    DGQuery.Visible = True
End Sub

Public Sub Listado_Facturas_Por_Meses(Por_FA As Boolean)
Dim AnioIni As Integer
Dim AnioFin As Integer
Dim AnioAct As Integer
Dim MesIni As Byte
Dim MesFin As Byte
Dim AnioI As String
Dim AnioF As String
Dim Anios() As Currency
Dim VerAnios() As Boolean
Dim VerMeses(1 To 12) As Boolean
Dim CxC_VerMeses(1 To 12) As Boolean
Dim SumMeses(1 To 12) As Currency

FechaValida MBFechaI
FechaValida MBFechaF

RatonReloj
MesIni = Month(MBFechaI)
MesFin = Month(MBFechaF)

FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)

AnioI = Format(Year(MBFechaI), "0000")
AnioF = Format(Year(MBFechaF), "0000")

'Periodos antes del actual
AnioFin = Year(MBFechaF) - 1
AnioIni = AnioFin - 7
ReDim Anios(AnioIni To AnioFin) As Currency
ReDim VerAnios(AnioIni To AnioFin) As Boolean
  
  Saldo = 0
  SaldoAnterior = 0
  For J = AnioIni To AnioFin
      Anios(J) = 0
      VerAnios(J) = False
  Next J
  
  For J = 1 To 12
      SumMeses(J) = 0
      VerMeses(J) = False
      CxC_VerMeses(J) = False
      sSQL = "UPDATE Detalle_Factura " _
           & "SET Mes_No = " & CInt(J) & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Mes = '" & MesesLetras(CInt(J)) & "' " _
           & "AND Mes_No = 0 "
      ConectarAdoExecute sSQL
  Next J
  
  Contador = 1
  Si_No = False
  NumFacturas = 12
  DGQuery.Visible = False
  DGQuery.Caption = "DEUDAS PENDIENTES DEL PERIODO " & AnioI
 'Lista de Clientes con deuda pendiente extracontable
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' "
  ConectarAdoExecute sSQL
  
 'Lista Clientes de la Consulta
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' "
  SelectAdodc AdoTipo, sSQL
 
 'Listado de Facturas Emitidas
  sSQL = "SELECT F.CodigoC,C.Cliente,C.Grupo,F.Fecha,"
  If Por_FA Then
     sSQL = sSQL _
          & "Total_MN,Saldo_MN " _
          & "FROM Facturas As F,Clientes As C "
  Else
     If OpcPend.value Then
        sSQL = sSQL & "Mes_No,(Total-Total_Desc+Total_IVA) As Saldo_MN "
     Else
        sSQL = sSQL & "Mes_No,(Total-Total_Desc+Total_IVA) As Total_MN "
     End If
     sSQL = sSQL & "FROM Detalle_Factura As F,Clientes As C "
  End If
  sSQL = sSQL _
      & "WHERE F.Item = '" & NumEmpresa & "' " _
      & "AND F.Fecha <= #" & FechaFin & "# " _
      & Tipo_De_Consulta() _
      & "AND F.CodigoC = C.Codigo "
  If Por_FA Then
     sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente,F.Fecha "
  Else
     sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente,F.Fecha,F.Mes_No "
  End If
  SelectAdodc AdoQuery, sSQL
 'MsgBox sSQL
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Codigo1 = .Fields("Grupo")
       CodigoCli = .Fields("CodigoC")
       K = 0
      'MsgBox .RecordCount
       Do While Not .EOF
          K = K + 1
          HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA - Insertando: " & Format(K / .RecordCount, "00%") & "..."
          If CodigoCli <> .Fields("CodigoC") Then
             If AdoTipo.Recordset.RecordCount > 0 Then AdoTipo.Recordset.MoveFirst
             AdoTipo.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
             If AdoTipo.Recordset.EOF Then
                AdoTipo.Recordset.AddNew
                AdoTipo.Recordset.Fields("CodigoC") = CodigoCli
                AdoTipo.Recordset.Fields("TP") = "CXCP"
                AdoTipo.Recordset.Fields("Ln") = Contador
                AdoTipo.Recordset.Fields("Item") = NumEmpresa
                AdoTipo.Recordset.Fields("CodigoU") = CodigoUsuario
                Contador = Contador + 1
             End If
             Total = 0
             AdoTipo.Recordset.Fields("PEN") = Saldo
             Total = Total + Saldo
             For I = AnioIni + 1 To AnioFin
                 AdoTipo.Recordset.Fields("P_" & CStr(I)) = Anios(I)
                 Total = Total + Anios(I)
             Next I
            'Deuda del año Actual
             For I = 1 To 12
                 AdoTipo.Recordset.Fields(MesesLetras(CInt(I))) = SumMeses(I)
                 Total = Total + SumMeses(I)
             Next I
            'Total Deuda
             AdoTipo.Recordset.Fields("Total") = Total
             AdoTipo.Recordset.Update
             
             Saldo = 0
             For J = AnioIni To AnioFin
                 Anios(J) = 0
             Next J
             For J = 1 To 12
                 SumMeses(J) = 0
             Next J
             Codigo1 = .Fields("Grupo")
             CodigoCli = .Fields("CodigoC")
          End If
         'Sumamos el total de deuda pendiente
          AnioAct = Year(.Fields("Fecha"))
          If Por_FA Then MesIni = Month(.Fields("Fecha")) Else MesIni = .Fields("Mes_No")
          If MesIni = 0 Then MesIni = Month(.Fields("Fecha"))
          If AnioIni <= AnioAct And AnioAct <= AnioFin Then
             If OpcPend.value Then
                Anios(AnioAct) = Anios(AnioAct) + .Fields("Saldo_MN")
             Else
                Anios(AnioAct) = Anios(AnioAct) + .Fields("Total_MN")
             End If
             VerAnios(AnioAct) = True
          ElseIf AnioAct < AnioIni Then
             If OpcPend.value Then
                Saldo = Saldo + .Fields("Saldo_MN")
             Else
                Saldo = Saldo + .Fields("Total_MN")
             End If
          Else
             If OpcPend.value Then
                SumMeses(MesIni) = SumMeses(MesIni) + .Fields("Saldo_MN")
             Else
                SumMeses(MesIni) = SumMeses(MesIni) + .Fields("Total_MN")
             End If
             VerMeses(MesIni) = True
          End If
          If OpcPend.value Then
             SaldoAnterior = SaldoAnterior + .Fields("Saldo_MN")
          Else
             SaldoAnterior = SaldoAnterior + .Fields("Total_MN")
          End If
         .MoveNext
       Loop
   End If
  End With
  
    If AdoTipo.Recordset.RecordCount > 0 Then AdoTipo.Recordset.MoveFirst
    AdoTipo.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
    If AdoTipo.Recordset.EOF Then
       AdoTipo.Recordset.AddNew
       AdoTipo.Recordset.Fields("CodigoC") = CodigoCli
       AdoTipo.Recordset.Fields("TP") = "CXCP"
       AdoTipo.Recordset.Fields("Ln") = Contador
       AdoTipo.Recordset.Fields("Item") = NumEmpresa
       AdoTipo.Recordset.Fields("CodigoU") = CodigoUsuario
    End If
    Total = 0
    AdoTipo.Recordset.Fields("PEN") = Saldo
    Total = Total + Saldo
    For I = AnioIni + 1 To AnioFin
        AdoTipo.Recordset.Fields("P_" & CStr(I)) = Anios(I)
        Total = Total + Anios(I)
    Next I
    'Deuda del año Actual
    For I = 1 To 12
        AdoTipo.Recordset.Fields(MesesLetras(CInt(I))) = SumMeses(I)
        Total = Total + SumMeses(I)
    Next I
    'Total Deuda
    AdoTipo.Recordset.Fields("Total") = Total
    AdoTipo.Recordset.Update
  
  
 'Lista Clientes de la Consulta
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' "
  SelectAdodc AdoTipo, sSQL
 
 'Listado CxC PreFacturable
  If CheqPreFa.value <> 0 Then
  sSQL = "SELECT C.Cliente,F.Codigo,C.Grupo,F.Num_Mes,SUM(F.Valor-F.Descuento) As Total_MN " _
       & "FROM Clientes_Facturacion As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & Buscar_x_Patron(True) _
       & "AND F.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente,F.Codigo,C.Grupo,F.Num_Mes " _
       & "ORDER BY C.Cliente,F.Codigo,C.Grupo,F.Num_Mes "
  SelectAdodc AdoQuery, sSQL
  For J = 1 To 12
      SumMeses(J) = 0
  Next J
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Codigo1 = .Fields("Grupo")
       CodigoCli = .Fields("Codigo")
       K = 0
       Do While Not .EOF
          K = K + 1
          HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA - Insertando: " & Format(K / .RecordCount, "00%") & "..."
          If CodigoCli <> .Fields("Codigo") Then
             If AdoTipo.Recordset.RecordCount > 0 Then AdoTipo.Recordset.MoveFirst
             AdoTipo.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
             If AdoTipo.Recordset.EOF Then
                AdoTipo.Recordset.AddNew
                AdoTipo.Recordset.Fields("CodigoC") = CodigoCli
                AdoTipo.Recordset.Fields("TP") = "CXCP"
                AdoTipo.Recordset.Fields("Ln") = Contador
                AdoTipo.Recordset.Fields("Item") = NumEmpresa
                AdoTipo.Recordset.Fields("CodigoU") = CodigoUsuario
                Contador = Contador + 1
             End If
            'Deuda del año Actual
             Total = 0
             For I = 1 To 12
                 AdoTipo.Recordset.Fields("CxC_" & Mid(MesesLetras(CInt(I)), 1, 3)) = SumMeses(I)
                 Total = Total + SumMeses(I)
             Next I
            'Total Deuda
             AdoTipo.Recordset.Fields("Total") = Total
             AdoTipo.Recordset.Update
             Saldo = 0
             For J = 1 To 12
                 SumMeses(J) = 0
             Next J
             Codigo1 = .Fields("Grupo")
             CodigoCli = .Fields("Codigo")
          End If
         'Sumamos el total de deuda pendiente
          MesIni = .Fields("Num_Mes")
          SumMeses(MesIni) = SumMeses(MesIni) + .Fields("Total_MN")
          CxC_VerMeses(MesIni) = True
         .MoveNext
       Loop
   End If
  End With
  
    If AdoTipo.Recordset.RecordCount > 0 Then AdoTipo.Recordset.MoveFirst
    AdoTipo.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
    If AdoTipo.Recordset.EOF Then
       AdoTipo.Recordset.AddNew
       AdoTipo.Recordset.Fields("CodigoC") = CodigoCli
       AdoTipo.Recordset.Fields("TP") = "CXCP"
       AdoTipo.Recordset.Fields("Ln") = Contador
       AdoTipo.Recordset.Fields("Item") = NumEmpresa
       AdoTipo.Recordset.Fields("CodigoU") = CodigoUsuario
    End If
  
    Total = 0
   'Deuda del año Actual
    For I = 1 To 12
        AdoTipo.Recordset.Fields("CxC_" & Mid(MesesLetras(CInt(I)), 1, 3)) = SumMeses(I)
        Total = Total + SumMeses(I)
    Next I
    'Total Deuda
    AdoTipo.Recordset.Fields("Total") = Total
    AdoTipo.Recordset.Update
  End If
  
  sSQL = "UPDATE Saldo_Diarios " _
       & "SET PEN = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND PEN IS NULL "
  ConectarAdoExecute sSQL

  For I = AnioIni To AnioFin
      If VerAnios(I) Then
         sSQL = "UPDATE Saldo_Diarios " _
              & "SET P_" & CStr(I) & " = 0 " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' " _
              & "AND P_" & CStr(I) & " IS NULL "
         ConectarAdoExecute sSQL
      End If
  Next I

  For IE = 1 To 12
      If VerMeses(IE) Then
         sSQL = "UPDATE Saldo_Diarios " _
              & "SET " & MesesLetras(IE) & " = 0 " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' " _
              & "AND " & MesesLetras(IE) & " IS NULL "
         ConectarAdoExecute sSQL
      End If
  Next IE
  
  For IE = 1 To 12
      If CxC_VerMeses(IE) Then
         sSQL = "UPDATE Saldo_Diarios " _
              & "SET CxC_" & Mid(MesesLetras(IE), 1, 3) & " = 0 " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' " _
              & "AND CxC_" & Mid(MesesLetras(IE), 1, 3) & " IS NULL "
         ConectarAdoExecute sSQL
      End If
  Next IE
  
  sSQL = "UPDATE Saldo_Diarios " _
       & "SET Total = "
  For I = AnioIni To AnioFin
      If VerAnios(I) Then sSQL = sSQL & "P_" & CStr(I) & " + "
  Next I
  For IE = 1 To 12
      If VerMeses(IE) Then sSQL = sSQL & MesesLetras(IE) & " + "
  Next IE
  For IE = 1 To 12
      If CxC_VerMeses(IE) Then sSQL = sSQL & "CxC_" & Mid(MesesLetras(IE), 1, 3) & " + "
  Next IE
  sSQL = sSQL & "PEN " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' "
  ConectarAdoExecute sSQL
  
 'Listado de Rubros de pensiones por meses
  sSQL = "SELECT C.Cliente,"
  If VerAnios(AnioIni) Then sSQL = sSQL & "PEN,"
  For I = AnioIni To AnioFin
      If VerAnios(I) Then sSQL = sSQL & "P_" & CStr(I) & ","
  Next I
  For IE = 1 To 12
      If VerMeses(IE) Then sSQL = sSQL & MesesLetras(IE) & ","
  Next IE
  For IE = 1 To 12
      If CxC_VerMeses(IE) Then sSQL = sSQL & "CxC_" & Mid(MesesLetras(IE), 1, 3) & ","
  Next IE
  sSQL = sSQL & "SD.Total,C.Direccion,C.Grupo,C.Plan_Afiliado As BUS_No,SD.Ln As No_ " _
       & "FROM Saldo_Diarios As SD,Clientes As C " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.TP = 'CXCP' " _
       & "AND SD.CodigoC = C.Codigo " _
       & "ORDER BY C.Grupo,C.Cliente,SD.TC "
  SelectDataGrid DGQuery, AdoQuery, sSQL, , True
  'MsgBox sSQL
  LabelSaldo.Caption = Format(SaldoAnterior, "#,##0.00")
  DGQuery.Visible = True
  HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
  Opcion = 11
  RatonNormal
End Sub

Public Sub Estado_Cuenta_Cliente()
Dim CaptionTemp As String
  CaptionTemp = HistorialFacturas.Caption
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE TP = 'CXCF' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Detalle_Factura " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  ConectarAdoExecute sSQL
    
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Clientes_Facturacion " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET X = 'P' " _
          & "FROM Facturas As F,Clientes As C "
  Else
     sSQL = "UPDATE Facturas As F,Clientes As C " _
          & "SET F.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & Buscar_x_Patron _
       & "AND F.CodigoC = C.Codigo "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes_Facturacion " _
          & "SET X = 'P' " _
          & "FROM Clientes_Facturacion As F,Clientes As C "
  Else
     sSQL = "UPDATE Clientes_Facturacion As F,Clientes As C " _
          & "SET F.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & Buscar_x_Patron(True) _
       & "AND F.Codigo = C.Codigo "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET X = F.X " _
          & "FROM Detalle_Factura As XF,Facturas As F "
  Else
     sSQL = "UPDATE Detalle_Factura As XF,Facturas As F " _
          & "SET XF.X = F.X "
  End If
  sSQL = sSQL _
       & "WHERE XF.Item = '" & NumEmpresa & "' " _
       & "AND XF.Periodo = '" & Periodo_Contable & "' " _
       & "AND XF.TC = F.TC " _
       & "AND XF.Serie = F.Serie " _
       & "AND XF.Autorizacion = F.Autorizacion " _
       & "AND XF.Factura = F.Factura " _
       & "AND XF.CodigoC = F.CodigoC " _
       & "AND XF.Item = F.Item " _
       & "AND XF.Periodo = F.Periodo "
  ConectarAdoExecute sSQL
    
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = F.X " _
          & "FROM Trans_Abonos As XF,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As XF,Facturas As F " _
          & "SET XF.X = F.X "
  End If
  sSQL = sSQL _
       & "WHERE XF.Item = '" & NumEmpresa & "' " _
       & "AND XF.Periodo = '" & Periodo_Contable & "' " _
       & "AND XF.TP = F.TC " _
       & "AND XF.Serie = F.Serie " _
       & "AND XF.Autorizacion = F.Autorizacion " _
       & "AND XF.Factura = F.Factura " _
       & "AND XF.CodigoC = F.CodigoC " _
       & "AND XF.Item = F.Item " _
       & "AND XF.Periodo = F.Periodo "
  ConectarAdoExecute sSQL
    
  SQL1 = "SELECT Fecha,Factura,Mid(Producto & ': ' & Mes,1,65) As Concepto,(Total-Total_Desc+Total_IVA) As CxC,0 As Abonos,CodigoC," _
       & "'" & NumEmpresa & "' As Items,'" & CodigoUsuario & "' As CodigoUs, 'CXCF' As TP " _
       & "FROM Detalle_Factura " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND X = 'P' "
  sSQL = "INSERT INTO Saldo_Diarios (Fecha,Numero,Comprobante,Ingresos,Egresos,CodigoC,Item,CodigoU,TP) " _
       & SQL1
  ConectarAdoExecute sSQL
  
  SQL1 = "SELECT Fecha,Factura,Mid(Banco & '-' & Cheque,1,65) As Concepto,0 As CxC,Abono,CodigoC," _
       & "'" & NumEmpresa & "' As Items,'" & CodigoUsuario & "' As CodigoUs, 'CXCF' As TP " _
       & "FROM Trans_Abonos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND X = 'P' "
  sSQL = "INSERT INTO Saldo_Diarios (Fecha,Numero,Comprobante,Ingresos,Egresos,CodigoC,Item,CodigoU,TP) " _
       & SQL1
  ConectarAdoExecute sSQL
  
  SQL1 = "SELECT Fecha,999999 As Factura,('Por Facturar de ' & Mes & '-' & Periodo) As Concepto,(Valor-Descuento) As CxC,0 As Abonos,Codigo," _
       & "'" & NumEmpresa & "' As Items,'" & CodigoUsuario & "' As CodigoUs, 'CXCF' As TP " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND X = 'P' "
  sSQL = "INSERT INTO Saldo_Diarios (Fecha,Numero,Comprobante,Ingresos,Egresos,CodigoC,Item,CodigoU,TP) " _
       & SQL1
  ConectarAdoExecute sSQL
  
  Contador = 0
  Saldo = 0
  SQL1 = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Fecha >= #" & FechaIni & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'CXCF' " _
       & "ORDER BY CodigoC,Numero,Fecha,Ingresos DESC,Egresos "
  SelectAdodc AdoFacturas, SQL1
  RatonReloj
  With AdoFacturas.Recordset
   If .RecordCount > 0 Then
       Codigo = .Fields("CodigoC")
       Factura_No = .Fields("Numero")
       Do While Not .EOF
          If Codigo <> .Fields("CodigoC") Or Factura_No <> .Fields("Numero") Then
             Codigo = .Fields("CodigoC")
             Factura_No = .Fields("Numero")
             Saldo = 0
          End If
          Saldo = Saldo + .Fields("Ingresos") - .Fields("Egresos")
         .Fields("Saldo_Actual") = Saldo
         .Fields("Ln") = Contador
         .Update
          HistorialFacturas.Caption = Format(Contador / .RecordCount, "00%") & " Mayorizando Saldos de Facturas..."
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
  HistorialFacturas.Caption = CaptionTemp
  SQL1 = "SELECT C.Cliente,F.Fecha,F.Numero,Comprobante,Ingresos As CxC,Egresos As Abonos,F.Saldo_Actual,C.Grupo,F.CodigoC " _
       & "FROM Saldo_Diarios As F,Clientes As C " _
       & "WHERE F.Fecha >= #" & FechaIni & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.TP = 'CXCF' " _
       & Buscar_x_Patron _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY C.Grupo,C.Cliente,F.Numero,F.Fecha,CxC DESC "
  SelectDataGrid DGQuery, AdoQuery, SQL1
  RatonNormal
End Sub

Public Function Buscar_x_Patron(Optional SinFactura As Boolean) As String
Dim SQL3X, Patron_Busqueda As String
   SQL3X = ""
   Patron_Busqueda = DCCliente.Text
   If Patron_Busqueda = "" Then Patron_Busqueda = Ninguno
   Select Case ListCliente.Text
     Case "Codigo"
          SQL3X = SQL3X & "AND C.Codigo = '" & Patron_Busqueda & "' "
     Case "CI_RUC"
          SQL3X = SQL3X & "AND C.CI_RUC = '" & Patron_Busqueda & "' "
     Case "Cliente"
          SQL3X = SQL3X & "AND Ucase(Mid(C.Cliente,1," & Len(Patron_Busqueda) & ")) = '" & Patron_Busqueda & "' "
     Case "Ciudad"
          SQL3X = SQL3X & "AND C.Ciudad = '" & Patron_Busqueda & "' "
     Case "Grupo"
          SQL3X = SQL3X & "AND C.Grupo = '" & Patron_Busqueda & "' "
     Case "Factura"
          If Not SinFactura Then SQL3X = SQL3X & "AND F.Factura = " & Val(Patron_Busqueda) & " "
     Case "Forma_Pago"
          If Not PorFactura Then SQL3X = SQL3X & "AND F.Forma_Pago = '" & Patron_Busqueda & "' "
     Case "Plan_Afiliado"
          SQL3X = SQL3X & "AND C.Plan_Afiliado = '" & Patron_Busqueda & "' "
   End Select
   Buscar_x_Patron = SQL3X
End Function

