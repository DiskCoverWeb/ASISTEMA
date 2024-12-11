VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form HistorialFacturas 
   Caption         =   "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheqIngreso 
      Caption         =   "Cuenta de Ingreso"
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
      Left            =   10710
      TabIndex        =   32
      Top             =   735
      Width           =   2085
   End
   Begin VB.Frame FrmPatronBusqueda 
      BackColor       =   &H00C00000&
      Caption         =   "| PATRON DE BUSQUEDA |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4320
      Left            =   1260
      TabIndex        =   27
      Top             =   1785
      Visible         =   0   'False
      Width           =   15765
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
         Height          =   3375
         Left            =   105
         TabIndex        =   28
         Top             =   735
         Width           =   2010
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "Histfact.frx":0000
         DataSource      =   "AdoCliente"
         Height          =   3780
         Left            =   2205
         TabIndex        =   30
         Top             =   315
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   6668
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
         Left            =   105
         TabIndex        =   29
         Top             =   315
         Width           =   2010
      End
   End
   Begin VB.TextBox TxtDocHasta 
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
      Height          =   330
      Left            =   4725
      TabIndex        =   26
      Text            =   "0"
      Top             =   1050
      Width           =   1695
   End
   Begin VB.TextBox TxtDocDesde 
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
      Height          =   330
      Left            =   2835
      TabIndex        =   25
      Text            =   "0"
      Top             =   1050
      Width           =   1800
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   2115
      Left            =   19110
      OleObjectBlob   =   "Histfact.frx":0019
      TabIndex        =   22
      Top             =   735
      Visible         =   0   'False
      Width           =   2745
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Histfact.frx":2698
      Height          =   4425
      Left            =   105
      TabIndex        =   20
      Top             =   1470
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   7805
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   11025
      TabIndex        =   1
      Top             =   0
      Width           =   6105
      Begin VB.CheckBox CheqPreFa 
         Caption         =   "Incluir Prefacturación"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4620
         TabIndex        =   6
         Top             =   105
         Width           =   1380
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "&Canceladas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   210
         Width           =   1095
      End
      Begin VB.OptionButton OpcTodas 
         Caption         =   "&Todas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         TabIndex        =   5
         Top             =   210
         Width           =   765
      End
      Begin VB.OptionButton OpcPend 
         Caption         =   "&Pendiente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OpcAnul 
         Caption         =   "&Anuladas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   210
         Width           =   975
      End
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
      Left            =   6510
      TabIndex        =   11
      Top             =   735
      Width           =   1980
   End
   Begin VB.CheckBox CheqAbonos 
      Caption         =   "Cuenta de Abono"
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
      Left            =   8610
      TabIndex        =   12
      Top             =   735
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   105
      TabIndex        =   21
      Top             =   7770
      Width           =   540
   End
   Begin MSComctlLib.ImageList ImageListFacturas 
      Left            =   12390
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":26AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":2F89
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3863
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":3E97
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":42E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":4BC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":4EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":57DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":5AF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":5E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":6129
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":6A03
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":72DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":9A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":9BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":9F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":A355
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":A4AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":A609
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":A923
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":AD75
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":13247
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":17EC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1B24C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1B566
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1B880
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1CD4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1D065
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Histfact.frx":1D4B7
            Key             =   ""
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
      Top             =   2835
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
      Top             =   2205
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
      Top             =   2520
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
      Top             =   3150
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
      Top             =   3465
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1470
      TabIndex        =   10
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   8
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
      Width           =   15960
      _ExtentX        =   28152
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageListFacturas"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir los resultados"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Facturas"
            Object.ToolTipText     =   "Presenta el Resumen de Facturas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resumen"
            Object.ToolTipText     =   "Resumen de Ventas"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen_Prod"
                  Text            =   "Resumen de Productos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen_Prod_Meses"
                  Text            =   "Resumen de Productos por Meses"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ResumenVentCost"
                  Text            =   "Resumen de Ventas/Costos"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen_Ventas_Vendedor"
                  Text            =   "Resumen Comisiones por Vendedor"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ventas_x_Cli"
                  Text            =   "Ventas por Cliente"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ventas_Cli_x_Mes"
                  Text            =   "Ventas Clientes por Meses"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VentasxProductos"
                  Text            =   "Ventas Clientes por Productos"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ventas_ResumindasxVendedor"
                  Text            =   "Ventas Resumidas por Vendedor"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Detalle_Abonos"
            Object.ToolTipText     =   "Detalle de Abonos de Facturas/Notas de Ventas"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SMAbonos_Anticipados"
                  Text            =   "Anticipados de Abonos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Contra_Cta"
                  Text            =   "Contrapartida del Abonos"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Abonos_Ant"
                  Text            =   "Errores en Abonos Anticipados"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Abonos_Erroneos"
                  Text            =   "Presenta Abonos mal Procesados"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Protestado"
            Object.ToolTipText     =   "Cheques protestados"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Facturas_Clientes"
            Object.ToolTipText     =   "Listado de Facturas"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Por_Clientes"
                  Text            =   "Ordenadas por Clientes"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Por_Facturas"
                  Text            =   "Ordenados por Facturas"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Por_Vendedor"
                  Text            =   "CxC Clientes por Vendedor"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen_Vent_x_Ejec"
                  Text            =   "Resumen de Ventas por Vendedor"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Resumen_Cartera"
                  Text            =   "Resumen de Cartera Detallado"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CxC_Tiempo_Credito"
                  Text            =   "Cuentas por Cobrar por Tiempo de Credito"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Tipo_Pago_Cliente"
                  Text            =   "Tipo de Pagos Clientes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Por_Buses"
            Object.ToolTipText     =   "Listado de Buses"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Listado_Tarjetas"
            Object.ToolTipText     =   "Listado de Clientes con Tarjetas de Credito"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tipo_de_Abonos"
            Object.ToolTipText     =   "Presentar Retenciones y Notas de Credito"
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Tipo_Abonos"
                  Text            =   "Abonos Procesados"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Retenciones_NC"
                  Text            =   "Notas de Creditos y Retenciones"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CxC_Clientes"
            Object.ToolTipText     =   "Listado de Cartera por Meses"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Listar_Por_Meses"
            Object.ToolTipText     =   "Listar Clientes por Rubro de Meses "
            ImageIndex      =   19
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Estado_Cuenta_Cliente"
            Object.ToolTipText     =   "Estado de cuenta de Clientes"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Listados_Medidor"
            Object.ToolTipText     =   "Lista de Medidores por Socios"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ventas_x_Excel"
            Object.ToolTipText     =   "Genera las Ventas por Excel"
            ImageIndex      =   22
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Bajar_Excel"
                  Text            =   "Bajar a Excel"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Reporte_Ventas"
                  Text            =   "Reporte de Ventas"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Reporte_Catastro"
                  Text            =   "Reporte de Catastro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Enviar_FA_Emails"
            Object.ToolTipText     =   "Enviar por mail Facturas Electronicas"
            ImageIndex      =   29
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Enviar_FA_Email"
                  Text            =   "Enviar por mail Facturas Electronicas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Enviar_RE_Email"
                  Text            =   "Enviar por mail Recibos de Pago"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Recibos_Anticipados"
                  Text            =   "Enviar por Mail Recibos Anticipados"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Deuda_x_Mail"
                  Text            =   "Enviar Resumen de Cartera por mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Base_Access"
            Object.ToolTipText     =   "Presenta un listado en Access"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Base_MySQL"
            Object.ToolTipText     =   "Presenta Listado de MySQL"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar_Malla"
            Object.ToolTipText     =   "Patron de busqueda"
            ImageIndex      =   30
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCxC 
      Bindings        =   "Histfact.frx":1DD91
      DataSource      =   "AdoCxC"
      Height          =   315
      Left            =   6510
      TabIndex        =   13
      Top             =   1050
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
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
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   315
      Top             =   1890
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
   Begin VB.Label LblPatronBusqueda 
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   12915
      TabIndex        =   31
      Top             =   735
      Width           =   6210
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Documento Hasta"
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
      TabIndex        =   24
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Documento  Desde"
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
      Left            =   2835
      TabIndex        =   23
      Top             =   735
      Width           =   1800
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
      TabIndex        =   7
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
      Left            =   1470
      TabIndex        =   9
      Top             =   735
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
      Left            =   10920
      TabIndex        =   18
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
      Left            =   10080
      TabIndex        =   19
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
      Left            =   8400
      TabIndex        =   14
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
      Left            =   7455
      TabIndex        =   15
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
      Left            =   5775
      TabIndex        =   16
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
      TabIndex        =   17
      Top             =   7770
      Width           =   1380
   End
End
Attribute VB_Name = "HistorialFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PorCxC As Boolean
Dim PresionoEsc As Boolean
Dim VetPEN As Boolean

Dim Con_Costeo As String
Dim DescItem As String

Dim CnExterna As Integer

Private Sub CheqAbonos_Click()
  If CheqAbonos.value = 1 Then
     sSQL = "SELECT (TA.Cta & ' - ' & CC.Cuenta) As NomCxC " _
          & "FROM Trans_Abonos As TA, Catalogo_Cuentas As CC " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND CC.Periodo = TA.Periodo " _
          & "AND CC.Item = TA.Item " _
          & "AND CC.Codigo = TA.Cta " _
          & "GROUP BY Cta,Cuenta " _
          & "ORDER BY Cta "
     SelectDB_Combo DCCxC, AdoCxC, sSQL, "NomCxC"
     CheqCxC.value = 0
     CheqIngreso.value = 0
     DCCxC.Visible = True
  Else
     DCCxC.Visible = False
  End If
End Sub

Private Sub CheqCxC_Click()
  If CheqCxC.value = 1 Then
     sSQL = "SELECT (F.Cta_CxP & ' - ' & CC.Cuenta) As NomCxC " _
          & "FROM Facturas As F,Catalogo_Cuentas As CC " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Item = F.Item " _
          & "AND CC.Codigo = F.Cta_CxP " _
          & "GROUP BY F.Cta_CxP,Cuenta " _
          & "ORDER BY F.Cta_CxP "
     SelectDB_Combo DCCxC, AdoCxC, sSQL, "NomCxC"
     CheqAbonos.value = 0
     CheqIngreso.value = 0
     DCCxC.Visible = True
  Else
     DCCxC.Visible = False
  End If
End Sub

Public Sub Historico_Facturas()
'Facturas
 Opcion = 1
 FechaIni = BuscarFecha(MBFechaI)
 FechaFin = BuscarFecha(MBFechaF)
 
 DGQuery.Caption = "HISTORIAL DE FACTURAS"
 Label2.Caption = "(" & Opcion & ") Facturado"
 Label4.Caption = " Cobrado"
 Label3.Caption = " Saldo"
 PorCxC = False
 If CheqCxC.value = 1 Then PorCxC = True
 RatonReloj
 Total = 0
 Abono = 0
 Saldo = 0
 sSQL = "SELECT C.Cliente, F.T, F.Serie, F.Factura, F.Fecha, Fecha_V, F.Total_MN As Total, F.Total_Efectivo, F.Total_Banco, " _
      & "F.Total_Ret_Fuente, F.Total_Ret_IVA_B, F.Total_Ret_IVA_S, F.Otros_Abonos, F.Total_Abonos,F.Saldo_Actual, " _
      & "F.Fecha_C As Abonado_El, F.CodigoC, C.CI_RUC, F.TC, F.Autorizacion, C.Grupo, A.Nombre_Completo As Ejecutivo, C.Ciudad, " _
      & "C.Plan_Afiliado As Sectorizacion, F.Cta_CxP, C.EMail, C.EMail2, C.EMailR, C.Representante " _
      & "FROM Facturas As F, Clientes As C, Accesos As A " _
      & "WHERE F.Item = '" & NumEmpresa & "' " _
      & "AND F.Periodo = '" & Periodo_Contable & "' " _
      & "AND F.Fecha <= #" & FechaFin & "# " _
      & Tipo_De_Consulta() _
      & "AND F.CodigoC = C.Codigo " _
      & "AND F.Cod_Ejec = A.Codigo " _
      & "ORDER BY C.Cliente,F.Serie,F.Factura,F.Fecha "
' MsgBox sSQL
 Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
 RatonNormal
 Totales_CxC_Abonos
End Sub

Public Sub Ventas_Productos()
'Resumen Ventas por Producto
Opcion = 8
RatonReloj
'MsgBox "Ventas_Productos"
DGQuery.Visible = False
  Si_No = False
  Con_Costeo = " "
  Mensajes = "Reporte Con Costeo "
  Titulo = "Formulario de Confirmación"
  If BoxMensaje = vbYes Then
     If ClaveAdministrador Then Si_No = True
  End If
Label2.Caption = "(" & Opcion & ") Ventas"
Label4.Caption = "  "
Label3.Caption = "  "
DGQuery.Visible = False

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT F.T, CL.Cliente, F.TC As Doc, F.Serie, F.Factura, F.Fecha, F.Codigo, F.Producto, F.Mes, F.Cantidad, F.Total, 0 As Total_NC, " _
     & "(Total_Desc+Total_Desc2) As Descuento, (F.Total-Total_Desc-Total_Desc2) As SubTotal, C.Marca, " _
     & "C.Desc_Item As Parte, F.Lote_No, F.Fecha_Fab, F.Fecha_Exp, C.Reg_Sanitario, F.Serie_No " & Con_Costeo
If Si_No Then sSQL = sSQL & ",F.Precio, Valor_Compra As Costos "
sSQL = sSQL _
     & "FROM Detalle_Factura As F, Catalogo_Productos As C, Clientes As CL " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & Tipo_De_Consulta(, , True) _
     & "AND C.INV <> " & Val(adFalse) & " " _
     & "AND F.T <> '" & Anulado & "' "
If CodigoInv <> Ninguno Then sSQL = sSQL & "AND C.Codigo_Inv = '" & CodigoInv & "' "
sSQL = sSQL _
     & "AND F.Item = C.Item " _
     & "AND F.Periodo = C.Periodo " _
     & "AND F.Codigo = C.Codigo_Inv " _
     & "AND F.CodigoC = CL.Codigo "
sSQL = sSQL _
     & "UNION ALL " _
     & "SELECT F.T, CL.Cliente, F.TP As Doc, F.Serie, F.Factura, F.Fecha, F.Cta As Codigo, (F.Banco + ' - ' + F.Cheque) AS Producto_Aux, F.Mes, 1 As Cantidad, 0 As Total, -F.Abono As Total_NC, " _
     & "0 As Descuento, -F.Abono As SubTotal, '.' As Marca, " _
     & "'.' As Parte, '.' As Lote_No, F.Fecha As Fecha_Fab, F.Fecha As Fecha_Exp, '.' As Reg_Sanitario, '.' As Serie_No "
If Si_No Then sSQL = sSQL & ", 0 As Precio, 0 As Costos "
sSQL = sSQL _
     & "FROM Trans_Abonos As F, Clientes As CL " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND F.Banco = 'NOTA DE CREDITO' " _
     & Tipo_De_Consulta() _
     & "AND F.T <> '" & Anulado & "' "
sSQL = sSQL _
     & "AND F.CodigoC = CL.Codigo " _
     & "ORDER BY Doc, F.Factura, F.Fecha "
'MsgBox sSQL
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
'MsgBox sSQL
RatonReloj
DGQuery.Visible = False
Total = 0: Abono = 0: Saldo = 0
'Total = T_Fields("Total")
'If Si_No Then
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            If Si_No Then Saldo = Saldo + .fields("Costos")
            Total = Total + .fields("Total")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
'End If
'3301738
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = "PVP"
Label3.Caption = "Costo"

LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Abono, "#,##0.00")
LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Abonos_Facturas(Optional Ret_NC As Boolean)
Dim IDMes As Byte
RatonReloj
Opcion = 6
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
If CheqCxC.value = 1 Then PorCxC = True
   DGQuery.Caption = "ABONOS DE FACTURAS"
   For IDMes = 1 To 12
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
            & "AND MONTH(TA.Fecha) = " & IDMes & " " _
            & "AND TA.Item = DF.Item " _
            & "AND TA.Periodo = DF.Periodo " _
            & "AND TA.Factura = DF.Factura " _
            & "AND TA.Serie = DF.Serie " _
            & "AND TA.Autorizacion = DF.Autorizacion " _
            & "AND TA.CodigoC = DF.CodigoC "
       Ejecutar_SQL_SP sSQL
   Next IDMes
   Total = 0
  'Asientos de CxC Cheque
   
   If Ret_NC Then
      sSQL = "SELECT F.TP, F.Fecha, C.Cliente, F.Serie, F.Factura, F.Banco, F.Cheque, F.Abono, F.Mes, F.Comprobante, F.Autorizacion, F.Serie_NC, " _
           & "Secuencial_NC, F.Autorizacion_NC, F.Base_Imponible, F.Porc, C.Representante As Razon_Social, F.Cta,F.Cta_CxP " _
           & "FROM Trans_Abonos As F, Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & Tipo_De_Consulta(True) _
           & "AND F.Banco LIKE 'RETENCION%' " _
           & "AND F.CodigoC = C.Codigo " _
           & "UNION " _
           & "SELECT F.TP, F.Fecha, C.Cliente, F.Serie, F.Factura, F.Banco, F.Cheque, F.Abono, F.Mes, F.Comprobante, F.Autorizacion, F.Serie_NC, " _
           & "Secuencial_NC, F.Autorizacion_NC, F.Base_Imponible, F.Porc, C.Representante As Razon_Social, F.Cta,F.Cta_CxP " _
           & "FROM Trans_Abonos As F,Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & Tipo_De_Consulta(True) _
           & "AND F.Banco = 'NOTA DE CREDITO' " _
           & "AND F.CodigoC = C.Codigo " _
           & "ORDER BY F.Banco, F.Cheque, C.Cliente, F.Serie, F.Factura, F.Fecha "
   Else
      sSQL = "SELECT F.TP, F.Fecha, C.Cliente, F.Serie, F.Factura, F.Recibo_No, F.Banco, F.Cheque, F.Abono, F.Mes, F.Comprobante, F.Autorizacion, F.Serie_NC, " _
           & "Secuencial_NC, F.Autorizacion_NC, C.Representante As Razon_Social, C.Grupo, C.Direccion As Ubicacion, F.Cta, F.Cta_CxP " _
           & "FROM Trans_Abonos As F,Clientes C " _
           & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & Tipo_De_Consulta(True) _
           & "AND F.CodigoC = C.Codigo " _
           & "ORDER BY F.Fecha, C.Cliente, F.Serie, F.Factura, F.Banco "
   End If
   Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   'Totales_CxC_Abonos
   LabelAbonado.Caption = Format$(Total, "#,##0.00")
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
   RatonNormal
End Sub

Public Sub Recibo_Abonos_Anticipados()
Dim SiEnviar As Boolean
  If Opcion = 20 Then
     SiEnviar = False
     DGQuery.Visible = False
     Titulo = "FORMULARIO DE ENVIO POR MAIL"
     Mensajes = "Enviar recibo de abono anticipado por mail?"
     If BoxMensaje = vbYes Then SiEnviar = True
     
     Co.Fecha = DGQuery.Columns(3)
     Co.TP = DGQuery.Columns(4)
     Co.Numero = DGQuery.Columns(5)
     
     sSQL = "SELECT C.Cliente, C.Email, C.Email2, C.CI_RUC, TS.Cta, TS.Fecha, TS.TP, TS.Numero, TS.Creditos As Abono, Co.Concepto " _
          & "FROM Trans_SubCtas As TS, Comprobantes AS Co, Clientes As C " _
          & "WHERE TS.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND TS.Item = '" & NumEmpresa & "' " _
          & "AND TS.Periodo = '" & Periodo_Contable & "' " _
          & "AND TS.TP = '" & Co.TP & "' " _
          & "AND TS.Numero = " & Co.Numero & " " _
          & "AND TS.T <> 'A' " _
          & Tipo_De_Consulta() _
          & "AND TS.Item = Co.Item " _
          & "AND TS.Periodo = Co.Periodo " _
          & "AND TS.TP = Co.TP " _
          & "AND TS.Numero = Co.Numero " _
          & "AND TS.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente, TS.Cta, TS.Fecha, TS.TP, TS.Numero "
     Select_Adodc AdoFacturas, sSQL
     With AdoFacturas.Recordset
      If .RecordCount > 0 Then
          Co.Beneficiario = .fields("Cliente")
          Co.RUC_CI = .fields("CI_RUC")
          Co.Concepto = .fields("Concepto")
          Co.Efectivo = .fields("Abono")
          Co.Email = ""
          If Len(.fields("Email")) > 3 Then Co.Email = .fields("Email")
          If Co.Email = "" And Len(.fields("Email2")) > 3 Then
             Co.Email = .fields("Email2")
          ElseIf Len(.fields("Email2")) > 3 And .fields("Email") <> .fields("Email2") Then
             Co.Email = Co.Email & ";" & .fields("Email2")
          End If
          Imprimir_Recibo_Anticipos Co, True
         'Procedemos a enviar por mail el recibo
          If Len(Co.Email) > 3 Then
             If SiEnviar Then
                TMail.para = Co.Email
                TMail.Asunto = "RECIBO ABONO ANTICIPADO No. " & Format$(Year(Co.Fecha), "0000") & "-" & Co.TP & "-" & Format$(Co.Numero, "000000000")
                TMail.Adjunto = RutaDocumentoPDF
                TMail.Mensaje = "Beneficiario: " & Co.Beneficiario & vbCrLf _
                              & "Fecha del Abono: " & Co.Fecha & vbCrLf _
                              & "Abono Anticipado por USD " & Format$(Co.Efectivo, "#,##0.00") & vbCrLf
                FEnviarCorreos.Show 1
             End If
          End If
      End If
     End With
     DGQuery.Visible = True
  End If
End Sub

Public Sub Abonos_Anticipados()
RatonReloj
Opcion = 6
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
   DGQuery.Caption = "ABONOS DE ANTICIPADOS"
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT TA.TP,F.Serie,F.Autorizacion,F.Fecha,F.Factura," _
        & "TA.Fecha As Fecha_Abono,TA.Abono " _
        & "FROM Facturas As F, Trans_Abonos As TA " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.Item = TA.Item " _
        & "AND F.Periodo = TA.Periodo " _
        & "AND F.TC = TA.TP " _
        & "AND F.Serie = TA.Serie " _
        & "AND F.Autorizacion = TA.Autorizacion " _
        & "AND F.Factura = TA.Factura " _
        & "AND F.Fecha > TA.Fecha " _
        & "ORDER BY TA.Fecha,F.Factura "
   Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   LabelAbonado.Caption = "0.00"
   LabelFacturado.Caption = "0.00"
   LabelSaldo.Caption = "0.00"
   DGQuery.Visible = True
   RatonNormal
End Sub

Public Sub Abonos_Erroneos()
RatonReloj
Opcion = 6
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False
PorCxC = False
   DGQuery.Caption = "ABONOS MAL PROCESADOS"
   Total = 0
   sSQL = "UPDATE Trans_Abonos " _
        & "SET X = 'E' " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL
   If SQL_Server Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET X = '.' " _
           & "FROM Trans_Abonos As TA,Facturas As F "
   Else
      sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
           & "SET TA.X = '.' "
   End If
   sSQL = sSQL & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.Item = TA.Item " _
        & "AND F.Periodo = TA.Periodo " _
        & "AND F.TC = TA.TP " _
        & "AND F.Serie = TA.Serie " _
        & "AND F.Autorizacion = TA.Autorizacion " _
        & "AND F.Factura = TA.Factura " _
        & "AND F.CodigoC = TA.CodigoC "
   Ejecutar_SQL_SP sSQL
  'Asientos de CxC Cheque
   sSQL = "SELECT TP,Serie,Autorizacion,Fecha,Factura,Abono,CodigoC " _
        & "FROM Trans_Abonos " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND X = 'E' " _
        & "AND TP <> 'CB' " _
        & "ORDER BY Fecha,Factura "
   Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
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
Label2.Caption = "(" & Opcion & ") I.V.A."
Label4.Caption = " VENTAS"
Label3.Caption = " TOTAL"
DGQuery.Visible = False

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT C.Cliente,SUM(F.Cantidad) As Cant_Prod,CP.Producto,F.Codigo,SUM(F.Total_IVA) As IVA," _
     & "SUM(F.Total) As Ventas,SUM(F.Cantidad*CP.Gramaje/1000) As Kilos,CP.Gramaje " _
     & "FROM Clientes As C, Detalle_Factura As F, Catalogo_Productos AS CP " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & Tipo_De_Consulta(, , True) _
     & "AND F.CodigoC = C.Codigo " _
     & "AND F.Item = CP.Item " _
     & "AND F.Periodo = CP.Periodo " _
     & "AND F.Codigo = CP.Codigo_Inv " _
     & "GROUP BY C.Cliente,F.Codigo,F.CodigoC,CP.Producto,CP.Gramaje " _
     & "ORDER BY C.Cliente,F.Codigo,F.CodigoC,CP.Producto,CP.Gramaje "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
DGQuery.Visible = False
Total = 0: Abono = 0
Total = T_Fields("IVA")
Abono = T_Fields("Ventas")
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Abono, "#,##0.00")
LabelSaldo.Caption = Format$(Total + Abono, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Ventas_Cliente()
'Resumen Ventas por Cliente
Opcion = 4
RatonReloj
Label2.Caption = "(" & Opcion & ") Ventas"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"
DGQuery.Visible = False

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT C.Cliente,F.TC, COUNT(F.CodigoC) As Cant_Fact,SUM(F.Total) As Ventas," _
     & "SUM(F.Total_IVA) As I_V_A, SUM(F.Total + F.Total_IVA) As Total_Facturado " _
     & "FROM Detalle_Factura As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & Tipo_De_Consulta(, , True) _
     & "AND F.CodigoC = C.Codigo " _
     & "GROUP BY C.Cliente,F.TC " _
     & "ORDER BY SUM(F.Total + F.Total_IVA) DESC,C.Cliente "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
RatonReloj
DGQuery.Visible = False
Total = 0: Abono = 0
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     Do While Not .EOF
        Total = Total + .fields("Ventas")
        Abono = Abono + .fields("I_V_A")
       .MoveNext
     Loop
    .MoveFirst
 End If
End With
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Abono, "#,##0.00")
LabelSaldo.Caption = Format$(Total - Abono, "#,##0.00")
DGQuery.Visible = True
RatonNormal
End Sub

Public Sub Resumen_Prod_Meses()
'Resumen Ventas por Cliente Mensual
Dim Nom_Mes(13) As String
Dim sSQLx As String
Dim PorCantidad As Boolean
    PorCantidad = False
    Mensajes = "(SI) Reporte por Cantidad" & vbCrLf & "(NO) Por Valor Economico"
    Titulo = "Formulario de Confirmación"
    If BoxMensaje = vbYes Then PorCantidad = True
    
    Opcion = 16
    RatonReloj
    Label2.Caption = "(" & Opcion & ") Ventas"
    Label4.Caption = " Cobrado"
    Label3.Caption = " Saldo"
    DGQuery.Visible = False
    
    
    
    DGQuery.Caption = "RESUMEN DE VENTAS DE PRODUCTOS MENSUALIZADO"
  
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' "
    Ejecutar_SQL_SP sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Catalogo_Productos " _
            & "SET X = 'X' " _
            & "FROM Catalogo_Productos As CP, Detalle_Factura As DF "
    Else
       sSQL = "UPDATE Catalogo_Productos As CP, Detalle_Factura As DF " _
            & "SET CP.X = 'X' "
    End If
    sSQL = sSQL _
         & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND CP.Item = '" & NumEmpresa & "' " _
         & "AND CP.Periodo = '" & Periodo_Contable & "' " _
         & "AND CP.TC = 'P' " _
         & "AND CP.Item = DF.Item " _
         & "AND CP.Periodo = DF.Periodo " _
         & "AND CP.Codigo_Inv = DF.Codigo "
    Ejecutar_SQL_SP sSQL
    
    Nom_Mes(1) = "Enero"
    Nom_Mes(2) = "Febrero"
    Nom_Mes(3) = "Marzo"
    Nom_Mes(4) = "Abril"
    Nom_Mes(5) = "Mayo"
    Nom_Mes(6) = "Junio"
    Nom_Mes(7) = "Julio"
    Nom_Mes(8) = "Agosto"
    Nom_Mes(9) = "Septiembre"
    Nom_Mes(10) = "Octubre"
    Nom_Mes(11) = "Noviembre"
    Nom_Mes(12) = "Diciembre"
  
    sSQLx = "Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre,Total"

    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'RPXM' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "INSERT INTO Saldo_Diarios (TC, Codigo_Aux, Item, CodigoU, TP, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, " _
         & "Agosto, Septiembre, Octubre, Noviembre, Diciembre, Total) " _
         & "SELECT TC,Codigo_Inv,'" & NumEmpresa & "' As Itemx,'" & CodigoUsuario & "' As CodigoUs,'RPXM' As TPs, " _
         & "0 As Enerox, 0 As Febrerox, 0 As Marzox, 0 As Abrilx, 0 As Mayox, 0 As Juniox, 0 As Juliox, " _
         & "0 As Agostox, 0 As Septiembrex, 0 As Octubrex, 0 As Noviembrex, 0 As Diciembrex, 0 As Totalx " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND X = 'X' "
    'If FA.Cod_Ejec <> Ninguno Then sSQL = sSQL & "AND Cod_Ejec = '" & FA.Cod_Ejec & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'RPXM' "
    Select_Adodc AdoQuery1, sSQL

    For NoMes = 1 To 12
        sSQL = "UPDATE Saldo_Diarios "
        If PorCantidad Then
           sSQL = sSQL & "SET " & Nom_Mes(NoMes) & " = (SELECT SUM(Cantidad) "
        Else
           sSQL = sSQL & "SET " & Nom_Mes(NoMes) & " = (SELECT SUM(Total-Total_Desc-Total_Desc2) "
        End If
        sSQL = sSQL _
             & "               FROM Detalle_Factura As F " _
             & "               WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
             & "               AND F.Item = '" & NumEmpresa & "' " _
             & "               AND F.Periodo = '" & Periodo_Contable & "' " _
             & "               AND F.T <> '" & Anulado & "' " _
             & "               AND MONTH(F.Fecha) = " & NoMes & " " _
             & "               AND F.Codigo = Saldo_Diarios.Codigo_Aux " _
             & "               AND F.Item = Saldo_Diarios.Item) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "AND TP = 'RPXM' "
        Ejecutar_SQL_SP sSQL
        sSQL = "UPDATE Saldo_Diarios " _
             & "SET " & Nom_Mes(NoMes) & " = 0 " _
             & "WHERE " & Nom_Mes(NoMes) & " IS NULL " _
             & "AND Item = '" & NumEmpresa & "' "
        Ejecutar_SQL_SP sSQL
    Next NoMes

    sSQLx = "Total=Enero+Febrero+Marzo+Abril+Mayo+Junio+Julio+Agosto+Septiembre+Octubre+Noviembre+Diciembre"
    sSQL = "UPDATE Saldo_Diarios " _
         & "SET " & sSQLx & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'RPXM' "
    Ejecutar_SQL_SP sSQL
        
    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Total = 0 " _
         & "AND TP = 'RPXM' "
    Ejecutar_SQL_SP sSQL
    
    sSQLx = ""
    For NoMes = 1 To Month(MBFechaF)
        sSQLx = sSQLx & ",SD." & Nom_Mes(NoMes)
    Next NoMes
    sSQL = "SELECT SD.Codigo_Aux As Codigos, CP.Producto, CP.Unidad " & sSQLx & ",SD.Total " _
         & "FROM Saldo_Diarios As SD,Catalogo_Productos As CP " _
         & "WHERE CP.Item = '" & NumEmpresa & "' " _
         & "AND CP.Periodo = '" & Periodo_Contable & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
         & "AND SD.TP = 'RPXM' " _
         & "AND SD.Codigo_Aux = CP.Codigo_Inv " _
         & "AND SD.Item = CP.Item " _
         & "ORDER BY SD.Total DESC, CP.Producto, SD.Codigo_Aux "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
    RatonReloj
    
    DGQuery.Visible = False
    Total = 0: Abono = 0
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Total = Total + .fields("Total")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    LabelFacturado.Caption = Format$(Total, "#,##0.00")
    LabelAbonado.Caption = Format$(Abono, "#,##0.00")
    LabelSaldo.Caption = Format$(Total - Abono, "#,##0.00")
    DGQuery.Visible = True
    RatonNormal
End Sub

Public Sub Ventas_Clientes_Por_Meses()
'Resumen Ventas por Cliente Mensual
Dim Nom_Mes(13) As String
Dim sSQLx As String
Dim CantProm As Integer

    Opcion = 14
    RatonReloj
    Label2.Caption = "(" & Opcion & ") Ventas"
    Label4.Caption = " Cobrado"
    Label3.Caption = " Saldo"
    
    DGQuery.Caption = "VENTAS POR CLIENTES MENSUALIZADO"
  
    sSQL = "UPDATE Clientes " _
         & "SET X = '.' " _
         & "WHERE FA <> " & Val(adFalse) & " "
    Ejecutar_SQL_SP sSQL
    
    DGQuery.Visible = False
    
    If SQL_Server Then
       sSQL = "UPDATE Clientes " _
            & "SET X = 'X', FA = 1 " _
            & "FROM Clientes As C, Facturas As F "
    Else
       sSQL = "UPDATE Clientes As C, Facturas As F " _
            & "SET C.X = 'X' "
    End If
    sSQL = sSQL _
         & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND C.FA <> " & Val(adFalse) & " " _
         & "AND C.Codigo = F.CodigoC "
    Ejecutar_SQL_SP sSQL
    
    DGQuery.Visible = False
    
    Nom_Mes(1) = "Enero"
    Nom_Mes(2) = "Febrero"
    Nom_Mes(3) = "Marzo"
    Nom_Mes(4) = "Abril"
    Nom_Mes(5) = "Mayo"
    Nom_Mes(6) = "Junio"
    Nom_Mes(7) = "Julio"
    Nom_Mes(8) = "Agosto"
    Nom_Mes(9) = "Septiembre"
    Nom_Mes(10) = "Octubre"
    Nom_Mes(11) = "Noviembre"
    Nom_Mes(12) = "Diciembre"
  
    sSQLx = "Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre,Total"

    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'VCXM' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "INSERT INTO Saldo_Diarios (Cta, CodigoC, Item, CodigoU, TP, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, " _
         & "Agosto, Septiembre, Octubre, Noviembre, Diciembre, Total) " _
         & "SELECT Cod_Ejec, Codigo,'" & NumEmpresa & "' As Itemx,'" & CodigoUsuario & "' As CodigoUs,'VCXM' As TPs, " _
         & "0 As Enerox, 0 As Febrerox, 0 As Marzox, 0 As Abrilx, 0 As Mayox, 0 As Juniox, 0 As Juliox, " _
         & "0 As Agostox, 0 As Septiembrex, 0 As Octubrex, 0 As Noviembrex, 0 As Diciembrex, 0 As Totalx " _
         & "FROM Clientes " _
         & "WHERE FA <> " & Val(adFalse) & " " _
         & "AND X = 'X' "
    If Len(FA.Cod_Ejec) > 1 Then sSQL = sSQL & "AND Cod_Ejec = '" & FA.Cod_Ejec & "' "
    If Len(FA.CodigoC) > 1 Then sSQL = sSQL & "AND Codigo = '" & FA.CodigoC & "' "
    If Len(FA.Grupo) > 1 Then sSQL = sSQL & "AND Grupo = '" & FA.Grupo & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'VCXM' "
    Select_Adodc AdoQuery1, sSQL
     
    For NoMes = 1 To 12
        sSQL = "UPDATE Saldo_Diarios " _
             & "SET " & Nom_Mes(NoMes) & " = (SELECT SUM(Total_MN-IVA) " _
             & "               FROM Facturas As F " _
             & "               WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
             & "               AND F.Item = '" & NumEmpresa & "' " _
             & "               AND F.Periodo = '" & Periodo_Contable & "' " _
             & "               AND F.T <> '" & Anulado & "' " _
             & "               AND MONTH(F.Fecha) = " & NoMes & " " _
             & "               AND F.CodigoC = Saldo_Diarios.CodigoC " _
             & "               AND F.Item = Saldo_Diarios.Item) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "AND TP = 'VCXM' "
        Ejecutar_SQL_SP sSQL
         
        sSQL = "UPDATE Saldo_Diarios " _
             & "SET " & Nom_Mes(NoMes) & " = 0 " _
             & "WHERE " & Nom_Mes(NoMes) & " IS NULL " _
             & "AND Item = '" & NumEmpresa & "' "
        Ejecutar_SQL_SP sSQL
    Next NoMes

    sSQLx = "Total=Enero+Febrero+Marzo+Abril+Mayo+Junio+Julio+Agosto+Septiembre+Octubre+Noviembre+Diciembre"
    sSQL = "UPDATE Saldo_Diarios " _
         & "SET " & sSQLx & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'VCXM' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Saldo_Diarios " _
         & "SET Grupo_No = RP.Cod_Ejec " _
         & "FROM Saldo_Diarios As SD, Catalogo_Rol_Pagos As RP " _
         & "WHERE RP.Item = '" & NumEmpresa & "' " _
         & "AND RP.Periodo = '" & Periodo_Contable & "' " _
         & "AND SD.Cta = RP.Codigo " _
         & "AND SD.Item = RP.Item "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Total = 0 " _
         & "AND TP = 'VCXM' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'VCXM' "
    Select_Adodc AdoQuery1, sSQL
    With AdoQuery1.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            CantProm = 0
            Total = 0
            For NoMes = 1 To 12
                If .fields(Nom_Mes(NoMes)) <> 0 Then
                    Total = Total + .fields(Nom_Mes(NoMes))
                    CantProm = CantProm + 1
                End If
            Next NoMes
            If CantProm <= 0 Then CantProm = 1
           .fields("Diferencia") = Redondear(Total / CantProm, 2)
           .Update
           .MoveNext
         Loop
     End If
    End With
    
    sSQLx = ""
    For NoMes = 1 To Month(MBFechaF)
        sSQLx = sSQLx & ",SD." & Nom_Mes(NoMes)
    Next NoMes
    sSQL = "SELECT SD.Grupo_No As Ejecutivo,C.Grupo,C.Cliente " _
         & sSQLx & ",SD.Total,SD.Diferencia As Promedio " _
         & "FROM Saldo_Diarios As SD,Clientes As C " _
         & "WHERE SD.Item = '" & NumEmpresa & "' " _
         & "AND SD.CodigoU = '" & CodigoUsuario & "' "
    If Len(FA.Cod_Ejec) > 1 Then sSQL = sSQL & "AND SD.Cta = '" & FA.Cod_Ejec & "' "
    If Len(FA.CodigoC) > 1 Then sSQL = sSQL & "AND SD.Codigo = '" & FA.CodigoC & "' "
    'If Len(FA.Grupo) > 1 Then sSQL = sSQL & "AND Grupo = '" & FA.Grupo & "' "
    sSQL = sSQL _
         & "AND SD.TP = 'VCXM' " _
         & "AND SD.CodigoC = C.Codigo " _
         & "ORDER BY SD.Total DESC,SD.Grupo_No,C.Grupo, C.Cliente "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   'MsgBox sSQL
    RatonReloj
    DGQuery.Visible = False
    Total = 0: Abono = 0
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Total = Total + .fields("Total")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    LabelFacturado.Caption = Format$(Total, "#,##0.00")
    LabelAbonado.Caption = Format$(Abono, "#,##0.00")
    LabelSaldo.Caption = Format$(Total - Abono, "#,##0.00")
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
Label2.Caption = "(" & Opcion & ") Ventas"
Label4.Caption = "  "
Label3.Caption = "  "
DGQuery.Visible = False

DGQuery.Caption = "HISTORIAL DE FACTURAS Y PRODUCTOS"
sSQL = "SELECT * " _
     & "FROM Catalogo_Productos " _
     & "WHERE TC = 'P' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv "
Select_Adodc AdoHistoria, sSQL

sSQL = "SELECT * " _
     & "FROM Trans_Kardex " _
     & "WHERE Fecha <= #" & FechaFin & "# " _
     & "AND T <> 'A' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv,Fecha,ID "
Select_Adodc AdoQuery, sSQL
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     
     Codigo = .fields("Codigo_Inv")
     TotalIngreso = 0
     Do While Not .EOF
        If Codigo <> .fields("Codigo_Inv") Then
           If AdoHistoria.Recordset.RecordCount > 0 Then
              AdoHistoria.Recordset.MoveFirst
              AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
              If Not AdoHistoria.Recordset.EOF Then
                 If AdoHistoria.Recordset.fields("Valor_Compra") <> TotalIngreso Then
                    AdoHistoria.Recordset.fields("Valor_Compra") = TotalIngreso
                    AdoHistoria.Recordset.Update
                 End If
              End If
           End If
           Codigo = .fields("Codigo_Inv")
           TotalIngreso = 0
        End If
        TotalIngreso = .fields("Valor_Unitario")
       .MoveNext
     Loop
     If AdoHistoria.Recordset.RecordCount > 0 Then
        AdoHistoria.Recordset.MoveFirst
        AdoHistoria.Recordset.Find ("Codigo_Inv = '" & Codigo & "' ")
        If Not AdoHistoria.Recordset.EOF Then
           If AdoHistoria.Recordset.fields("Valor_Compra") <> TotalIngreso Then
              AdoHistoria.Recordset.fields("Valor_Compra") = TotalIngreso
              AdoHistoria.Recordset.Update
           End If
        End If
     End If
     
 End If
End With
sSQL = "SELECT F.Codigo,CP.Producto,SUM(F.Cantidad) As Cant_Prod,SUM(F.Total) As Ventas," _
     & "SUM(F.Cantidad*CP.Gramaje/1000) As Kilos,CP.Desc_Item " & Con_Costeo
If Si_No Then sSQL = sSQL & ",AVG(F.Precio) As PVP,Valor_Compra As Costos "
sSQL = sSQL & "FROM Detalle_Factura As F,Catalogo_Productos As CP,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND CP.INV <> " & Val(adFalse) & " " _
     & Tipo_De_Consulta(, , True)
If DescItem <> Ninguno Then sSQL = sSQL & "AND CP.Desc_Item = '" & DescItem & "' "
sSQL = sSQL & "AND F.Item = CP.Item " _
     & "AND F.Periodo = CP.Periodo " _
     & "AND F.Codigo = CP.Codigo_Inv " _
     & "AND F.CodigoC = C.Codigo "
If Si_No Then
   If DescItem <> Ninguno Then
      sSQL = sSQL & "GROUP BY CP.Desc_Item,F.Codigo,CP.Valor_Compra "
   Else
      sSQL = sSQL & "GROUP BY F.Codigo,CP.Valor_Compra,CP.Producto,CP.Desc_Item "
   End If
Else
   If DescItem <> Ninguno Then
      sSQL = sSQL & "GROUP BY CP.Desc_Item,F.Codigo,CP.Producto "
   Else
      sSQL = sSQL & "GROUP BY F.Codigo,CP.Producto,CP.Desc_Item "
   End If
End If
If DescItem <> Ninguno Then
   sSQL = sSQL & "ORDER BY CP.Desc_Item,F.Codigo,SUM(F.Total) DESC "
Else
   sSQL = sSQL & "ORDER BY F.Codigo,SUM(F.Total) DESC,CP.Producto "
End If
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
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
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = " PVP"
Label3.Caption = " Costo"

LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Abono, "#,##0.00")
LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
DGQuery.Visible = True

RatonNormal
End Sub

Public Sub Resumen_Ventas_Vendedor()
    sSQL = "SELECT C.Grupo,C.Cliente, F.Fecha, TA.Fecha As Fecha_A, F.Serie, TA.Factura, CONVERT(Money,TA.Abono/(1+F.Porc_IVA)) As Abonos, " _
         & "DATEDIFF(day,F.Fecha,TA.Fecha) As Dias_T, A.Nombre_Completo " _
         & "FROM Clientes As C, Facturas As F, Trans_Abonos As TA, Accesos As A " _
         & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND TA.Item = '" & NumEmpresa & "' " _
         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
         & "AND NOT SUBSTRING(TA.Banco,1,9) IN ('RETENCION','NOTA DE C') " _
         & Tipo_De_Consulta() _
         & "AND C.Codigo = F.CodigoC " _
         & "AND A.Codigo = F.Cod_Ejec " _
         & "AND F.Item = TA.Item " _
         & "AND F.Periodo = TA.Periodo " _
         & "AND F.TC = TA.TP " _
         & "AND F.Serie = TA.Serie " _
         & "AND F.Autorizacion = TA.Autorizacion " _
         & "AND F.Factura = TA.Factura " _
         & "AND F.CodigoC = TA.CodigoC " _
         & "ORDER BY C.Grupo,C.Cliente,F.Serie,F.Factura "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   Opcion = 15
End Sub

Public Sub Ventas_Resumindas_x_Vendedor()
Dim sSQLGrupo As String
Dim sSQLSubTotal As String
    sSQL = "UPDATE Accesos " _
         & "SET Cuota_Venta = 1 " _
         & "WHERE Cuota_Venta = 0 "
    Ejecutar_SQL_SP sSQL
    
    sSQLGrupo = "SELECT A.Cod_Ejec,A.Nombre_Completo As Nombre_Vendedor,C.Grupo,CC.Cuenta,SUM(F.SubTotal - F.Descuento - F.Descuento2) As Cantidad, ' ' As Cuota " _
              & "FROM Facturas As F, Catalogo_Cuentas As CC, Accesos As A, Clientes As C " _
              & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
              & "AND F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "AND F.T <> '" & Anulado & "' " _
              & "AND A.Codigo = F.Cod_Ejec " _
              & "AND C.Codigo = F.CodigoC " _
              & "AND F.Item = CC.Item " _
              & "AND F.Periodo = CC.Periodo " _
              & "AND F.Cta_CxP = CC.Codigo " _
              & "GROUP BY C.Grupo, A.Cod_Ejec,A.Nombre_Completo, A.Cuota_Venta, CC.Cuenta "
    sSQLSubTotal = "SELECT A.Cod_Ejec, ' ' As Nombre_Vendedor, ' ' As Grupo, 'SUBTOTAL VENDEDOR' As Cuenta,SUM(F.SubTotal - F.Descuento - F.Descuento2) As Cantidad, " _
                 & "STR((SUM(F.SubTotal - F.Descuento - F.Descuento2)/A.Cuota_Venta)*100)+'%' As Cuota " _
                 & "FROM Facturas As F, Catalogo_Cuentas As CC, Accesos As A, Clientes As C " _
                 & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                 & "AND F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' " _
                 & "AND F.T <> '" & Anulado & "' " _
                 & "AND A.Codigo = F.Cod_Ejec " _
                 & "AND C.Codigo = F.CodigoC " _
                 & "AND F.Item = CC.Item " _
                 & "AND F.Periodo = CC.Periodo " _
                 & "AND F.Cta_CxP = CC.Codigo " _
                 & "GROUP BY A.Cod_Ejec, A.Cuota_Venta "
    sSQL = sSQLGrupo _
         & " UNION " _
         & sSQLSubTotal _
         & "ORDER BY A.Cod_Ejec, A.Nombre_Completo DESC, C.Grupo, CC.Cuenta "
   ' sSQL = sSQLSubTotal
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
    DGQuery.Visible = False
    Total = 0
     
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Total = Total + .fields("Cantidad")
           .MoveNext
         Loop
     End If
    End With
    LabelFacturado.Caption = Format(Total, "#,##0.00")
    DGQuery.Visible = True
   Opcion = 17
End Sub

Public Sub CxC_Tiempo_Credito()
Dim sSQLV As String
Dim sSQLT As String

    Mifecha = BuscarFecha(FechaSistema)
    
    sSQL = "UPDATE Facturas " _
         & "SET Venc_0_60=0,Venc_61_90=0,Venc_91_120=0,Venc_121_360=0,Venc_mas_360=0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Venc_0_60 = Saldo_MN " _
         & "WHERE DATEDIFF(DAY,Fecha, '" & Mifecha & "') BETWEEN 0 and 60 " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T = '" & Pendiente & "' "
    Ejecutar_SQL_SP sSQL
   
    sSQL = "UPDATE Facturas " _
         & "SET Venc_61_90 = Saldo_MN " _
         & "WHERE DATEDIFF(DAY,Fecha, '" & Mifecha & "') BETWEEN 61 and 90 " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T = '" & Pendiente & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "UPDATE Facturas " _
         & "SET Venc_91_120 = Saldo_MN " _
         & "WHERE DATEDIFF(DAY,Fecha, '" & Mifecha & "') BETWEEN 91 and 120 " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T = '" & Pendiente & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "UPDATE Facturas " _
         & "SET Venc_121_360 = Saldo_MN " _
         & "WHERE DATEDIFF(DAY,Fecha, '" & Mifecha & "') BETWEEN 121 and 360 " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T = '" & Pendiente & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "UPDATE Facturas " _
         & "SET Venc_mas_360 = Saldo_MN " _
         & "WHERE DATEDIFF(DAY,Fecha, '" & Mifecha & "') > 360 " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T = '" & Pendiente & "' "
    Ejecutar_SQL_SP sSQL

    sSQLV = "SELECT A.Nombre_Completo As Nombre_Vendedor, C.Cliente As Clientes, YEAR(F.Fecha) As Año, MONTH(F.Fecha) As Mes, " _
          & "SUM(Venc_0_60) As _0_60, " _
          & "SUM(Venc_61_90) As _61_90, " _
          & "SUM(Venc_91_120) As _61_120, " _
          & "SUM(Venc_121_360) As _121_360, " _
          & "SUM(Venc_mas_360) As _mas_360, " _
          & "SUM(Venc_0_60+Venc_61_90+Venc_91_120+Venc_121_360+Venc_mas_360) As Saldo_Total, " _
          & "SUM(F.Total_MN) As Total_Facturado " _
          & "FROM Facturas As F, Accesos As A, Clientes As C " _
          & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND F.Item = '" & NumEmpresa & "' " _
          & "AND F.Periodo = '" & Periodo_Contable & "' " _
          & "AND F.T = '" & Pendiente & "' "
    If Len(FA.Cod_Ejec) > 1 Then sSQLV = sSQLV & "AND C.Cod_Ejec = '" & FA.Cod_Ejec & "' "
    sSQLV = sSQLV _
          & "AND A.Codigo = F.Cod_Ejec " _
          & "AND C.Codigo = F.CodigoC " _
          & "GROUP BY A.Nombre_Completo, C.Cliente, YEAR(F.Fecha), MONTH(F.Fecha) "
    
    sSQLT = "SELECT A.Nombre_Completo As Nombre_Vendedor, 'zz" & String(40, " ") & "SUBTOTALES' As Clientes, " & Year(FechaSistema) & " As Año, " & Month(FechaSistema) & " As Mes, " _
          & "SUM(Venc_0_60) As T_Venc_0_60, " _
          & "SUM(Venc_61_90) As T_Venc_61_90, " _
          & "SUM(Venc_91_120) As T_Venc_61_90, " _
          & "SUM(Venc_121_360) As T_Venc_121_360, " _
          & "SUM(Venc_mas_360) As T_Venc_mas_360, " _
          & "SUM(Venc_0_60+Venc_61_90+Venc_91_120+Venc_121_360+Venc_mas_360) As Saldo_Total, " _
          & "SUM(F.Total_MN) As Total_Facturado " _
          & "FROM Facturas As F, Accesos As A, Clientes As C " _
          & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND F.Item = '" & NumEmpresa & "' " _
          & "AND F.Periodo = '" & Periodo_Contable & "' " _
          & "AND F.T = '" & Pendiente & "' "
    If Len(FA.Cod_Ejec) > 1 Then sSQLT = sSQLT & "AND C.Cod_Ejec = '" & FA.Cod_Ejec & "' "
    sSQLT = sSQLT _
          & "AND A.Codigo = F.Cod_Ejec " _
          & "AND C.Codigo = F.CodigoC " _
          & "GROUP BY A.Nombre_Completo "
         
    sSQL = sSQLV & "UNION " & sSQLT _
         & "ORDER BY A.Nombre_Completo, Clientes "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL
    DGQuery.Visible = False
    Total = 0
    Saldo = 0
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            If .fields("Clientes") = "zz" & String(40, " ") & "SUBTOTALES" Then
                Total = Total + .fields("Total_Facturado")
                Saldo = Saldo + .fields("Saldo_Total")
            End If
           .MoveNext
         Loop
     End If
    End With
    LabelFacturado.Caption = Format(Total, "#,##0.00")
    LabelSaldo.Caption = Format(Saldo, "#,##0.00")
    DGQuery.Visible = True
    Opcion = 18
End Sub

Public Sub Cheques_Protestados()
'Cheques protestado
RatonReloj
Label2.Caption = "(" & Opcion & ") Facturado"
Label4.Caption = " Cobrado"
Label3.Caption = " Saldo"

   DGQuery.Caption = "ABONOS DE FACTURAS"
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT F.TP,F.Fecha,C.Cliente,F.Factura,F.Banco,F.Cheque,F.Abono,F.Comprobante,F.Cta,F.Cta_CxP " _
        & "FROM Trans_Abonos As F,Clientes C " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & Tipo_De_Consulta() _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.Protestado <> " & Val(adFalse) & " " _
        & "ORDER BY C.Cliente,F.Factura,F.Fecha,F.Banco "
   Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   Opcion = 7
   Totales_CxC_Abonos
   RatonNormal
End Sub

Private Sub CheqIngreso_Click()
  If CheqIngreso.value = 1 Then
     sSQL = "SELECT (DF.Cta_Venta & ' - ' & CC.Cuenta) As NomCxC " _
          & "FROM Detalle_Factura As DF, Catalogo_Cuentas As CC " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND CC.Periodo = DF.Periodo " _
          & "AND CC.Item = DF.Item " _
          & "AND CC.Codigo = DF.Cta_Venta " _
          & "GROUP BY DF.Cta_Venta, CC.Cuenta " _
          & "ORDER BY DF.Cta_Venta "
     SelectDB_Combo DCCxC, AdoCxC, sSQL, "NomCxC"
     CheqCxC.value = 0
     CheqAbonos.value = 0
     DCCxC.Visible = True
  Else
     DCCxC.Visible = False
  End If
End Sub

Private Sub Command1_Click()
  Unload HistorialFacturas
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
   SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  Select Case KeyCode
    Case vbKeyEscape
         LblPatronBusqueda.Caption = ""
         FrmPatronBusqueda.Visible = False
         PresionoEsc = True
    Case vbKeyReturn
         Pulsar_Tecla (vbKeyTab)
         FrmPatronBusqueda.Visible = False
    Case vbKeyTab
         FrmPatronBusqueda.Visible = False
  End Select
End Sub

Private Sub DCCliente_LostFocus()
 FA.Cod_Ejec = Ninguno
 FA.CodigoC = Ninguno
 FA.Cliente = Ninguno
 FA.CI_RUC = Ninguno
 FA.Grupo = Ninguno
 FA.CiudadC = Ninguno
 FA.Autorizacion = Ninguno
 FA.Forma_Pago = Ninguno
 FA.TC = Ninguno
 FA.Serie = Ninguno
 FA.Factura = 0
 CodigoInv = Ninguno
 Cod_Marca = Ninguno
 DescItem = Ninguno
 If Not PresionoEsc Then
    With AdoCliente.Recordset
     If .RecordCount > 0 Then
         Select Case ListCliente.Text
           Case "Codigo"
                FA.CodigoC = DCCliente
           Case "CI_RUC"
               .MoveFirst
               .Find ("CI_RUC = '" & DCCliente & "' ")
                If Not .EOF Then
                   FA.CodigoC = .fields("Codigo")
                   FA.Cliente = .fields("Cliente")
                   FA.CI_RUC = .fields("CI_RUC")
                End If
           Case "Ciudad"
                FA.CiudadC = DCCliente
           Case "Cliente"
               .MoveFirst
               .Find ("Cliente = '" & DCCliente & "' ")
                If Not .EOF Then
                   FA.CodigoC = .fields("CodigoC")
                   FA.Cliente = .fields("Cliente")
                End If
           Case "Vendedor"
               .MoveFirst
               .Find ("Cliente = '" & DCCliente & "' ")
                If Not .EOF Then FA.Cod_Ejec = .fields("Codigo")
           Case "Grupo"
                FA.Grupo = DCCliente
           Case "Factura"
                FA.TC = SinEspaciosIzq(DCCliente)
                FA.Serie = MidStrg(DCCliente, 4, 6)
                FA.Factura = Val(SinEspaciosDer(DCCliente))
                LblPatronBusqueda.Caption = "P A T R O N   D E   B U S Q U E D A:" & vbCrLf _
                                          & ListCliente & " = " & FA.TC & ": " & FA.Serie & "-" & FA.Factura
           Case "Serie"
                FA.Serie = DCCliente
           Case "Autorizacion"
                FA.Autorizacion = DCCliente
           Case "Forma_Pago"
                FA.Forma_Pago = DCCliente
           Case "Plan_Afiliado"
           Case "Tipo Documento"
                FA.TC = DCCliente
           Case "Marca"
                DescItem = SinEspaciosIzq(DCCliente)
           Case "DescItem"
                Cod_Marca = DCCliente
           Case "Producto"
                CodigoInv = TrimStrg(SinEspaciosIzq(DCCliente))
                Producto = TrimStrg(MidStrg(DCCliente, Len(CodigoInv) + 1, Len(DCCliente)))
         End Select
     End If
    End With
    If ListCliente <> "Factura" Then LblPatronBusqueda.Caption = "P A T R O N   D E   B U S Q U E D A:" & vbCrLf _
                                                               & ListCliente & " = " & DCCliente
    FrmPatronBusqueda.Visible = False
 End If
End Sub

Private Sub DCCxC_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
Dim OpcBusqueda() As String
Dim TempOpcBusqueda As String
Dim CantOpcBusqueda As Byte

   CnExterna = 0
   HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
   Actualizar_Datos_Representantes_SP Mas_Grupos
      
   CantOpcBusqueda = 16
   ReDim OpcBusqueda(CantOpcBusqueda) As String
   OpcBusqueda(0) = "Cliente"
   OpcBusqueda(1) = "CI_RUC"
   OpcBusqueda(2) = "Ciudad"
   OpcBusqueda(3) = "Codigo"
   OpcBusqueda(4) = "Plan_Afiliado"
   OpcBusqueda(5) = "Tipo Documento"
   OpcBusqueda(6) = "Autorizacion"
   OpcBusqueda(7) = "Serie"
   OpcBusqueda(8) = "Factura"
   OpcBusqueda(9) = "Forma_Pago"
   OpcBusqueda(10) = "Cuenta_No"
   OpcBusqueda(11) = "Vendedor"
   OpcBusqueda(12) = "Grupo/Zona"
   OpcBusqueda(13) = "Producto"
   OpcBusqueda(14) = "DescItem"
   OpcBusqueda(15) = "Marca"
   
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
   ListCliente.AddItem "Todos"
   ListCliente.Text = "Todos"
   For I = 0 To CantOpcBusqueda - 1
       ListCliente.AddItem OpcBusqueda(I)
   Next I
   
   If TipoFactura = "" Then TipoFactura = Ninguno
   
''   sSQL = "SELECT (Codigo & ' - ' & Concepto) As NomCxC " _
''        & "FROM Catalogo_Lineas " _
''        & "WHERE TL <> " & Val(adFalse) & " " _
''        & "AND Item = '" & NumEmpresa & "' " _
''        & "AND Periodo = '" & Periodo_Contable & "' " _
''        & "ORDER BY Codigo "
''   SelectDB_Combo DCCxC, AdoCxC, sSQL, "NomCxC"
   
   DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
   DGQuery.width = MDI_X_Max - 100
   LblPatronBusqueda.width = MDI_X_Max - LblPatronBusqueda.Left - 100
   AdoQuery.Top = DGQuery.Top + DGQuery.Height
   Command1.Top = DGQuery.Top + DGQuery.Height
   Label2.Top = DGQuery.Top + DGQuery.Height
   Label3.Top = DGQuery.Top + DGQuery.Height
   Label4.Top = DGQuery.Top + DGQuery.Height
   LabelFacturado.Top = DGQuery.Top + DGQuery.Height
   LabelAbonado.Top = DGQuery.Top + DGQuery.Height
   LabelSaldo.Top = DGQuery.Top + DGQuery.Height
   
   FA.Cliente = Ninguno
   FA.CI_RUC = Ninguno
   FA.Factura = 0
   FA.Cod_Ejec = Ninguno
   FA.CodigoC = Ninguno
   FA.Cliente = Ninguno
   FA.CI_RUC = Ninguno
   FA.Grupo = Ninguno
   FA.CiudadC = Ninguno
   FA.Autorizacion = Ninguno
   FA.Forma_Pago = Ninguno
   FA.TC = Ninguno
   FA.Serie = Ninguno
   FA.Factura = 0
   CodigoInv = Ninguno
   Cod_Marca = Ninguno
   DescItem = Ninguno

   RatonNormal
   MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoCxC
   ConectarAdodc AdoTipo
   ConectarAdodc AdoQuery
   ConectarAdodc AdoQuery1
   ConectarAdodc AdoCliente
   ConectarAdodc AdoHistoria
   ConectarAdodc AdoFacturas
End Sub

Private Sub ListCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then
     LblPatronBusqueda.Caption = ""
     FrmPatronBusqueda.Visible = False
     PresionoEsc = True
  End If
  PresionoEnter KeyCode
End Sub

Private Sub ListCliente_LostFocus()
Dim Nombre_Campo As String
  If Not PresionoEsc Then
     Nombre_Campo = ListCliente.Text
     Select Case ListCliente.Text
       Case "Autorizacion"
            sSQL = "SELECT Autorizacion " _
                 & "FROM Facturas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "GROUP BY Autorizacion " _
                 & "ORDER BY Autorizacion DESC "
            Nombre_Campo = "Autorizacion"
       Case "Serie"
            sSQL = "SELECT Serie " _
                 & "FROM Facturas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "GROUP BY Serie " _
                 & "ORDER BY Serie "
            Nombre_Campo = "Serie"
       Case "Codigo"
            sSQL = "SELECT Codigo,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C,Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY Codigo " _
                 & "ORDER BY Codigo "
            Nombre_Campo = "Codigo"
       Case "CI_RUC"
            sSQL = "SELECT CI_RUC,Cliente,Codigo,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C, Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY CI_RUC,Cliente,Codigo " _
                 & "ORDER BY CI_RUC "
            Nombre_Campo = "CI_RUC"
       Case "Cliente"
            sSQL = "SELECT F.CodigoC,Cliente,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C,Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY F.CodigoC,Cliente " _
                 & "ORDER BY Cliente "
            Nombre_Campo = "Cliente"
       Case "Grupo/Zona"
            sSQL = "SELECT C.Grupo,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C, Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY C.Grupo " _
                 & "ORDER BY C.Grupo "
            Nombre_Campo = "Grupo"
       Case "Vendedor"
            sSQL = "SELECT C.Codigo, C.Cliente, Count(CR.Codigo) As Fact_Proc " _
                 & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
                 & "WHERE CR.Item = '" & NumEmpresa & "' " _
                 & "AND CR.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND CR.Codigo = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = CR.Codigo " _
                 & "GROUP BY C.Codigo,C.Cliente " _
                 & "ORDER BY Cliente "
            Nombre_Campo = "Cliente"
       Case "Ciudad"
            sSQL = "SELECT Ciudad,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C, Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY Ciudad " _
                 & "ORDER BY Ciudad "
            Nombre_Campo = "Ciudad"
       Case "Factura"
            sSQL = "SELECT (TC + ' ' + Serie + ' ' + CAST(Factura As VARCHAR)) As TipoFactura " _
                 & "FROM Facturas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC NOT IN ('C','P') "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "GROUP BY TC, Serie, Factura " _
                 & "ORDER BY TC, Serie, Factura DESC "
            Nombre_Campo = "TipoFactura"
       Case "Forma_Pago"
            sSQL = "SELECT Forma_Pago " _
                 & "FROM Facturas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC NOT IN ('C','P') "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "GROUP BY Forma_Pago " _
                 & "ORDER BY Forma_Pago "
            Nombre_Campo = "Forma_Pago"
       Case "Tipo Documento"
            sSQL = "SELECT TC,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C,Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC NOT IN ('C','P') "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY TC " _
                 & "ORDER BY TC "
            Nombre_Campo = "TC"
       Case "Plan_Afiliado"
            sSQL = "SELECT Plan_Afiliado,Count(Factura) As Fact_Proc " _
                 & "FROM Clientes As C, Facturas As F " _
                 & "WHERE F.Item = '" & NumEmpresa & "' " _
                 & "AND F.Periodo = '" & Periodo_Contable & "' " _
                 & "AND LEN(C.Plan_Afiliado) > 3 " _
                 & "AND TC NOT IN ('C','P') "
            If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND F.Cod_Ejec = '" & CodigoUsuario & "' "
            sSQL = sSQL _
                 & "AND C.Codigo = F.CodigoC " _
                 & "GROUP BY Plan_Afiliado " _
                 & "ORDER BY Plan_Afiliado "
            Nombre_Campo = "Plan_Afiliado"
       Case "Cuenta_No"
            sSQL = "SELECT Cuenta_No " _
                 & "FROM Clientes_Datos_Extras " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "GROUP BY Cuenta_No " _
                 & "ORDER BY Cuenta_No "
            Nombre_Campo = "Cuenta_No"
       Case "Producto"
            sSQL = "SELECT (Codigo_Inv & ' - ' & Producto) As Codigos " _
                 & "FROM Catalogo_Productos " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Codigo_Inv "
            Nombre_Campo = "Codigos"
       Case "DescItem"
            sSQL = "SELECT Desc_Item " _
                 & "FROM Catalogo_Productos " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Desc_Item <> '" & Ninguno & "' " _
                 & "GROUP BY Desc_Item " _
                 & "ORDER BY Desc_Item "
            Nombre_Campo = "Desc_Item"
       Case "Marca"
            sSQL = "SELECT (CodMar & ' - ' & Marca) As NomMarca " _
                 & "FROM Catalogo_Marcas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodMar <> '" & Ninguno & "' " _
                 & "ORDER BY Marca "
            Nombre_Campo = "NomMarca"
       Case Else
            sSQL = "SELECT Codigo,Cliente " _
                 & "FROM Clientes " _
                 & "WHERE Codigo = '-' "
            Nombre_Campo = "Cliente"
     End Select
    'MsgBox Modulo & vbCrLf & sSQL
     SelectDB_Combo DCCliente, AdoCliente, sSQL, Nombre_Campo
     If AdoCliente.Recordset.RecordCount > 0 Then DCCliente.SetFocus
  End If
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
   SetAdoFields "Recibo", TipoProc & ": " & Format$(Factura_No, "0000000")
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
DGQuery.Caption = "CXC CLIENTES POR VENDEDOR"
FA.Fecha_Corte = MBFechaF
Actualizar_Abonos_Facturas_SP FA
RatonReloj
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)
If Actualiza_Buses Then
   If SQL_Server Then
      sSQL = "UPDATE Facturas " _
           & "SET Forma_Pago = MidStrg(DF.Producto,1,10) " _
           & "FROM Facturas AS F,Detalle_Factura AS DF "
   Else
      sSQL = "UPDATE Facturas AS F,Detalle_Factura AS DF " _
           & "SET F.Forma_Pago = MidStrg(DF.Producto,1,10) "
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
   Ejecutar_SQL_SP sSQL
End If

'''    If Tipo = "V" Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Chq_Posf = (SELECT SUM(TA.Abono) " _
'''            & "                FROM Trans_Abonos As TA " _
'''            & "                WHERE TA.Item = '" & NumEmpresa & "' " _
'''            & "                AND TA.Periodo = '" & Periodo_Contable & "' " _
'''            & "                AND TA.Fecha > #" & FechaFin & "# " _
'''            & "                AND TA.TP = Facturas.TC " _
'''            & "                AND TA.Item = Facturas.Item " _
'''            & "                AND TA.Periodo = Facturas.Periodo " _
'''            & "                AND TA.Factura = Facturas.Factura " _
'''            & "                AND TA.Serie = Facturas.Serie " _
'''            & "                AND TA.Autorizacion = Facturas.Autorizacion) " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND T <> 'A' "
'''       Ejecutar_SQL_SP sSQL
'''
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Chq_Posf = 0 " _
'''            & "WHERE Chq_Posf IS NULL " _
'''            & "AND Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' "
'''       Ejecutar_SQL_SP sSQL
'''    End If

If TipoFactura = "" Then TipoFactura = Ninguno
sSQL = "SELECT F.T,F.Razon_Social,"
If SiUnidadEducativa Then sSQL = sSQL & "C.Cliente,"
sSQL = sSQL & "F.Fecha,F.Fecha_V,F.TC,F.Serie,F.Factura,"
If Tipo = "R" Then
   sSQL = sSQL _
        & "F.Con_IVA,F.Sin_IVA,F.SubTotal,F.IVA,(F.Descuento + F.Descuento2) As Total_Descuento,F.Servicio,F.Total_MN," _
        & "F.Total_Abonos,F.Saldo_MN,F.Autorizacion,F.Cta_CxP,F.Total_Ret_Fuente,F.Total_Ret_IVA_B,F.Total_Ret_IVA_S,"
Else
   sSQL = sSQL & "F.Total_MN,F.Abonos_MN,F.Saldo_MN,F.Total_ME,F.Saldo_ME,F.Autorizacion,F.RUC_CI As RUC_CI_SRI,"
End If
If SiUnidadEducativa Then sSQL = sSQL & "C.CI_RUC,"
sSQL = sSQL & "F.Forma_Pago,C.Telefono,C.Celular,C.Ciudad,C.Direccion,C.DireccionT,C.Email,C.Grupo,"
If SQL_Server Then
   sSQL = sSQL & "DATEDIFF(day,'" & BuscarFecha(MBFechaF) & "',F.Fecha_V) As Dias_De_Mora,"
Else
   sSQL = sSQL & "DATEDIFF('d',#" & BuscarFecha(MBFechaF) & "#,F.Fecha_V) As Dias_De_Mora,"
End If
sSQL = sSQL & "A.Nombre_Completo As Ejecutivo,C.Plan_Afiliado As Sectorizacion,A.Cod_Ejec,F.Chq_Posf " _
     & "FROM Facturas As F,Clientes As C,Accesos As A " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & Tipo_De_Consulta() _
     & "AND C.Codigo = F.CodigoC " _
     & "AND A.Codigo = F.Cod_Ejec " _
     & "AND F.TC NOT IN ('C','P') "
If Tipo = "V" Then
   Opcion = 13
   sSQL = sSQL & "ORDER BY A.Nombre_Completo,C.Grupo,"
   If SiUnidadEducativa Then sSQL = sSQL & "C.Cliente,F.Razon_Social," Else sSQL = sSQL & "F.Razon_Social,"
End If
If Tipo = "C" Then
   Opcion = 9
   If SiUnidadEducativa Then sSQL = sSQL & "ORDER BY C.Cliente,F.Razon_Social," Else sSQL = sSQL & "ORDER BY F.Razon_Social,"
End If
If Tipo = "F" Then
   Opcion = 10
   sSQL = sSQL & "ORDER BY "
End If
If Tipo = "R" Then
   Opcion = 19
   If SiUnidadEducativa Then sSQL = sSQL & "ORDER BY C.Cliente,F.Razon_Social," Else sSQL = sSQL & "ORDER BY F.Razon_Social "
End If
sSQL = sSQL & "F.TC,F.Serie,F.Fecha,F.Factura "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True, "CxC Cartera"
DGQuery.Visible = False
Total = 0: Saldo = 0
With AdoQuery.Recordset
 If .RecordCount > 0 Then
     Do While Not .EOF
        Total = Total + .fields("Total_MN")
        Saldo = Saldo + .fields("Saldo_MN")
       .MoveNext
     Loop
 End If
End With
'MsgBox sSQL
DGQuery.Visible = True
TipoDoc = Tipo
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Total - Saldo, "#,##0.00")
LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
Totales_CxC_Abonos
RatonNormal
End Sub

Public Function Tipo_De_Consulta(Optional Opcion_TP As Boolean, Optional Opcion_Email As Boolean, Optional Opcion_DF As Boolean) As String
Dim SQL3X As String
Dim Patron_Busqueda As String

   Patron_Busqueda = DCCliente.Text
   If Patron_Busqueda = "" Then Patron_Busqueda = Ninguno
   Cta_Cobrar = TrimStrg(SinEspaciosIzq(DCCxC.Text))
  'Encabezado de Factura
  'MsgBox Opcion
   SQL3X = ""
   If OpcPend.value Then
      If Opcion > 0 Then
        '& "AND F.T = '" & Pendiente & "' "
         Select Case Opcion
           Case 1: SQL3X = SQL3X _
                         & "AND F.Saldo_Actual <> 0 " _
                         & "AND F.T <> 'A' "
           Case 2: SQL3X = SQL3X & "AND F.T = 'P' "
           Case 9, 10, 13: SQL3X = SQL3X & "AND F.Saldo_MN <> 0 "
           Case 14: SQL3X = SQL3X & "AND F.T <> 'A' "
         End Select
      Else
         SQL3X = SQL3X & "AND F.T = '" & Pendiente & "' "
      End If
   ElseIf OpcCanc.value Then
      SQL3X = SQL3X & "AND F.T = '" & Cancelado & "' "
   ElseIf OpcAnul.value Then
      SQL3X = SQL3X & "AND F.T = '" & Anulado & "' "
   End If
  'MsgBox ListCliente.Text
   Select Case ListCliente.Text
     Case "Codigo"
          SQL3X = SQL3X & "AND C.Codigo = '" & Patron_Busqueda & "' "
     Case "Grupo/Zona"
          SQL3X = SQL3X & "AND C.Grupo = '" & Patron_Busqueda & "' "
     Case "CI_RUC"
          SQL3X = SQL3X & "AND C.CI_RUC = '" & Patron_Busqueda & "' "
     Case "Cliente"
          LongStrg = Len(Patron_Busqueda)
          SQL3X = SQL3X & "AND C.Cliente LIKE '" & Patron_Busqueda & "%' "
     Case "Vendedor"
          LongStrg = Len(Patron_Busqueda)
          SQL3X = SQL3X & "AND A.Nombre_Completo LIKE '" & Patron_Busqueda & "%' "
     Case "Ciudad"
          SQL3X = SQL3X & "AND C.Ciudad = '" & Patron_Busqueda & "' "
     Case "Factura"
          SQL3X = SQL3X & "AND F.Factura = " & Val(Patron_Busqueda) & " "
     Case "Serie"
          SQL3X = SQL3X & "AND F.Serie = '" & Patron_Busqueda & "' "
     Case "Autorizacion"
          SQL3X = SQL3X & "AND F.Autorizacion = '" & Patron_Busqueda & "' "
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

  'Detalle Factura:
   If DescItem <> Ninguno Then SQL3X = SQL3X & "AND MidStrg(F.Codigo,1," & Len(Codigo) & ") = '" & Codigo & "' "
   If Cod_Marca <> Ninguno Then SQL3X = SQL3X & "AND F.CodMarca = '" & Cod_Marca & "' "
   If CheqCxC.value = 1 Then SQL3X = SQL3X & "AND F.Cta_CxP = '" & Cta_Cobrar & "' "
   If CheqIngreso.value = 1 And Opcion_DF = True Then SQL3X = SQL3X & "AND F.Cta_Venta = '" & Cta_Cobrar & "' "
   If CheqAbonos.value = 1 Then
      If Opcion_Email Then SQL3X = SQL3X & "AND TA.Cta = '" & Cta_Cobrar & "' " Else SQL3X = SQL3X & "AND F.Cta = '" & Cta_Cobrar & "' "
   End If
   
  'MsgBox SQL3X
   Tipo_De_Consulta = SQL3X
End Function

Public Sub Totales_CxC_Abonos()
  RatonReloj
  DGQuery.Visible = False
  Total = 0
  Abono = 0
  Saldo = 0
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .fields("T") <> Anulado Then
              Select Case Opcion
                Case 1
                     Total = Total + .fields("Total")
                     Abono = Abono + .fields("Total_Abonos")
                Case 6, 7
                     Total = Total + .fields("Abono")
                Case 9, 10
                     Total = Total + .fields("Total_MN")
                     Saldo = Saldo + .fields("Saldo_MN")
              End Select
          End If
         .MoveNext
       Loop
   End If
  End With
  Select Case Opcion
    Case 1
         Saldo = Total - Abono
         LabelFacturado.Caption = Format$(Total, "#,##0.00")
         LabelAbonado.Caption = Format$(Abono, "#,##0.00")
         LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
    Case 7
         LabelFacturado.Caption = Format$(Total, "#,##0.00")
         LabelAbonado.Caption = Format$(Abono, "#,##0.00")
         LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
    Case 9, 10
         Abono = Total - Saldo
         LabelFacturado.Caption = Format$(Total, "#,##0.00")
         LabelAbonado.Caption = Format$(Abono, "#,##0.00")
         LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
  End Select
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub ToolbarMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  FechaValida MBFechaI
  FechaValida MBFechaF
  
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Mifecha = FechaIni
  FechaTexto = FechaFin
  
  FA.Fecha_Corte = MBFechaF
  FA.Fecha_Desde = MBFechaI
  FA.Fecha_Hasta = MBFechaF
  
  PorCxC = False
  
  If CheqCxC.value = 1 Then PorCxC = True
  If ListCliente.Text = "Todos" Then
     FA.TC = Ninguno
     FA.Serie = Ninguno
     FA.Factura = 0
  End If
  
  Total = 0
  Abono = 0

  DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
  DGQuery.width = MDI_X_Max - 100
  ''FrmAnticipos.Visible = False
  '"CxC_Clientes",
  'MsgBox "Principal: " & Button.key
  Select Case Button.key
    Case "Salir"
         Unload HistorialFacturas
    Case "Imprimir"
         Impresiones
    Case "Facturas"
         Actualizar_Abonos_Facturas_SP FA
         Historico_Facturas
    Case "Protestado"
         Cheques_Protestados
    Case "Por_Buses"
         Por_Buses DCCliente
    Case "Listado_Tarjetas"
         sSQL = "SELECT CM.Tipo_Cta,C.Grupo,CM.Representante,CM.Cedula_R,CM.Telefono_R,C.Cliente,C.Direccion,CM.Cta_Numero,CM.Cod_Banco,CM.Caducidad " _
              & "FROM Clientes As C,Clientes_Matriculas As CM " _
              & "WHERE C.FA <> " & Val(adFalse) & " " _
              & "AND CM.Item = '" & NumEmpresa & "' " _
              & "AND CM.Periodo = '" & Periodo_Contable & "' " _
              & "AND LEN(CM.Tipo_Cta) > 1 " _
              & "AND C.Codigo = CM.Codigo " _
              & "ORDER BY CM.Tipo_Cta, C.Grupo, C.Cliente "
         Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
    Case "CxC_Clientes"
         Actualizar_Abonos_Facturas_SP FA
         Listado_Facturas_Por_Meses True
    Case "Listar_Por_Meses"
         Listado_Facturas_Por_Meses False
    Case "Estado_Cuenta_Cliente"
         If ListCliente.Text = "Todos" Then FA.CodigoC = "Todos"
         Reporte_Cartera_Clientes_SP PrimerDiaMes(MBFechaI), UltimoDiaMes(FechaSistema), FA.CodigoC
         sSQL = "SELECT C.Cliente, RCC.T, RCC.TC, RCC.Serie, RCC.Factura, RCC.Fecha, RCC.Detalle, RCC.Anio, RCC.Mes, RCC.Cargos, RCC.Abonos, RCC.Saldo, RCC.CodigoC, " _
              & "C.Email, C.EmailR, C.Direccion " _
              & "FROM Reporte_Cartera_Clientes As RCC, Clientes As C " _
              & "WHERE RCC.Item = '" & NumEmpresa & "' " _
              & "AND RCC.CodigoU = '" & CodigoUsuario & "' " _
              & "AND RCC.T <> 'A' " _
              & "AND RCC.CodigoC = C.Codigo " _
              & "ORDER BY C.Cliente, RCC.TC, RCC.Serie, RCC.Factura, RCC.Anio, RCC.Mes, RCC.ID "
         Select_Adodc_Grid DGQuery, AdoQuery, sSQL
         DGQuery.Visible = False
         With AdoQuery.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Total = Total + .fields("Cargos")
                 Abono = Abono + .fields("Abonos")
                .MoveNext
              Loop
          End If
         End With
         Opcion = 19
    Case "Listados_Medidor"
         'Listados_Medidor
    Case "Base_Access"
         'Listar_Base_Externa
    Case "Base_MySQL"
         'Leer_Datos_MySQL
         'Listar_Base_MySQL
    Case "Buscar_Malla"
         sSQL = "SELECT Codigo, Cliente " _
              & "FROM Clientes " _
              & "WHERE Codigo = '-' "
         SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
         PresionoEsc = False
         FrmPatronBusqueda.Top = Label5.Top
         FrmPatronBusqueda.Left = Label5.Left
         FrmPatronBusqueda.Visible = True
         ListCliente.Text = "Todos"
         ListCliente.SetFocus
        'Buscar Datos
'''    Case "Buscar_Malla"
'''          Buscar_Datos DGQuery, AdoQuery
         
  End Select
  If Button.key <> "Salir" Then
     Command1.Caption = "&S(" & Opcion & ")"
     DGQuery.Visible = True
     LabelFacturado.Caption = Format$(Total, "#,##0.00")
     LabelAbonado.Caption = Format$(Abono, "#,##0.00")
     LabelSaldo.Caption = Format$(Total - Abono, "#,##0.00")
  End If
  RatonNormal
End Sub

Private Sub ToolbarMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   FechaValida MBFechaI
   FechaValida MBFechaF
   FechaIni = BuscarFecha(MBFechaI)
   FechaFin = BuscarFecha(MBFechaF)
   
   DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
   DGQuery.width = MDI_X_Max - 100
   
   If CheqCxC.value = 1 Then PorCxC = True
   Mifecha = FechaIni
   FechaTexto = FechaFin
   FA.Fecha_Corte = MBFechaF
   FA.Factura = 0
   FA.Fecha_Hasta = MBFechaF
   TMail.Volver_Envial = False
    
   If ListCliente.Text = "Todos" Then
      FA.TC = Ninguno
      FA.Serie = Ninguno
      FA.Factura = 0
   End If
   
  'MsgBox "Secundario: " & ButtonMenu.key
   Select Case ButtonMenu.key
    'Resumen de Ventas
     Case "Resumen_Prod"
          Resumen_Productos
     Case "Resumen_Prod_Meses"
          Resumen_Prod_Meses
     Case "ResumenVentCost"
          Resumen_Ventas_Costos
     Case "Resumen_Ventas_Vendedor"
          Resumen_Ventas_Vendedor
     Case "Ventas_x_Cli"
          Ventas_Cliente
     Case "Ventas_Cli_x_Mes"
          Ventas_Clientes_Por_Meses
     Case "VentasxProductos"
          Ventas_Productos
     Case "Ventas_ResumindasxVendedor"
          Ventas_Resumindas_x_Vendedor
    'Detalle de Abonos
     Case "SMAbonos_Anticipados"
          SMAbonos_Anticipados
     Case "Abonos_Ant"
          Abonos_Anticipados
     Case "Abonos_Erroneos"
          Abonos_Erroneos
     Case "Contra_Cta"
          Contra_Cta_Abonos
    'Ventas CxC por
     Case "Por_Clientes"
          Tipo_Consulta_CxC "C"
     Case "Por_Facturas"
          Tipo_Consulta_CxC "F"
     Case "Resumen_Cartera"
          Tipo_Consulta_CxC "R"
     Case "Por_Vendedor"
          Tipo_Consulta_CxC "V"
     Case "Resumen_Vent_x_Ejec"
          'Resumen_Ventas_x_Ejec
     Case "CxC_Tiempo_Credito"
          CxC_Tiempo_Credito
     Case "Tipo_Pago_Cliente"
          Tipo_Pago_Cliente
    'Reportes por Excel
     Case "Bajar_Excel"
          DGQuery.Visible = False
          Exportar_AdoDB_Excel AdoQuery.Recordset
          DGQuery.Visible = True
     Case "Reporte_Ventas"
          Ventas_x_Excel
     Case "Reporte_Catastro"
          Catastro_Registro_Datos_Clientes
    'Tipo de Abonos
     Case "Tipo_Abonos"
          'Actualizar_Abonos_Facturas FA
          Abonos_Facturas
     Case "Retenciones_NC"
          'Actualizar_Abonos_Facturas FA
          Abonos_Facturas True
    'Envios por mail
     Case "Enviar_FA_Email"
          Enviar_Emails_Facturas_Recibos "FA", Val(TxtDocDesde.Text), Val(TxtDocHasta.Text)
     Case "Enviar_RE_Email"
          Enviar_Emails_Facturas_Recibos "RE", Val(TxtDocDesde.Text), Val(TxtDocHasta.Text)
     Case "Recibos_Anticipados"
          Recibo_Abonos_Anticipados
     Case "Deuda_x_Mail"
          Actualizar_Abonos_Facturas_SP FA
          Historico_Facturas
          Titulo = "Pregunta de Envio de Mails"
          Mensajes = "Esta seguro de querer enviar por mail Cartera de Clientes?"
          If BoxMensaje = vbYes Then Deuda_x_Mail "FA"
   End Select
End Sub

Private Sub TxtDocDesde_GotFocus()
  MarcarTexto TxtDocDesde
End Sub

Private Sub TxtDocDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocDesde_LostFocus()
  TextoValido TxtDocDesde, True, , 0
End Sub

Private Sub TxtDocHasta_GotFocus()
  MarcarTexto TxtDocHasta
End Sub

Private Sub TxtDocHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocHasta_LostFocus()
  TextoValido TxtDocHasta, True, , 0
End Sub

Public Sub Impresiones()
Dim Resultado  As Boolean
    DGQuery.Visible = False
    'MsgBox Opcion
    Select Case Opcion
      Case 1
           MensajeEncabData = "ESTADO DE CUENTA DE CLIENTES"
           SQLMsg1 = "Corte al " & MBFechaF.Text
           Mifecha = MBFechaF.Text
           Imprimir_Saldo_Factura AdoQuery
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
      Case 9, 10, 13
           Codigo4 = Ninguno
           If CheqCxC.value = 1 Then Codigo4 = DCCxC
           If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
           If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
           If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
           If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
           Mifecha = MBFechaF
           If TipoDoc = "C" Then Imprimir_Resumen_Cartera AdoQuery, Codigo4
           If TipoDoc = "F" Then ImprimirCtasCob AdoQuery, sSQL, True
           If TipoDoc = "V" Then Imprimir_Resumen_Cartera_Vendedor AdoQuery
      Case 11
           If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
           If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
           If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
           If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
           Mifecha = MBFechaF.Text
           Imprimir_Pendientes_Facturacion AdoQuery, Opcion, True
      Case 12
           Imprimir_Por_Buses AdoQuery, DCCliente
      Case 15
           MensajeEncabData = "RESUMEN DE COMISIONES POR VENDEDORES"
           SQLMsg1 = "CORTE DEL " & MBFechaI & " AL " & MBFechaF
           Mifecha = MBFechaF
           Orientacion_Pagina = 2
           ImprimirAdo AdoQuery, True, 2, 7, True
      Case 16
           MensajeEncabData = DGQuery.Caption
           SQLMsg1 = "CORTE DEL " & MBFechaI & " AL " & MBFechaF
           Mifecha = MBFechaF
           Orientacion_Pagina = 2
           ImprimirAdo AdoQuery, True, 2, 7, True
      Case 17
           MensajeEncabData = "VENTAS RESUMIDAS POR VENDEDOR"
           SQLMsg1 = "CORTE DEL " & MBFechaI & " AL " & MBFechaF
           Mifecha = MBFechaF
           Orientacion_Pagina = 1
           Imprimir_Ventas_Resumidas_Vendedor AdoQuery, MSChart, 2, 7, True
      Case 18
           MensajeEncabData = "TOTAL CUENTAS POR COBRAR POR TIEMPO DE CREDITO"
           SQLMsg1 = "CORTE DEL " & MBFechaI & " AL " & MBFechaF
           Mifecha = MBFechaF
           Orientacion_Pagina = 2
           Imprimir_Tiempo_Credito AdoQuery, True, 2, 10, True
      Case 19
           RatonReloj
           Resultado = Reporte_Cartera_Clientes_PDF(PrimerDiaMes(MBFechaI), FA.CodigoC, False, True)
           'Presenta__Archivo_PDF HistorialFacturas, RutaDocumentoPDF
    End Select
    HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA"
    DGQuery.Visible = True
End Sub

Public Sub Listado_Facturas_Por_Meses(Por_FA As Boolean)
Dim AnioIni As Integer
Dim AnioFin As Integer
Dim AnioAct As Integer
Dim MesIni As Integer
Dim MesFin As Integer
Dim AnioI As String
Dim AnioF As String
Dim MesS As String
Dim Patron_Busqueda As String
Dim TPEN As Currency
Dim TAnios() As Currency
Dim TMeses(1 To 12) As Currency
Dim VerPEN As Boolean
Dim VerAnios() As Boolean
Dim VerMeses(1 To 12) As Boolean
Dim CxC_VerMeses(1 To 12) As Boolean
Dim SQLSubTotal As String
Dim Por_Fecha As Boolean
Dim Valor_Total As Currency

  Por_Fecha = False
  Mensajes = "Listar Reporte por Fecha?" & vbCrLf
  Titulo = "PREGUNTA DE CONFIRMACION"
  If BoxMensaje = vbYes Then
     Por_Fecha = True
  Else
     MBFechaI = "01/01/2000"
  End If

TiempoSistema = Time

FechaValida MBFechaI
FechaValida MBFechaF

RatonReloj

Valor_Total = 0

MesIni = Month(MBFechaI)
MesFin = Month(MBFechaF)

FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)

AnioI = Format$(Year(MBFechaI), "0000")
AnioF = Format$(Year(MBFechaF), "0000")

'Periodos antes del actual

AnioFin = Year(MBFechaF) - 1
AnioIni = AnioFin - 6
ReDim TAnios(AnioIni To AnioFin) As Currency
ReDim VerAnios(AnioIni To AnioFin) As Boolean
  

  Progreso_Barra.Mensaje_Box = "RESUMEN HISTORICO: Consultando..."
  Progreso_Iniciar
  Progreso_Barra.Valor_Maximo = 100
  
  DGQuery.Visible = False
  DGQuery.Caption = "DEUDAS PENDIENTES DEL PERIODO " & AnioI
  Saldo = 0
  SaldoAnterior = 0
  VerPEN = False
  TPEN = 0
  
  For JE = AnioIni To AnioFin
      TAnios(JE) = 0
      VerAnios(JE) = False
  Next JE
  
  For JE = 1 To 12
      TMeses(JE) = 0
      VerMeses(JE) = False
      CxC_VerMeses(JE) = False
      MesS = MesesLetras(CByte(JE))
      sSQL = "UPDATE Detalle_Factura " _
           & "SET Mes_No = " & JE & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Mes = '" & MesS & "' " _
           & "AND Mes_No = 0 "
      Ejecutar_SQL_SP sSQL
  Next JE
  
  Si_No = False
  NumFacturas = 12
   'Actualizamos Clienes
    sSQL = "UPDATE Clientes " _
         & "SET X = '.' " _
         & "WHERE X <> '.' "
    Ejecutar_SQL_SP sSQL
   
   'Patron de busqueda
    Patron_Busqueda = DCCliente.Text
    If Patron_Busqueda = "" Then Patron_Busqueda = Ninguno
    
   'Insertamos los Clientes que estan en procesos
    sSQL = "DELETE * " _
         & "FROM Saldo_Diarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND TP = 'CXCP' "
    Ejecutar_SQL_SP sSQL
   
   'Actualizamos patrones de busqueda Facturado
    If SQL_Server Then
       sSQL = "UPDATE Clientes " _
            & "SET X = 'A' " _
            & "FROM Clientes As C,Facturas As F "
    Else
       sSQL = "UPDATE Clientes As C,Facturas As F " _
            & "SET C.X = 'A' "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & Buscar_x_Patron(True) _
         & "AND C.Codigo = F.CodigoC "
    Ejecutar_SQL_SP sSQL
   
  'Actualizamos patrones de busqueda Facturado
   If CheqPreFa.value <> 0 Then
      If SQL_Server Then
         sSQL = "UPDATE Clientes " _
              & "SET X = 'A' " _
              & "FROM Clientes As C,Clientes_Facturacion As F "
      Else
         sSQL = "UPDATE Clientes As C,Clientes_Facturacion As F " _
              & "SET C.X = 'A' "
      End If
      sSQL = sSQL _
           & "WHERE F.Item = '" & NumEmpresa & "' " _
           & Buscar_x_Patron() _
           & "AND C.Codigo = F.Codigo "
      Ejecutar_SQL_SP sSQL
   End If
   
   Progreso_Barra.Mensaje_Box = "Actualizando Abonos por meses"
   Progreso_Esperar
  sSQL = "INSERT INTO Saldo_Diarios (CodigoC,Item,CodigoU,TP) " _
       & "SELECT Codigo,'" & NumEmpresa & "' As Item,'" & CodigoUsuario & "' As CodigoUs,'CXCP' As TP " _
       & "FROM Clientes " _
       & "WHERE X = 'A' " _
       & "GROUP BY Codigo "
  Ejecutar_SQL_SP sSQL
  
  Progreso_Barra.Mensaje_Box = "RESUMEN HISTORICO: Determinando Nulos"
  Progreso_Esperar
  Eliminar_Nulos_SP "Saldo_Diarios"

  Progreso_Barra.Mensaje_Box = "RESUMEN HISTORICO: Actualizando Totales"
  Progreso_Esperar
 'Listado de Facturas Emitidas
  sSQL = "SELECT F.CodigoC,C.Cliente,C.Grupo,F.Fecha,F.T,"
  If Por_FA Then
     sSQL = sSQL _
          & "Total_MN,Saldo_MN " _
          & "FROM Facturas As F,Clientes As C "
  Else
     If OpcPend.value Then
        sSQL = sSQL & "Mes_No,(Total-Total_Desc-Total_Desc2+Total_IVA) As Saldo_MN,Mes_No,Ticket "
     Else
        sSQL = sSQL & "Mes_No,(Total-Total_Desc-Total_Desc2+Total_IVA) As Total_MN,Mes_No,Ticket "
     End If
     sSQL = sSQL & "FROM Detalle_Factura As F,Clientes As C "
  End If
  If Por_Fecha Then
     sSQL = sSQL & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = sSQL & "WHERE F.Fecha <= #" & FechaFin & "# "
  End If
  sSQL = sSQL _
      & "AND F.Item = '" & NumEmpresa & "' " _
      & "AND F.Periodo = '" & Periodo_Contable & "' " _
      & Tipo_De_Consulta(, , True) _
      & "AND F.CodigoC = C.Codigo " _
      & "ORDER BY C.Cliente, F.Fecha "
 'MsgBox sSQL
  Select_Adodc AdoHistoria, sSQL
  DGQuery.Visible = False
 'Actualizamos valores de los datos consultados
  K = 0
  With AdoHistoria.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount
      .MoveFirst
       CodigoCli = .fields("CodigoC")
       Do While Not .EOF
          If CodigoCli <> .fields("CodigoC") Then
             SQLSubTotal = ""
             Total = 0
             For JE = AnioIni To AnioFin
                 Total = Total + TAnios(JE)
                 SQLSubTotal = SQLSubTotal & "P_" & CStr(JE) & " = " & TAnios(JE) & ", "
                 TAnios(JE) = 0
             Next JE
             For JE = 1 To 12
                 Total = Total + TMeses(JE)
                 SQLSubTotal = SQLSubTotal & MesesLetras(CByte(JE)) & " = " & TMeses(JE) & ", "
                 TMeses(JE) = 0
             Next JE
             SQLSubTotal = SQLSubTotal & "PEN = " & TPEN & " "
             Total = Total + TPEN
             TPEN = 0
             If Total <> 0 Then
                sSQL = "UPDATE Saldo_Diarios " _
                     & "SET " & SQLSubTotal _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND CodigoU = '" & CodigoUsuario & "' " _
                     & "AND CodigoC = '" & CodigoCli & "' " _
                     & "AND TP = 'CXCP' "
                Ejecutar_SQL_SP sSQL
             End If
             CodigoCli = .fields("CodigoC")
          End If
          K = K + 1
          Progreso_Barra.Mensaje_Box = "RESUMEN HISTORICO: Actualizando Valores"
          Progreso_Esperar

          If OpcPend.value Then Total = .fields("Saldo_MN") Else Total = .fields("Total_MN")
          Total = Redondear(Total, 2)
          Valor_Total = Valor_Total + Total
         'Seteamos años y meses de actualizacion
          If Por_FA Then
             MesIni = Month(.fields("Fecha"))
             AnioAct = Year(.fields("Fecha"))
          Else
             MesIni = .fields("Mes_No")
             AnioAct = Val(.fields("Ticket"))
          End If
          If MesIni = 0 Then MesIni = Month(.fields("Fecha"))
          
         'Actualizamo los valores en los campos respectivos
          If Total <> 0 Then
             If AnioIni <= AnioAct And AnioAct <= AnioFin Then
                TAnios(AnioAct) = TAnios(AnioAct) + Total
                 VerAnios(AnioAct) = True
             ElseIf AnioAct < AnioIni Then
                TPEN = TPEN + Total
                VerPEN = True
             Else
                TMeses(MesIni) = TMeses(MesIni) + Total
                VerMeses(MesIni) = True
             End If
          End If
         .MoveNext
       Loop
       SQLSubTotal = ""
       Total = 0
       
       For JE = AnioIni To AnioFin
           Total = Total + TAnios(JE)
           SQLSubTotal = SQLSubTotal & "P_" & CStr(JE) & " = " & TAnios(JE) & ", "
           TAnios(JE) = 0
       Next JE
       For JE = 1 To 12
           Total = Total + TMeses(JE)
           SQLSubTotal = SQLSubTotal & MesesLetras(CByte(JE)) & " = " & TMeses(JE) & ", "
           TMeses(JE) = 0
       Next JE
       SQLSubTotal = SQLSubTotal & "PEN = " & TPEN & " "
       Total = Total + TPEN
       TPEN = 0
       If Total <> 0 Then
          sSQL = "UPDATE Saldo_Diarios " _
               & "SET " & SQLSubTotal _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "AND CodigoC = '" & CodigoCli & "' " _
               & "AND TP = 'CXCP' "
          Ejecutar_SQL_SP sSQL
       End If
   End If
  End With
  
 '=============================================================================================================
 'Listado CxC PreFacturable
  If CheqPreFa.value <> 0 Then
     K = 0
     sSQL = "SELECT C.Cliente,F.Codigo,C.Grupo,F.Fecha,SUM(F.Valor-F.Descuento) As Total_MN " _
          & "FROM Clientes_Facturacion As F,Clientes As C " _
          & "WHERE F.Item = '" & NumEmpresa & "' " _
          & "AND F.Fecha <= #" & FechaFin & "# " _
          & Buscar_x_Patron(True) _
          & "AND F.Codigo = C.Codigo " _
          & "GROUP BY C.Cliente,F.Codigo,C.Grupo,F.Fecha " _
          & "ORDER BY C.Cliente,F.Codigo,C.Grupo,F.Fecha "
     Select_Adodc AdoHistoria, sSQL
     DGQuery.Visible = False

    'Actualizamos valores de los datos consultados Pre-Factura
     With AdoHistoria.Recordset
      If .RecordCount > 0 Then
          Progreso_Barra.Valor_Maximo = .RecordCount
         .MoveFirst
          CodigoCli = .fields("Codigo")
          Do While Not .EOF
             If CodigoCli <> .fields("Codigo") Then
                Total = 0
                SQLSubTotal = ""
                For JE = 1 To 12
                    Total = Total + TMeses(JE)
                    SQLSubTotal = SQLSubTotal & "CxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3) & " = " & TMeses(JE) & ", "
                    TMeses(JE) = 0
                Next JE
                SQLSubTotal = MidStrg(SQLSubTotal, 1, Len(SQLSubTotal) - 2)
                'MsgBox SQLSubTotal
                If Total <> 0 Then
                   sSQL = "UPDATE Saldo_Diarios " _
                        & "SET " & SQLSubTotal _
                        & "WHERE Item = '" & NumEmpresa & "' " _
                        & "AND CodigoU = '" & CodigoUsuario & "' " _
                        & "AND CodigoC = '" & CodigoCli & "' " _
                        & "AND TP = 'CXCP' "
                   Ejecutar_SQL_SP sSQL
                End If
                CodigoCli = .fields("Codigo")
             End If
             K = K + 1
             Progreso_Barra.Mensaje_Box = "RESUMEN HISTORICO: Actualizando Valores Pre-CxC"
             Progreso_Esperar
             CodigoCli = .fields("Codigo")
             Total = Redondear(.fields("Total_MN"), 2)
             Valor_Total = Valor_Total + Total
            'Seteamos años y meses de actualizacion
             MesIni = Month(.fields("Fecha"))
             AnioAct = Year(.fields("Fecha"))
             
            'Actualizamo los valores en los campos respectivos
             If Total <> 0 And AnioAct >= AnioIni Then
                TMeses(MesIni) = TMeses(MesIni) + Total
             End If
            .MoveNext
          Loop
          SQLSubTotal = ""
          For JE = 1 To 12
              Total = Total + TMeses(JE)
              SQLSubTotal = SQLSubTotal & "CxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3) & " = " & TMeses(JE) & ", "
              TMeses(JE) = 0
          Next JE
          SQLSubTotal = MidStrg(SQLSubTotal, 1, Len(SQLSubTotal) - 2)
          If Total <> 0 Then
             sSQL = "UPDATE Saldo_Diarios " _
                  & "SET " & SQLSubTotal _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND CodigoC = '" & CodigoCli & "' " _
                  & "AND TP = 'CXCP' "
             Ejecutar_SQL_SP sSQL
          End If
      End If
     End With
  End If
    
 'Totalizamos los meses y años pendientes
  SQLSubTotal = ""
  For JE = AnioIni To AnioFin
      SQLSubTotal = SQLSubTotal & "P_" & CStr(JE) & " + "
  Next JE
  For JE = 1 To 12
      SQLSubTotal = SQLSubTotal & MesesLetras(CByte(JE)) & " + "
  Next JE
  For JE = 1 To 12
      SQLSubTotal = SQLSubTotal & "CxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3) & " + "
  Next JE
  SQLSubTotal = SQLSubTotal & "PEN "
    
  sSQL = "UPDATE Saldo_Diarios " _
       & "SET Total = " & SQLSubTotal _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' "
  Ejecutar_SQL_SP sSQL
      
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Total = 0 " _
       & "AND TP = 'CXCP' "
  Ejecutar_SQL_SP sSQL
       
  SQLSubTotal = ""
  For JE = AnioIni To AnioFin
      SQLSubTotal = SQLSubTotal & "SUM(P_" & CStr(JE) & ") As TP_" & CStr(JE) & ","
  Next JE
  For JE = 1 To 12
      SQLSubTotal = SQLSubTotal & "SUM(" & MesesLetras(CByte(JE)) & ") As " & "T" & MesesLetras(CByte(JE)) & ","
  Next JE
  For JE = 1 To 12
      SQLSubTotal = SQLSubTotal & "SUM(CxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3) & ") As " & "TCxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3) & ","
  Next JE
  SQLSubTotal = SQLSubTotal & "SUM(PEN) As TPEN "
      
  sSQL = "SELECT TP," & SQLSubTotal _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'CXCP' " _
       & "GROUP BY TP "
  Select_Adodc AdoQuery1, sSQL
  With AdoQuery1.Recordset
   If .RecordCount > 0 Then
       If .fields("TPEN") <> 0 Then VetPEN = True
       For JE = AnioIni To AnioFin
           If .fields("TP_" & CStr(JE)) <> 0 Then VerAnios(JE) = True
       Next JE
       For JE = 1 To 12
           If .fields("T" & MesesLetras(CByte(JE))) <> 0 Then VerMeses(JE) = True
       Next JE
       For JE = 1 To 12
           If .fields("TCxC_" & MidStrg(MesesLetras(CByte(JE)), 1, 3)) <> 0 Then CxC_VerMeses(JE) = True
       Next JE
   End If
  End With
 'Listado de Rubros de pensiones por meses
  sSQL = "SELECT C.Cliente,"
  If VetPEN Then sSQL = sSQL & "PEN,"
  For IE = AnioIni To AnioFin
      If VerAnios(IE) Then sSQL = sSQL & "P_" & CStr(IE) & ","
  Next IE
  For IE = 1 To 12
      If VerMeses(IE) Then sSQL = sSQL & MesesLetras(CByte(IE)) & ","
  Next IE
  For IE = 1 To 12
      If CxC_VerMeses(IE) Then sSQL = sSQL & "CxC_" & MidStrg(MesesLetras(CByte(IE)), 1, 3) & ","
  Next IE
  sSQL = sSQL & "SD.Total,C.Direccion,C.Grupo,C.Plan_Afiliado As BUS_No,SD.Ln As No_ " _
       & "FROM Saldo_Diarios As SD,Clientes As C " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.TP = 'CXCP' " _
       & "AND SD.CodigoC = C.Codigo " _
       & "ORDER BY C.Grupo,C.Cliente,SD.TC "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  
  'MsgBox sSQL
  LabelSaldo.Caption = Format$(Valor_Total, "#,##0.00")
  DGQuery.Visible = True
  Minutos = Time
  HistorialFacturas.Caption = "RESUMEN HISTORICO DE FACTURAS/NOTAS DE VENTA "
  If Por_FA Then
     HistorialFacturas.Caption = HistorialFacturas.Caption & " EMITIDAS "
  Else
     HistorialFacturas.Caption = HistorialFacturas.Caption & " POR MESES"
  End If
  HistorialFacturas.Caption = HistorialFacturas.Caption & Format$(Minutos - TiempoSistema, "HH:MM:SS")
  Opcion = 11
  RatonNormal
End Sub

Public Function Buscar_x_Patron(Optional PorFactura As Boolean, Optional Opcion_TP As Boolean) As String
Dim SQL3X
Dim Patron_Busqueda As String
   SQL3X = ""
   Patron_Busqueda = DCCliente.Text
   If Patron_Busqueda = "" Then Patron_Busqueda = Ninguno
   Select Case ListCliente.Text
     Case "Factura"
          If PorFactura Then SQL3X = SQL3X & "AND F.Factura = " & Val(Patron_Busqueda) & " "
     Case "Forma_Pago"
          If PorFactura Then SQL3X = SQL3X & "AND F.Forma_Pago = '" & Patron_Busqueda & "' "
     Case "Tipo Documento"
          If PorFactura Then
             If Opcion_TP Then
                SQL3X = SQL3X & "AND F.TP = '" & Patron_Busqueda & "' "
             Else
                SQL3X = SQL3X & "AND F.TC = '" & Patron_Busqueda & "' "
             End If
             TipoFactura = Patron_Busqueda
          End If
     Case "Codigo"
          SQL3X = SQL3X & "AND C.Codigo = '" & Patron_Busqueda & "' "
     Case "CI_RUC"
          SQL3X = SQL3X & "AND C.CI_RUC = '" & Patron_Busqueda & "' "
     Case "Cliente"
          SQL3X = SQL3X & "AND UCaseStrg(MidStrg(C.Cliente,1," & Len(Patron_Busqueda) & ")) = '" & Patron_Busqueda & "' "
     Case "Ciudad"
          SQL3X = SQL3X & "AND C.Ciudad = '" & Patron_Busqueda & "' "
     Case "Grupo"
          SQL3X = SQL3X & "AND C.Grupo = '" & Patron_Busqueda & "' "
     Case "Plan_Afiliado"
          SQL3X = SQL3X & "AND C.Plan_Afiliado = '" & Patron_Busqueda & "' "
     Case Else
          SQL3X = SQL3X & "AND C.Codigo <> ' ' "
   End Select
   Buscar_x_Patron = SQL3X
End Function

Public Sub Ventas_x_Excel()
   DGQuery.Visible = False
   sSQL = "SELECT T,TC,Fecha,'" & Empresa & "' As Razon_Social,'" & RUC & "' As RUC,Serie,Autorizacion," _
        & "Factura,Con_IVA,Sin_IVA,SubTotal,IVA,Total_MN,'999999' As Serie_R,'0' As Secuencial_R," _
        & "'" & RUC & "' As Autorizacion_R,'312' As Cod_Ret,'0' As Total_Retenido, '5' As Cta_Gasto " _
        & "FROM Facturas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
        & "AND T <> 'A' " _
        & "ORDER BY TC,Serie,Factura "
   Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
   GenerarDataTexto HistorialFacturas, AdoQuery
   DGQuery.Visible = True
End Sub

Public Sub Catastro_Registro_Datos_Clientes()
Dim AdoCatastro As ADODB.Recordset
Dim NFila As Integer
Dim NCelda As Integer
Dim RutaGeneraFile As String
Dim Dias_Morosidad As Integer
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
  RatonReloj
  DGQuery.Caption = "HISTORIAL DE FACTURAS"
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.Workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  Contador = 0
  RutaGeneraFile = RutaSysBases & "\Excel\Catastro Cliente del "
  If MBFechaI = MBFechaF Then
     RutaGeneraFile = RutaGeneraFile & Replace(MBFechaF, "/", "-") & ".xls"
  Else
     RutaGeneraFile = RutaGeneraFile & Replace(MBFechaI, "/", "-") & " al " & Replace(MBFechaF, "/", "-") & ".xls"
  End If
  If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
  FA.Fecha_Corte = MBFechaF
  Actualizar_Abonos_Facturas_SP FA
  RatonReloj
   oSheet.Range("A1").value = "CATASTRO REGISTRO DE DATOS CREDITICIOS"
   oSheet.Range("A2").value = "DATOS DEL CLIENTE"
   oSheet.Range("G2").value = "CAMPOS SOCIODEMOGRÁFICOS"
   oSheet.Range("M2").value = "DATOS DE ENDEUDAMIENTO"
   oSheet.Range("A3").value = "Código de la Entidad"
   oSheet.Range("B3").value = "Fecha de Datos"
   oSheet.Range("C3").value = "Tipo de Identificación del Sujeto"
   oSheet.Range("D3").value = "Identificación del Sujeto"
   oSheet.Range("E3").value = "Nombres y Apellidos del Sujeto"
   oSheet.Range("F3").value = "Clase del Sujeto"
   oSheet.Range("G3").value = "Provincia"
   oSheet.Range("H3").value = "Cantón"
   oSheet.Range("I3").value = "Parroquia"
   oSheet.Range("J3").value = "Sexo"
   oSheet.Range("K3").value = "Estado Civil"
   oSheet.Range("L3").value = "Origen de Ingresos"
   oSheet.Range("M3").value = "Número de Operación"
   oSheet.Range("N3").value = "Valor de la Operación"
   oSheet.Range("O3").value = "Saldo de Operación"
   oSheet.Range("P3").value = "Fecha de Concesión"
   oSheet.Range("Q3").value = "Fecha de Vencimiento"
   oSheet.Range("R3").value = "Fecha que es Exigible"
   oSheet.Range("S3").value = "Plazo Operación (días)"
   oSheet.Range("T3").value = "Periodicidad de Pago (días)"
   oSheet.Range("U3").value = "Días de Morosidad"
   oSheet.Range("V3").value = "Monto de Morosidad"
   oSheet.Range("W3").value = "Monto de Interés en Mora"
   oSheet.Range("X3").value = "Valor por vencer de 1 a 30 días"
   oSheet.Range("Y3").value = "Valor por vencer de 31 a 90 días"
   oSheet.Range("Z3").value = "Valor por vencer de 91 a 180 días"
   oSheet.Range("AA3").value = "Valor por vencer de 181 a 360 días"
   oSheet.Range("AB3").value = "Valor por vencer de mas 360 días"
   oSheet.Range("AC3").value = "Valor vencido 1 A 30 dias"
   oSheet.Range("AD3").value = "Valor vencido de 31 a 90 días"
   oSheet.Range("AE3").value = "Valor vencido de 91 a 180  días"
   oSheet.Range("AF3").value = "Valor vencido de 181 a 360 días"
   oSheet.Range("AG3").value = "Valor vencido de más de 360 días"
   oSheet.Range("AH3").value = "Valor en Demanda Judicial"
   oSheet.Range("AI3").value = "Cartera Castigada"
   oSheet.Range("AJ3").value = "Cuota del Crédito"
   oSheet.Range("AK3").value = "Fecha de Cancelación"
   oSheet.Range("AL3").value = "Forma de Cancelación"
  RatonReloj
  sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Est_Civil,C.Sexo,C.Ciudad,C.Prov,C.Pais,F.T,F.Serie,F.Factura," _
       & "F.Fecha,F.Fecha_C,F.Fecha_V,F.Total_MN As Total,F.Total_Efectivo,F.Total_Banco,F.Total_Ret_Fuente," _
       & "F.Total_Ret_IVA_B,F.Total_Ret_IVA_S,F.Otros_Abonos,F.Total_Abonos,F.Saldo_Actual,F.CodigoC,F.TC,F.Autorizacion " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & Tipo_De_Consulta() _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY C.Cliente,F.Factura,F.Fecha "
  Select_AdoDB AdoCatastro, sSQL
  With AdoCatastro
   If .RecordCount > 0 Then
       RatonReloj
       NFila = 4
       Do While Not .EOF
          Dias_Morosidad = 0
          If .fields("T") <> "C" Then Dias_Morosidad = CFechaLong(FechaSistema) - CFechaLong(.fields("Fecha_V"))
          
          oSheet.Range("A" & NFila).value = "SP10101"
          oSheet.Range("B" & NFila).value = MBFechaF
          oSheet.Range("C" & NFila).value = .fields("TD")
          oSheet.Range("D" & NFila).value = .fields("CI_RUC")
          oSheet.Range("E" & NFila).value = .fields("Cliente")
          If MidStrg(.fields("CI_RUC"), 3, 1) = "9" Then
             oSheet.Range("F" & NFila).value = "J"
          Else
             oSheet.Range("F" & NFila).value = "N"
          End If
          oSheet.Range("G" & NFila).value = .fields("Prov")
          oSheet.Range("H" & NFila).value = "C" & .fields("Ciudad")
          oSheet.Range("I" & NFila).value = "P" & .fields("Ciudad")
          oSheet.Range("J" & NFila).value = .fields("Sexo")
          oSheet.Range("K" & NFila).value = .fields("Est_Civil")
          oSheet.Range("L" & NFila).value = "I"
          oSheet.Range("M" & NFila).value = "'" & .fields("Serie") & Format$(.fields("Factura"), "000000000")
          oSheet.Range("N" & NFila).value = Format$(.fields("Total"), "#,##0.00")
          oSheet.Range("O" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00")
          oSheet.Range("P" & NFila).value = Format$(.fields("Fecha"), FormatoFechas)
          oSheet.Range("Q" & NFila).value = Format$(.fields("Fecha_V"), FormatoFechas)
          oSheet.Range("R" & NFila).value = Format$(.fields("Fecha_V"), FormatoFechas)
          oSheet.Range("S" & NFila).value = " "  'Plazo Operación (días)
          oSheet.Range("T" & NFila).value = " "  'Periodicidad de Pago (días)
          If Dias_Morosidad > 0 Then oSheet.Range("U" & NFila).value = Dias_Morosidad 'Días de Morosidad
          If .fields("T") <> "C" Then
              oSheet.Range("V" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00")  'Monto de Morosidad
          End If
          oSheet.Range("W" & NFila).value = " "  'Monto de Interés en Mora
          
          If -30 <= Dias_Morosidad And Dias_Morosidad <= -1 Then oSheet.Range("X" & NFila).value = Format$(.fields("Total"), "#,##0.00") 'Valor por vencer de 1 a 30 días
          If -90 <= Dias_Morosidad And Dias_Morosidad <= -31 Then oSheet.Range("Y" & NFila).value = Format$(.fields("Total"), "#,##0.00")  'Valor por vencer de 31 a 90 días
          If -180 <= Dias_Morosidad And Dias_Morosidad <= -91 Then oSheet.Range("Z" & NFila).value = Format$(.fields("Total"), "#,##0.00")  'Valor por vencer de 91 a 180 días
          If -360 <= Dias_Morosidad And Dias_Morosidad <= -181 Then oSheet.Range("AA" & NFila).value = Format$(.fields("Total"), "#,##0.00") 'Valor por vencer de 181 a 360 días
          If -360 <= Dias_Morosidad Then oSheet.Range("AB" & NFila).value = Format$(.fields("Total"), "#,##0.00")  'Valor por vencer de mas 360 días
          
          If 1 <= Dias_Morosidad And Dias_Morosidad <= 30 Then oSheet.Range("AC" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00") 'Valor vencido 1 A 30 dias
          If 31 <= Dias_Morosidad And Dias_Morosidad <= 90 Then oSheet.Range("AD" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00") 'Valor vencido de 31 a 90 días
          If 91 <= Dias_Morosidad And Dias_Morosidad <= 180 Then oSheet.Range("AE" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00") 'Valor vencido de 91 a 180  días
          If 181 <= Dias_Morosidad And Dias_Morosidad <= 360 Then oSheet.Range("AF" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00") 'Valor vencido de 181 a 360 días
          If Dias_Morosidad > 360 Then oSheet.Range("AG" & NFila).value = Format$(.fields("Saldo_Actual"), "#,##0.00") 'Valor vencido de más de 360 días
          oSheet.Range("AH" & NFila).value = " " 'Valor en Demanda Judicial
          oSheet.Range("AI" & NFila).value = " " 'Cartera Castigada
          oSheet.Range("AJ" & NFila).value = " " 'Cuota del Crédito
          If .fields("T") = "C" Then
              oSheet.Range("AK" & NFila).value = Format$(.fields("Fecha_V"), FormatoFechas)
              If .fields("Total_Efectivo") > 0 Then
                  oSheet.Range("AL" & NFila).value = "E"
              ElseIf .fields("Total_Banco") > 0 Then
                  oSheet.Range("AL" & NFila).value = "C"
              Else
                  oSheet.Range("AL" & NFila).value = "T"
              End If
          End If
          NFila = NFila + 1
         .MoveNext
       Loop
''      'Bloqueamos las celdas que no se puden cambiar
''       For NCelda = 66 To 76
''           oSheet.Columns(Chr(NCelda)).ColumnWidth = 8
''       Next NCelda
''       For NCelda = 1 To 14
''           With oSheet.Cells(1, NCelda)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 41    ' Color fondo = azul '41
''               .Font.Size = 9             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''           With oSheet.Cells(NFila + 1, NCelda)  ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 41    ' Color fondo = azul '41
''               .Font.Size = 9             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''       Next NCelda
''       For NCelda = 2 To NFila
''           With oSheet.Cells(NCelda, 1)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 42    ' Color fondo = azul '41
''               .Font.Size = 10             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''           With oSheet.Cells(NCelda, 10)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 42    ' Color fondo = azul '41
''               .Font.Size = 8             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''           With oSheet.Cells(NCelda, 11)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 42    ' Color fondo = azul '41
''               .Font.Size = 8             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''           With oSheet.Cells(NCelda, 12)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 42    ' Color fondo = azul '41
''               .Font.Size = 8             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''           With oSheet.Cells(NCelda, 13)    ' seleccionamos la 1ª celda
''               .Interior.ColorIndex = 42    ' Color fondo = azul '41
''               .Font.Size = 8             ' tamaño de letra
''               .Font.Bold = True           ' Fuente en negrita
''               .Font.ColorIndex = 2       ' Color fuente = blanco
''           End With
''       Next NCelda
''       oSheet.Unprotect "DiskCoverEducativo"
''       oSheet.Range("B2:B" & CStr(NFila)).Locked = False
''       oSheet.Range("C2:C" & CStr(NFila)).Locked = False
''       oSheet.Range("D2:D" & CStr(NFila)).Locked = False
''       oSheet.Range("E2:E" & CStr(NFila)).Locked = False
''       oSheet.Range("F2:F" & CStr(NFila)).Locked = False
''       oSheet.Range("G2:F" & CStr(NFila)).Locked = False
''       oSheet.Range("H2:F" & CStr(NFila)).Locked = False
''       oSheet.Range("I2:F" & CStr(NFila)).Locked = False
''       oSheet.Protect "DiskCoverEducativo"

      'Salvamos el reporte de excel
       oBook.SaveAs RutaGeneraFile
       oExcel.Quit
       RatonNormal
       MsgBox "ARCHIVO GENERADO EN:" & vbCrLf & RutaGeneraFile
   Else
       RatonNormal
       MsgBox "ARCHIVO NO GENERADO"
   End If
  End With
  AdoCatastro.Close
End Sub

Public Sub Enviar_Emails_Facturas_Recibos(TipoEnvio As String, DocDesde As Long, DocHasta As Long)
Dim MesNo As Byte
Dim Cta_Aux_Mail As String
Dim Detalle_Abono As String
''Dim Archivo_PDF As Boolean

    RatonReloj
    TMail.ListaMail = 255
    Cta_Aux_Mail = Ninguno
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    If CheqAbonos.value <> 0 Then
       Cta_Aux_Mail = SinEspaciosIzq(DCCxC.Text)
       
       sSQL = "UPDATE Facturas " _
            & "SET X = 'X' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
       Ejecutar_SQL_SP sSQL
       
       sSQL = "UPDATE Facturas " _
            & "SET X = '.' " _
            & "FROM Facturas As F, Trans_Abonos As TA " _
            & "WHERE F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
            & "AND TA.Cta = '" & Cta_Aux_Mail & "' " _
            & "AND F.Item = TA.Item " _
            & "AND F.Periodo = TA.Periodo " _
            & "AND F.TC = TA.TP " _
            & "AND F.Serie = TA.Serie " _
            & "AND F.Factura = TA.Factura "
       Ejecutar_SQL_SP sSQL
       CheqAbonos.value = 0
    End If
    If TipoEnvio = "FA" Then
        sSQL = "SELECT C.Cliente, F.CodigoC, F.Estado_SRI, F.TC, F.Fecha, F.Fecha_V, F.Serie, F.Factura, F.Hora_Aut, F.Fecha_Aut, F.Autorizacion, " _
             & "F.Saldo_MN, C.Email, C.Email2, C.EmailR, C.CI_RUC, C.Grupo " _
             & "FROM Facturas As F, Clientes As C " _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND F.TC IN ('FA','NV') " _
             & Tipo_De_Consulta(, True) & " "
        If DocDesde > 0 And DocHasta > 0 And DocDesde <= DocHasta Then sSQL = sSQL & "AND F.Factura BETWEEN " & DocDesde & " and " & DocHasta & " "
        If Cta_Aux_Mail <> Ninguno Then sSQL = sSQL & "AND F.X = '.' "
    Else
        Opcion = 14
        sSQL = "SELECT F.TP, F.Fecha, C.Cliente, F.Serie, F.Factura, F.Recibo_No, F.Banco, F.Cheque, F.Abono, F.Mes, F.Comprobante, F.Autorizacion, " _
             & "F.Serie_NC, Secuencial_NC, F.Autorizacion_NC, C.Representante As Razon_Social, C.Grupo, C.Direccion As Ubicacion, F.Cta, F.Cta_CxP, " _
             & "F.CodigoC, F.Hora_Aut, F.Fecha_Aut, C.CI_RUC, C.Email, C.Email2, C.EmailR " _
             & "FROM Trans_Abonos As F, Clientes C " _
             & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & Tipo_De_Consulta(True) & " "
        If DocDesde > 0 And DocHasta > 0 And DocDesde <= DocHasta Then sSQL = sSQL & "AND F.Factura BETWEEN " & DocDesde & " and " & DocHasta & " "
        If Cta_Aux_Mail <> Ninguno Then sSQL = sSQL & "AND F.Cta = '" & Cta_Aux_Mail & "' "
    End If
    sSQL = sSQL _
         & "AND F.CodigoC = C.Codigo " _
         & "ORDER BY F.Fecha, C.Cliente, F.Serie, F.Factura "
    Select_Adodc AdoQuery, sSQL
    RatonReloj
    TMail.TipoDeEnvio = "CO"
    TMail.ListaMail = 255
    
    With AdoQuery.Recordset
    'MsgBox "Total Registros: " & .RecordCount
     If .RecordCount > 0 Then
         Titulo = "Pregunta de Envio de Mails"
         Mensajes = "Esta seguro de querer enviar por mail los documentos?"
         If BoxMensaje = vbYes Then
            Do While Not .EOF
               FA.CodigoC = .fields("CodigoC")
               FA.ClaveAcceso = .fields("Autorizacion")
               If TipoEnvio = "FA" Then
                  FA.Estado_SRI = .fields("Estado_SRI")
                  FA.Fecha_V = .fields("Fecha_V")
                  FA.TC = .fields("TC")
               Else
                  FA.Estado_SRI = "OK"
                  FA.TC = .fields("TP")
                  FA.DireccionC = .fields("Ubicacion")
               End If
               FA.Fecha = .fields("Fecha")
               FA.Serie = .fields("Serie")
               FA.CI_RUC = .fields("CI_RUC")
               FA.Factura = .fields("Factura")
               FA.Autorizacion = .fields("Autorizacion")
               FA.Hora_FA = .fields("Hora_Aut")
               FA.Fecha_Aut = .fields("Fecha_Aut")
               FA.EmailC = .fields("Email")
               FA.EmailR = .fields("Email2")
               FA.Cliente = .fields("Cliente")
               FA.Comercial = .fields("Cliente")
               
               SRI_Autorizacion.Hora_Autorizacion = .fields("Hora_Aut")
               
               If TipoEnvio = "FA" Then
                  FA.Recibo_No = "000000000"
                  SRI_Enviar_Mails FA, SRI_Autorizacion, "FA"
               Else
                  TMail.Adjunto = ""
                  ValorTotal = .fields("Abono")
                  TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_recibo.html")
                  FA.Recibo_No = Format(FA.Fecha, "yyyymm") & .fields("Recibo_No")
                  TMail.Credito_No = "" ' "R" & FA.Recibo_No
                  TMail.Asunto = FA.Cliente & ", " & "Documento No " & FA.Serie & "-" & Format$(FA.Factura, "000000000")
                  Detalle_Abono = ""
                  If Len(.fields("Banco")) > 1 Then Detalle_Abono = Detalle_Abono & .fields("Banco") & ", "
                  If Len(.fields("Cheque")) > 1 Then Detalle_Abono = Detalle_Abono & .fields("Cheque") & ", "
                  html_Detalle_adicional = "<tr>" _
                                         & "<td>vSerie_Documento_FA</td>" _
                                         & "<td>" & FA.Fecha & "</td>" _
                                         & "<td>" & Detalle_Abono & "</td>" _
                                         & "<td class='row text-right'>" & Format(.fields("Abono"), "#,##0.00") & "</td>" _
                                         & "</tr>"
                  
                  If FA.Cliente <> Ninguno And FA.Cliente <> FA.Razon_Social Then
                     html_Informacion_adicional = ""
                     If Len(FA.Cliente) > 1 Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Beneficiario:</strong> " & FA.Cliente & "</a>"
                     If Len(FA.CI_RUC) > 1 Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Codigo:</strong> " & FA.CI_RUC & "</a>"
                     If Len(FA.Curso) > 1 Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Ubicacion:</strong> " & FA.Grupo & " - " & FA.Curso & "</a>"
                     If Len(FA.DireccionC) > 1 Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Direccion:</strong> " & FA.DireccionC & "</a>"
                     If Len(FA.TelefonoC) > 1 Then Insertar_Campo_XML html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Telefono:</strong> " & FA.TelefonoC & "</a>"
                     If EsUnEmail(FA.EmailC) Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Email:</strong> " & FA.EmailC & "</a>"
                     If EsUnEmail(FA.EmailR) And InStr(FA.EmailC, FA.EmailR) = 0 Then html_Informacion_adicional = html_Informacion_adicional & "<a class='col-6'><strong>Email 2:</strong> " & FA.EmailR & "</a>"
                     If html_Informacion_adicional <> "" Then html_Informacion_adicional = "<strong>INFORMACION ADICIONAL:</strong><br>" & html_Informacion_adicional
                  Else
                     html_Informacion_adicional = ""
                  End If
                 'Enviamos lista de mails
                  TMail.para = ""
                  Insertar_Mail TMail.para, .fields("Email")
                  Insertar_Mail TMail.para, .fields("Email2")
                  Insertar_Mail TMail.para, .fields("EmailR")
                  If Email_CE_Copia Then Insertar_Mail TMail.para, EmailProcesos
                  FEnviarCorreos.Show 1
               End If
              'MsgBox ">>>>>>>"
              .MoveNext
            Loop
            RatonNormal
            MsgBox "Proceso terminado exitosamente"
         End If
     Else
         RatonNormal
         If TipoEnvio = "FA" Then
            MsgBox "No hay Facturas Pendientes para enviar"
         Else
            MsgBox "No hay Recibos Pendientes para enviar"
         End If
     End If
    End With
End Sub

Public Sub Deuda_x_Mail(TipoEnvio As String)
Dim Si_Envia As Boolean
Dim posPuntoComa As Integer

  DGQuery.Visible = False

  TMail.ListaMail = 0
  TMail.ListaError = ""
  TMail.Adjunto = ""
  TMail.Credito_No = ""
  Total = 0
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Codigo = .fields("CodigoC")
       TBeneficiario.Cliente = .fields("Cliente")
       TBeneficiario.Representante = .fields("Representante")
       TBeneficiario.Grupo_No = .fields("Grupo")
       TBeneficiario.Email1 = .fields("Email")
       TBeneficiario.Email2 = .fields("Email2")
       TBeneficiario.EmailR = .fields("EmailR")
       'If Len(TBeneficiario.Representante) <= 1 And Len(TBeneficiario.Cliente) > 1 Then TBeneficiario.Representante = TBeneficiario.Cliente
        
       html_Detalle_adicional = "<thead><tr><strong><td>TC</td><td>FECHA EMIS</td><td>SERIE</td><td>DOCUMENTO</td><td>SALDO ACTUTAL</td></strong></tr></thead><tbody>"
       
      'NombreRepresentante = TBeneficiario.Representante
       Do While Not .EOF
          If Codigo <> .fields("CodigoC") Then
             ComunicadoEntidad = ""
             TMail.Asunto = "Envio automatizado de su cartera pendiente por USD " & Format$(Total, "#,#0.00")
             
             html_Detalle_adicional = html_Detalle_adicional _
                                    & "<tr>" _
                                    & "<td colspan='4'>TOTAL PENDIENTE POR CANCELAR</td>" _
                                    & "<td class='row text-right'>" & Format$(Total, "#,#0.00") & "</td>" _
                                    & "</tr></tbody>"
             TMail.Mensaje = "Estimado(a): "
             If Len(TBeneficiario.Representante) > 1 Then TMail.Mensaje = TMail.Mensaje & TBeneficiario.Representante & vbCrLf Else TMail.Mensaje = TMail.Mensaje & TBeneficiario.Cliente & vbCrLf
             If Len(TBeneficiario.Representante) > 1 And TBeneficiario.Representante <> TBeneficiario.Cliente Then TMail.Mensaje = TMail.Mensaje & "Beneficiario: " & TBeneficiario.Cliente & vbCrLf
             TMail.Mensaje = TMail.Mensaje _
                           & "Grupo " & TBeneficiario.Grupo_No & vbCrLf _
                           & "Usted tiene los siguientes pendientes por cancelar:" & vbCrLf _
                           & "<table BORDER CELLPADDING=10 CELLSPACING=0 class='content-table'>" & html_Detalle_adicional & "</table>" _
                           & "<N>NOTA:</N> En caso de tener inconformidad con los valores detallados en su Estado de Cuenta, comuniquese con atencion al Cliente." & vbCrLf
            'Enviamos lista de mails
             TMail.para = ""
             Insertar_Mail TMail.para, TBeneficiario.EmailR
             Insertar_Mail TMail.para, TBeneficiario.Email1
             Insertar_Mail TMail.para, TBeneficiario.Email2
             Insertar_Mail TMail.para, EmailProcesos
            'MsgBox TBeneficiario.Representante & " -> " & TBeneficiario.Cliente & vbCrLf & TMail.Mensaje
             FEnviarCorreos.Show vbModal
             TMail.Mensaje = ""
             TMail.MensajeHTML = ""
             html_Detalle_adicional = "<thead><tr><strong><td>TC</td><td>FECHA EMIS</td><td>SERIE</td><td>DOCUMENTO</td><td>SALDO ACTUTAL</td></strong></tr></thead>"
             Codigo = .fields("CodigoC")
             TBeneficiario.Cliente = .fields("Cliente")
             TBeneficiario.Representante = .fields("Representante")
             TBeneficiario.Grupo_No = .fields("Grupo")
             TBeneficiario.Email1 = .fields("Email")
             TBeneficiario.Email2 = .fields("Email2")
             TBeneficiario.EmailR = .fields("EmailR")
             Total = 0
          End If
          html_Detalle_adicional = html_Detalle_adicional _
                                 & "<tr>" _
                                 & "<td>" & .fields("TC") & "</td>" _
                                 & "<td>" & .fields("Fecha") & "</td>" _
                                 & "<td>" & .fields("Serie") & "</td>" _
                                 & "<td>" & Format(.fields("Factura"), "000000000") & "</td>" _
                                 & "<td class='row text-right'>" & Format$(.fields("Saldo_Actual"), "#,#0.00") & "</td>" _
                                 & "</tr>"
          Total = Total + .fields("Saldo_Actual")
         .MoveNext
       Loop
       ComunicadoEntidad = ""
       TMail.Asunto = "Envio automatizado de su cartera pendiente por USD " & Format$(Total, "#,#0.00")
       html_Detalle_adicional = html_Detalle_adicional _
                              & "<tr>" _
                              & "<td colspan='4'>TOTAL PENDIENTE POR CANCELAR</td>" _
                              & "<td class='row text-right'>" & Format$(Total, "#,#0.00") & "</td>" _
                              & "</tr></tbody>"
       TMail.Mensaje = "Estimado(a): "
       If Len(TBeneficiario.Representante) > 1 Then TMail.Mensaje = TMail.Mensaje & TBeneficiario.Representante & vbCrLf Else TMail.Mensaje = TMail.Mensaje & TBeneficiario.Cliente & vbCrLf
       If Len(TBeneficiario.Representante) > 1 And TBeneficiario.Representante <> TBeneficiario.Cliente Then TMail.Mensaje = TMail.Mensaje & "Beneficiario: " & TBeneficiario.Cliente & vbCrLf
       TMail.Mensaje = TMail.Mensaje _
                     & "Grupo " & TBeneficiario.Grupo_No & vbCrLf _
                     & "Usted tiene los siguientes pendientes por cancelar:" & vbCrLf _
                     & "<table BORDER CELLPADDING=10 CELLSPACING=0 class='content-table'>" & html_Detalle_adicional & "</table>" _
                     & "<N>NOTA:</N> En caso de tener inconformidad con los valores detallados en su Estado de Cuenta, comuniquese con atencion al Cliente." & vbCrLf
       TMail.para = ""
       Insertar_Mail TMail.para, TBeneficiario.Email1
       Insertar_Mail TMail.para, TBeneficiario.Email2
       Insertar_Mail TMail.para, TBeneficiario.EmailR
'      MsgBox TBeneficiario.Representante & " -> " & TBeneficiario.Cliente & vbCrLf & TMail.Mensaje
       FEnviarCorreos.Show vbModal
       TMail.Mensaje = ""
       TMail.MensajeHTML = ""
      'If Len(TMail.ListaError) > 1 Then Lista_Error = Lista_Error & TBeneficiario.Representante & " - Email: " & TMail.para & " => " & TMail.ListaError & vbCrLf
   End If
  End With
  DGQuery.Visible = True
  If Len(TMail.ListaError) > 1 Then
     MsgBox "Rebice en su correo los errores "
     TMail.para = Lista_De_Correos(0).Correo_Electronico
     TMail.Asunto = "CORREOS CON ERRORES"
     TMail.Mensaje = TMail.ListaError
     FEnviarCorreos.Show 1
  End If
End Sub

Public Sub Por_Buses(Patron_Busqueda As String)
'Por Buses
 Opcion = 12
 DGQuery.Caption = "HISTORIAL DE FACTURAS"
 Label2.Caption = "(" & Opcion & ") Facturado"
 Label4.Caption = " Cobrado"
 Label3.Caption = " Saldo"
 PorCxC = False

 If CheqCxC.value = 1 Then PorCxC = True
 RatonReloj
 Total = 0
 Abono = 0
 Saldo = 0
 sSQL = "SELECT Cliente,Telefono,Direccion As Curso,DireccionT As Direccion_Ruta,Contacto As Ruta " _
      & "FROM Clientes " _
      & "WHERE Plan_Afiliado = '" & Patron_Busqueda & "' " _
      & "AND FA <> " & Val(adFalse) & " " _
      & "ORDER BY Cliente "
 Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
End Sub

Public Sub SMAbonos_Anticipados()
'Abonos Anticipados de Clientes
    Opcion = 20
    RatonReloj
    Label2.Caption = "(" & Opcion & ") Ventas"
    Label4.Caption = " Cobrado"
    Label3.Caption = " Saldo"
    
    DGQuery.Visible = False
    SubCtaGen = Leer_Seteos_Ctas("Cta_Anticipos_Clientes")
'''    DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
'''    DGQuery.width = ((MDI_X_Max / 3) * 2) - 100
'''
'''    FrmAnticipos.Left = DGQuery.width + 200
'''    FrmAnticipos.Height = MDI_Y_Max - FrmAnticipos.Top - 400
'''    FrmAnticipos.width = (MDI_X_Max / 3) - 100
'''
'''    WBPDF.Height = FrmAnticipos.Height - 500
'''    WBPDF.width = FrmAnticipos.width - 400
'''
'''    FrmAnticipos.Visible = True
    
    DGQuery.Caption = "ABONOS ANTICIPADOS DE CLIENTES"
    sSQL = "SELECT C.Cliente, C.CI_RUC, TS.Cta, TS.Fecha, TS.TP, TS.Numero, TS.Creditos As Abono " _
         & "FROM Trans_SubCtas As TS, Clientes As C " _
         & "WHERE TS.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND TS.Item = '" & NumEmpresa & "' " _
         & "AND TS.Periodo = '" & Periodo_Contable & "' " _
         & "AND TS.T <> 'A' " _
         & Tipo_De_Consulta() _
         & "AND TS.Codigo = C.Codigo " _
         & "ORDER BY C.Cliente, TS.Cta, TS.Fecha, TS.TP, TS.Numero "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
    RatonReloj
    DGQuery.Visible = False
    Total = 0: Abono = 0
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Abono = Abono + .fields("Abono")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    LabelFacturado.Caption = Format$(Total, "#,##0.00")
    LabelAbonado.Caption = Format$(Abono, "#,##0.00")
    LabelSaldo.Caption = "0.00"
    DGQuery.Visible = True
    RatonNormal
End Sub

Public Sub Contra_Cta_Abonos()
Dim ContraCta As String
'Abonos Anticipados de Clientes
    Opcion = 21
    RatonReloj
    Label2.Caption = "(" & Opcion & ") Debitos"
    Label4.Caption = " Creditos"
    Label3.Caption = " Saldo"
    
    DGQuery.Visible = False
    DGQuery.Caption = "ABONOS ANTICIPADOS DE CLIENTES"
    
    sSQL = "SELECT CC.Cuenta, C.Cliente, TS.Fecha, TS.TP, TS.Numero, TS.Debitos, TS.Creditos, T.Cta AS Contra_Cta, TS.Cta " _
         & "FROM Trans_SubCtas AS TS, Transacciones AS T, Catalogo_Cuentas AS CC, Clientes AS C " _
         & "WHERE TS.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND TS.Item = '" & NumEmpresa & "' " _
         & "AND TS.Periodo = '" & Periodo_Contable & "' " _
         & "AND TS.T <> 'A' "
    If CheqAbonos.value = 1 Then
       ContraCta = SinEspaciosIzq(DCCxC)
       sSQL = sSQL & "AND TS.Cta = '" & ContraCta & "' "
       If MidStrg(ContraCta, 1, 1) = "1" Then sSQL = sSQL & "AND TS.Debitos > 0  "
       If MidStrg(ContraCta, 1, 1) = "2" Then sSQL = sSQL & "AND TS.Creditos > 0 "
    End If
    sSQL = sSQL _
         & "AND TS.Periodo = T.Periodo " _
         & "AND TS.Periodo = CC.Periodo " _
         & "AND TS.Item = T.Item " _
         & "AND TS.Item = CC.Item " _
         & "AND TS.TP = T.TP " _
         & "AND TS.Numero = T.Numero " _
         & "AND T.Cta = CC.Codigo " _
         & "AND TS.Codigo = C.Codigo " _
         & "AND TS.Cta <> T.Cta " _
         & "ORDER BY T.Cta, C.Cliente, TS.Fecha, TS.TP, TS.Numero "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
    RatonReloj
    DGQuery.Visible = False
    Total = 0: Abono = 0
    With AdoQuery.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Total = Total + .fields("Debitos")
            Abono = Abono + .fields("Creditos")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    LabelFacturado.Caption = Format$(Total, "#,##0.00")
    LabelAbonado.Caption = Format$(Abono, "#,##0.00")
    LabelSaldo.Caption = "0.00"
    DGQuery.Visible = True
    RatonNormal
End Sub

Public Sub Tipo_Pago_Cliente()
    sSQL = "SELECT C.Grupo, C.Cliente, C.CI_RUC, CM.Representante, CM.Cedula_R, CM.Telefono_R, " _
         & "CM.Tipo_Cta, CM.Cta_Numero, CM.Caducidad, CM.Cod_Banco, TRSRI.Descripcion As Institucion_Financiera " _
         & "FROM Clientes As C, Clientes_Matriculas As CM, Tabla_Referenciales_SRI As TRSRI " _
         & "WHERE CM.Item = '" & NumEmpresa & "' " _
         & "AND CM.Periodo = '" & Periodo_Contable & "' " _
         & "AND TRSRI.Tipo_Referencia = 'BANCOS Y COOP' " _
         & "AND C.Codigo = CM.Codigo " _
         & "AND CM.Cod_Banco = TRSRI.Codigo " _
         & "ORDER BY CM.Tipo_Cta, C.Grupo, C.Cliente "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , , True
End Sub
