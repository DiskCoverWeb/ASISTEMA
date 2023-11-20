VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FResumenFletes 
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11490
   WindowState     =   1  'Minimized
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FRMFlete.frx":0000
      Height          =   4635
      Left            =   105
      TabIndex        =   5
      Top             =   1995
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8176
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   105
      Top             =   6720
      Width           =   3480
      _ExtentX        =   6138
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FRMFlete.frx":0017
      DataSource      =   "AdoListCtas"
      Height          =   1155
      Left            =   5040
      TabIndex        =   4
      Top             =   315
      Width           =   5160
      _ExtentX        =   9102
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4305
      TabIndex        =   17
      Top             =   1470
      Width           =   4110
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
         Left            =   2940
         TabIndex        =   20
         Top             =   105
         Width           =   885
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
         Left            =   1470
         TabIndex        =   19
         Top             =   105
         Width           =   1395
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
         Left            =   105
         TabIndex        =   18
         Top             =   105
         Value           =   -1  'True
         Width           =   1290
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
      Left            =   10290
      Picture         =   "FRMFlete.frx":0031
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1995
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
      Height          =   855
      Left            =   10290
      Picture         =   "FRMFlete.frx":0A27
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   105
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
      Height          =   855
      Left            =   10290
      Picture         =   "FRMFlete.frx":0E69
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1050
      Width           =   1065
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
      Value           =   -1  'True
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
      Width           =   4965
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
      Left            =   7035
      Top             =   420
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
      Left            =   7035
      Top             =   735
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
      Left            =   5145
      Top             =   735
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
      Left            =   5145
      Top             =   1050
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
      Left            =   7035
      Top             =   1050
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
      Left            =   5145
      Top             =   420
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2940
      TabIndex        =   10
      Top             =   1575
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
      TabIndex        =   11
      Top             =   1575
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
      Left            =   8400
      TabIndex        =   6
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
      Left            =   6930
      TabIndex        =   7
      Top             =   6720
      Width           =   1485
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
      Left            =   5145
      TabIndex        =   8
      Top             =   6720
      Width           =   1800
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
      Left            =   3570
      TabIndex        =   9
      Top             =   6720
      Width           =   1590
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
      Left            =   2205
      TabIndex        =   13
      Top             =   1575
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
      TabIndex        =   12
      Top             =   1575
      Width           =   750
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
            Picture         =   "FRMFlete.frx":1733
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":1A4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":1D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":2081
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":239B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":26B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":29CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":2CE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":3003
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":331D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":3637
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":3951
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":126E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":129FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":12D17
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":13031
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FRMFlete.frx":131E3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FResumenFletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub TipoConsultaCxC(Tipo As String)
If Tipo = "C" Then
   sSQL = "SELECT F.T,C.Cliente,F.Fecha,F.Factura,F.Total_MN,F.Total_ME,F.Saldo_MN,F.Saldo_ME," _
        & "C.CI_RUC,C.Telefono,C.Celular,C.FAX,C.Ciudad,C.Direccion,C.Email,C.Prov,C.Codigo "
ElseIf Tipo = "F" Then
   sSQL = "SELECT F.T,F.Fecha,F.Factura,C.Cliente,F.Total_MN," _
        & "(F.Total_MN-F.Saldo_MN) As Abono_MN,F.Saldo_MN,C.Telefono "
End If
sSQL = sSQL & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND C.Codigo = F.CodigoC " _
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
                TextoBusqueda = "MidStrg(C." & LstCampos.Text & ", 1," & Len(TextoBusqueda) & ") = '" & TextoBusqueda & "' "
           Case Else
                TextoBusqueda = "MidStrg(C." & LstCampos.Text & ", 1," & Len(TextoBusqueda) & ") = '" & TextoBusqueda & "' "
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
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
Total = 0: Saldo = 0
DGQuery.Visible = False
With AdoQuery.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
    
  If Tipo = "C" Then Codigo4 = .Fields("Direccion") Else Codigo4 = TxtCIRUC.Text
  Do While Not .EOF
     Contador = Contador + 1
     FCarteraCli.Caption = "Recalculando " & Format$(Contador / .RecordCount, "00%")
     If Tipo = "C" Then
        Total = Total + .Fields("Total_MN")
     Else
        Total = Total + .Fields("Abono_MN")
     End If
     Saldo = Saldo + .Fields("Saldo_MN")
    .MoveNext
  Loop
  .MoveFirst
  End If
End With
DGQuery.Visible = True
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Saldo, "#,##0.00")
FCarteraCli.Caption = "CARTERA POR CLIENTES"
RatonNormal
End Sub

Public Sub ListarClientes(Optional LlenarCliente As Boolean)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  If LlenarCliente Then sSQL = sSQL & "AND FA = " & Val(adTrue) & " "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
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

Private Sub Command2_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  DGQuery.Caption = "LISTADO DE FLETES"
  RatonReloj
  Contador = 0
  sSQL = "SELECT Ok,Cliente,Numero,Fecha_I,Fecha_F,Factura,Producto,Carga,Flete," _
       & "Km_Inicial,Km_Final,(Km_Final-Km_Inicial) As Recorrido,TF.Referencia," _
       & "TF.CodigoC,TF.Ayudante,TF.Conductor " _
       & "FROM Clientes As C,Catalogo_Productos As CP,Trans_Fletes As TF " _
       & "WHERE TF.Fecha_I BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TF.Item = '" & NumEmpresa & "' " _
       & "AND TF.Periodo = '" & Periodo_Contable & "' "
  If OpcBusq.value Then
     TextoBusqueda = TxtCIRUC.Text
     If LstCampos.Text <> "Ninguno" Then
        Select Case LstCampos.Text
          Case "Cliente": TextoBusqueda = "C." & TextoBusqueda
          Case "Ayudante": TextoBusqueda = "TF." & TextoBusqueda
          Case "Conductor": TextoBusqueda = "TF." & TextoBusqueda
          Case "CI_RUC": TextoBusqueda = "C." & TextoBusqueda
          Case "Numero": TextoBusqueda = "TF." & TextoBusqueda
          Case "Ruta": TextoBusqueda = "TF." & TextoBusqueda
        End Select
        sSQL = sSQL & "AND " & TextoBusqueda & " "
     End If
  Else
     sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
  End If
  If OpcPend.value Then sSQL = sSQL & "AND TF.T = '" & Pendiente & "' "
  If OpcCanc.value Then sSQL = sSQL & "AND TF.T = '" & Cancelado & "' "
  sSQL = sSQL & "AND TF.T = '" & Normal & "' " _
       & "AND C.Codigo = TF.CodigoC " _
       & "AND CP.Item = TF.Item " _
       & "AND CP.Periodo = TF.Periodo " _
       & "AND CP.Codigo_Inv = TF.Codigo_Inv " _
       & "ORDER BY Cliente,Fecha_I DESC,TF.Numero "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL
  Opcion = 1
End Sub

Private Sub Command4_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBFechaF.Text
   If Opcion = 1 Then
      ImprimirCtasCob AdoQuery, sSQL, True
   Else
      If CheqCom.value = 1 Then
         Imprimir_Pendientes_Facturacion AdoQuery, Opcion, True
      Else
         Imprimir_Pendientes_Facturacion AdoQuery, Opcion, True
      End If
   End If
   DGQuery.Visible = True
End Sub

Private Sub Command6_Click()
  Unload FResumenFletes
End Sub


Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente.Text
  TipoDoc = "M"
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FResumenFletes, AdoQuery
     DGQuery.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  FResumenFletes.Caption = "CREACION DEL CLIENTE"
  ListarClientes CliFact
  LstCampos.Clear
  LstCampos.AddItem "Ninguno"
  LstCampos.AddItem "Cliente"
  LstCampos.AddItem "Ayudante"
  LstCampos.AddItem "Conductor"
  LstCampos.AddItem "CI_RUC"
  LstCampos.AddItem "Numero"
  LstCampos.AddItem "Ruta"
  LstCampos.Text = "Ninguno"
  DCCliente.SetFocus
  RatonNormal
  FResumenFletes.WindowState = vbMaximized
  If Nuevo Then
     TxtApellidosS = NombreCliente
     LblCodigo.Caption = "Ninguno"
     TxtGrupo.Text = NumEmpresa
     TxtApellidosS.SetFocus
  Else
     ListarCuenta DCCliente.Text
     DCCliente.SetFocus
  End If
End Sub

Private Sub Form_Deactivate()
  FResumenFletes.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
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
                RatonNormal
               .MoveFirst
               .Find (TextoBusqueda)
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

