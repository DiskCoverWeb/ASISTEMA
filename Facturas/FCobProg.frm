VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FCobrosProgramados 
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11790
   WindowState     =   1  'Minimized
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
      Height          =   750
      Left            =   4935
      Picture         =   "FCobProg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1575
      Width           =   1275
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
      Height          =   750
      Left            =   6300
      Picture         =   "FCobProg.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1575
      Width           =   1275
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
      Height          =   750
      Left            =   3570
      Picture         =   "FCobProg.frx":12C0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1575
      Width           =   1275
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
      Height          =   750
      Left            =   2205
      Picture         =   "FCobProg.frx":1956
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1575
      Width           =   1275
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
      Left            =   5355
      TabIndex        =   1
      Top             =   0
      Width           =   5490
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
      Width           =   5175
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
      Width           =   5160
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
      Width           =   5160
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FCobProg.frx":1D98
      DataSource      =   "AdoListCtas"
      Height          =   1155
      Left            =   5355
      TabIndex        =   4
      Top             =   315
      Width           =   5895
      _ExtentX        =   10398
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
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FCobProg.frx":1DB2
      Height          =   4215
      Left            =   105
      TabIndex        =   5
      Top             =   2415
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   7435
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   14
      Top             =   1995
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
      TabIndex        =   15
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
      TabIndex        =   17
      Top             =   1995
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
      TabIndex        =   16
      Top             =   1575
      Width           =   750
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
            Picture         =   "FCobProg.frx":1DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":20E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":23FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":2717
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":2A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":2D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":3065
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":337F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":3699
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":39B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":3CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":3FE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":12D79
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":13093
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":133AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":136C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCobProg.frx":13879
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FCobrosProgramados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Tipo_De_Grabado(Dias() As Double, Venc() As Double)
        SetAdoAddNew "Saldo_Diarios"
        SetAdoFields "T", Normal
        SetAdoFields "TP", "FPFA"
        SetAdoFields "Fecha", FechaSistema
        SetAdoFields "CodigoC", CodigoCliente
        SetAdoFields "Numero", Factura_No
        SetAdoFields "Ven_mas_de_360", Venc(7)
        SetAdoFields "Ven_181_a_360", Venc(6)
        SetAdoFields "Ven_91_a_180", Venc(5)
        SetAdoFields "Ven_61_a_90", Venc(4)
        SetAdoFields "Ven_31_a_60", Venc(3)
        SetAdoFields "Ven_8_a_30", Venc(2)
        SetAdoFields "Ven_1_a_7", Venc(1)
        SetAdoFields "CxC_Hoy", Dias(0)
        SetAdoFields "De_1_a_7", Dias(1)
        SetAdoFields "De_8_a_30", Dias(2)
        SetAdoFields "De_31_a_60", Dias(3)
        SetAdoFields "De_61_a_90", Dias(4)
        SetAdoFields "De_91_a_180", Dias(5)
        SetAdoFields "De_181_a_360", Dias(6)
        SetAdoFields "Mas_De_360", Dias(7)
        'SetAdoFields "Total", 0
        'SetAdoFields "Total_Mora", 0
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        
End Sub

Public Sub TipoConsultaCxC(Tipo As String)
Dim Dias(10) As Double
Dim Venc(10) As Double
FechaValida MBFechaI
FechaValida MBFechaF
FechaTexto = MBFechaI.Text
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
DGQuery.Caption = ""
  RatonReloj
  For I = 0 To 9
     Dias(I) = 0
     Venc(I) = 0
  Next I
  Contador = 0
  SQL1 = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE TP = 'FPFA' "
  Ejecutar_SQL_SP SQL1
  
  sSQL = "SELECT F.T,C.Cliente,F.Fecha,F.Pagos,F.Numero,C.Codigo " _
       & "FROM Trans_Prestamos As F,Clientes As C " _
       & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND F.TP = 'FPFA' " _
       & "AND Codigo = Cuenta_No "
  If OpcBusq.value Then
     TextoBusqueda = TxtCIRUC.Text
     If TextoBusqueda <> Ninguno Then
        If LstCampos.Text <> "Ninguno" Then
           TipoDatoBusq = AdoListCtas.Recordset.fields(LstCampos.Text).Type
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
     sSQL = sSQL & "AND F.Codigo = '" & CodigoCli & "' "
  End If
  sSQL = sSQL & "AND F.T = '" & Pendiente & "' "
  If Tipo = "C" Then sSQL = sSQL & "ORDER BY C.Cliente,F.Fecha,F.Numero "
  If Tipo = "F" Then sSQL = sSQL & "ORDER BY F.Numero,F.Fecha "
  Select_Adodc AdoAux, sSQL
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Saldo = 0
       CodigoCliente = .fields("Codigo")
       Factura_No = .fields("Numero")
       Do While Not .EOF
          If Tipo = "C" Then
             If CodigoCliente <> .fields("Codigo") Then
                Tipo_De_Grabado Dias, Venc
                Saldo = 0
                'MsgBox CodigoCliente
                CodigoCliente = .fields("Codigo")
                Factura_No = .fields("Numero")
                For I = 0 To 9
                    Dias(I) = 0
                    Venc(I) = 0
                Next I
             End If
          End If
          If Tipo = "F" Then
             If Factura_No <> .fields("Numero") Then
                Tipo_De_Grabado Dias, Venc
                Saldo = 0
                CodigoCliente = .fields("Codigo")
                Factura_No = .fields("Numero")
                For I = 0 To 9
                    Dias(I) = 0
                    Venc(I) = 0
                Next I
             End If
          End If
          Mifecha = .fields("Fecha")
          J = CFechaLong(FechaSistema)
          I = CFechaLong(Mifecha)
          Total = .fields("Pagos")
          Saldo = Saldo + Total
          K = I - J
          'MsgBox MiFecha & " - " & K & " - " & Total
          If K = 0 Then Dias(0) = Dias(0) + Total
          Select Case K
            Case Is < -360: Venc(7) = Venc(7) + Total
            Case -360 To -181: Venc(6) = Venc(6) + Total
            Case -180 To -91: Venc(5) = Venc(5) + Total
            Case -90 To -61: Venc(4) = Venc(4) + Total
            Case -60 To -31: Venc(3) = Venc(3) + Total
            Case -30 To -8: Venc(2) = Venc(2) + Total
            Case -7 To -1: Venc(1) = Venc(1) + Total
            Case 0: Dias(0) = Dias(0) + Total
            Case 1 To 7: Dias(1) = Dias(1) + Total
            Case 8 To 30: Dias(2) = Dias(2) + Total
            Case 31 To 60: Dias(3) = Dias(3) + Total
            Case 61 To 90: Dias(4) = Dias(4) + Total
            Case 91 To 180: Dias(5) = Dias(5) + Total
            Case 181 To 360: Dias(6) = Dias(6) + Total
            Case Is > 360: Dias(7) = Dias(7) + Total
          End Select
         .MoveNext
       Loop
       Tipo_De_Grabado Dias, Venc
   End If
  End With
  'MsgBox "..."
  Total = 0: Saldo = 0
  DGQuery.Visible = False
  sSQL = "SELECT C.Cliente,Numero As Factura," _
       & "Total_Mora," _
       & "Ven_mas_de_360," _
       & "Ven_181_a_360," _
       & "Ven_91_a_180," _
       & "Ven_61_a_90," _
       & "Ven_31_a_60," _
       & "Ven_8_a_30," _
       & "Ven_1_a_7," _
       & "CxC_Hoy," _
       & "De_1_a_7," _
       & "De_8_a_30," _
       & "De_31_a_60," _
       & "De_61_a_90," _
       & "De_91_a_180," _
       & "De_181_a_360," _
       & "Mas_de_360," _
       & "Total " _
       & "FROM Saldo_Diarios As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.TP = 'FPFA' " _
       & "AND C.Codigo = F.CodigoC "
  If Tipo = "C" Then
     sSQL = sSQL & "ORDER BY C.Cliente "
  Else
     sSQL = sSQL & "ORDER BY F.Numero "
  End If
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          FCobrosProgramados.Caption = "Recalculando " & Format$(Contador / .RecordCount, "00%")
          Total = Total + .fields("CxC_Hoy")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGQuery.Visible = True
  LabelFacturado.Caption = Format$(Total, "#,##0.00")
  LabelAbonado.Caption = Format$(Saldo, "#,##0.00")
  FCobrosProgramados.Caption = "COBROS PROGRAMADOS"
RatonNormal
End Sub

Public Sub ListarClientes(Optional LlenarCliente As Boolean)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  If LlenarCliente Then sSQL = sSQL & "AND FA = " & Val(adTrue) & " "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
  'Label2.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "'")
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
  TipoDoc = "F"
  TipoConsultaCxC TipoDoc
End Sub

Private Sub Command2_Click()
  TipoDoc = "C"
  TipoConsultaCxC TipoDoc
End Sub

Private Sub Command4_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBFechaF.Text
   ImprimirCtasCob AdoQuery, sSQL, True
   DGQuery.Visible = True
End Sub

Private Sub Command5_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBFechaF.Text
   Imprimir_Resumen_Cartera AdoQuery, Codigo4
   DGQuery.Visible = True
End Sub

Private Sub Command6_Click()
  Unload Me
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
     Select_Adodc AdoListCtas, sSQL
     RatonReloj
     With AdoListCtas.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
            .fields("Cliente") = UCaseStrg(CompilarString(.fields("Cliente")))
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

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  
  If KeyCode = vbKeyF1 Then GenerarDataTexto FCobrosProgramados, AdoQuery
  'If CtrlDown And KeyCode = vbKeyB Then Buscar_Datos DGQuery, AdoQuery
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdo AdoQuery, True, 2, 8
End Sub

Private Sub Form_Activate()
  FCobrosProgramados.Caption = "COBROS PROGRAMADOS"
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
  DCCliente.SetFocus
  RatonNormal
  FCobrosProgramados.WindowState = vbMaximized
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
  FCobrosProgramados.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   CentrarForm FCobrosProgramados
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
          TipoDatoBusq = .fields(LstCampos.Text).Type
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
                DCCliente.Text = .fields("Cliente")
                CodigoCli = .fields("Codigo")
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

