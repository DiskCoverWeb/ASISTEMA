VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FNotaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANULACION DE FACTURAS"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "FNotaCrd.frx":0000
      Height          =   2010
      Left            =   105
      TabIndex        =   14
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3545
      _Version        =   393216
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
      Picture         =   "FNotaCrd.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1995
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Diario &Caja"
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
      Picture         =   "FNotaCrd.frx":08E3
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar"
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
      Picture         =   "FNotaCrd.frx":0D25
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1050
      Width           =   1065
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Mayorizar"
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
      Left            =   105
      Picture         =   "FNotaCrd.frx":1167
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   2
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
   Begin MSAdodcLib.Adodc AdoVentas 
      Height          =   330
      Left            =   525
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
      Caption         =   "Ventas"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   4515
      Top             =   1365
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
      Caption         =   "Asiento"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   2520
      Top             =   1365
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
      Caption         =   "Inv"
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
      Left            =   525
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
   Begin MSAdodcLib.Adodc AdoSQL 
      Height          =   330
      Left            =   2520
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
      Caption         =   "SQL"
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
      Left            =   2520
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   525
      Top             =   1365
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
      Caption         =   "Clientes"
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
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   2940
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc AdoVentaAct 
      Height          =   330
      Left            =   4515
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
      Caption         =   "VentaAct"
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
   Begin MSAdodcLib.Adodc AdoInv1 
      Height          =   330
      Left            =   4515
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
      Caption         =   "Inv1"
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
   Begin MSAdodcLib.Adodc AdoFactAnul 
      Height          =   330
      Left            =   6510
      Top             =   1365
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
      Caption         =   "FactAnul"
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
   Begin MSAdodcLib.Adodc AdoAnticipos 
      Height          =   330
      Left            =   6510
      Top             =   1785
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
      Caption         =   "Anticipos"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Periodo de Cierre"
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
      TabIndex        =   15
      Top             =   105
      Width           =   6630
   End
   Begin VB.Label LblConcepto 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   3570
      TabIndex        =   5
      Top             =   420
      Width           =   6630
   End
   Begin VB.Label LabelHaber 
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
      Left            =   8400
      TabIndex        =   6
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label LabelDebe 
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
      Left            =   6615
      TabIndex        =   7
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label LblDiferencia 
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
      Left            =   3780
      TabIndex        =   10
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diferencia "
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
      TabIndex        =   9
      Top             =   2940
      Width           =   1065
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTALES "
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
      TabIndex        =   8
      Top             =   2940
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Periodo de Cierre"
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
      Left            =   945
      TabIndex        =   1
      Top             =   105
      Width           =   2535
   End
End
Attribute VB_Name = "FNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ErrorInventario As String
Dim CtasProc() As CtasAsiento
Dim ContCtas As Integer
Dim Combos As String

Private Sub Command1_Click()
   ErrorInventario = ""
   GrabarAsientosFacturacion Normal
   If Redondear(Debe - Haber, 2) <> 0 Then
      MsgBox "Las Transacciones no cuadran," & vbCrLf & "verifique de que anulo las facturas correctas"
      Command1.SetFocus
   Else
      If Command2.Enabled Then Command2.SetFocus Else Command5.SetFocus
   End If
End Sub

Private Sub Command2_Click()
   FechaValida MBFechaI
   FechaValida MBFechaF
   If MBFechaI.Text = MBFechaF.Text Then
      Cadena = "Cierre de Caja del " & MBFechaI
   Else
      Cadena = "Cierre de Caja del " & MBFechaI & " al " & MBFechaF
   End If
   If ((NuevoDiario) And Redondear(Debe - Haber, 2) = 0) Then
       NumComp = ReadSetDataNum("Diario", True, False)
       Mensajes = "Esta seguro de Grabar el Comprobante No. " & NumComp & "]"
       Titulo = "Pregunta de grabación"
       If BoxMensaje = vbYes Then
          If AdoAsiento.Recordset.RecordCount > 0 Then
             RatonReloj
             Mensajes = "Imprimir Diario de Caja."
             Titulo = "Pregunta de Impresion"
             If BoxMensaje = vbYes Then Imprimir_Diario_Caja AdoVentas, AdoCxC, AdoInv, AdoInv1, AdoAnticipos, MBFechaI, MBFechaF
             FechaTexto = MBFechaF.Text
             FechaIni = BuscarFecha(MBFechaF.Text)
             NumComp = ReadSetDataNum("Diario", True, True)
             DiarioCaja = NumComp
            'Inventario
             FechaTexto = MBFechaF.Text
             sSQL = "SELECT * " _
                  & "FROM Trans_Kardex " _
                  & "WHERE Item = '" & NumEmpresa & "' "
             SelectAdodc AdoSQL, sSQL
             With AdoInv.Recordset
              If .RecordCount > 0 Then
                  RatonReloj
                  CodigoBenef = SinEspaciosIzq(DCBenef.Text)
                 .MoveFirst
                  Cta_Inventario = .Fields("CTA_INVENTARIO")
                  Contra_Cta = .Fields("CONTRA_CTA")
                  ValorDH = 0
                 'Llenamos los datos ingresados al Kardex
                  Do While Not .EOF
                     SetAddNew AdoSQL
                     SetFields AdoSQL, "T", Normal
                     SetFields AdoSQL, "Codigo_Inv", .Fields("Codigo_Inv")
                     SetFields AdoSQL, "Codigo_P", Ninguno
                     SetFields AdoSQL, "Fecha", FechaTexto
                     SetFields AdoSQL, "TP", CompDiario
                     SetFields AdoSQL, "Numero", NumComp
                     SetFields AdoSQL, "Descuento", .Fields("P_DESC")
                     If .Fields("DH") = 1 Then
                         SetFields AdoSQL, "Entrada", .Fields("CANT_ES")
                     Else
                         SetFields AdoSQL, "Salida", .Fields("CANT_ES")
                     End If
                     SetFields AdoSQL, "Valor_Unitario", .Fields("VALOR_UNIT")
                     SetFields AdoSQL, "Valor_Total", .Fields("VALOR_TOTAL")
                     SetFields AdoSQL, "Existencia", .Fields("CANTIDAD")
                     SetFields AdoSQL, "Total", .Fields("SALDO")
                     SetFields AdoSQL, "Cta_Inv", .Fields("CTA_INVENTARIO")
                     SetFields AdoSQL, "Contra_Cta", .Fields("CONTRA_CTA")
                     SetFields AdoSQL, "Item", NumEmpresa
                     SetUpdate AdoSQL
                     Asiento = Asiento + 1
                     ValorDH = ValorDH + .Fields("VALOR_TOTAL")
                    .MoveNext
                  Loop
              End If
             End With
            'Grabacion del Comprobante
             Co.T = Normal
             Co.TP = CompDiario
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             If MBFechaI.Text = MBFechaF.Text Then
                Co.Concepto = "Cierre de Caja del " & MBFechaI.Text & ", Diario No. " & NumComp
             Else
                Co.Concepto = "Cierre de Caja del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. " & NumComp
             End If
             Co.CodigoB = Ninguno
             Co.Efectivo = 0
             Co.Monto_Total = Debe
             Co.T_No = Trans_No
             Co.Usuario = CodigoUsuario
             Co.Item = NumEmpresa
             GrabarComprobante Co
             Control_Procesos Normal, Co.Concepto
             ImprimirComprobantesDe False, Co
             IniciarAsientosDe DGAsiento, AdoAsiento
             LabelDebe.Caption = Format$(0, "#,##0.00")
             LabelHaber.Caption = Format$(0, "#,##0.00")
             RatonNormal
          End If
          Mifecha = BuscarFecha(FechaSistema)
          FechaIni = BuscarFecha(MBFechaI.Text)
          FechaFin = BuscarFecha(MBFechaF.Text)
             sSQL = "UPDATE Trans_Abonos " _
                  & "SET C = " & Val(adTrue) & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Fecha <= #" & FechaFin & "# "
             Conectar_Ado_Execute sSQL
             'MsgBox sSQL
             sSQL = "UPDATE Facturas " _
                  & "SET C = " & Val(adTrue) & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Fecha <= #" & FechaFin & "# "
             Conectar_Ado_Execute sSQL
        End If
   End If
End Sub

Private Sub Command3_Click()
  Unload FCierreCaja
End Sub

Private Sub Command8_Click()
  RatonReloj
  If Inv_Promedio Then
     FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO PRECIO PROMEDIO"
  Else
     FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO ULTIMO PRECIO"
  End If
  MayorizarInv.Show 1
  RatonNormal
  MBFechaI.SetFocus
End Sub


Private Sub DGAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGAsiento.Visible = False
     GenerarDataTexto FCierreCaja, AdoAsiento
     DGAsiento.Visible = True
  End If
End Sub

Private Sub Form_Activate()
   RatonNormal
   NuevoDiario = False
   IniciarAsientosDe DGAsiento, AdoAsiento
   Mifecha = BuscarFecha(FechaSistema)
   MayorizarInv.Show 1
   RatonNormal
   MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FNotaCredito
   ConectarAdodc AdoAux
   ConectarAdodc AdoSQL
   ConectarAdodc AdoCxC
   ConectarAdodc AdoInv
   ConectarAdodc AdoInv1
   ConectarAdodc AdoVentas
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoClientes
   ConectarAdodc AdoVentaAct
   ConectarAdodc AdoFactAnul
   ConectarAdodc AdoAnticipos
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF.Text = MBFechaI.Text
  LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI.Text) & " al " & FechaStrgDias(MBFechaF.Text)
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyF10 Then
     FechaIni = BuscarFecha(MBFechaI.Text)
     FechaFin = BuscarFecha(MBFechaF.Text)
     sSQL = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND (Total_MN * 0.12) <> IVA AND TC <> 'C' AND TC <> 'P' "
    SelectAdodc AdoVentas, sSQL
  End If
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
  LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI.Text) & " al " & FechaStrgDias(MBFechaF.Text)
End Sub

Public Sub GrabarAsientosFacturacion(TipoConsulta As String)
Dim VentasDia As Boolean
   Beneficiario = Ninguno
   DGVentas.Visible = False
   DGCxC.Visible = False
   DGInv.Visible = False
   DGAsiento.Visible = False
   FechaValida MBFechaI
   FechaValida MBFechaF
   ErrorInventario = ""
   Trans_No = 96
   VentasDia = False
   RatonReloj
   Combos = Ninguno
   FechaIni = BuscarFecha(MBFechaI.Text)
   FechaFin = BuscarFecha(MBFechaF.Text)
   FechaFinal = BuscarFecha("31/12/" & FechaAnio(MBFechaF.Text))
   If SQL_Server Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Cta_CxP = F.Cta_CxP, T = F.T " _
           & "FROM Trans_Abonos TA,Facturas As F "
   Else
      sSQL = "UPDATE Trans_Abonos TA,Facturas As F " _
           & "SET TA.Cta_CxP = F.Cta_CxP, TA.T = F.T "
   End If
   sSQL = sSQL & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC <> 'C' " _
        & "AND F.TC <> 'P' " _
        & "AND TA.Factura = F.Factura " _
        & "AND TA.Periodo = F.Periodo " _
        & "AND TA.Item = F.Item "
   Conectar_Ado_Execute sSQL
   
   If SQL_Server Then
      sSQL = "UPDATE Detalle_Factura " _
           & "SET T = F.T " _
           & "FROM Detalle_Factura TA,Facturas As F "
   Else
      sSQL = "UPDATE Detalle_Factura TA,Facturas As F " _
           & "SET TA.T = F.T "
   End If
   sSQL = sSQL & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC <> 'C' " _
        & "AND F.TC <> 'P' " _
        & "AND TA.Factura = F.Factura " _
        & "AND TA.Periodo = F.Periodo " _
        & "AND TA.Item = F.Item "
   Conectar_Ado_Execute sSQL
   ContCtas = 0
   sSQL = "SELECT Codigo " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND CxC <> '.' " _
        & "ORDER BY TL,Codigo "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Codigo1 = .Fields("Codigo")
      ' Facturas
        sSQL = "UPDATE Facturas " _
             & "SET Cod_CxC = '" & Codigo1 & "' " _
             & "WHERE Cod_CxC = '" & Ninguno & "' "
        Conectar_Ado_Execute sSQL
      ' Detalle Facturas
        sSQL = "UPDATE Detalle_Factura " _
             & "SET CodigoL = '" & Codigo1 & "' " _
             & "WHERE CodigoL = '" & Ninguno & "' "
        Conectar_Ado_Execute sSQL
    End If
   End With
   sSQL = "SELECT CxC " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND CxC <> '.' " _
        & "GROUP BY CxC "
   SelectAdodc AdoAux, sSQL
   ContCtas = AdoAux.Recordset.RecordCount
   sSQL = "SELECT Cta " _
        & "FROM Trans_Abonos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta <> '.' " _
        & "GROUP BY Cta "
   SelectAdodc AdoSQL, sSQL
   ContCtas = ContCtas + AdoSQL.Recordset.RecordCount
   
   sSQL = "SELECT Cta_Inventario,Cta_Costo_Venta,Cta_Ventas,Cta_Ventas_0,Cta_Venta_Anticipada " _
        & "FROM Catalogo_Productos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "GROUP BY Cta_Inventario,Cta_Costo_Venta,Cta_Ventas,Cta_Ventas_0,Cta_Venta_Anticipada "
   SelectAdodc AdoSQL, sSQL
   
   ContCtas = ContCtas + (AdoSQL.Recordset.RecordCount * 5) + 3
   
   ReDim CtasProc(ContCtas) As CtasAsiento
   For IE = 0 To ContCtas - 1
       CtasProc(IE).Cta = "0"
       CtasProc(IE).Valor = 0
   Next IE
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("CxC")
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT Cta " _
        & "FROM Trans_Abonos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta <> '.' " _
        & "GROUP BY Cta "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta")
          .MoveNext
        Loop
    End If
   End With
   SetearCtasCierre Cta_IVA
   SetearCtasCierre Cta_Desc
   SetearCtasCierre Cta_Desc2
   SetearCtasCierre Cta_Servicio
   With AdoSQL.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta_Inventario")
           SetearCtasCierre .Fields("Cta_Costo_Venta")
           SetearCtasCierre .Fields("Cta_Ventas")
           SetearCtasCierre .Fields("Cta_Ventas_0")
           SetearCtasCierre .Fields("Cta_Venta_Anticipada")
          .MoveNext
        Loop
    End If
   End With
   Total = 0
   Select Case TipoConsulta
     Case Procesado: NuevoDiario = False
     Case Normal:    NuevoDiario = True
   End Select
   IniciarAsientosDe DGAsiento, AdoAsiento
 ' Ventas Anticipadas
    sSQL = "SELECT P.Cta,P.TP,P.Fecha,AVG(P.Pagos) As Valor_Ed " _
        & "FROM Prestamos As P,Detalle_Factura As F " _
        & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & "" & FechaFin & "# " _
        & "AND P.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.T <> 'A' " _
        & "AND P.Pagos > 0 " _
        & "AND P.Cuenta_No = F.CodigoC " _
        & "AND P.Item = F.Item "
    If CheqCajero.value = 1 Then
       sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
       Beneficiario = SinEspaciosIzq(DCBenef.Text)
    End If
    sSQL = sSQL & "GROUP BY P.Cta,P.TP,P.Fecha "
   SelectAdodc AdoVentaAct, sSQL
   Total = 0
  'Asientos de CxC Cheque
   sSQL = "SELECT TA.TP,TA.Fecha,C.Cliente,TA.Factura,TA.Banco,TA.Cheque,TA.Abono,TA.Comprobante,TA.Cta,TA.Cta_CxP " _
        & "FROM Trans_Abonos As TA,Clientes C " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.TP <> 'C' " _
        & "AND TA.TP <> 'P' " _
        & "AND TA.T <> 'A' " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.CodigoC = C.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TA.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL & "ORDER BY TA.TP,TA.Fecha,TA.Cta,TA.Banco,C.Cliente,TA.Factura "
   SelectAdodc AdoCxC, sSQL
   With AdoCxC.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           InsValorCta .Fields("Cta"), .Fields("Abono")
           InsValorCta .Fields("Cta_CxP"), -.Fields("Abono")
           Total = Total + .Fields("Abono")
          .MoveNext
        Loop
    End If
   End With
   LabelCheque.Caption = Format$(Total, "#,##0.00")
   Total = 0
  'Asientos de CxC Efectivo
   sSQL = "SELECT F.TC,F.Fecha,C.Cliente,F.Factura,F.IVA As Total_IVA,F.Descuento,F.Servicio,F.Total_MN,F.Saldo_MN,F.Cta_CxP " _
        & "FROM Facturas F,Clientes C " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC IN ('FA','FM','NV','PV') " _
        & "AND F.T <> 'A' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.CodigoC = C.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL & "ORDER BY F.TC,F.Fecha,F.Cta_CxP,C.Cliente,F.Factura "
   SelectAdodc AdoVentas, sSQL
   With AdoVentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           InsValorCta Cta_IVA, -.Fields("Total_IVA")
           InsValorCta Cta_Desc, .Fields("Descuento")
           InsValorCta Cta_Desc2, .Fields("Descuento2")
           InsValorCta Cta_Servicio, -.Fields("Servicio")
           InsValorCta .Fields("Cta_CxP"), .Fields("Total_MN")
           Total = Total + .Fields("Total_MN")
          .MoveNext
        Loop
    End If
   End With
   LabelAbonos.Caption = Format$(Total, "#,##0.00")
  'Salida de Inventario
   sSQL = "SELECT * " _
        & "FROM Asiento_K " _
        & "WHERE CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' "
   SelectAdodc AdoInv, sSQL

   sSQL = "SELECT DF.TC,DF.Codigo,A.Cta_Inventario,A.Cta_Costo_Venta,A.Cta_Ventas,A.Cta_Ventas_0,A.Cta_Venta_Anticipada," _
        & "DF.Cantidad,DF.Precio,DF.Total,A.Unidad,A.Producto,A.PVP " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As A " _
        & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND DF.Codigo = A.Codigo_Inv " _
        & "AND DF.Item = A.Item " _
        & "AND DF.Periodo = A.Periodo " _
        & "AND A.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.T <> '" & Anulado & "' " _
        & "AND DF.TC IN ('FA','FM','NV','PV') " _
        & "AND DF.Item = '" & NumEmpresa & "' "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND DF.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL & "ORDER BY DF.Codigo,DF.Fecha,DF.Precio "
   SelectAdodc AdoAux, sSQL
   Total = 0
   TotalIngreso = 0
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Entrada = 0
        Combos = Ninguno
        Precio = .Fields("Precio")
        CodigoInv = .Fields("Codigo")
        Producto = .Fields("Producto")
        Unidad = .Fields("Unidad")
        Cta_Inventario = .Fields("Cta_Inventario")
        Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
        Cta_Ventas = .Fields("Cta_Ventas")
        Cta_Ventas_0 = .Fields("Cta_Ventas_0")
        TipoProc = .Fields("TC")
        Cta_Provision = .Fields("Cta_Venta_Anticipada")
        Total = 0
        If CodigoInv = "" Then CodigoInv = Ninguno
        Do While Not .EOF
           If Cta_Ventas <> .Fields("Cta_Ventas") Then
              'Cta_Provision = "."
              If Len(Cta_Provision) > 1 Then
                 SaldoDisp = 0: SaldoCont = 0
                 If AdoVentaAct.Recordset.RecordCount > 0 Then
                    AdoVentaAct.Recordset.MoveFirst
                    AdoVentaAct.Recordset.Find ("Cta = '" & Cta_Ventas & "' ")
                    If Not AdoVentaAct.Recordset.EOF Then
                       Contador = 0
                       Do While Not AdoVentaAct.Recordset.EOF
                          If AdoVentaAct.Recordset.Fields("Cta") = Cta_Ventas Then
                             FechaTexto = AdoVentaAct.Recordset.Fields("Fecha")
                             FechaTexto1 = "31/12/" & FechaAnio(MBFechaF.Text)
                             Select Case AdoVentaAct.Recordset.Fields("TP")
                               Case "MENS": Documento = DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1
                               Case "QUNC": Documento = (DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1) * 2
                               Case "SEMA": Documento = DatePart("ww", FechaTexto1) - DatePart("ww", FechaTexto) - 1
                             End Select
                             'MsgBox "Fecha: " & FechaTexto & vbCrLf & FechaTexto1 & vbCrLf & "Valor Ed = " & AdoVentaAct.Recordset.Fields("Valor_Ed") & vbCrLf & Documento
                             SaldoCont = SaldoCont + Redondear(AdoVentaAct.Recordset.Fields("Valor_Ed") * Documento, 2)
                             Contador = Contador + 1
                          End If
                          AdoVentaAct.Recordset.MoveNext
                       Loop
                    End If
                 End If
                 'MsgBox "Total(" & Contador & ") = " & Total & vbCrLf & SaldoCont
                 InsValorCta Cta_Provision, -SaldoCont
                 If SaldoCont > 0 Then Total = Total - SaldoCont
              End If
              If TipoProc = "NV" Then Cta_Ventas = Cta_Ventas_0
              'MsgBox Cta_Ventas & vbCrLf & Total
              InsValorCta Cta_Ventas, -Total
              Cta_Ventas = .Fields("Cta_Ventas")
              Cta_Ventas_0 = .Fields("Cta_Ventas_0")
              Cta_Provision = .Fields("Cta_Venta_Anticipada")
              Total = 0
           End If
           If Cta_Inventario <> .Fields("Cta_Inventario") _
              Or CodigoInv <> .Fields("Codigo") Then
              EgresosArtInv
              InsValorCta Cta_Inventario, -ValorTotal
              InsValorCta Cta_Costo_Ventas, ValorTotal
              Combos = Ninguno
              Codigo = .Fields("Codigo")
              Precio = .Fields("Precio")
              CodigoInv = .Fields("Codigo")
              Producto = .Fields("Producto")
              Unidad = .Fields("Unidad")
              Cta_Inventario = .Fields("Cta_Inventario")
              Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
              If Codigo = "" Then Codigo = Ninguno
              If CodigoInv = "" Then CodigoInv = Ninguno
              Entrada = 0
           End If
           'Total = Total + (.Fields("Cantidad") * .Fields("Precio"))
           Total = Total + .Fields("Total")
           Entrada = Entrada + .Fields("Cantidad")
           TipoProc = .Fields("TC")
           TotalIngreso = TotalIngreso + .Fields("Total")
          .MoveNext
        Loop
        If Cta_Inventario <> Ninguno Then EgresosArtInv
        InsValorCta Cta_Inventario, -ValorTotal
        InsValorCta Cta_Costo_Ventas, ValorTotal
        'Cta_Provision = "."
        If Len(Cta_Provision) > 1 Then
           SaldoDisp = 0: SaldoCont = 0
           If AdoVentaAct.Recordset.RecordCount > 0 Then
              AdoVentaAct.Recordset.MoveFirst
              AdoVentaAct.Recordset.Find ("Cta = '" & Cta_Ventas & "' ")
              If Not AdoVentaAct.Recordset.EOF Then
                 Contador = 0
                 Do While Not AdoVentaAct.Recordset.EOF
                    If AdoVentaAct.Recordset.Fields("Cta") = Cta_Ventas Then
                       FechaTexto = AdoVentaAct.Recordset.Fields("Fecha")
                       FechaTexto1 = "31/12/" & FechaAnio(MBFechaF.Text)
                       Select Case AdoVentaAct.Recordset.Fields("TP")
                         Case "MENS": Documento = DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1
                         Case "QUNC": Documento = (DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1) * 2
                         Case "SEMA": Documento = DatePart("ww", FechaTexto1) - DatePart("ww", FechaTexto) - 1
                       End Select
                       'MsgBox "Fecha: " & FechaTexto & vbCrLf & FechaTexto1 & vbCrLf & "Valor Ed = " & AdoVentaAct.Recordset.Fields("Valor_Ed") & vbCrLf & Documento
                       SaldoCont = SaldoCont + Redondear(AdoVentaAct.Recordset.Fields("Valor_Ed") * Documento, 2)
                       'MsgBox Cta_Ventas & vbCrLf & SaldoCont
                       Contador = Contador + 1
                    End If
                    AdoVentaAct.Recordset.MoveNext
                 Loop
              End If
           End If
           'MsgBox "Total(" & Contador & ") = " & Total & vbCrLf & SaldoCont
           InsValorCta Cta_Provision, -SaldoCont
           If SaldoCont > 0 Then Total = Total - SaldoCont
        End If
        If TipoProc = "NV" Then Cta_Ventas = Cta_Ventas_0
        InsValorCta Cta_Ventas, -Total
    End If
  End With
  sSQL = "SELECT C.Codigo_Inv,DF.TC,DF.Codigo,A.Cta_Inventario,A.Cta_Costo_Venta,A.Cta_Ventas,A.Cta_Ventas_0,A.Cta_Venta_Anticipada," _
       & "DF.Cantidad,DF.Precio,DF.Total,A.Unidad,A.Producto,A.PVP " _
       & "FROM Detalle_Factura As DF,Catalogo_Combos As C,Catalogo_Productos As A " _
       & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND A.Periodo = '" & Periodo_Contable & "' " _
       & "AND A.Periodo = DF.Periodo " _
       & "AND A.Periodo = C.Periodo " _
       & "AND DF.Codigo = C.Codigo_Cmb " _
       & "AND A.Codigo_Inv = C.Codigo_Inv " _
       & "AND DF.Item = A.Item " _
       & "AND DF.Item = C.Item " _
       & "AND DF.T <> '" & Anulado & "' " _
       & "AND DF.TC IN ('FA','FM','NV','PV') " _
       & "AND DF.Item = '" & NumEmpresa & "' "
  If CheqCajero.value = 1 Then sSQL = sSQL & "AND DF.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
  sSQL = sSQL & "ORDER BY DF.Codigo,C.Codigo_Inv,DF.Fecha,DF.Precio "
  SelectAdodc AdoAux, sSQL
  Total = 0
  TotalIngreso = 0
  'MsgBox "."
  Entrada = 0
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CodigoInv = .Fields("Codigo_Inv")
       Cta_Inventario = .Fields("Cta_Inventario")
       Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
       Producto = .Fields("Producto")
       Do While Not .EOF
          If CodigoInv <> .Fields("Codigo_Inv") Then
             EgresosArtInv
             InsValorCta Cta_Inventario, -ValorTotal
             InsValorCta Cta_Costo_Ventas, ValorTotal
             Entrada = 0
             CodigoInv = .Fields("Codigo_Inv")
             Cta_Inventario = .Fields("Cta_Inventario")
             Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
             Producto = .Fields("Producto")
          End If
          Entrada = Entrada + .Fields("Cantidad")
          'MsgBox CodigoInv & vbCrLf & Entrada
         .MoveNext
       Loop
       EgresosArtInv
   End If
  End With
  RatonNormal
  If ErrorInventario <> "" Then
     MsgBox "Warning: " & vbCrLf _
            & Space(10) & "Falta de Ingresar codigo(s):" & vbCrLf _
            & Space(10) & ErrorInventario & vbCrLf _
            & Space(10) & "La entrada inicial"
  End If
  For IE = 0 To ContCtas - 1
   If CtasProc(IE).Valor >= 0 Then
      InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, CtasProc(IE).Valor, 0
   Else
      InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, 0, -CtasProc(IE).Valor
   End If
  Next IE
 'Verificacion SubTotal
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelVentas.Caption = Format$(TotalIngreso, "#,##0.00")
  LabelDebe.Caption = Format$(Debe, "#,##0.00")
  LabelHaber.Caption = Format$(Haber, "#,##0.00")
  LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
  If MBFechaI.Text = MBFechaF.Text Then
     LblConcepto.Caption = "Cierre Diario de Caja del " & MBFechaI.Text & ", Diario No. ?"
  Else
     LblConcepto.Caption = "Cierre Diario de Caja del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. ?"
  End If
 'Listado de Facturas anuladas
  sSQL = "SELECT F.T,F.TC,F.Fecha,C.Cliente,F.Factura,F.IVA As Total_IVA,F.Total_MN,F.Cta_CxP " _
       & "FROM Facturas F,Clientes C " _
       & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND F.T = 'A' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = C.Codigo "
  If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
  sSQL = sSQL & "ORDER BY F.TC,F.Fecha,F.Cta_CxP,C.Cliente,F.Factura "
  SelectAdodc AdoFactAnul, sSQL
  DGVentas.Visible = True
  DGCxC.Visible = True
  DGInv.Visible = True
  DGAsiento.Visible = True
End Sub

Public Sub EgresosArtInv()
 'Fecha <= #" & BuscarFecha(MBFechaF.Text) & "#
 If Len(Cta_Inventario) > 1 Then
    If CodigoInv <> Ninguno Then
       ValorUnit = 0: Total_Desc = 0: Saldo = 0
       'If AdoBodega.Recordset.RecordCount > 1 Then
       '   sSQL = "SELECT TOP 1 Codigo_Inv,Costo_Bod As V_Unit,Existencia,Total,T "
       'Else
          sSQL = "SELECT TOP 1 Codigo_Inv,Costo As V_Unit,Existencia,Total,T "
       'End If
       sSQL = sSQL & "FROM Trans_Kardex " _
            & "WHERE Fecha <= #" & BuscarFecha(MBFechaF.Text) & "# " _
            & "AND Codigo_Inv = '" & CodigoInv & "' " _
            & "AND T <> 'A' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND CodBodega = '" & Cod_Bodega & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "ORDER BY Fecha DESC,TP DESC, Numero DESC,ID DESC "
      'MsgBox sSQL
       SelectData AdoSQL, sSQL
       If AdoSQL.Recordset.RecordCount > 0 Then ValorUnit = Redondear(AdoSQL.Recordset.Fields("V_Unit"), 4)
      'Precio = ValorUnit
       If ValorUnit <= 0 Then
          ErrorInventario = ErrorInventario & CodigoInv & vbCrLf & Space(10)
       Else
          ValorTotal = Redondear(ValorUnit * Entrada, 4)
         'Llenamos el ultimo saldo del kardex
         'If Entrada > 0 And Precio > 0 Then
          Cantidad = Cantidad - Entrada
          Saldo = Redondear(SaldoAnterior - ValorTotal)
          SetAddNew AdoInv
          SetFields AdoInv, "DH", 1
          SetFields AdoInv, "CODIGO_INV", CodigoInv
          SetFields AdoInv, "P_DESC", 0
          SetFields AdoInv, "PRODUCTO", Producto
          SetFields AdoInv, "CANT_ES", Entrada
          SetFields AdoInv, "VALOR_UNIT", ValorUnit
          SetFields AdoInv, "VALOR_TOTAL", ValorTotal
          SetFields AdoInv, "CTA_INVENTARIO", Cta_Inventario
          SetFields AdoInv, "CONTRA_CTA", Cta_Costo_Ventas
          SetFields AdoInv, "CANTIDAD", 0
          SetFields AdoInv, "SALDO", 0
          SetFields AdoInv, "CodigoU", CodigoUsuario
          SetFields AdoInv, "T_No", Trans_No
          SetFields AdoInv, "Item", NumEmpresa
          SetUpdate AdoInv
       End If
    End If
 End If
End Sub

Public Sub SetearCtasCierre(CtaFields As String)
  Si_No = True
  For IE = 0 To ContCtas - 1
      If CtaFields = CtasProc(IE).Cta Then Si_No = False
  Next IE
  If Si_No Then
     IE = 0
     While IE < ContCtas
        If CtasProc(IE).Cta = "0" Then
           CtasProc(IE).Cta = CtaFields
           IE = ContCtas + 1
        End If
        IE = IE + 1
     Wend
  End If
End Sub

Public Sub InsValorCta(NCta As String, _
                       NValor As Currency)
  For IE = 0 To ContCtas - 1
      If CtasProc(IE).Cta = NCta Then
         CtasProc(IE).Valor = CtasProc(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

