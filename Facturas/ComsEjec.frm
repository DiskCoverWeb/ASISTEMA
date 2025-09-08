VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FComisiones 
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11490
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
      Left            =   105
      Picture         =   "ComsEjec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3045
      Width           =   1065
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
      Left            =   105
      Picture         =   "ComsEjec.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3885
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
      Height          =   750
      Left            =   105
      Picture         =   "ComsEjec.frx":12C0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2205
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Por Ejecutivo"
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
      Left            =   105
      Picture         =   "ComsEjec.frx":1956
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1365
      Width           =   1065
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por E&jecutivo"
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
      TabIndex        =   1
      Top             =   0
      Width           =   2340
   End
   Begin VB.OptionButton OpcBusq 
      Caption         =   "Todos los &Ejecutivos"
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
      Left            =   3990
      TabIndex        =   0
      Top             =   0
      Value           =   -1  'True
      Width           =   2130
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
      Height          =   1065
      Left            =   6195
      TabIndex        =   3
      Top             =   210
      Width           =   5160
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
         Height          =   315
         Left            =   105
         TabIndex        =   8
         Top             =   630
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
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         Top             =   630
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
         Height          =   315
         Left            =   2940
         TabIndex        =   10
         Top             =   630
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
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   630
         Width           =   885
      End
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   2940
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
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   840
         TabIndex        =   5
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
         TabIndex        =   4
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
         Left            =   2205
         TabIndex        =   6
         Top             =   210
         Width           =   750
      End
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "ComsEjec.frx":1D98
      DataSource      =   "AdoListCtas"
      Height          =   960
      Left            =   105
      TabIndex        =   2
      Top             =   315
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   1693
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
      Left            =   1995
      Top             =   315
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
      Left            =   1995
      Top             =   630
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
      Left            =   105
      Top             =   630
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
      Left            =   3990
      Top             =   630
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
      Left            =   3990
      Top             =   315
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
      Left            =   105
      Top             =   315
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
      Bindings        =   "ComsEjec.frx":1DB2
      Height          =   5265
      Left            =   1260
      TabIndex        =   12
      Top             =   1365
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9287
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
      Left            =   1260
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   6720
      Width           =   1800
   End
End
Attribute VB_Name = "FComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub TipoConsultaCxC(Tipo As String)
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
DGQuery.Caption = "LISTADO DE FACTURAS"
RatonReloj
SQL1 = "UPDATE Facturas " _
     & "SET Comision = 0 " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TC NOT IN ('C','P') " _
     & "AND Comision <> 0 "
Ejecutar_SQL_SP SQL1

SQL1 = "UPDATE Detalle_Factura " _
     & "SET Comision = 0 " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TC NOT IN ('C','P') " _
     & "AND Comision <> 0 "
Ejecutar_SQL_SP SQL1

SQL1 = "UPDATE Facturas " _
     & "SET Comision=(Porc_C*SubTotal) " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND TC NOT IN ('C','P') " _
     & "AND Porc_C > 0 "
Ejecutar_SQL_SP SQL1

SQL1 = "UPDATE Detalle_Factura " _
     & "SET Comision=(Porc_C*Total) " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND TC NOT IN ('C','P') " _
     & "AND Porc_C > 0 "
Ejecutar_SQL_SP SQL1

SQL1 = "SELECT CR.Ejecutivo,F.Fecha,F.Factura,C.Cliente,F.SubTotal As Total,F.Porc_C," _
     & "F.Comision,F.T " _
     & "FROM Facturas As F,Clientes As C,Catalogo_Rol_Pagos As CR " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND F.Cod_Ejec = CR.Codigo " _
     & "AND CR.Item = F.Item " _
     & "AND CR.Periodo = F.Periodo " _
     & "AND F.TC NOT IN ('C','P') " _
     & "AND F.Porc_C > 0 "
If Option2.value Then SQL1 = SQL1 & "AND F.Cod_Ejec = '" & CodigoCli & "' "
If OpcPend.value Then SQL1 = SQL1 & "AND F.T = '" & Pendiente & "' "
If OpcAnul.value Then SQL1 = SQL1 & "AND F.T = '" & Anulado & "' "
If OpcCanc.value Then SQL1 = SQL1 & "AND F.T = '" & Cancelado & "' "
SQL2 = "SELECT CR.Ejecutivo,F.Fecha,F.Factura,C.Cliente,F.Total,F.Porc_C," _
     & "F.Comision,F.T " _
     & "FROM Detalle_Factura As F,Clientes As C,Catalogo_Rol_Pagos As CR " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND F.Cod_Ejec = CR.Codigo " _
     & "AND CR.Item = F.Item " _
     & "AND F.TC NOT IN('C','P') " _
     & "AND F.Porc_C > 0 "
If Option2.value Then SQL2 = SQL2 & "AND F.Cod_Ejec = '" & CodigoCli & "' "
If OpcPend.value Then SQL2 = SQL2 & "AND F.T = '" & Pendiente & "' "
If OpcAnul.value Then SQL2 = SQL2 & "AND F.T = '" & Anulado & "' "
If OpcCanc.value Then SQL2 = SQL2 & "AND F.T = '" & Cancelado & "' "
sSQL = SQL1 & " UNION " & SQL2
If Tipo = "C" Then sSQL = sSQL & "ORDER BY CR.Ejecutivo,F.Fecha,F.Factura "
If Tipo = "F" Then sSQL = sSQL & "ORDER BY F.Factura,F.Fecha "
'MsgBox sSQL
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
Total = 0: Saldo = 0
DGQuery.Visible = False
With AdoQuery.Recordset
 If .RecordCount > 0 Then
  .MoveFirst
  Do While Not .EOF
    Total = Total + .fields("Total")
     Saldo = Saldo + .fields("Comision")
    .MoveNext
  Loop
  .MoveFirst
  End If
End With
DGQuery.Visible = True
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Saldo, "#,##0.00")
RatonNormal
End Sub

Public Sub ListarClientes(Optional LlenarCliente As Boolean)
  sSQL = "SELECT CR.Codigo,C.Cliente,C.CI_RUC,C.Porc_C " _
        & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
        & "WHERE CR.Item = '" & NumEmpresa & "' " _
        & "AND CR.Periodo = '" & Periodo_Contable & "' " _
        & "AND C.Codigo = CR.Codigo " _
        & "ORDER BY C.Cliente "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "'")
       If Not .EOF Then
          'MsgBox .Fields("Cliente")
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
  'MsgBox "........."
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
   'If TipoDoc = "C" Then
   ImprimirAdo AdoQuery, True, 1, 9
   'If TipoDoc = "F" Then ImprimirResumenCartera AdoQuery, Codigo4
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
     sSQL = "SELECT * " _
          & "FROM Clientes " _
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
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto FComisiones, AdoQuery
     DGQuery.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  FComisiones.Caption = "CREACION DEL CLIENTE"
  If SQL_Server Then
     sSQL = "UPDATE Catalogo_Rol_Pagos " _
          & "SET Ejecutivo = P.Cliente " _
          & "FROM Catalogo_Rol_Pagos As TS,Clientes As P "
  Else
     sSQL = "UPDATE Catalogo_Rol_Pagos As TS,Clientes As P " _
          & "SET TS.Ejecutivo = P.Cliente "
  End If
  sSQL = sSQL & "WHERE TS.Item = '" & NumEmpresa & "' " _
       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
       & "AND TS.Codigo = P.Codigo "
  Ejecutar_SQL_SP sSQL
  
  ListarClientes CliFact
  DCCliente.SetFocus
  RatonNormal
  FComisiones.WindowState = vbMaximized
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
  FClientes.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   CentrarForm FComisiones
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

