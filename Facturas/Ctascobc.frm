VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CtasCobC 
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   17865
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Ctascobc.frx":0000
      Height          =   7470
      Left            =   1470
      TabIndex        =   9
      Top             =   105
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   13176
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
   Begin VB.Frame Frame1 
      Caption         =   " Saldo: "
      Height          =   960
      Left            =   105
      TabIndex        =   14
      Top             =   3255
      Width           =   1275
      Begin VB.OptionButton OpcHistorico 
         Caption         =   "Historico"
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
         TabIndex        =   16
         Top             =   525
         Width           =   1080
      End
      Begin VB.OptionButton OpcActual 
         Caption         =   "Actual"
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
         TabIndex        =   15
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
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
      TabIndex        =   4
      Top             =   1575
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
      Height          =   210
      Left            =   105
      TabIndex        =   6
      Top             =   2205
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
      Height          =   210
      Left            =   105
      TabIndex        =   5
      Top             =   1890
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
      Height          =   210
      Left            =   105
      TabIndex        =   7
      Top             =   2520
      Width           =   885
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
      Left            =   105
      Picture         =   "Ctascobc.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7140
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clientes"
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
      Picture         =   "Ctascobc.frx":0A0D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4305
      Width           =   1275
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
      Left            =   105
      Picture         =   "Ctascobc.frx":0E4F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6195
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Facturas"
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
      Picture         =   "Ctascobc.frx":1719
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5250
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   1470
      Top             =   7665
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
   Begin MSAdodcLib.Adodc AdoCiudad 
      Height          =   330
      Left            =   2520
      Top             =   3885
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
      Caption         =   "Ciudad"
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
      Left            =   2520
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
   Begin MSAdodcLib.Adodc AdoNiveles 
      Height          =   330
      Left            =   2520
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
      Caption         =   "Niveles"
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
      TabIndex        =   3
      Top             =   1155
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
      TabIndex        =   1
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "Ctascobc.frx":1B5B
      DataSource      =   "AdoTipo"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   2835
      Width           =   1275
      _ExtentX        =   2249
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   2520
      Top             =   4200
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   12915
      TabIndex        =   21
      Top             =   7665
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total CxC"
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
      Left            =   11865
      TabIndex        =   22
      Top             =   7665
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9975
      TabIndex        =   17
      Top             =   7665
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Abonos"
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
      Left            =   8610
      TabIndex        =   18
      Top             =   7665
      Width           =   1380
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   6720
      TabIndex        =   19
      Top             =   7665
      Width           =   1800
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturas"
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
      TabIndex        =   20
      Top             =   7665
      Width           =   1380
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
      Width           =   1275
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
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "CtasCobC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  TipoDoc = "C"
  FechaValida MBFechaI
  FechaValida MBFechaF
  TipoFactura = DCTipo.Text
  If TipoFactura = "" Then TipoFactura = Ninguno
  FA.Fecha_Corte = MBFechaF
  Actualizar_Abonos_Facturas FA
  TipoConsultaCxC TipoDoc
  Totales_CxC_Abonos
  Opcion = 1
End Sub

Private Sub Command3_Click()
  Unload CtasCobC
End Sub

Private Sub Command4_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBFechaF.Text
   If TipoDoc = "C" Then ImprimirCtasCob AdoQuery, sSQL, True
   If TipoDoc = "F" Then ImprimirResumenCartera AdoQuery, Codigo4
   DGQuery.Visible = True
End Sub

Private Sub Command5_Click()
  TipoDoc = "F"
  FechaValida MBFechaI
  FechaValida MBFechaF
  TipoFactura = DCTipo.Text
  If TipoFactura = "" Then TipoFactura = Ninguno
  Actualizar_Abonos_Facturas FA,true MBFechaF
  TipoConsultaCxC TipoDoc
  Totales_CxC_Abonos
  Opcion = 2
End Sub

Private Sub DCTipo_LostFocus()
  TipoFactura = DCTipo.Text
  If TipoFactura = "" Then TipoFactura = Ninguno
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto CtasCobC, AdoQuery
     DGQuery.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyB Then BuscarDatos DGQuery, AdoQuery
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC DESC "
  SelectDBCombo DCTipo, AdoTipo, sSQL, "TC"
  TipoFactura = DCTipo.Text
  If TipoFactura = "" Then TipoFactura = Ninguno
  RatonNormal
  Opcion = 0
End Sub

Private Sub Form_Load()
  'CentrarForm CtasCobC
  ConectarAdodc AdoTipo
  ConectarAdodc AdoQuery
  ConectarAdodc AdoCiudad
  ConectarAdodc AdoNiveles
  ConectarAdodc AdoCliente
  DGQuery.Height = MDI_Y_Max - DGQuery.Top - 300
  DGQuery.width = MDI_X_Max - DGQuery.Left
  Label1.Top = DGQuery.Top + DGQuery.Height + 50
  Label3.Top = DGQuery.Top + DGQuery.Height + 50
  Label7.Top = DGQuery.Top + DGQuery.Height + 50
  Label8.Top = DGQuery.Top + DGQuery.Height + 50
  Label9.Top = DGQuery.Top + DGQuery.Height + 50
  Label10.Top = DGQuery.Top + DGQuery.Height + 50
  AdoQuery.Top = DGQuery.Top + DGQuery.Height + 50
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

Public Sub TipoConsultaCxC(Tipo As String)
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)
DGQuery.Caption = ""
RatonReloj
TipoFactura = DCTipo.Text
If TipoFactura = "" Then TipoFactura = Ninguno
If Tipo = "C" Then
   sSQL = "SELECT F.T,C.Cliente,F.Fecha,F.Factura,F.Total_MN,F.Total_ME,F.Saldo_MN,F.Saldo_ME," _
        & "C.CI_RUC,C.Telefono,C.Celular,C.FAX,C.Ciudad,C.Direccion,C.Email,"
ElseIf Tipo = "F" Then
   sSQL = "SELECT F.T,F.Fecha,F.Factura,C.Cliente,F.Total_MN As Total," _
        & "(F.Total_MN-F.Saldo_MN) As Abono,F.Saldo_MN As Saldo,C.Telefono,"
End If
If SQL_Server Then
   sSQL = sSQL & "F.Fecha_V,DATEDIFF(day,F.Fecha,'" & BuscarFecha(MBFechaF) & "') As Dias_De_Mora,"
Else
   sSQL = sSQL & "F.Fecha_V,DATEDIFF('d',F.Fecha,#" & BuscarFecha(MBFechaF) & "#) As Dias_De_Mora,"
End If
sSQL = sSQL & "A.Nombre_Completo As Ejecutivo " _
     & "FROM Facturas As F,Clientes As C,Accesos As A " _
     & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND F.Item = '" & NumEmpresa & "' " _
     & "AND C.Codigo = F.CodigoC " _
     & "AND A.Codigo = F.Cod_Ejec " _
     & "AND F.TC NOT IN ('C','P') "
Select Case TipoFactura
  Case "NV": sSQL = sSQL & "AND F.TC = 'NV' "
  Case Else: sSQL = sSQL & "AND F.TC <> 'NV' "
End Select

sSQL = sSQL & "AND F.Periodo = '" & Periodo_Contable & "' "
If OpcPend.value Then sSQL = sSQL & "AND F.T = '" & Pendiente & "' "
If OpcAnul.value Then sSQL = sSQL & "AND F.T = '" & Anulado & "' "
If OpcCanc.value Then sSQL = sSQL & "AND F.T = '" & Cancelado & "' "
If OpcTodas.value Then sSQL = sSQL & "AND F.T <> '" & Anulado & "' "
If Tipo = "C" Then sSQL = sSQL & "ORDER BY C.Cliente,F.Fecha,F.Factura "
If Tipo = "F" Then sSQL = sSQL & "ORDER BY F.Factura,F.Fecha "
SelectDataGrid DGQuery, AdoQuery, sSQL
DGQuery.Visible = False
Total = 0: Saldo = 0
With AdoQuery.Recordset
  Do While Not .EOF
     If Tipo = "C" Then
        Total = Total + .Fields("Total_MN")
        Saldo = Saldo + .Fields("Saldo_MN")
     Else
        Total = Total + .Fields("Total")
        Saldo = Saldo + .Fields("Saldo")
     End If
    .MoveNext
  Loop
End With
DGQuery.Visible = True
'LabelFacturado.Caption = Format$(Total, "#,##0.00")
'LabelAbonado.Caption = Format$(Saldo, "#,##0.00")
RatonNormal
End Sub

Public Sub Totales_CxC_Abonos()
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "SELECT Periodo,SUM(Total_MN) As Total_CxC " _
       & "FROM Facturas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "AND T <> 'A' " _
       & "GROUP BY Periodo "
  SelectAdodc AdoNiveles, sSQL
  If AdoNiveles.Recordset.RecordCount Then
     Label7.Caption = Format$(AdoNiveles.Recordset.Fields("Total_CxC"), "#,##0.00")
  End If
  sSQL = "SELECT Periodo,SUM(Abono) As Total_Abonos " _
       & "FROM Trans_Abonos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> 'A' " _
       & "GROUP BY Periodo "
  SelectAdodc AdoNiveles, sSQL
  If AdoNiveles.Recordset.RecordCount Then
     Label1.Caption = Format$(AdoNiveles.Recordset.Fields("Total_Abonos"), "#,##0.00")
  End If
  Label9.Caption = Format$(CCur(Label7.Caption) - CCur(Label1.Caption), "#,##0.00")
  RatonNormal
End Sub
