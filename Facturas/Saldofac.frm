VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ListarFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Saldofac.frx":0000
      Height          =   4320
      Left            =   135
      TabIndex        =   21
      Top             =   1470
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   7620
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
   Begin MSAdodcLib.Adodc AdoCiudad 
      Height          =   375
      Left            =   480
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Height          =   375
      Left            =   480
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   480
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Linea"
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
      Left            =   120
      Top             =   5880
      Width           =   10935
      _ExtentX        =   19288
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
   Begin VB.CheckBox CheqLinea 
      Caption         =   "Por &Linea de Producc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Top             =   1050
      Width           =   2310
   End
   Begin MSDBCtls.DBCombo DBCLinea 
      Bindings        =   "Saldofac.frx":0017
      DataSource      =   "DataLinea"
      Height          =   315
      Left            =   4830
      TabIndex        =   20
      Top             =   1050
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
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
   Begin VB.CheckBox CheqDir 
      Caption         =   "Ordenar por Dirección"
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
      TabIndex        =   19
      Top             =   1155
      Width           =   2325
   End
   Begin VB.CheckBox CheqCiudad 
      Caption         =   "C&onsultar por Ciudad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   18
      Top             =   630
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2310
      TabIndex        =   12
      Top             =   0
      Width           =   5580
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
         Height          =   270
         Left            =   210
         TabIndex        =   16
         Top             =   210
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
         Height          =   270
         Left            =   1680
         TabIndex        =   15
         Top             =   210
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
         Height          =   270
         Left            =   3150
         TabIndex        =   14
         Top             =   210
         Width           =   1200
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
         Height          =   270
         Left            =   4410
         TabIndex        =   13
         Top             =   210
         Width           =   870
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&onsultar"
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
      Left            =   7980
      Picture         =   "Saldofac.frx":002F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
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
      Height          =   855
      Left            =   9030
      Picture         =   "Saldofac.frx":0471
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
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
      Height          =   855
      Left            =   10080
      Picture         =   "Saldofac.frx":0CF3
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   630
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   840
      TabIndex        =   0
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSDBCtls.DBCombo DBCCiudad 
      Bindings        =   "Saldofac.frx":16E9
      DataSource      =   "DataCiudad"
      Height          =   315
      Left            =   4830
      TabIndex        =   17
      Top             =   630
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
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
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " hasta:"
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
      TabIndex        =   5
      Top             =   630
      Width           =   750
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desde:"
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
      Top             =   105
      Width           =   750
   End
   Begin VB.Label LabelAbonado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   9030
      TabIndex        =   10
      Top             =   6300
      Width           =   2010
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total x Cobrar    S/."
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
      Left            =   6720
      TabIndex        =   9
      Top             =   6300
      Width           =   2220
   End
   Begin VB.Label LabelFacturado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   4620
      TabIndex        =   8
      Top             =   6300
      Width           =   2010
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado S/."
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
      Left            =   2310
      TabIndex        =   7
      Top             =   6300
      Width           =   2220
   End
End
Attribute VB_Name = "ListarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqCiudad_Click()
 If CheqCiudad.Value = 1 Then DCCiudad.Visible = True Else DCCiudad.Visible = False
End Sub

Private Sub CheqLinea_Click()
  If CheqLinea.Value = 1 Then DCLinea.Visible = True Else DCLinea.Visible = False
End Sub

Private Sub Command1_Click()
FechaValida MBoxFechaI, False
FechaValida MBoxFechaF, False
FechaIni = BuscarFecha(MBoxFechaI.Text)
FechaFin = BuscarFecha(MBoxFechaF.Text)
DGQuery.Caption = "LISTADO DE FACTURAS"
RatonReloj
sSQL = "SELECT T,C.Cliente,Factura,Fecha,Total_MN As Total,(Total_MN-Saldo_MN) As Abono,Saldo_MN As Saldo,C.Telefono "
sSQL = sSQL & "FROM Facturas As F,Clientes As C "
sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
sSQL = sSQL & "AND C.Codigo = F.Codigo_C "
If OpcPend Then sSQL = sSQL & "AND T = '" & Pendiente & "' "
If OpcAnul Then sSQL = sSQL & "AND T = '" & Anulado & "' "
If OpcCanc Then sSQL = sSQL & "AND T = '" & Cancelado & "' "
If CheqCiudad.Value = 1 Then sSQL = sSQL & "AND UCase(Ciudad) = '" & UCase(DCCiudad.Text) & "' "
If CheqLinea.Value = 1 Then sSQL = sSQL & "AND Cta_CxC = '" & SinEspaciosIzq(DCLinea.Text) & "' "
If CheqDir.Value = 1 Then
   sSQL = sSQL & "ORDER BY Direccion,Factura,Fecha,C.Cliente "
Else
   sSQL = sSQL & "ORDER BY Factura,Fecha,C.Cliente "
End If
SelectDataGrid DGQuery, AdoQuery, sSQL
Total = 0: Saldo = 0
Do While Not AdoQuery.Recordset.EOF
   Total = Total + AdoQuery.Recordset.Fields("Total")
   Saldo = Saldo + AdoQuery.Recordset.Fields("Saldo")
   AdoQuery.Recordset.MoveNext
Loop
Pagina = AdoQuery.Recordset.RecordCount / 25
If AdoQuery.Recordset.RecordCount > 0 And Pagina <= 0 Then Pagina = 1
Cadena = "Facturas No. " & AdoQuery.Recordset.RecordCount
Cadena = Cadena & Space(80) & "Aproximadamente: " & Format(Pagina, "#,##0") & "  Página(s)."
AdoQuery.Caption = Cadena
LabelFacturado.Caption = Format(Total, "#,##0.00")
LabelAbonado.Caption = Format(Saldo, "#,##0.00")
RatonNormal
End Sub

Private Sub Command2_Click()
   DGQuery.Visible = False
   If OpcPend.Value Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcCanc.Value Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcAnul.Value Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcTodas.Value Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   If CheqCiudad.Value = 1 Then SQLMsg1 = SQLMsg1 & " DE " & UCase(DCCiudad.Text)
   MiFecha = MBoxFechaF.Text
   ImprimirResumenCartera AdoQuery, 9
   DGQuery.Visible = True
End Sub

Private Sub Command3_Click()
  Unload ListarFacturas
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto ListarFacturas, AdoQuery
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT Ciudad " _
        & "FROM Clientes " _
        & "GROUP BY Ciudad "
   'SelectDBCombo DCCiudad, AdoCiudad, sSQL, "Ciudad", False
   sSQL = "SELECT Cta_CxC & '     ' & Linea As NomLinea FROM Linea_Producto ORDER BY Linea "
   'SelectDBCombo DCLinea, AdoLinea, sSQL, "NomLinea", False
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm ListarFacturas
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCliente
   ConectarAdodc AdoCiudad
   ConectarAdodc AdoLinea
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub
