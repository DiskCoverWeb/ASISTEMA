VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CtasComisiones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FCtasCom.frx":0000
      Height          =   4950
      Left            =   105
      TabIndex        =   19
      Top             =   1785
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   8731
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
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   105
      TabIndex        =   14
      Top             =   735
      Width           =   7680
      Begin VB.OptionButton OpcTodos 
         Caption         =   "Todods"
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
         TabIndex        =   20
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton OpcE 
         Caption         =   "Ejecutivo"
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
         TabIndex        =   17
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton OpcR 
         Caption         =   "Responsable"
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
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "FCtasCom.frx":0017
         DataSource      =   "AdoCliente"
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DEL RESPONSABLE/EJECUTIVO"
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
         Left            =   1680
         TabIndex        =   18
         Top             =   210
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   8925
      Picture         =   "FCtasCom.frx":0030
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
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
      Left            =   7875
      Picture         =   "FCtasCom.frx":08B2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
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
      Left            =   9975
      Picture         =   "FCtasCom.frx":0CF4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Suscripcion"
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
      Left            =   2835
      TabIndex        =   4
      Top             =   0
      Width           =   8310
      Begin VB.OptionButton OpcSusp 
         Caption         =   "Suspensas"
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
         Left            =   5985
         TabIndex        =   11
         Top             =   210
         Width           =   1290
      End
      Begin VB.OptionButton OpcRenov 
         Caption         =   "Renovaciones"
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
         Left            =   4305
         TabIndex        =   10
         Top             =   210
         Width           =   1605
      End
      Begin VB.OptionButton OpcTodas 
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
         Height          =   375
         Left            =   7350
         TabIndex        =   8
         Top             =   210
         Width           =   885
      End
      Begin VB.OptionButton OpcAnul 
         Caption         =   "Anuladas"
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
         Left            =   1575
         TabIndex        =   6
         Top             =   210
         Width           =   1200
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "Nueva / Cancelada"
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
         Left            =   2835
         TabIndex        =   7
         Top             =   210
         Width           =   1395
      End
      Begin VB.OptionButton OpcPend 
         Caption         =   "Pendientes"
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
         Left            =   210
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1470
      TabIndex        =   3
      Top             =   315
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
      TabIndex        =   2
      Top             =   315
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
   Begin MSAdodcLib.Adodc AdoCli 
      Height          =   330
      Left            =   315
      Top             =   3465
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
      Caption         =   "Cli"
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
      Top             =   3150
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2835
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
      Left            =   105
      Top             =   6825
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
      Caption         =   "Query"
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
      Caption         =   " Hasta:"
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
      Left            =   1470
      TabIndex        =   1
      Top             =   105
      Width           =   1275
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
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "CtasComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
Codigo = SinEspaciosIzq(DCCliente.Text)
DGQuery.Caption = "LISTADO DE FACTURAS"
RatonReloj
If OpcCanc.value Then
   sSQL = "SELECT F.ME,F.T,Cl.Cliente As Ejecutivo,F.Fecha_C As Fecha,F.Factura,"
Else
   sSQL = "SELECT F.ME,F.T,Cl.Cliente As Ejecutivo,F.Fecha,F.Factura,"
End If
If OpcR.value Then
   sSQL = sSQL & "F.SubTotal As Monto,F.Porc_C,(F.SubTotal * F.Porc_C/100) As Total_Comision "
   sSQL = sSQL & "FROM Facturas As F,Clientes As Cl "
   If OpcCanc.value Then
      sSQL = sSQL & "WHERE F.Fecha_C BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
   Else
      sSQL = sSQL & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
   End If
   sSQL = sSQL & "AND F.Cod_Ejec = Cl.Codigo "
Else
   sSQL = sSQL & "DF.Total As Monto,DF.Porc_C, (DF.Total * DF.Porc_C/100) As Total_Comision "
   sSQL = sSQL & "FROM Facturas As F,Detalle_Factura AS DF,Clientes As Cl "
   If OpcCanc.value Then
      sSQL = sSQL & "WHERE F.Fecha_C BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
   Else
      sSQL = sSQL & "WHERE F.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
   End If
   sSQL = sSQL & "AND DF.Cod_Ejec = Cl.Codigo "
   sSQL = sSQL & "AND F.Factura = DF.Factura_No "
End If
If CheckTodos.value = 0 Then sSQL = sSQL & "AND Cl.Codigo = '" & Codigo & "' "
If OpcPend Then sSQL = sSQL & "AND F.T = '" & Pendiente & "' "
If OpcAnul Then sSQL = sSQL & "AND F.T = '" & Anulado & "' "
If OpcCanc Then sSQL = sSQL & "AND F.T = '" & Cancelado & "' "
If OpcCanc.value Then
   sSQL = sSQL & "ORDER BY Cl.Cliente,F.Fecha_C,F.Factura "
Else
   sSQL = sSQL & "ORDER BY Cl.Cliente,F.Fecha,F.Factura "
End If
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
Total = 0: Saldo = 0
Do While Not AdoQuery.Recordset.EOF
   Total = Total + AdoQuery.Recordset.fields("Monto")
   Saldo = Saldo + AdoQuery.Recordset.fields("Total_Comision")
   AdoQuery.Recordset.MoveNext
Loop
LabelFacturado.Caption = Format$(Total, "#,##0.00")
LabelAbonado.Caption = Format$(Saldo, "#,##0.00")
RatonNormal
End Sub

Private Sub Command2_Click()
   DGQuery.Visible = False
   If OpcPend Then SQLMsg1 = "LISTADO DE FACTURAS PENDIENTES"
   If OpcAnul Then SQLMsg1 = "LISTADO DE FACTURAS ANULADAS"
   If OpcCanc Then SQLMsg1 = "LISTADO DE FACTURAS CANCELADAS"
   If OpcTodas Then SQLMsg1 = "LISTADO DE TODAS LAS FACTURAS"
   Mifecha = MBoxFechaF.Text
   'ImprimirData AdoQuery, True, 1, 8
   DGQuery.Visible = True
End Sub

Private Sub Command3_Click()
  Unload CtasComisiones
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto CtasComisiones, AdoQuery
End Sub

Private Sub Form_Activate()
   'CTAsientos_F
   sSQL = "SELECT Codigo & '   ' & Cliente As Ejecutivos " _
        & "FROM Clientes " _
        & "WHERE T= 'R' " _
        & "ORDER BY Cliente "
   SelectDB_Combo DCCliente, AdoCliente, sSQL, "Ejecutivos", False
   RatonNormal
   MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm CtasComisiones
  ConectarAdodc AdoAux
  ConectarAdodc AdoCli
  ConectarAdodc AdoQuery
  ConectarAdodc AdoCliente
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF, False
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI, False
End Sub

Private Sub OpcE_Click()
  sSQL = "SELECT Codigo & '   ' & Cliente As Ejecutivos "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE T = 'E' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Ejecutivos", False
  RatonNormal
End Sub

Private Sub OpcR_Click()
  sSQL = "SELECT Codigo & '   ' & Cliente As Ejecutivos "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE T = 'R' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Ejecutivos", False
  RatonNormal
End Sub

