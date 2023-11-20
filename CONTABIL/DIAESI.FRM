VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form ListDiarioESI 
   Caption         =   "DIARIO DE ESTADO DE SITUACION INICIAL"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   105
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDBGrid.DBGrid DBGTrans 
      Bindings        =   "DiaESI.frx":0000
      Height          =   5685
      Left            =   105
      OleObjectBlob   =   "DiaESI.frx":0019
      TabIndex        =   0
      Top             =   420
      Width           =   10830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCELAR"
      Height          =   330
      Left            =   9555
      TabIndex        =   6
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR"
      Height          =   330
      Left            =   7875
      TabIndex        =   5
      Top             =   105
      Width           =   1695
   End
   Begin VB.Data DataFechaBal 
      Caption         =   "FechaBal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6510
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ESTADO DE SITUACION INICIAL"
      Height          =   330
      Left            =   4830
      TabIndex        =   7
      Top             =   105
      Width           =   3060
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6825
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataEstInicial 
      Caption         =   "EstInicial"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6090
      Width           =   10830
   End
   Begin VB.Label LabelTotHaber 
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
      Height          =   330
      Left            =   8925
      TabIndex        =   4
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Pasivos:"
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
      Left            =   7140
      TabIndex        =   3
      Top             =   6510
      Width           =   1695
   End
   Begin VB.Label LabelTotDebe 
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
      Height          =   330
      Left            =   4935
      TabIndex        =   2
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Activos:"
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
      Left            =   3150
      TabIndex        =   1
      Top             =   6510
      Width           =   1695
   End
End
Attribute VB_Name = "ListDiarioESI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Unload ListDiarioESI
End Sub

Private Sub Command2_Click()
  DBGTrans.Visible = False
  sSQL = "SELECT * FROM FechaBalance "
  DataFechaBal.RecordSource = sSQL: DataFechaBal.Refresh
  MiFecha = DataFechaBal.Recordset.Fields("Fecha_Final")
  SQLMsg1 = "ESTADO DE SITUACION INICIAL"
  ImprimirBalance DataEstInicial
  DBGTrans.Visible = True
End Sub

Private Sub Command3_Click()
  MousePointer = vbHourglass
  SumaDebe = 0: SumaHaber = 0
  DBGTrans.Visible = False
  MiTiempo1 = Time
  sSQL = "SELECT SI.TC,SI.Codigo,Cta.Cuenta,SI.Activos,SI.Pasivos,SI.TOTALES "
  sSQL = sSQL & "FROM Situacion_Inicial As SI,Cuentas As Cta "
  sSQL = sSQL & "WHERE Cta.Codigo = SI.Codigo "
  sSQL = sSQL & "AND (SI.Activos <> 0 OR SI.Pasivos <> 0  OR TOTALES <> 0) "
  sSQL = sSQL & "ORDER BY SI.Codigo "
  DataEstInicial.RecordSource = sSQL: DataEstInicial.Refresh
  If DataEstInicial.Recordset.RecordCount > 0 Then
  TotalBarra = DataEstInicial.Recordset.RecordCount
    ProgBar.Min = 1: ProgBar.Max = TotalBarra
    TotalBarra = 1
    Do While Not DataEstInicial.Recordset.EOF
       SumaDebe = SumaDebe + DataEstInicial.Recordset.Fields("Activos")
       SumaHaber = SumaHaber + DataEstInicial.Recordset.Fields("Pasivos")
       ProgBar.Value = TotalBarra: TotalBarra = TotalBarra + 1
       DataEstInicial.Recordset.MoveNext
    Loop
  End If
  DBGTrans.Visible = True
  DBGTrans.Columns(DataEstInicial.Recordset.Fields.Count - 3).NumberFormat = "#,##0.00"
  DBGTrans.Columns(DataEstInicial.Recordset.Fields.Count - 2).NumberFormat = "#,##0.00"
  DBGTrans.Columns(DataEstInicial.Recordset.Fields.Count - 1).NumberFormat = "#,##0.00"
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  Cadena = "Cuentas Iniciales: " & Format(DataEstInicial.Recordset.RecordCount, "#,##0")
  Cadena = Cadena & "      Páginas: " & Format((DataEstInicial.Recordset.RecordCount / 50) + 1, "#,##0")
  DataEstInicial.Caption = Cadena
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  'Centrar el formulario
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 - 600
  DataEstInicial.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataFechaBal.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  MDIConta.MousePointer = vbDefault
End Sub

