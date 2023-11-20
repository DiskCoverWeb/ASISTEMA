VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ListMayorizacion2 
   Caption         =   "Mayor Analitico"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   6090
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSDBCtls.DBCombo DBCCtas 
      Bindings        =   "Mayoriz2.frx":0000
      DataSource      =   "DataCtas"
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cuentas"
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
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   4620
      TabIndex        =   4
      Top             =   420
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
      Left            =   1680
      TabIndex        =   2
      Top             =   420
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
      Height          =   960
      Left            =   10290
      Picture         =   "Mayoriz2.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir Mayor actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   9030
      Picture         =   "Mayoriz2.frx":0299
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   7770
      Picture         =   "Mayoriz2.frx":0903
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   1170
   End
   Begin VB.Data DataTrans 
      Caption         =   "Base de Mayorizacion"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Width           =   11355
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1365
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
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
      Left            =   3045
      TabIndex        =   3
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicial"
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
      TabIndex        =   1
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label LabelTotSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   9450
      TabIndex        =   15
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe - Haber:"
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
      Left            =   7770
      TabIndex        =   14
      Top             =   6510
      Width           =   1695
   End
   Begin VB.Label LabelTotHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   5565
      TabIndex        =   13
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Haber:"
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
      TabIndex        =   12
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label LabelTotDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   2520
      TabIndex        =   11
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe:"
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
      TabIndex        =   10
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHAS (DD/MM/AA):"
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
      Width           =   5895
   End
End
Attribute VB_Name = "ListMayorizacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  Codigo1 = SinEspaciosIzq(DBCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  SumaDebe = 0: SumaHaber = 0
  sSQL = "SELECT ID,Trans.Fecha,Trans.TP,Trans.Numero,Concepto,"
  If OpcCoop = False Then sSQL = sSQL & "(Debe_ME-Haber_ME) As Parcial_ME,"
  sSQL = sSQL & "Debe,Haber,Saldo,Debe_ME,Haber_ME,Saldo_ME "
  sSQL = sSQL & "FROM Transacciones As Trans,Comprobantes As Comp "
  sSQL = sSQL & "WHERE Trans.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  sSQL = sSQL & "AND Cta = '" & Codigo1 & "' "
  sSQL = sSQL & "AND Trans.TP = Comp.TP "
  sSQL = sSQL & "AND Trans.Numero = Comp.Numero "
  sSQL = sSQL & "AND Trans.Fecha = Comp.Fecha "
  sSQL = sSQL & "AND Trans.Item = Comp.Item "
  sSQL = sSQL & "AND Trans.T = '" & Normal & "' "
  sSQL = sSQL & "ORDER BY Trans.Fecha,Trans.TP,Trans.Numero,Debe DESC,Haber,ID "
  SelectDBGrid DBGTrans, DataTrans, sSQL
  DBGTrans.Visible = False
  With DataTrans.Recordset
  If .RecordCount > 0 Then
      ProgBar.Min = 0: ProgBar.Max = .RecordCount: I = 0
      Do While Not .EOF
         ProgBar.Value = I
         SumaDebe = SumaDebe + .Fields("Debe")
         SumaHaber = SumaHaber + .Fields("Haber")
         SaldoTotal = .Fields("Saldo")
         I = I + 1
        .MoveNext
      Loop
  End If
  End With
  DBGTrans.Visible = True
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SaldoTotal, "#,##0.00")
  Cadena = "TRANSACCIONES MAYORIZADAS: Cantidad de Registros: " & Format(DataTrans.Recordset.RecordCount, "#,##0")
  Cadena = Cadena & Space(30) & "Páginas No. " & Format((DataTrans.Recordset.RecordCount / 50) + 1, "#,##0")
  DataTrans.Caption = Cadena
  RatonNormal
  DBCCtas.SetFocus
End Sub

Private Sub Command2_Click()
  RatonReloj
  Codigo1 = SinEspaciosIzq(DBCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  FechaCorte = "Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
  sSQL = "SELECT ID,Cta,Ctas.Cuenta,Trans.Fecha,Trans.TP,Trans.Numero,Concepto,"
  If OpcCoop = False Then sSQL = sSQL & "(Debe_ME-Haber_ME) As Parcial_ME,"
  sSQL = sSQL & "Debe,Haber,Saldo,Debe_ME,Haber_ME,Saldo_ME,Ctas.ME "
  sSQL = sSQL & "FROM Transacciones As Trans,Catalogo As Ctas,Comprobantes As Comp "
  sSQL = sSQL & "WHERE Trans.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  sSQL = sSQL & "AND Cta = '" & Codigo1 & "' "
  sSQL = sSQL & "AND Cta = Ctas.Codigo "
  sSQL = sSQL & "AND Trans.TP = Comp.TP "
  sSQL = sSQL & "AND Trans.Numero = Comp.Numero "
  sSQL = sSQL & "AND Trans.Fecha = Comp.Fecha "
  sSQL = sSQL & "AND Trans.Item = Comp.Item "
  sSQL = sSQL & "AND Trans.T = '" & Normal & "' "
  sSQL = sSQL & "ORDER BY Trans.Fecha,Trans.TP,Trans.Numero,Debe DESC,Haber,ID "
  SelectData DataTrans, sSQL, False
  DBGTrans.Visible = False
  ImprimirMayor DataTrans
  sSQL = "SELECT ID,Trans.Fecha,Trans.TP,Trans.Numero,Concepto,"
  If OpcCoop = False Then sSQL = sSQL & "(Debe_ME-Haber_ME) As Parcial_ME,"
  sSQL = sSQL & "Debe,Haber,Saldo,Debe_ME,Haber_ME,Saldo_ME "
  sSQL = sSQL & "FROM Transacciones As Trans,Comprobantes As Comp "
  sSQL = sSQL & "WHERE Trans.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  sSQL = sSQL & "AND Cta = '" & Codigo1 & "' "
  sSQL = sSQL & "AND Trans.TP = Comp.TP "
  sSQL = sSQL & "AND Trans.Numero = Comp.Numero "
  sSQL = sSQL & "AND Trans.Fecha = Comp.Fecha "
  sSQL = sSQL & "AND Trans.Item = Comp.Item "
  sSQL = sSQL & "AND Trans.T = '" & Normal & "' "
  sSQL = sSQL & "ORDER BY Trans.Fecha,Trans.TP,Trans.Numero,Debe DESC,Haber,ID "
  SelectDBGrid DBGTrans, DataTrans, sSQL
  DBGTrans.Visible = True
  RatonNormal
End Sub

Private Sub Command3_Click()
  Unload ListMayorizacion2
End Sub

Private Sub DBCCtas_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub DBGTrans_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto ListMayorizacion2, DataTrans
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel_3 Or CNivel_6 Then
        Command2.Enabled = False
     End If
  End If
  sSQL = "SELECT Codigo & '" & Space(20) & "' & Cuenta As Nombre_Cta "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE DG = 'D' AND Cuenta <> '" & Ninguno & "' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectData DataCtas, sSQL, False
  DBCCtas.ListField = "Nombre_Cta"
  With DataCtas.Recordset
       If .RecordCount > 0 Then DBCCtas.Text = .Fields("Nombre_Cta")
  End With
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ListMayorizacion2
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub
