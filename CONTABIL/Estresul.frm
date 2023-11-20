VERSION 5.00
Begin VB.Form EstadoResultado 
   Caption         =   "ESTADO DE RESULTADOS"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "TIPO DE PRESENTACION"
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
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   10200
      Begin VB.OptionButton OpcDG 
         Caption         =   "Cuentas de Grupo y de Detalle"
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
         Left            =   210
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   3165
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "Solo Cuentas de Grupo"
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
         Left            =   4200
         TabIndex        =   8
         Top             =   210
         Width           =   2535
      End
      Begin VB.OptionButton OpcD 
         Caption         =   "Solo Cuentas de Detalle"
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
         Left            =   7245
         TabIndex        =   7
         Top             =   210
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir &Total"
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
      Left            =   10395
      Picture         =   "Estresul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
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
      Left            =   10395
      Picture         =   "Estresul.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir &Parcial"
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
      Left            =   10395
      Picture         =   "Estresul.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3150
      Width           =   1065
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
      Left            =   10395
      Picture         =   "Estresul.frx":0D2E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Data DataFechaBal 
      Caption         =   "FechaBal"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Data DataCtas 
      Caption         =   "Balance"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1995
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data DataResultado 
      Caption         =   "Resultado"
      Connect         =   "Access 2000;"
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
      Top             =   6720
      Width           =   10200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   9690
   End
End
Attribute VB_Name = "EstadoResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  DBGBalance.Visible = False
  RatonReloj
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Estado_Resultado "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDBGrid DBGBalance, DataResultado, sSQL
  With DataResultado.Recordset
       Cadena = "Registros: " & Format(.RecordCount, "#,##0")
       Cadena = Cadena & ".   Páginas: " & Format((.RecordCount / 45) + 1, "#,##0") & "."
       DataResultado.Caption = Cadena
  End With
  Opcion = 1
  DBGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command2_Click()
  DBGBalance.Visible = False
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Estado_Resultado "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDBGrid DBGBalance, DataResultado, sSQL
  DBGBalance.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS"
  SQLMsg2 = "AL " & FechaStrg(FechaFin)
  If OpcCoop Then
     ImprimirGeneralCon DataResultado, 1, False
  Else
     ImprimirGeneral DataResultado, 1
  End If
  DBGBalance.Visible = True
End Sub

Private Sub Command3_Click()
  DBGBalance.Visible = False
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Estado_Resultado "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDBGrid DBGBalance, DataResultado, sSQL
  DBGBalance.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS"
  SQLMsg2 = "AL " & FechaStrg(FechaFin)
  If OpcCoop Then
     ImprimirGeneralCon DataResultado, 2, False
  Else
     ImprimirGeneral DataResultado, 2
  End If
  DBGBalance.Visible = True
End Sub

Private Sub Command4_Click()
  Unload EstadoResultado
End Sub

Private Sub DBGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto EstadoResultado, DataResultado
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * FROM FechaBalance "
  SelectAdodc DataFechaBal, sSQL, False
  If DataFechaBal.Recordset.RecordCount > 0 Then
     FechaIni = DataFechaBal.Recordset.Fields("Fecha_Inicial")
     FechaFin = DataFechaBal.Recordset.Fields("Fecha_Final")
  End If
  Label5.Caption = Empresa & ": ESTADO DE RESULTADOS" & vbCrLf & "AL " & FechaStrg(FechaFin)
  If OpcCoop Then Command3.Visible = False
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm EstadoResultado
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataFechaBal.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataResultado.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

