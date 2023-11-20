VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BalanceGeneral 
   Caption         =   "BALANCE GENERAL"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11670
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
      Height          =   960
      Left            =   7350
      TabIndex        =   6
      Top             =   0
      Width           =   3165
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
         Left            =   105
         TabIndex        =   2
         Top             =   630
         Width           =   2535
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
         Left            =   105
         TabIndex        =   8
         Top             =   420
         Width           =   2535
      End
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
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   2955
      End
   End
   Begin MSDataGridLib.DataGrid DGBalanceG 
      Bindings        =   "BalanceG.frx":0000
      Height          =   5685
      Left            =   105
      TabIndex        =   9
      Top             =   1050
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   10028
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin MSAdodcLib.Adodc AdoBalanceG 
      Height          =   330
      Left            =   105
      Top             =   6720
      Width           =   10410
      _ExtentX        =   18362
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
      Caption         =   "BalanceG"
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
      Height          =   855
      Left            =   10605
      Picture         =   "BalanceG.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2940
      Width           =   960
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
      Height          =   855
      Left            =   10605
      Picture         =   "BalanceG.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   960
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
      Height          =   855
      Left            =   10605
      Picture         =   "BalanceG.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1995
      Width           =   960
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
      Height          =   855
      Left            =   10605
      Picture         =   "BalanceG.frx":1188
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "X"
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
      Height          =   975
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   7185
   End
End
Attribute VB_Name = "BalanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  DGBalanceG.Visible = False
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Balance_General "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDataGrid DGBalanceG, AdoBalanceG, sSQL
  DGBalanceG.Visible = True
  RatonNormal
End Sub

Private Sub Command2_Click()
  DGBalance.Visible = False
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Balance_General "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDataGrid DGBalanceG, AdoBalanceG, sSQL
  DGBalanceG.Visible = False
  SQLMsg1 = "BALANCE GENERAL"
  SQLMsg2 = "AL " & FechaStrg(FechaFin)
  If OpcCoop Then
     ImprimirGeneralCon AdoBalanceG, 1, True
  Else
     ImprimirGeneral AdoBalanceG, 1
  End If
  DGBalanceG.Visible = True
End Sub

Private Sub Command3_Click()
  DGBalanceG.Visible = False
  If OpcCoop Then
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Balance_General "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDataGrid DGBalanceG, AdoBalanceG, sSQL
  DGBalanceG.Visible = False
  SQLMsg1 = "BALANCE GENERAL"
  SQLMsg2 = "AL " & FechaStrg(FechaFin)
  If OpcCoop Then
     ImprimirGeneralCon AdoBalanceG, 2, True
  Else
     ImprimirGeneral AdoBalanceG, 2
  End If
  DGBalanceG.Visible = True
End Sub

Private Sub Command4_Click()
  Unload BalanceGeneral
End Sub

Private Sub DGBalanceG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto BalanceGeneral, AdoBalanceG
End Sub

Private Sub Form_Activate()
  RatonReloj
  DGBalanceG.Visible = False
  sSQL = "SELECT * FROM Fechas_Balance " _
       & "WHERE Detalle = 'Balance' "
  SelectAdodc AdoBalanceG, sSQL
  If AdoBalanceG.Recordset.RecordCount > 0 Then
     FechaIni = AdoBalanceG.Recordset.Fields("Fecha_Inicial")
     FechaFin = AdoBalanceG.Recordset.Fields("Fecha_Final")
  End If
  Label5.Caption = Empresa & " : BALANCE GENERAL" & vbCrLf & "AL " & FechaStrg(FechaFin)
  If OpcCoop Then
     Command3.Visible = False
     sSQL = "SELECT Codigo,Cuenta,Analitico As Saldo_ME,Parcial As Saldo_MN, TOTAL "
  Else
     sSQL = "SELECT * "
  End If
  sSQL = sSQL & "FROM Balance_General "
  If OpcG.Value Then sSQL = sSQL & "WHERE DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "WHERE DG = 'D' "
  SelectDataGrid DGBalanceG, AdoBalanceG, sSQL
  DGBalanceG.Visible = True
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm BalanceGeneral
  ConectarAdodc AdoBalanceG
End Sub
