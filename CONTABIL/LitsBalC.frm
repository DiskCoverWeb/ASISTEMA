VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ListarBalanceComp 
   Caption         =   "BALANCE DE COMPROBACION"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoBalance 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   11460
      _ExtentX        =   20214
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
      Caption         =   "Balance"
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
   Begin MSDataGridLib.DataGrid DGBalance 
      Bindings        =   "LitsBalC.frx":0000
      Height          =   5370
      Left            =   105
      TabIndex        =   15
      Top             =   1260
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   9472
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Left            =   8505
      Picture         =   "LitsBalC.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   960
   End
   Begin VB.CheckBox CheckRango 
      Caption         =   "Rango de Cuentas"
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
      Top             =   840
      Width           =   2010
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
      Left            =   9555
      Picture         =   "LitsBalC.frx":045D
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   10605
      Picture         =   "LitsBalC.frx":0AC7
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxCtaF 
      Height          =   330
      Left            =   6615
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBoxCtaI 
      Height          =   330
      Left            =   3570
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Final"
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
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Inicial"
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
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label5 
      Caption         =   "Balance de Comprobacion"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8310
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
      Left            =   9765
      TabIndex        =   13
      Top             =   7035
      Width           =   1800
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
      Left            =   8085
      TabIndex        =   12
      Top             =   7035
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
      Left            =   6195
      TabIndex        =   11
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Haber:"
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
      TabIndex        =   10
      Top             =   7035
      Width           =   1485
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
      Left            =   2835
      TabIndex        =   9
      Top             =   7035
      Width           =   1800
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Debe:"
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
      Left            =   1470
      TabIndex        =   8
      Top             =   7035
      Width           =   1380
   End
End
Attribute VB_Name = "ListarBalanceComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  'DBGBalance.Visible = False
  Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
  Codigo2 = CambioCodigoCta(MBoxCtaF.Text)
  sSQL = "SELECT * FROM FechaBalance "
  SelectAdodc AdoBalance, sSQL
  AdodcBalance.Recordset.MoveFirst
  FechaIni = AdodcBalance.Recordset.Fields("Fecha_Inicial")
  FechaFin = AdodcBalance.Recordset.Fields("Fecha_Final")
  Label5.Caption = "Balance de Comprobacion del " & FechaStrgCorta(FechaIni) & " al " & FechaStrgCorta(FechaFin)
  sSQL = "SELECT DG,Codigo,Cuenta, Saldo_Anterior, Debitos, Creditos, Saldo_Total " _
       & "FROM Catalogo " _
       & "WHERE (Debitos<>0 OR Creditos<>0 OR Saldo_Total<>0) "
  If CheckRango.Value = 1 Then sSQL = sSQL & "AND Codigo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDataGrid DGBalance, AdoBalance, sSQL
  DataGBalance.Visible = False
  SumaDebe = 0: SumaHaber = 0
  With AdodcBalance.Recordset
    If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("Debitos")
          SumaHaber = SumaHaber + .Fields("Creditos")
         .MoveNext
       Loop
    End If
  End With
  DataGBalance.Visible = True
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
End Sub

Private Sub Command2_Click()
  DBGBalance.Visible = False
  RatonReloj
  ImprimirBalance AdoBalance
  RatonNormal
  DBGBalance.Visible = True
End Sub

Private Sub Command3_Click()
  Unload ListarBalanceComp
End Sub

Private Sub CheckRango_Click()
 If CheckRango.Value = 1 Then Si_No = True Else Si_No = False
 Label3.Visible = Si_No
 Label4.Visible = Si_No
 MBoxCtaI.Visible = Si_No
 MBoxCtaF.Visible = Si_No
 If Si_No Then MBoxCtaI.SetFocus
End Sub

Private Sub DBGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto ListarBalanceComp, AdoBalance
End Sub

Private Sub Form_Activate()
  FormatoMaskCta MBoxCtaI
  FormatoMaskCta MBoxCtaF
  ListarBalanceComp.Caption = "BALANCE DE COMPROBACION"
  RatonNormal
'  sSQL = "SELECT * FROM Tipos_Ado "
'  SelectAdodc AdodcBalance, sSQL
'  Cadena = "Campos y Tipos:" & vbCrLf
'  With AdodcBalance.Recordset
'  For I = 0 To .Fields.Count - 1
'     .AddNew
'     .Fields("CTexto") = FieldType(.Fields(I).Name, .Fields(I).Type)
'     .Update
'  Next I
'  End With
End Sub

Private Sub Form_Load()
  CentrarForm ListarBalanceComp
' DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  ConectarAdodc AdoBalance
End Sub
