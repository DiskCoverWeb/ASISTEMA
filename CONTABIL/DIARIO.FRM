VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ListDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DIARIO GENERAL"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR"
      Height          =   492
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGTrans 
      Bindings        =   "Diario.frx":0000
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "Diario.frx":0014
      TabIndex        =   10
      Top             =   720
      Width           =   10815
   End
   Begin VB.Data DataTrans 
      Caption         =   "Registros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   6240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INICIAR CONSULTA"
      Height          =   492
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox TextFechaAF 
      Height          =   285
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TextFechaMF 
      Height          =   285
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "12"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TextFechaDF 
      Height          =   285
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "31"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TextFechaD 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "01"
      Top             =   360
      Width           =   372
   End
   Begin VB.TextBox TextFechaM 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "01"
      Top             =   360
      Width           =   372
   End
   Begin VB.TextBox TextFechaA 
      Alignment       =   2  'Center
      Height          =   288
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   360
      Width           =   372
   End
   Begin VB.Label LabelTotPag 
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
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label LabelTotalPagina 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Página(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label LabelTotalReg 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Registros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label LabelTotReg 
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
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label LabelTotSaldo 
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
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe - Haber:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   6240
      Width           =   2055
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
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Haber:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   6240
      Width           =   2175
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
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Debe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha final"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHAS DD/MM/AA"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha inicial"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "ListDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  MiTiempo1 = Time
  SumaDebe = 0: SumaHaber = 0
  FechaIni = BuscarFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
  FechaFin = BuscarFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
  MousePointer = vbHourglass
  DBGTrans.Visible = False
  SSQL = "SELECT SUM(Debe),SUM(Haber) FROM Transacciones "
  SSQL = SSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  SSQL = SSQL & "AND T <> 'A' "
  DataTrans.RecordSource = SSQL: DataTrans.Refresh
  If DataTrans.Recordset.RecordCount > 0 Then
     Cadena = DataTrans.Recordset.Fields(0).Name
     SumaDebe = DataTrans.Recordset.Fields(Cadena)
     Cadena = DataTrans.Recordset.Fields(1).Name
     SumaHaber = DataTrans.Recordset.Fields(Cadena)
  End If
  SSQL = "SELECT Trans.Fecha,TP,COMP,Cta,Ctas.Cuenta,Concepto,Debe,Haber "
  SSQL = SSQL & "FROM Transacciones As Trans,Cuentas As Ctas "
  SSQL = SSQL & "WHERE Trans.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  SSQL = SSQL & "AND Trans.Cta = Ctas.Codigo "
  SSQL = SSQL & "AND T <> 'A' "
  SSQL = SSQL & "ORDER BY Trans.Fecha "
  DataTrans.RecordSource = SSQL: DataTrans.Refresh
  FechaIni = FormatoFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
  FechaFin = FormatoFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
  DBGTrans.Caption = "DIARIO GENERAL DESDE EL " & FechaIni & " AL " & FechaFin
  DBGTrans.Columns(DataTrans.Recordset.Fields.Count - 3).NumberFormat = "#,##0.00"
  DBGTrans.Columns(DataTrans.Recordset.Fields.Count - 2).NumberFormat = "#,##0.00"
  DBGTrans.Columns(DataTrans.Recordset.Fields.Count - 1).NumberFormat = "#,##0.00"
  DBGTrans.Visible = True
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  LabelTotReg.Caption = Format(DataTrans.Recordset.RecordCount, "#,##0")
  LabelTotPag.Caption = Format((DataTrans.Recordset.RecordCount / 50) + 1, "#,##0")
  MiTiempo = Time - MiTiempo1
  ListDiario.Caption = "DIARIO GENERAL.       Tiempo de proceso: " & Format(MiTiempo, "HH:MM:SS")
  MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  MousePointer = vbHourglass
  DBGTrans.Visible = False
  MiFecha = FormatoFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
  SQLMsg1 = "DIARIO GENERAL"
  Call ImprimirDataDiario(0.5, 0.5, DataTrans)
  DBGTrans.Visible = True
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  'Centrar el formulario
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 - 600
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  MDIConta.MousePointer = vbDefault
End Sub

Private Sub TextCS1_GotFocus()
  TextCS1.Text = ""
End Sub

Private Sub TextCS2_GotFocus()
   TextCS2.Text = ""
End Sub

Private Sub TextCS3_GotFocus()
   TextCS3.Text = ""
End Sub

Private Sub TextCS4_GotFocus()
  TextCS4.Text = ""
End Sub

Private Sub TextCS5_GotFocus()
   TextCS5.Text = ""
End Sub

Private Sub TextCS6_GotFocus()
   TextCS6.Text = ""
End Sub

Private Sub TextFechaA_Change()
  If Len(TextFechaA.Text) >= TextFechaA.MaxLength Then
     TextFechaDF.SetFocus
  End If
End Sub

Private Sub TextFechaA_LostFocus()
  If TextFechaA.Text = "" Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaD_Change()
  If Len(TextFechaD.Text) >= TextFechaD.MaxLength Then
     TextFechaM.SetFocus
  End If
End Sub

Private Sub TextFechaD_LostFocus()
  If TextFechaD.Text = "" Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaM_Change()
  If Len(TextFechaM.Text) >= TextFechaM.MaxLength Then
     TextFechaA.SetFocus
  End If
End Sub

Private Sub TextFechaM_LostFocus()
  If TextFechaM.Text = "" Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaAF_Change()
  If Len(TextFechaAF.Text) >= TextFechaAF.MaxLength Then
     Command1.SetFocus
  End If
End Sub

Private Sub TextFechaAF_LostFocus()
  If TextFechaAF.Text = "" Then TextFechaAF.Text = Anio
End Sub

Private Sub TextFechaDF_Change()
  If Len(TextFechaDF.Text) >= TextFechaDF.MaxLength Then
     TextFechaMF.SetFocus
  End If
End Sub

Private Sub TextFechaDF_LostFocus()
  If TextFechaDF.Text = "" Then TextFechaDF.Text = Dia
End Sub

Private Sub TextFechaMF_Change()
  If Len(TextFechaMF.Text) >= TextFechaMF.MaxLength Then
     TextFechaAF.SetFocus
  End If
End Sub

Private Sub TextFechaMF_LostFocus()
  If TextFechaMF.Text = "" Then TextFechaMF.Text = Mes
End Sub

Private Sub TextC1_Change()
   If Len(TextC1.Text) >= TextC1.MaxLength Then TextC2.SetFocus
End Sub

Private Sub TextC2_Change()
If Len(TextC2.Text) >= TextC2.MaxLength Then TextC3.SetFocus
End Sub

Private Sub TextC2_GotFocus()
  TextC2.Text = ""
End Sub

Private Sub TextC3_Change()
If Len(TextC3.Text) >= TextC3.MaxLength Then TextC4.SetFocus
End Sub

Private Sub TextC3_GotFocus()
  TextC3.Text = ""
End Sub

Private Sub TextC4_Change()
If Len(TextC4.Text) >= TextC4.MaxLength Then TextC5.SetFocus
End Sub

Private Sub TextC4_GotFocus()
  TextC4.Text = ""
End Sub

Private Sub TextC5_Change()
If Len(TextC5.Text) >= TextC5.MaxLength Then TextC6.SetFocus
End Sub

Private Sub TextC5_GotFocus()
  TextC5.Text = ""
End Sub

Private Sub TextC6_GotFocus()
  TextC6.Text = ""
End Sub

Private Sub TextC6_LostFocus()
  TextCS1.SetFocus
End Sub

Private Sub TextCS1_Change()
   If Len(TextCS1.Text) >= TextCS1.MaxLength Then TextCS2.SetFocus
End Sub

Private Sub TextCS2_Change()
If Len(TextCS2.Text) >= TextCS2.MaxLength Then TextCS3.SetFocus
End Sub

Private Sub TextCS3_Change()
If Len(TextCS3.Text) >= TextCS3.MaxLength Then TextCS4.SetFocus
End Sub

Private Sub TextCS4_Change()
If Len(TextCS4.Text) >= TextCS4.MaxLength Then TextCS5.SetFocus
End Sub

Private Sub TextCS5_Change()
If Len(TextCS5.Text) >= TextCS5.MaxLength Then TextCS6.SetFocus
End Sub

