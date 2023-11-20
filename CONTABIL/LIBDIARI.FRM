VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LibroDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIBRO DIARIO"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Crystal.CrystalReport RptLibroDiario 
      Bindings        =   "LibDiari.frx":0000
      Left            =   3780
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "D:\SISTEMA\librod1.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "D:\SISTEMA\HOLA"
      PrintFileType   =   17
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR LIBRO DIARIO"
      Height          =   960
      Left            =   3990
      TabIndex        =   4
      Top             =   1155
      Width           =   2010
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   3150
      TabIndex        =   3
      Top             =   525
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
      OLEDragMode     =   1
      OLEDropMode     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "0"
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1050
      TabIndex        =   2
      Top             =   525
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
      OLEDragMode     =   1
      OLEDropMode     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "0"
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   2310
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   582
      _Version        =   327680
      Appearance      =   1
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      TabIndex        =   1
      Top             =   525
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
      Height          =   330
      Left            =   1890
      TabIndex        =   11
      Top             =   1890
      Width           =   1905
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
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   1890
      Width           =   1695
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
      Left            =   1890
      TabIndex        =   9
      Top             =   1470
      Width           =   1905
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
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1470
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
      Left            =   1890
      TabIndex        =   7
      Top             =   1050
      Width           =   1905
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
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1050
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "LibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  MousePointer = vbHourglass
  SumaDebe = 0: SumaHaber = 0
  SSQL = "SELECT Transacciones.*,Cuentas.* "
  SSQL = SSQL & "FROM Transacciones, Cuentas "
  SSQL = SSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  SSQL = SSQL & "AND Cuentas.Codigo = Cta "
  SSQL = SSQL & "AND T <> '" & Anulado & "' "
  'SSQL = SSQL & "ORDER BY Fecha,TP,Numero "
  SelectData DataTrans, SSQL, False
  RptLibroDiario.Connect = RutaEmpresa & "\CONTABIL.MDB"
  If DataTrans.Recordset.RecordCount > 0 Then
     SetProgBar ProgBarra, DataTrans.Recordset.RecordCount
     DataTrans.Recordset.MoveFirst
     'Do While Not DataTrans.Recordset.EOF
     '   SumaDebe = SumaDebe + DataTrans.Recordset.Fields("Debe")
     '   SumaHaber = SumaHaber + DataTrans.Recordset.Fields("Haber")
     '   IncProgBar ProgBarra
     '   DataTrans.Recordset.MoveNext
     'Loop
  End If
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0")
  MousePointer = vbDefault
  'RptLibroDiario.Action = 1
End Sub

Private Sub Command2_Click()
  MousePointer = vbHourglass
  FechaIni = MBoxFechaI.Text
  FechaFin = MBoxFechaF.Text
  'ImprimirDataBalance DataCtas
  MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  Label5.Caption = Empresa
End Sub

Private Sub Form_Load()
  CentrarForm LibroDiario
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  MDIConta.MousePointer = vbDefault
End Sub
