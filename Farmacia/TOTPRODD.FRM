VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form TotalProdDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VENTAS DEL DIA"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   420
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
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
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   105
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
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
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCELAR"
      Height          =   645
      Left            =   6615
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR"
      Height          =   645
      Left            =   5145
      TabIndex        =   5
      Top             =   105
      Width           =   1485
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   4200
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   582
      _Version        =   327680
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
      Height          =   645
      Left            =   3675
      TabIndex        =   4
      Top             =   105
      Width           =   1485
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TotProdD.frx":0000
      Height          =   3375
      Left            =   105
      OleObjectBlob   =   "TotProdD.frx":0019
      TabIndex        =   8
      Top             =   840
      Width           =   7995
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Query2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2205
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2310
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Left            =   1785
      TabIndex        =   2
      Top             =   420
      Width           =   750
   End
   Begin VB.Label Label3 
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
      Left            =   1785
      TabIndex        =   9
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FORMATO DE FECHA (DD/MM/AA):"
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1590
   End
End
Attribute VB_Name = "TotalProdDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   RatonReloj MDIFact
   FechaIni = BuscarFecha(MBoxFechaI.Text): FechaFin = BuscarFecha(MBoxFechaF.Text)
   'Detalle de facturacion
   DBGrid1.Visible = False
   sSQL = "SELECT Fecha,Detalle_Factura.Codigo,Articulo.Articulo, SUM(Cantidad) As Total_Ventas "
   sSQL = sSQL & "FROM Detalle_Factura,Articulo "
   sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
   sSQL = sSQL & "AND Articulo.Codigo = Detalle_Factura.Codigo "
   sSQL = sSQL & "AND T <> '" & Anulado & "' "
   sSQL = sSQL & "GROUP BY Fecha,Detalle_Factura.Codigo,Articulo.Articulo "
   SelectData DataProduccion, sSQL, False
   DBAnchoMallas TotalProdDia, DBGrid1, 0, DataProduccion
   DBGrid1.Visible = True
   RatonNormal MDIFact
End Sub

Private Sub Command2_Click()
  MousePointer = vbHourglass
  SQLMsg1 = "REPORTE DE VENTAS"
  SQLMsg2 = ""
  SQLMsg3 = "Produccion desde: " & MBoxFechaI.Text & " Hasta: " & MBoxFechaF.Text
  ImprimirData MDIFact, DataProduccion, True, 1, 9, 0
  MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   Unload TotalProdDia
End Sub

Private Sub Form_Activate()
  RatonNormal MDIFact
End Sub

Private Sub Form_Load()
   CentrarForm TotalProdDia
   'Abriendo bases relacionadas
   DataProduccion.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub

Private Sub MBoxFechaF_LostFocus()
  If MBoxFechaF.Text = "00/00/00" Then MBoxFechaF.Text = FechaSistema
End Sub

Private Sub MBoxFechaI_LostFocus()
  If MBoxFechaI.Text = "00/00/00" Then MBoxFechaI.Text = FechaSistema
End Sub

