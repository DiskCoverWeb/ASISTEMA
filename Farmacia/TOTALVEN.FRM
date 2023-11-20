VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form TotalesVentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VENTAS MENSUALES POR CLIENTES"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   2955
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   2100
         TabIndex        =   10
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         _Version        =   327680
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####"
         Mask            =   "####"
         PromptChar      =   "0"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año a procesar:"
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
         TabIndex        =   11
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   3150
      TabIndex        =   5
      Top             =   0
      Width           =   8205
      Begin VB.OptionButton OpcCant 
         Caption         =   "En Cantidad"
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
         Left            =   4830
         TabIndex        =   9
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton OpcSucres 
         Caption         =   "En Dinero"
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
         Left            =   6720
         TabIndex        =   8
         Top             =   210
         Width           =   1170
      End
      Begin MSDBCtls.DBCombo DBCProd 
         Bindings        =   "TotalVen.frx":0000
         DataSource      =   "DataProd"
         Height          =   315
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   327680
         Text            =   "DBCombo1"
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
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCELAR"
      Height          =   330
      Left            =   10080
      TabIndex        =   2
      Top             =   735
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR"
      Height          =   330
      Left            =   8820
      TabIndex        =   1
      Top             =   735
      Width           =   1275
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   6300
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   582
      _Version        =   327680
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
      Height          =   330
      Left            =   7455
      TabIndex        =   0
      Top             =   735
      Width           =   1380
   End
   Begin MSDBGrid.DBGrid DBGDatas 
      Bindings        =   "TotalVen.frx":0013
      Height          =   5265
      Left            =   105
      OleObjectBlob   =   "TotalVen.frx":0029
      TabIndex        =   3
      Top             =   1050
      Width           =   11250
   End
   Begin VB.Data DataFacturas 
      Caption         =   "Query2"
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
      Top             =   5985
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Data DataResumen 
      Caption         =   "Query2"
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
      Top             =   735
      Width           =   7395
   End
   Begin VB.Data DataProd 
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
      Top             =   5985
      Visible         =   0   'False
      Width           =   2145
   End
End
Attribute VB_Name = "TotalesVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim SumaMeses(13) As Double
Dim TotalMeses As Double
Dim MesActual As Integer
Dim Meses As Integer
   MousePointer = vbHourglass
   DBGDatas.Visible = False
   If OpcCant Then
      DataResumen.Caption = "RESUMEN DE VENTAS DE PRODUCTOS EN: CANTIDAD - REPOSICION"
   Else
      DataResumen.Caption = "RESUMEN DE VENTAS DE PRODUCTOS EN: DINERO + IVA"
   End If
   For I = 0 To 11
      SumaMeses(I) = 0
   Next I
   FechaIni = "01/01/" & MBoxFechaI.Text: FechaFin = "12/31/" & MBoxFechaI.Text
   'Detalle de facturacion
   sSQL = "DELETE * FROM Total_Meses1 "
   DataResumen.Database.Execute sSQL
   DataResumen.RecordSource = "Total_Meses1 ": DataResumen.Refresh
   sSQL = "SELECT Codigo_C,Cliente,Fecha,Cantidad,Reposicion,Total,Total_Desc,Total_IVA "
   sSQL = sSQL & "FROM Detalle_Factura,Clientes "
   sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
   sSQL = sSQL & "AND Detalle_Factura.Codigo = '" & SinEspaciosIzq(DBCProd.Text) & "' "
   sSQL = sSQL & "AND T <> '" & Anulado & "' "
   sSQL = sSQL & "AND Clientes.Codigo = Detalle_Factura.Codigo_C "
   sSQL = sSQL & "ORDER BY Codigo_C,Fecha "
   SelectData DataFacturas, sSQL, False
   With DataFacturas.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       TotalBarra = 0
       ProgBar.Min = TotalBarra: ProgBar.Max = .RecordCount + 10
       ProgBar.Value = TotalBarra
      .MoveFirst
       Producto = .Fields("Cliente")
       Codigo = .Fields("Codigo_C")
       Do While Not .EOF
         If Codigo <> .Fields("Codigo_C") Then
            DataResumen.Recordset.AddNew
            DataResumen.Recordset.Fields("Codigo") = Mid(Producto, 1, 30)
            TotalMeses = 0
            For Meses = 1 To 12
                TotalMeses = TotalMeses + SumaMeses(Meses - 1)
                Cadena = MesesLetras(Meses)
                DataResumen.Recordset.Fields(Cadena) = SumaMeses(Meses - 1)
                SumaMeses(Meses - 1) = 0
            Next Meses
            DataResumen.Recordset.Fields("TOTAL") = TotalMeses
            DataResumen.Recordset.Update
            Producto = .Fields("Cliente")
            Codigo = .Fields("Codigo_C")
         End If
         MesActual = FechaMes(.Fields("Fecha")) - 1
         If OpcCant Then
            SumaMeses(MesActual) = SumaMeses(MesActual) + .Fields("Cantidad") - .Fields("Reposicion")
         Else
            SumaMeses(MesActual) = SumaMeses(MesActual) + .Fields("Total") - .Fields("Total_Desc") + .Fields("Total_IVA")
         End If
         ProgBar.Value = TotalBarra
         TotalBarra = TotalBarra + 1
        .MoveNext
       Loop
       DataResumen.Recordset.AddNew
       DataResumen.Recordset.Fields("Codigo") = Codigo
       TotalMeses = 0
       For Meses = 1 To 12
           TotalMeses = TotalMeses + SumaMeses(Meses - 1)
           Cadena = MesesLetras(Meses)
           DataResumen.Recordset.Fields(Cadena) = SumaMeses(Meses - 1)
           SumaMeses(Meses - 1) = 0
       Next Meses
       DataResumen.Recordset.Fields("TOTAL") = TotalMeses
       DataResumen.Recordset.Update
       ProgBar.Value = ProgBar.Max
   End If
   End With
   sSQL = "SELECT Codigo As Nombre_Cliente, "
   For Meses = 1 To 12
       sSQL = sSQL & MesesLetras(Meses) & ","
   Next Meses
   sSQL = sSQL & "TOTAL FROM Total_Meses1 "
   sSQL = sSQL & "ORDER BY TOTAL DESC "
   SelectDBGrid TotalesVentas, DBGDatas, True, DataResumen, sSQL
   DBGDatas.Visible = True
   MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  DBGDatas.Visible = False
  MousePointer = vbHourglass
  SQLMsg1 = "VENTAS DE: " & DBCProd.Text
  If OpcCant Then
     SQLMsg2 = "EN CANTIDAD - REPOSICION DEL AÑO: " & MBoxFechaI.Text
  Else
     SQLMsg2 = "EN DINERO CON IVA DEL AÑO: " & MBoxFechaI.Text
  End If
  ImprimirMeses TotalesVentas, DataResumen, 8, False
  DBGDatas.Visible = True
  MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   Unload TotalesVentas
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & ' ->'& Articulo As Nom_Art "
  sSQL = sSQL & "FROM Articulo ORDER BY Articulo "
  SelectDBCombo DBCProd, DataProd, sSQL, "Nom_Art", False
  RatonNormal MDIFact
End Sub

Private Sub Form_Load()
   CentrarForm TotalesVentas
   'Abriendo bases relacionadas
   DataFacturas.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataProd.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataResumen.DatabaseName = RutaEmpresa & "\RESULT.MDB"
End Sub

