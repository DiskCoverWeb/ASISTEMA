VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form TotalProduccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESUMEN DE PRODUCCION Y VENTAS"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TotProd.frx":0000
      Height          =   1485
      Left            =   105
      OleObjectBlob   =   "TotProd.frx":0019
      TabIndex        =   13
      Top             =   4620
      Width           =   11250
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALIR"
      Height          =   645
      Left            =   10395
      TabIndex        =   12
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR"
      Height          =   645
      Left            =   9240
      TabIndex        =   4
      Top             =   105
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   3675
      TabIndex        =   8
      Top             =   0
      Width           =   5475
      Begin VB.OptionButton Opc3 
         Caption         =   "Detalle de Facturas por Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3255
         TabIndex        =   11
         Top             =   210
         Width           =   2010
      End
      Begin VB.OptionButton Opc2 
         Caption         =   "Ventas del Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1995
         TabIndex        =   10
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton Opc1 
         Caption         =   "Producción contra Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
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
      Left            =   2205
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
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
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   6090
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   582
      _Version        =   327680
      Appearance      =   1
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TotProd.frx":09CB
      Height          =   3795
      Left            =   105
      OleObjectBlob   =   "TotProd.frx":09E4
      TabIndex        =   6
      Top             =   840
      Width           =   11250
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Produccion"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   945
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Data DataFacturas 
      Caption         =   "Facturas"
      Connect         =   "Access"
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
      Top             =   945
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Data DataDiarioCaja 
      Caption         =   "DiarioCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   945
      Visible         =   0   'False
      Width           =   2040
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
      Left            =   1470
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
      Left            =   1470
      TabIndex        =   7
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FORMATO FECHA (DD/MM/AA)"
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
      Width           =   1275
   End
End
Attribute VB_Name = "TotalProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
  MousePointer = vbHourglass
  DBGrid1.Visible = False
  SQLMsg3 = "Produccion desde: " & FechaIni & " al: " & FechaFin & "."
  If Opc1 Then
     SQLMsg1 = "REPORTE DE PRODUCCION"
     SQLMsg2 = "CONTRA VENTAS"
     ImprimirProd DataProduccion, DataDiarioCaja, 1, 9, 0
  ElseIf Opc2 Then
     SQLMsg1 = "REPORTE DE VENTAS"
     SQLMsg2 = ""
     ImprimirVentas DataProduccion, True, 1, 9, 0
  ElseIf Opc3 Then
     SQLMsg1 = "REPORTE DE VENTAS"
     SQLMsg2 = "POR GRUPO DE PRODUCTO"
     ImprimirVentas DataProduccion, True, 1, 9, 0
  End If
  DBGrid1.Visible = True
  MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
   Unload TotalProduccion
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm TotalProduccion
   'Abriendo bases relacionadas
   DataFacturas.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataDiarioCaja.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataProduccion.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub Opc1_DblClick()
   RatonReloj
   FechaIni = BuscarFecha(MBoxFechaI.Text): FechaFin = BuscarFecha(MBoxFechaF.Text)
     'Detalle de facturacion
      sSQL = "SELECT Fecha,Codigo, SUM(Cantidad) As Total_Ventas, SUM(Reposicion) As Total_Rep "
      sSQL = sSQL & "FROM Detalle_Factura "
      sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
      sSQL = sSQL & "AND T <> '" & Anulado & "' "
      sSQL = sSQL & "GROUP BY Fecha,Codigo "
      SelectData DataFacturas, sSQL, False
      With DataFacturas.Recordset
        If .RecordCount > 0 Then
           .MoveLast: TotalBarra = 1
            ProgBar.Min = 0: ProgBar.Max = .RecordCount
            ProgBar.Value = TotalBarra
           .MoveFirst
            Do While Not .EOF
               Cantidad = .Fields("Total_Ventas")
               Reposicion = .Fields("Total_Rep")
               Codigo = .Fields("Codigo")
               MiFecha = BuscarFecha(.Fields("Fecha"))
               sSQL = "UPDATE Stock_Articulo SET Ventas_Dia = " & Cantidad & ", "
               sSQL = sSQL & "Reposicion = " & Reposicion & ", "
               sSQL = sSQL & "Saldo_Actual = Saldo_Anterior + Produccion + Sobrantes "
               sSQL = sSQL & " - " & Cantidad & " - Rotos - Faltantes "
               sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# "
               sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
               UpdateData DataProduccion, sSQL
               ProgBar.Value = TotalBarra
               TotalBarra = TotalBarra + 1
              .MoveNext
            Loop
            ProgBar.Value = ProgBar.Max
        End If
      End With
      sSQL = "SELECT Articulo.Articulo,Fecha,Saldo_Anterior,Produccion,Sobrantes,"
      sSQL = sSQL & "Ventas_Dia,Rotos,Faltantes,Saldo_Actual,Reposicion "
      sSQL = sSQL & "FROM Stock_Articulo,Articulo "
      sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
      sSQL = sSQL & "AND Articulo.Codigo = Stock_Articulo.Codigo "
      sSQL = sSQL & "ORDER BY Articulo.Articulo,Fecha "
      SelectDBGrid DBGrid1, True, DataProduccion, sSQL
      
      sSQL = "SELECT SUM(Monto) As Monto_Total,"
      sSQL = sSQL & "SUM(Efectivo) As Total_Efectivo,"
      sSQL = sSQL & "SUM(Retencion) As Total_Retencion,"
      sSQL = sSQL & "SUM(Cheque) As Total_Cheque,"
      sSQL = sSQL & "SUM(Total_Abono) As Total_Abonos,"
      sSQL = sSQL & "SUM(Saldo) As Saldo_Total "
      sSQL = sSQL & "FROM Diario_Caja "
      sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
      sSQL = sSQL & "AND TP ='" & Ventas & "' "
      
      'sSQL = "SELECT Articulo.Articulo,"
      'sSQL = sSQL & "SUM(Produccion),"
      'sSQL = sSQL & "SUM(Sobrantes), "
      'sSQL = sSQL & "SUM(Ventas_Dia),"
      'sSQL = sSQL & "SUM(Rotos),"
      'sSQL = sSQL & "SUM(Faltantes),"
      'sSQL = sSQL & "SUM(Saldo_Actual),"
      'sSQL = sSQL & "SUM(Reposicion) "
      'sSQL = sSQL & "FROM Stock_Articulo,Articulo "
      'sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
      'sSQL = sSQL & "AND Articulo.Codigo = Stock_Articulo.Codigo "
      'sSQL = sSQL & "ORDER BY Articulo.Articulo,Fecha "
      'sSQL = sSQL & "GROUP BY Fecha "
      'SelectData DataDiarioCaja, sSQL, False
      SelectDBGrid DBGrid2, 0, DataDiarioCaja, sSQL
   RatonNormal MDIFact
End Sub

Private Sub Opc2_DblClick()
   RatonReloj MDIFact
   FechaIni = BuscarFecha(MBoxFechaI.Text)
   FechaFin = BuscarFecha(MBoxFechaF.Text)
   DBGrid1.Visible = False
   'Detalle de facturacion
   sSQL = "SELECT Articulo.Articulo,Fecha,Count(Factura_No) As Cant_Facturas, SUM(Cantidad) As Cantidad_Vendida,SUM(Reposicion) As Reposiciones,SUM(Total-Total_Desc) As Total_Ventas "
   sSQL = sSQL & "FROM Detalle_Factura,Articulo "
   sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
   sSQL = sSQL & "AND Articulo.Codigo = Detalle_Factura.Codigo "
   sSQL = sSQL & "AND T <> '" & Anulado & "' "
   sSQL = sSQL & "GROUP BY Articulo.Articulo,Fecha "
   SelectDBGrid TotalProduccion, DBGrid1, False, DataProduccion, sSQL
   DBGrid1.Visible = True
   RatonNormal MDIFact
End Sub

Private Sub Opc3_DblClick()
   RatonReloj MDIFact
   FechaIni = BuscarFecha(MBoxFechaI.Text): FechaFin = BuscarFecha(MBoxFechaF.Text)
   'Detalle de facturacion
   sSQL = "SELECT Articulo.Articulo,Fecha, Factura_No, Cantidad, Reposicion, Total "
   sSQL = sSQL & "FROM Detalle_Factura,Articulo "
   sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
   sSQL = sSQL & "AND Articulo.Codigo = Detalle_Factura.Codigo "
   sSQL = sSQL & "AND T <> '" & Anulado & "' "
   sSQL = sSQL & "ORDER BY Articulo.Articulo,Fecha,Factura_No "
   '   SelectData DataProduccion, sSQL, True
   SelectDBGrid TotalProduccion, DBGrid1, 0, DataProduccion, sSQL
   RatonNormal MDIFact
End Sub
