VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ProdTerm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos Terminados"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextRotos 
      Alignment       =   1  'Right Justify
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
      Left            =   5670
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "ProdTerm.frx":0000
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox TextFaltante 
      Alignment       =   1  'Right Justify
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
      Left            =   5670
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "ProdTerm.frx":0002
      Top             =   1785
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   2625
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
   Begin MSDBCtls.DBList DBLArt 
      Bindings        =   "ProdTerm.frx":0004
      DataSource      =   "DataArt"
      Height          =   2790
      Left            =   105
      TabIndex        =   3
      Top             =   840
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   4921
      _Version        =   327680
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
   Begin VB.TextBox TextSobrante 
      Alignment       =   1  'Right Justify
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
      Left            =   5670
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "ProdTerm.frx":0016
      Top             =   1155
      Width           =   1275
   End
   Begin VB.TextBox TextCantidad 
      Alignment       =   1  'Right Justify
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
      Left            =   5670
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "ProdTerm.frx":0018
      Top             =   840
      Width           =   1275
   End
   Begin VB.Data DataArt1 
      Caption         =   "Art1"
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
      Top             =   3045
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   5670
      TabIndex        =   21
      Top             =   3045
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
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
      Left            =   4305
      TabIndex        =   20
      Top             =   3045
      Width           =   1275
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
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
      Top             =   2730
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataStock 
      Caption         =   "Stock"
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
      Top             =   2415
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label LabelSaldo 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5670
      TabIndex        =   19
      Top             =   2415
      Width           =   1275
   End
   Begin VB.Label LabelVentas 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5670
      TabIndex        =   13
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4095
      TabIndex        =   18
      Top             =   2415
      Width           =   1590
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Rotos"
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
      Left            =   4095
      TabIndex        =   16
      Top             =   2100
      Width           =   1590
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Faltante"
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
      Left            =   4095
      TabIndex        =   14
      Top             =   1785
      Width           =   1590
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ventas del Dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4095
      TabIndex        =   12
      Top             =   1470
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N O M B R E    P R O D U C T O"
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
      TabIndex        =   2
      Top             =   525
      Width           =   3900
   End
   Begin VB.Label LabelStock 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5670
      TabIndex        =   7
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
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
      Left            =   5670
      TabIndex        =   5
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sobrante"
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
      Left            =   4095
      TabIndex        =   10
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Producción"
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
      Left            =   4095
      TabIndex        =   8
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4095
      TabIndex        =   6
      Top             =   525
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Codigo del Prod"
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
      Left            =   4095
      TabIndex        =   4
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha de Producción:"
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
      Width           =   2430
   End
End
Attribute VB_Name = "ProdTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Saldos = Val(LabelStock.Caption)
   SaldoProd = Val(LabelStock.Caption)
   Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
   LabelSaldo.Caption = Str(Saldos)
   Mensajes = "Desea Grabar Produccion del Dia"
   Titulo = "Formulario de impresion."
   TipoDeCaja = 4 + 32: K = MsgBox(Mensajes, TipoDeCaja, Titulo)
   If K = 6 And Codigo <> "" Then
      GrabarProduccion MBoxFecha.Text, SaldoProd
      TextRotos.Text = "0"
      TextFaltante.Text = "0"
      TextSobrante.Text = "0"
   End If
End Sub

Private Sub Command2_Click()
  Unload ProdTerm
End Sub

Private Sub DBLArt_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub DBLArt_LostFocus()
Volver_Buscar:
  Produccion = 0: Sobrantes = 0: Faltantes = 0
  Rotos = 0: Saldos = 0: TotalVentas = 0
  Codigo = SinEspaciosIzq(DBLArt.Text)
  LabelCodigo.Caption = Codigo
  FechaEntI = CFechaLong(MBoxFecha.Text) - 30
  FechaIni = BuscarFecha(CLongFecha(FechaEntI))
  FechaFin = BuscarFecha(MBoxFecha.Text)
  MiFecha = FechaFin
  sSQL = "SELECT Fecha,Codigo,Saldo_Actual FROM Stock_Articulo "
  sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
  sSQL = sSQL & "ORDER BY Fecha "
  SelectData DataStock, sSQL, False
  With DataStock.Recordset
     If .RecordCount > 0 Then
        .MoveLast
         Saldos = .Fields("Saldo_Actual")
         MiFecha = .Fields("Fecha")
         FechaEntI = CFechaLong(MiFecha)
         FechaEntF = CFechaLong(MBoxFecha.Text)
         If (FechaEntF - FechaEntI) > 1 Then
            SaldoProd = Saldos
            Do While FechaEntI < FechaEntF - 1
               FechaEntI = FechaEntI + 1
               TotalVentas = VentasDia(BuscarFecha(CLongFecha(FechaEntI)))
               GrabarProduccion CLongFecha(FechaEntI), SaldoProd
               SaldoProd = SaldoProd - TotalVentas
               GoTo Volver_Buscar
            Loop
         End If
     Else
         LabelStock.Caption = "0"
     End If
  End With
  TotalVentas = VentasDia(BuscarFecha(MBoxFecha.Text))
  LabelStock.Caption = Saldos
  LabelVentas.Caption = Str(TotalVentas)
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & ' ->' & Articulo As Nom_Art "
  sSQL = sSQL & "FROM Articulo ORDER BY Articulo "
  SelectDBList DBLArt, DataArt, sSQL, "Nom_Art"
  MBoxFecha.Text = "00/00/0000"
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm ProdTerm
  'Abriendo bases relacionadas
   DataArt.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataArt1.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataStock.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
   Produccion = 0: Sobrantes = 0: Faltantes = 0
   Rotos = 0: Reposicion = 0: Saldos = 0: TotalVentas = 0
End Sub

Private Sub TextCantidad_GotFocus()
   MarcarTexto TextCantidad
End Sub

Private Sub TextCantidad_LostFocus()
 If TextCantidad.Text = "" Then TextCantidad.Text = "0"
 Produccion = Val(TextCantidad.Text)
 Saldos = Val(LabelStock.Caption)
 Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
 LabelSaldo.Caption = Str(Saldos)
End Sub

Private Sub TextFaltante_GotFocus()
  MarcarTexto TextFaltante
End Sub

Private Sub TextFaltante_LostFocus()
  If TextFaltante.Text = "" Then TextFaltante.Text = "0"
  Faltantes = Val(TextFaltante.Text)
  Saldos = Val(LabelStock.Caption)
  Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
  LabelSaldo.Caption = Str(Saldos)
End Sub

Private Sub TextRotos_GotFocus()
  MarcarTexto TextRotos
End Sub

Private Sub TextRotos_LostFocus()
 If TextRotos.Text = "" Then TextRotos.Text = "0"
 Rotos = Val(TextRotos.Text)
 Saldos = Val(LabelStock.Caption)
 Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
 LabelSaldo.Caption = Str(Saldos)
End Sub

Private Sub TextSobrante_GotFocus()
  MarcarTexto TextSobrante
End Sub

Private Sub TextSobrante_LostFocus()
 If TextSobrante.Text = "" Then TextSobrante.Text = "0"
 Sobrantes = Val(TextSobrante.Text)
 Saldos = Val(LabelStock.Caption)
 Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
 LabelSaldo.Caption = Str(Saldos)
End Sub

Public Sub GrabarProduccion(FechaDia As String, SaldosProdDia As Long)
   Saldos = SaldosProdDia
   Saldos = Saldos + Produccion + Sobrantes - TotalVentas - Faltantes - Rotos
   sSQL = "SELECT * FROM Stock_Articulo "
   sSQL = sSQL & "WHERE Fecha = #" & BuscarFecha(FechaDia) & "# "
   sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
   SelectData DataStock, sSQL, False
   If DataStock.Recordset.RecordCount > 0 Then
      MsgBox "Esta produccion ya esta ingresada"
      MBoxFecha.SetFocus
   Else
    With DataStock.Recordset
         sSQL = "UPDATE Articulo SET Stock = " & Saldos & " "
         sSQL = sSQL & "WHERE Codigo ='" & Codigo & "' "
         UpdateData DataArt1, sSQL
        .AddNew
        .Fields("Fecha") = FechaDia
        .Fields("Codigo") = Codigo
        .Fields("Saldo_Anterior") = SaldosProdDia
        .Fields("Produccion") = Produccion
        .Fields("Sobrantes") = Sobrantes
        .Fields("Ventas_Dia") = TotalVentas
        .Fields("Reposicion") = 0
        .Fields("Rotos") = Rotos
        .Fields("Faltantes") = Faltantes
        .Fields("Saldo_Actual") = Saldos
        .Update
    End With
   End If
End Sub

Function VentasDia(FechaDia As String) As Long
Dim TotalVenta As Long
  TotalVenta = 0
  sSQL = "SELECT Cantidad FROM Detalle_Factura "
  sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaDia & "# and #" & FechaDia & "# "
  sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
  sSQL = sSQL & "AND T <> '" & Anulado & "' "
  SelectData DataArt1, sSQL, False
  With DataArt1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          TotalVenta = TotalVenta + .Fields("Cantidad")
         .MoveNext
       Loop
   End If
  End With
  VentasDia = TotalVenta
End Function
