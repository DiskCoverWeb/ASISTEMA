VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Kard_Egr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTROL DE INVENTARIO PARA EGRESOS"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "Kard_Egr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextConcepto 
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
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   3
      Top             =   525
      Width           =   6315
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Kard_Egr.frx":0442
      Height          =   1800
      Left            =   105
      OleObjectBlob   =   "Kard_Egr.frx":0457
      TabIndex        =   19
      Top             =   3465
      Width           =   7575
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
      Height          =   435
      Left            =   6300
      TabIndex        =   18
      Top             =   5355
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Egreso de Inventario"
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
      Left            =   3045
      TabIndex        =   17
      Top             =   5355
      Width           =   3165
   End
   Begin VB.Data DataIngArt 
      Caption         =   "IngArt"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3885
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataInv 
      Caption         =   "Inv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3885
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3885
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox TextEntrada 
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
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Kard_Egr.frx":0E09
      Top             =   2205
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   2100
      TabIndex        =   1
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   327680
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
   Begin MSDBCtls.DBCombo DBCArt 
      Bindings        =   "Kard_Egr.frx":0E0B
      DataSource      =   "DataArt"
      Height          =   315
      Left            =   2625
      TabIndex        =   7
      Top             =   1365
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   327680
      Text            =   ""
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
   Begin VB.Data DataKardex 
      Caption         =   "Kardex"
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
      Top             =   3885
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3885
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4515
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQLs"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5460
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3885
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBCtls.DBCombo DBCInv 
      Bindings        =   "Kard_Egr.frx":0E1D
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   2625
      TabIndex        =   5
      Top             =   945
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   556
      _Version        =   327680
      Text            =   ""
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO"
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
      Width           =   1170
   End
   Begin VB.Label LabelVUnit 
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
      Left            =   1680
      TabIndex        =   20
      Top             =   2625
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE INVENTARIO"
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
      TabIndex        =   4
      Top             =   945
      Width           =   2430
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3360
      TabIndex        =   11
      Top             =   1785
      Width           =   1275
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      TabIndex        =   10
      Top             =   1785
      Width           =   1065
   End
   Begin VB.Label LabelTotal 
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
      Left            =   1680
      TabIndex        =   16
      Top             =   3045
      Width           =   1590
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR TOTAL"
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
      TabIndex        =   14
      Top             =   3045
      Width           =   1485
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR UNIT."
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
      TabIndex        =   13
      Top             =   2625
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALIDA"
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
      TabIndex        =   12
      Top             =   2205
      Width           =   1485
   End
   Begin VB.Label LabelUnidad 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1155
      TabIndex        =   9
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD"
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
      TabIndex        =   8
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA (DD/MM/AA):"
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
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL ARTICULO"
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
      TabIndex        =   6
      Top             =   1365
      Width           =   2430
   End
End
Attribute VB_Name = "Kard_Egr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Mensajes = "Esta Seguro que desea grabar."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = 6 Then
     If DataIngArt.Recordset.RecordCount > 0 Then
        RatonReloj
        FechaTexto = MBoxFecha.Text
        'Numero = ReadSetDataNum(DataCodigos, "Diario", False)
        Numero = ReadSetDataNum("Asiento_Kardex", True, True)
        SelectData DataKardex, "Kardex", False
        With DataIngArt.Recordset
            .MoveFirst
            'Llenamos los datos ingresados al Kardex
             Do While Not .EOF
                DataKardex.Recordset.AddNew
                DataKardex.Recordset.Fields("T") = Normal
                DataKardex.Recordset.Fields("Codigo") = .Fields("Codigo")
                DataKardex.Recordset.Fields("Fecha") = FechaTexto
                DataKardex.Recordset.Fields("Numero") = Numero
                DataKardex.Recordset.Fields("TP") = CompDiario
                DataKardex.Recordset.Fields("Concepto") = TextConcepto.Text
                DataKardex.Recordset.Fields("Entrada") = 0
                DataKardex.Recordset.Fields("Valor_UnitE") = 0
                DataKardex.Recordset.Fields("Valor_TotalE") = 0
                DataKardex.Recordset.Fields("Salida") = .Fields("Cantidad_ES")
                DataKardex.Recordset.Fields("Valor_UnitS") = .Fields("Valor_Unit")
                DataKardex.Recordset.Fields("Valor_TotalS") = .Fields("Valor_Total")
                DataKardex.Recordset.Fields("Cantidad") = .Fields("Cantidad")
                DataKardex.Recordset.Fields("Valor_Unitario") = .Fields("Valor_Unitario")
                DataKardex.Recordset.Fields("Saldo") = .Fields("Saldo")
                DataKardex.Recordset.Fields("Cta") = .Fields("Cta")
                DataKardex.Recordset.Update
               .MoveNext
             Loop
        End With
        RatonNormal
        Unload Kard_Egr
        'FDiario.Show
     End If
  End If
End Sub

Private Sub Command2_Click()
  Unload Kard_Egr
End Sub

Private Sub DBCTipoInv_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub Command3_Click()
End Sub

Private Sub DBCArt_GotFocus()
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE TP = '" & TipoInventario & "' "
  sSQL = sSQL & "AND Articulo "
  SelectDBCombo DBCArt, DataArt, sSQL, "Articulo", False
  Salida = 0: Entrada = 0: ValorUnitario = 0: Total = 0
End Sub

Private Sub DBCArt_LostFocus()
  FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha.Text) - 30))
  FechaFin = BuscarFecha(MBoxFecha.Text)
  CantidadAnt = 0: ValorUnitAnt = 0: SaldoAnterior = 0
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Articulo= '" & DBCArt.Text & "' "
  SelectData DataSQL, sSQL, False
  If DataSQL.Recordset.RecordCount > 0 Then
     Codigo = DataSQL.Recordset.Fields("Codigo")
     LabelCodigo.Caption = Codigo
     LabelUnidad.Caption = DataSQL.Recordset.Fields("Unidad")
     Producto = DataSQL.Recordset.Fields("Articulo")
     Cuenta = DataSQL.Recordset.Fields("Codigo_Cta")
     sSQL = "SELECT * FROM Kardex "
     sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
     sSQL = sSQL & "ORDER BY Fecha,Numero "
     SelectData DataSQL, sSQL, False
     If DataSQL.Recordset.RecordCount > 0 Then
        DataSQL.Recordset.MoveLast
        CantidadAnt = DataSQL.Recordset.Fields("Cantidad")
        ValorUnitAnt = DataSQL.Recordset.Fields("Valor_Unitario")
        SaldoAnterior = DataSQL.Recordset.Fields("Saldo")
     End If
  Else
     DBCInv.SetFocus
  End If
End Sub

Private Sub DBCInv_LostFocus()
  sSQL = "SELECT * FROM Inventario "
  sSQL = sSQL & "WHERE Concepto = '" & DBCInv.Text & "' "
  SelectData DataSQL, sSQL, False
  If DataSQL.Recordset.RecordCount > 0 Then TipoInventario = DataSQL.Recordset.Fields("TP")
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataIngArt)
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, True
End Sub

Private Sub TextConcepto_GotFocus()
  TextConcepto.Text = ""
End Sub

Private Sub TextConcepto_LostFocus()
  If TextConcepto.Text = "" Then TextConcepto.Text = Ninguno
End Sub

Private Sub TextEntrada_GotFocus()
  TextEntrada.Text = ""
End Sub

Private Sub TextEntrada_LostFocus()
  If TextEntrada.Text = "" Then TextEntrada.Text = "0"
  Salida = Val(TextEntrada.Text)
  LabelVUnit.Caption = Format(ValorUnitAnt, "###.00")
  ValorUnit = Val(LabelVUnit.Caption)
  ValorTotal = Round(Salida * ValorUnit)
  LabelTotal.Caption = Format(ValorTotal, "#,##0.00")
  'llenamos el ultimo saldo del kardex
  If Salida > 0 And ValorUnit > 0 Then
     With DataIngArt.Recordset
         .AddNew
         .Fields("Codigo") = Codigo
         .Fields("Articulo") = Producto
         .Fields("Cantidad_ES") = Salida
         .Fields("Valor_Unit") = ValorUnit
         .Fields("Valor_Total") = ValorTotal
         .Fields("Cantidad") = CantidadAnt - Salida
         .Fields("Valor_Unitario") = (SaldoAnterior - ValorTotal) / (CantidadAnt - Salida)
         .Fields("Saldo") = Round(SaldoAnterior - ValorTotal)
         .Fields("Cta") = Cuenta
         .Update
     End With
  End If
  DBCArt.SetFocus
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * FROM Inventario ORDER BY Concepto "
  SelectDBCombo DBCInv, DataInv, sSQL, "Concepto", False
  SelectData DataKardex, "Kardex", False
  DeleteData DataIngArt, "DELETE * FROM Asiento_Kardex "
  SelectDBGrid DBGrid1, True, DataIngArt, "Asiento_Kardex"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Kard_Egr
  DataInv.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
  DataArt.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
  DataSQL.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
  DataIngArt.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
  DataKardex.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
  'DataCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  'DataAsiento.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextVUnit_GotFocus()
  TextVUnit.Text = ""
End Sub

Private Sub TextVUnit_LostFocus()
End Sub
