VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Existen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Existencias de Inventario"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "Existen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   9975
      Picture         =   "Existen.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2310
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
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
      Height          =   750
      Left            =   9975
      Picture         =   "Existen.frx":06C4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1575
      Width           =   1065
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Existen.frx":0F46
      Height          =   3375
      Left            =   105
      OleObjectBlob   =   "Existen.frx":0F5F
      TabIndex        =   9
      Top             =   3150
      Width           =   10935
   End
   Begin VB.Data DataSaldos 
      Caption         =   "Saldos"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
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
      Height          =   750
      Left            =   9975
      Picture         =   "Existen.frx":1911
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   840
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
      Height          =   750
      Left            =   9975
      Picture         =   "Existen.frx":1F7B
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   105
      Width           =   1065
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBCtls.DBCombo DBCInv 
      Bindings        =   "Existen.frx":23BD
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   105
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   9780
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   8295
         TabIndex        =   8
         Top             =   735
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
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
         Left            =   5985
         TabIndex        =   6
         Top             =   735
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
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
         Bindings        =   "Existen.frx":23D3
         DataSource      =   "DataArt"
         Height          =   1620
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   2858
         _Version        =   393216
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
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   330
         Left            =   105
         TabIndex        =   25
         Top             =   2100
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   582
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label LabelMaximo 
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
         Left            =   8295
         TabIndex        =   18
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label LabelExitencia 
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
         Left            =   8295
         TabIndex        =   21
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label LabelMinimo 
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
         Left            =   8295
         TabIndex        =   19
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label LabelBodega 
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
         Left            =   5985
         TabIndex        =   17
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label LabelUnidad 
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
         Left            =   5985
         TabIndex        =   16
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label LabelCodigo 
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
         Left            =   5985
         TabIndex        =   15
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Maximo:"
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
         Left            =   7455
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Existe.:"
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
         Left            =   7455
         TabIndex        =   20
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Minimo:"
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
         Left            =   7455
         TabIndex        =   13
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   7455
         TabIndex        =   7
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bodega:"
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
         Left            =   5145
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Unidad:"
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
         Left            =   5145
         TabIndex        =   11
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo:"
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
         Left            =   5145
         TabIndex        =   10
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Left            =   5145
         TabIndex        =   5
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FORMATO DE FECHAS: DD/MM/AAAA"
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
         Left            =   5145
         TabIndex        =   4
         Top             =   315
         Width           =   4530
      End
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
      Top             =   6510
      Width           =   10935
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataInv 
      Caption         =   "Inv"
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
      Top             =   5775
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5775
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      TabIndex        =   0
      Top             =   105
      Width           =   2220
   End
End
Attribute VB_Name = "Existen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  If Codigo = "" Then Codigo = "."
  sSQL = "SELECT K.Fecha,K.TP & ' : ' & K.Numero As Comp_No,"
  sSQL = sSQL & "Concepto As Detalle,"
  sSQL = sSQL & "Entrada,Salida,Valor_Unitario As Valor_Unit,Valor_Total,"
  sSQL = sSQL & "Cantidad,Valor_Unitario,Saldo_Total "
  sSQL = sSQL & "FROM Kardex As K, Comprobantes As C "
  sSQL = sSQL & "WHERE K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  sSQL = sSQL & "AND Codigo_Inv = '" & Codigo & "' "
  sSQL = sSQL & "AND K.T = '" & Normal & "' "
  sSQL = sSQL & "AND K.TP = C.TP "
  sSQL = sSQL & "AND K.Numero = C.Numero "
  sSQL = sSQL & "ORDER BY K.Fecha,K.TP,K.Numero,K.Kardex "
  SelectDBGrid DBGrid1, DataKardex, sSQL
  If DataKardex.Recordset.RecordCount > 0 Then
     DataKardex.Recordset.MoveLast
     LabelExitencia.Caption = Format(DataKardex.Recordset.Fields("Cantidad"), "#,##0.00")
  End If
End Sub

Private Sub Command2_Click()
 Unload Existen
End Sub

Private Sub Command3_Click()
  ImprimirKardex DataKardex, DataArticulo
End Sub

Private Sub DBCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLArt_DblClick()
  SiguienteControl
End Sub

Private Sub DBLArt_GotFocus()
  Codigo = SinEspaciosIzq(DBCInv.Text)
  LongStrg = Len(Codigo)
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomInv "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE Mid(Codigo_Inv,1," & LongStrg & ") = '" & Codigo & "' "
  sSQL = sSQL & "AND TP = 'P' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBList DBLArt, DataArt, sSQL, "NomInv"
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLArt_LostFocus()
  Codigo = ""
  sSQL = "SELECT * FROM Productos "
  sSQL = sSQL & "WHERE Codigo_Inv = '" & SinEspaciosIzq(DBLArt.Text) & "' "
  SelectData DataArticulo, sSQL, False
  If DataArticulo.Recordset.RecordCount > 0 Then
     Codigo = DataArticulo.Recordset.Fields("Codigo_Inv")
     LabelCodigo.Caption = DataArticulo.Recordset.Fields("Codigo_Inv")
     LabelUnidad.Caption = DataArticulo.Recordset.Fields("Unidad")
     LabelBodega.Caption = DataArticulo.Recordset.Fields("Bodega")
     'LabelProveedor.Caption = DataArticulo.Recordset.Fields("Proveedor")
     LabelMinimo.Caption = Format(DataArticulo.Recordset.Fields("Minimo"), "#,##0.00")
     LabelMaximo.Caption = Format(DataArticulo.Recordset.Fields("Maximo"), "#,##0.00")
     LabelExitencia.Caption = "0.00"
  End If
End Sub

Private Sub Form_Activate()
  Existen.Caption = "EXISTENCIA DE INVENTARIO: Espere un momento..."
  Cantidad = 0: Contador = 0
  sSQL = "SELECT * FROM Kardex "
  sSQL = sSQL & "WHERE T <> '" & Anulado & "' "
  sSQL = sSQL & "ORDER BY Codigo_Inv,Fecha,TP,Numero,Kardex "
  SelectData DataSaldos, sSQL, False
  With DataSaldos.Recordset
   If .RecordCount > 0 Then
       ProgBar.Max = .RecordCount
       ProgBar.Min = 0
       Contador = 0
       Codigo1 = .Fields("Codigo_Inv")
       Si_No = True
       Do While Not .EOF
         .Edit
          If Codigo1 <> .Fields("Codigo_Inv") Then
             Codigo1 = .Fields("Codigo_Inv")
             Cantidad = 0
          End If
          ValorUnitAnt = .Fields("Valor_Unitario")
          If .Fields("Entrada") <> 0 Then
              Cantidad = Cantidad + .Fields("Entrada")
          Else
              Cantidad = Cantidad - .Fields("Salida")
          End If
          SaldoAnterior = Round(Cantidad * ValorUnitAnt)
         .Fields("Cantidad") = Cantidad
         .Fields("Valor_Unitario") = ValorUnitAnt
         .Fields("Saldo_Total") = SaldoAnterior
         .Update
         .MoveNext
          Contador = Contador + 1
          ProgBar.Value = Contador
       Loop
   End If
  End With
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomInv "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE TP = 'I' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBCombo DBCInv, DataInv, sSQL, "NomInv", False
  RatonNormal
  Existen.Caption = "EXISTENCIA DE INVENTARIO"
End Sub

Private Sub Form_Load()
  CentrarForm Existen
  DataArt.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataInv.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSaldos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataKardex.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataArticulo.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub
