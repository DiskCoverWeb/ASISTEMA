VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Kard_Ing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "Kard_ing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10215
   Begin MSDBCtls.DBList DBLArt 
      Bindings        =   "Kard_ing.frx":0442
      DataSource      =   "DataArt"
      Height          =   1620
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   2858
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
   Begin MSDBCtls.DBCombo DBCInv 
      Bindings        =   "Kard_ing.frx":0454
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   6840
      _ExtentX        =   12065
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
   Begin VB.TextBox TextTotal 
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
      Left            =   8505
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Kard_ing.frx":0466
      Top             =   1995
      Width           =   1590
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Kard_ing.frx":0468
      Height          =   3060
      Left            =   105
      OleObjectBlob   =   "Kard_ing.frx":047D
      TabIndex        =   13
      Top             =   2415
      Width           =   9990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   5565
      Width           =   1380
   End
   Begin VB.Data DataIngArt 
      Caption         =   "IngArt"
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
      Top             =   2835
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataInv 
      Caption         =   "Inv"
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
      Top             =   2835
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TextVUnit 
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
      Left            =   8505
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Kard_ing.frx":0E2F
      Top             =   1680
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
      IMEMode         =   3  'DISABLE
      Left            =   8505
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   10
      Text            =   "Kard_ing.frx":0E31
      Top             =   1365
      Width           =   1590
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5250
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4935
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataSQL 
      Caption         =   "SQLs"
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
      Top             =   3255
      Visible         =   0   'False
      Width           =   1800
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
      Left            =   7035
      TabIndex        =   7
      Top             =   1995
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
      Left            =   7035
      TabIndex        =   6
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CANTIDAD"
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
      Left            =   7035
      TabIndex        =   5
      Top             =   1365
      Width           =   1485
   End
   Begin VB.Label LabelUnidad 
      BackColor       =   &H80000005&
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
      Left            =   8505
      TabIndex        =   9
      Top             =   1050
      Width           =   1590
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
      Left            =   7035
      TabIndex        =   4
      Top             =   1050
      Width           =   1485
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
      TabIndex        =   0
      Top             =   105
      Width           =   6840
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000005&
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
      Left            =   8505
      TabIndex        =   8
      Top             =   735
      Width           =   1590
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
      Left            =   7035
      TabIndex        =   3
      Top             =   735
      Width           =   1485
   End
End
Attribute VB_Name = "Kard_Ing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Unload Kard_Ing
End Sub

Private Sub DBCInv_LostFocus()
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomProd "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE Mid(Codigo_Inv,1,2) = '" & SinEspaciosIzq(DBCInv.Text) & "' "
  sSQL = sSQL & "AND TP = 'P' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBList DBLArt, DataArt, sSQL, "NomProd"
  SelectDBGrid DBGrid1, True, DataIngArt, "Asientos_K"
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataIngArt)
End Sub


Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLArt_LostFocus()
  sSQL = "SELECT * FROM Productos "
  sSQL = sSQL & "WHERE Codigo_Inv = '" & SinEspaciosIzq(DBLArt.Text) & "' "
  sSQL = sSQL & "AND TP = 'P' "
  SelectData DataSQL, sSQL, False
  Codigo = ""
  Cantidad = 0
  ValorUnitAnt = 0
  SaldoAnterior = 0
  Saldo = 0
  Producto = Ninguno
  With DataSQL.Recordset
   If .RecordCount > 0 Then
       Unidad = .Fields("Unidad")
       Codigo = .Fields("Codigo_Inv")
       Producto = .Fields("Producto")
   End If
  End With
  If Codigo <> "" Then
     sSQL = "SELECT * FROM Kardex "
     sSQL = sSQL & "WHERE Codigo_Inv = '" & Codigo & "' "
     sSQL = sSQL & "ORDER BY Fecha,Kardex "
     SelectData DataSQL, sSQL, False
     With DataSQL.Recordset
      If .RecordCount > 0 Then
         .MoveLast
          Cantidad = .Fields("Unidad")
          ValorUnitAnt = .Fields("Valor_Unitario")
          SaldoAnterior = .Fields("Saldo")
      End If
     End With
  End If
  LabelCodigo.Caption = Codigo
  LabelUnidad.Caption = Unidad
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True
  Entrada = Val(TextEntrada.Text)
End Sub

Private Sub Form_Activate()
  DeleteData DataIngArt, "DELETE * FROM Asientos_K "
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomInv "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE Len(Codigo_Inv) = 2 "
  sSQL = sSQL & "AND TP = 'I' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBCombo DBCInv, DataInv, sSQL, "NomInv", False
  SelectDBGrid DBGrid1, True, DataIngArt, "Asientos_K"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Kard_Ing
  DataInv.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataArt.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataIngArt.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextTotal_LostFocus()
   Entrada = Val(TextEntrada.Text)
   If OpcDH = 1 Then
      ValorUnit = Val(TextVUnit.Text)
   Else
      ValorUnit = ValorUnitAnt
   End If
   ValorTotal = ValorUnit * Entrada
  'llenamos el ultimo saldo del kardex
   sSQL = "SELECT * FROM Asientos_K "
   sSQL = sSQL & "WHERE Codigo_Inv = '" & Codigo & "' "
   SelectData DataSQL, sSQL, False
   If Entrada > 0 And DataSQL.Recordset.RecordCount <= 0 Then
      With DataIngArt.Recordset
          .AddNew
          .Fields("DH") = OpcDH
          .Fields("Codigo_Inv") = Codigo
          .Fields("Producto") = Producto
          .Fields("Cantidad_ES") = Entrada
          .Fields("Valor_Unit") = ValorUnit
          .Fields("Valor_Total") = ValorTotal
           If OpcDH = 1 Then
              Cantidad = Cantidad + Entrada
              Saldo = Round(SaldoAnterior + ValorTotal)
           Else
              Cantidad = Cantidad - Entrada
              Saldo = Round(SaldoAnterior - ValorTotal)
           End If
          .Fields("Cta") = SubCtaGen
          
          .Fields("Cantidad") = Cantidad
          .Fields("Valor_Unitario") = ValorUnit
          .Fields("Saldo") = Saldo
          .Update
      End With
   End If
   TextTotal.Text = Format(ValorTotal, "#,##0")
   DBLArt.SetFocus
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True
End Sub
