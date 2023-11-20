VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form IngBodega 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de Bodegas"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtCodigo 
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
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2415
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "IngBodeg.frx":0000
      DataSource      =   "AdoInv"
      Height          =   1935
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   3413
      _Version        =   393216
      Style           =   1
      BackColor       =   16744576
      ForeColor       =   16777215
      Text            =   "Productos"
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
   Begin VB.TextBox TextSubCta 
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
      Left            =   1050
      MaxLength       =   25
      TabIndex        =   4
      Top             =   2415
      Width           =   3375
   End
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
      Height          =   855
      Left            =   4515
      Picture         =   "IngBodeg.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Left            =   4515
      Picture         =   "IngBodeg.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   420
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   525
      Top             =   735
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   525
      Top             =   1050
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "Inv"
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C&oncepto"
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
      Left            =   1050
      TabIndex        =   3
      Top             =   2100
      Width           =   3375
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Codigo:"
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
      Top             =   2100
      Width           =   855
   End
End
Attribute VB_Name = "IngBodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  GrabarInv
End Sub

Private Sub Command2_Click()
  Unload IngBodega
End Sub

Private Sub DCInv_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    If KeyCode = vbKeyF1 Then GenerarDataTexto IngBodega, AdoInv
    If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(DCInv.Text)
     Cuenta = SinEspaciosDer(DCInv.Text)
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodBodega = '" & Codigo & "' "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        MsgBox "No se puede eliminar este codigo: " & Codigo & vbCrLf _
               & "Detalle: " & Cuenta & vbCrLf _
               & "existen datos procesados"
     Else
        Mensajes = "Seguro de Eliminar el Codigo:" & Codigo & vbCrLf _
                 & "De " & Cuenta & "?"
        Titulo = "ELIMINACION"
        If BoxMensaje = vbYes Then
           sSQL = "DELETE * " _
                & "FROM Catalogo_Bodegas " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND CodBod = '" & Codigo & "' "
           Ejecutar_SQL_SP sSQL
           ListarBodegas "P"
        End If
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Bodegas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Bodega "
     Select_Adodc AdoInv, sSQL
     ImprimirAdodc AdoInv, 1, 8
     ListarBodegas "P"
  End If
  If KeyCode = vbKeyReturn Then LlenarInv
End Sub

Private Sub DCInv_LostFocus()
  LlenarInv
End Sub

Private Sub Form_Activate()
  ListarBodegas "P"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm IngBodega
  ConectarAdodc AdoAux
  ConectarAdodc AdoInv
  ConectarAdodc AdoTInv
End Sub

Private Sub TxtCodigo_GotFocus()
  MarcarTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodigo_LostFocus()
   TxtCodigo.Text = UCase(TxtCodigo.Text)
End Sub

Private Sub TextSubCta_GotFocus()
  MarcarTexto TextSubCta
End Sub

Private Sub TextSubCta_LostFocus()
  TextoValido TextSubCta
End Sub

Public Sub LlenarInv()
   TextSubCta.Text = ""
   With AdoInv.Recordset
    If .RecordCount > 0 Then
        Codigo = SinEspaciosIzq(DCInv.Text)
       .MoveFirst
        TextoBusqueda = "CodBod Like '" & Codigo & "' "
       .Find (TextoBusqueda)
        If Not .EOF Then
           TextSubCta.Text = .Fields("Bodega")
           TxtCodigo.Text = .Fields("CodBod")
        End If
    Else
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
End Sub

Public Sub GrabarInv()
  RatonReloj
  Nuevo = False
  Codigo = UCase(TxtCodigo.Text)
  sSQL = "SELECT * " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodBod "
  Select_Adodc AdoInv, sSQL
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       TextoBusqueda = "CodBod Like '" & Codigo & "' "
      .Find (TextoBusqueda)
       If .EOF Then
           SetAddNew AdoInv
           Nuevo = True
       End If
   Else
      SetAddNew AdoInv
      Nuevo = True
   End If
   SetFields AdoInv, "CodBod", Codigo
   SetFields AdoInv, "Bodega", TextSubCta.Text
   SetFields AdoInv, "Periodo", Periodo_Contable
   SetFields AdoInv, "Item", NumEmpresa
   SetUpdate AdoInv
  End With
  ListarBodegas "P"
  RatonNormal
End Sub

Public Sub ListarBodegas(OpcVista As String)
   sSQL = "SELECT CodBod & '  ' & Bodega As NomProd,* " _
        & "FROM Catalogo_Bodegas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND CodBod <> '.' " _
        & "ORDER BY CodBod "
   SelectDB_Combo DCInv, AdoInv, sSQL, "NomProd"
End Sub

