VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ListCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catalogo de Cuentas Contables.                             <Esc> para anular seleccion"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "Buscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDBCtls.DBList DBLCtas 
      Bindings        =   "Buscar.frx":0442
      DataSource      =   "DataCuentas"
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   5609
      _Version        =   327680
      ForeColor       =   -2147483646
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
   Begin VB.Data DataCuentas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5880
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1890
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "ListCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBLCtas_DblClick()
   Cadena = ObtenerPalabra(DBLCtas.Text, 4)
   LeerCta DataCuentas, Cadena
   ListCuentas.Hide
End Sub

Private Sub DBLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case Key_Esc
         Codigo = "": Cuenta = Ninguno: TipoCta = "G"
         SubCta = "N": Moneda_US = False
         ListCuentas.Hide
    Case Key_Enter
         Cadena = ObtenerPalabra(DBLCtas.Text, 4)
         LeerCta DataCuentas, Cadena
         ListCuentas.Hide
  End Select
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT 'Clave = ' & Format(Clave,'0000') & '=> ' & Codigo & Space(15) & Cuenta As Nombre_Cuenta "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE DG <> 'G' AND Cuenta <> '" & Ninguno & "' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectData DataCuentas, sSQL, False
  If DataCuentas.Recordset.RecordCount > 0 Then
     DBLCtas.ListField = "Nombre_Cuenta"
     DBLCtas.Text = DataCuentas.Recordset.Fields("Nombre_Cuenta")
  Else
     ListCuentas.Hide
  End If
End Sub

Private Sub Form_Load()
  CentrarForm ListCuentas
  DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub
