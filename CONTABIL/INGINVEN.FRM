VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form IngProdInv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de SubCuentas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Elimina una Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nueva Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primera Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultima Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "IngInven.frx":0000
   End
   Begin VB.TextBox TextUnidad 
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
      Left            =   2730
      MaxLength       =   6
      TabIndex        =   15
      Top             =   3675
      Width           =   1275
   End
   Begin VB.TextBox TextBodega 
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
      MaxLength       =   8
      TabIndex        =   14
      Top             =   3675
      Width           =   1275
   End
   Begin VB.TextBox TextMinimo 
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
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4410
      Width           =   1275
   End
   Begin VB.TextBox TextMaximo 
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
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   12
      Top             =   4410
      Width           =   1275
   End
   Begin VB.TextBox TextCodigo 
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
      MaxLength       =   12
      TabIndex        =   3
      Text            =   "00.000000000"
      Top             =   2835
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   5985
      TabIndex        =   10
      Top             =   4095
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   105
      TabIndex        =   6
      Top             =   3255
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "Producto final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   945
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo de Inventario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin MSDBCtls.DBList DBLCtas 
      Bindings        =   "IngInven.frx":001C
      DataSource      =   "DataSubCta"
      Height          =   1620
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   7260
      _ExtentX        =   12806
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
      Height          =   435
      Left            =   5985
      TabIndex        =   9
      Top             =   3570
      Width           =   1170
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
      Left            =   1680
      MaxLength       =   35
      TabIndex        =   5
      Text            =   "0"
      Top             =   2835
      Width           =   5685
   End
   Begin VB.Data DataSubCta1 
      Caption         =   "SubCta1"
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
      Top             =   1680
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataSubCta 
      Caption         =   "SubCta"
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
      Top             =   1365
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD:"
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
      Left            =   2730
      TabIndex        =   19
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MAXIMO:"
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
      Left            =   2730
      TabIndex        =   18
      Top             =   4095
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA:"
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
      TabIndex        =   17
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MINIMO:"
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
      Top             =   4095
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Concepto"
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
      TabIndex        =   4
      Top             =   2520
      Width           =   5685
   End
   Begin VB.Label Label2 
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
      Left            =   105
      TabIndex        =   2
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE INVENTARIO"
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
      Top             =   525
      Width           =   7260
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0031
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0143
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0255
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0367
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0879
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":0D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":129D
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "IngProdInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  GrabarInv
  DataSubCta.Refresh
End Sub

Private Sub Command2_Click()
  Unload IngProdInv
End Sub

Private Sub DBLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DBLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLCtas_LostFocus()
    Cadena = SinEspaciosIzq(DBLCtas.Text)
    sSQL = "SELECT * FROM Productos WHERE Codigo_Inv = '" & Cadena & "' "
    SelectData DataSubCta1, sSQL, False
    LlenarInv
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As Nombre_Prod "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBList DBLCtas, DataSubCta, sSQL, "Nombre_Prod"
  DBLCtas.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm IngProdInv
  DataSubCta.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSubCta1.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextBodega_LostFocus()
  TextoValido TextBodega, True
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, True
End Sub

Private Sub TextMaximo_LostFocus()
  TextoValido TextMaximo, True
End Sub

Private Sub TextMinimo_LostFocus()
  TextoValido TextMinimo
End Sub

Private Sub TextSubCta_LostFocus()
  TextoValido TextSubCta
End Sub

Private Sub TextUnidad_LostFocus()
  TextoValido TextUnidad, True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 With DataSubCta.Recordset
 Select Case Button.Key
   Case "Eliminar"
      If DBLCtas.Enabled Then
         Cadena = SinEspaciosIzq(DBLCtas.Text)
         sSQL = "SELECT Codigo_Inv FROM Kardex "
         sSQL = sSQL & "WHERE Codigo_Inv = '" & Cadena & "' "
         SelectData DataSubCta1, sSQL, False
         If DataSubCta1.Recordset.RecordCount > 0 Then
            MsgBox "No se pude eliminar este producto"
         Else
            Mensajes = "Esta seguro que desea eliminar" & Chr(13)
            Mensajes = Mensajes & "El Producto [" & Cadena & "]"
            Titulo = "Pregunta de Eliminacion"
            If BoxMensaje = 6 Then
               sSQL = "DELETE * FROM Productos "
               sSQL = sSQL & "WHERE Codigo_Inv = '" & Cadena & "' "
               DeleteData DataSubCta1, sSQL
            End If
         End If
         DataSubCta.Refresh
         DBLCtas.SetFocus
      End If
   Case "Nuevo"
       NuevoInv
       Nuevo = True
       DBLCtas.Enabled = False
       Option1.Value = True
       TextSubCta.SetFocus
   Case "Grabar"
       GrabarInv
   Case "Primero"
       Nuevo = False
      .MoveFirst
   Case "Anterior"
       Nuevo = False
      .MovePrevious
       If .BOF Then .MoveFirst
   Case "Siguiente"
       Nuevo = False
      .MoveNext
       If .EOF Then .MoveLast
   Case "Ultimo"
       Nuevo = False
      .MoveLast
 End Select
 End With
 If Nuevo = False Then
    If DataSubCta.Recordset.RecordCount > 0 Then
       DBLCtas.Text = DataSubCta.Recordset.Fields(0)
       Cadena = SinEspaciosIzq(DBLCtas.Text)
       sSQL = "SELECT * FROM Productos WHERE Codigo_Inv = '" & Cadena & "' "
       SelectData DataSubCta1, sSQL, False
       LlenarInv
    End If
 End If
End Sub

Public Sub LlenarInv()
   With DataSubCta1.Recordset
    If .RecordCount > 0 Then
        TextSubCta.Text = .Fields("Producto")
        TextCodigo.Text = .Fields("Codigo_Inv")
        TextBodega.Text = .Fields("Bodega")
        TextUnidad.Text = .Fields("Unidad")
        TextMaximo.Text = .Fields("Maximo")
        TextMinimo.Text = .Fields("Minimo")
        If .Fields("TP") = "P" Then
            Option1.Value = False
            Option2.Value = True
        Else
            Option1.Value = True
            Option2.Value = False
        End If
    Else
        DBLCtas.Enabled = False
        TextSubCta.Text = ""
        TextBodega.Text = "0"
        TextUnidad.Text = "0"
        TextMaximo.Text = "0"
        TextMinimo.Text = "0"
        TextCodigo.Text = "00.000000000"
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
   DBLCtas.Enabled = True
End Sub

Public Sub NuevoInv()
  DBLCtas.Enabled = False
  TextSubCta.Text = ""
  TextBodega.Text = "0"
  TextUnidad.Text = "0"
  TextMaximo.Text = "0"
  TextMinimo.Text = "0"
  TextCodigo.Text = "0"
  Option1.Value = True
  TextCodigo.SetFocus
End Sub

Public Sub GrabarInv()
  sSQL = "SELECT * FROM Productos "
  sSQL = sSQL & "WHERE Codigo_Inv = '" & TextCodigo.Text & "' "
  SelectData DataSubCta1, sSQL, False
  With DataSubCta1.Recordset
   If .RecordCount > 0 Then
      .Edit
   Else
      .AddNew
      .Fields("Codigo_Inv") = TextCodigo.Text
   End If
  .Fields("Producto") = TextSubCta.Text
   Cadena = "P": If Option1.Value Then Cadena = "I"
  .Fields("TP") = Cadena
  .Fields("Bodega") = TextBodega.Text
  .Fields("Unidad") = TextUnidad.Text
  .Fields("Maximo") = TextMaximo.Text
  .Fields("Minimo") = TextMinimo.Text
  .Update
   Nuevo = False
   DataSubCta.Refresh
   DBLCtas.Enabled = True
  End With
End Sub

