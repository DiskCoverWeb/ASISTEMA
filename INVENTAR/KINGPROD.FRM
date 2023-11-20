VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form IngProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Articulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Articulo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Articulo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Ir al inicio de los Articulos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Retroceder un Articulo"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Avanzar un Articulo"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ir ak final de los productos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MouseIcon       =   "KIngProd.frx":0000
   End
   Begin MSDBCtls.DBList DBList1 
      Height          =   2400
      Left            =   105
      TabIndex        =   11
      Top             =   1050
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   4233
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCELAR"
      Height          =   435
      Left            =   5880
      TabIndex        =   9
      Top             =   2310
      Width           =   1380
   End
   Begin VB.TextBox TextProducto 
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
      Left            =   1890
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3990
      Width           =   3480
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
      Top             =   420
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataUnidad 
      Caption         =   "Unidad"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBCtls.DBCombo DBCUnidad 
      Bindings        =   "KIngProd.frx":001C
      DataSource      =   "DataUnidad"
      Height          =   315
      Left            =   1890
      TabIndex        =   6
      Top             =   4830
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo2"
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
   Begin VB.CommandButton Command2 
      Caption         =   "GRABAR"
      Height          =   435
      Left            =   5880
      TabIndex        =   8
      Top             =   1785
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   1890
      TabIndex        =   10
      Top             =   5250
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   420
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
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
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBCtls.DBCombo DBCTP 
      Bindings        =   "KIngProd.frx":0031
      DataSource      =   "DataUnidad"
      Height          =   315
      Left            =   1890
      TabIndex        =   13
      Top             =   4410
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo2"
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Proceso"
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
      Top             =   4410
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombre del Producto"
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
      Top             =   735
      Width           =   5265
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Producto"
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
      TabIndex        =   3
      Top             =   3990
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4200
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
            Picture         =   "KIngProd.frx":0046
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":0158
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":026A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":037C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":088E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":0DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":12B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Contable"
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
      TabIndex        =   7
      Top             =   5250
      Width           =   1695
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
      Left            =   1890
      TabIndex        =   1
      Top             =   3570
      Width           =   1380
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Codigo"
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
      Top             =   3570
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Unidad de Medida"
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
      TabIndex        =   5
      Top             =   4830
      Width           =   1695
   End
End
Attribute VB_Name = "IngProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LlenarArticulo(Clientes As String)
   If Nuevo Then
      TextC1.Text = ""
      TextC2.Text = ""
      TextC3.Text = ""
      TextC4.Text = ""
      TextC5.Text = ""
      TextC6.Text = ""
      TextProducto.Text = ""
    Else
      sSQL = "SELECT * FROM Materia_Prima WHERE Producto = '" & Clientes & "' "
      DataArt.RecordSource = sSQL: DataArt.Refresh
      If DataArt.Recordset.RecordCount > 0 Then
         Codigos = DataArt.Recordset.Fields("Cuenta")
         TextC1.Text = Mid(Codigos, 1, 1)
         TextC2.Text = Mid(Codigos, 3, 3)
         TextC3.Text = Mid(Codigos, 7, 2)
         TextC4.Text = Mid(Codigos, 10, 2)
         TextC5.Text = Mid(Codigos, 13, 2)
         TextC6.Text = Mid(Codigos, 16, 2)
         LabelCodigo.Caption = DataArt.Recordset.Fields("Codigo")
         TextProducto.Text = DataArt.Recordset.Fields("Producto")
         DBCUnidad.Text = DataArt.Recordset.Fields("Unidad")
      End If
   End If
   TextProducto.SetFocus
End Sub

Public Sub GrabarArticulo()
   Codigo1 = "0"
   If TextC1.Text <> "" Then Codigo1 = TextC1.Text
   If TextC2.Text <> "" Then Codigo1 = Codigo1 & "." & TextC2.Text
   If TextC3.Text <> "" Then Codigo1 = Codigo1 & "." & TextC3.Text
   If TextC4.Text <> "" Then Codigo1 = Codigo1 & "." & TextC4.Text
   If TextC5.Text <> "" Then Codigo1 = Codigo1 & "." & TextC5.Text
   If TextC6.Text <> "" Then Codigo1 = Codigo1 & "." & TextC6.Text
   If Nuevo Then
      DBCArticulo.Text = TextProducto.Text
      sSQL = "SELECT * FROM Materia_Prima "
      DataArt.RecordSource = sSQL: DataArt.Refresh
      DataArt.Recordset.AddNew
      sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Materia Prima' "
      DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
      I = DataCodigos.Recordset.Fields("Numero")
      LabelCodigo.Caption = FormatoCodigo(DBCArticulo.Text, I)
      DataArt.Recordset.Fields("Codigo") = LabelCodigo.Caption
      sSQL = "UPDATE Codigos SET Numero = Numero + 1 WHERE Concepto = 'Materia Prima' "
      DataCodigos.Database.Execute sSQL
   Else
      DataArt.Recordset.Edit
   End If
   Mensajes = "Esta seguro que desea grabar El Producto: " & Chr(13)
   Mensajes = Mensajes & DBCArticulo.Text
   Titulo = "Pregunta de Grabación"
   TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
   If J = 6 Then
      DataArt.Recordset.Fields("Producto") = TextProducto.Text
      DataArt.Recordset.Fields("Unidad") = DBCUnidad.Text
      DataArt.Recordset.Fields("Cuenta") = Codigo1
      DataArt.Recordset.Update
      Nuevo = False
   End If
End Sub

Private Sub Command1_Click()
   Unload IngProductos
End Sub

Private Sub Command2_Click()
  GrabarArticulo
  sSQL = "SELECT * FROM Materia_Prima ORDER BY Producto "
  DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
End Sub

Private Sub DBCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then
     If DataArticulo.Recordset.RecordCount > 0 Then LlenarArticulo DBCArticulo.Text
  End If
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * FROM Unidades ORDER BY Unidad "
   DataUnidad.RecordSource = sSQL: DataUnidad.Refresh
   DBCUnidad.ListField = "Unidad"
   If DataUnidad.Recordset.RecordCount > 0 Then DBCUnidad.Text = DataUnidad.Recordset.Fields("Unidad")
   sSQL = "SELECT * FROM Materia_Prima ORDER BY Producto "
   DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
   DBCArticulo.ListField = "Producto"
   If DataArticulo.Recordset.RecordCount > 0 Then
      DBCArticulo.Text = DataArticulo.Recordset.Fields("Producto")
      Nuevo = False
   Else
      Nuevo = True
   End If
   LlenarArticulo DBCArticulo.Text
   MDIKardex.MousePointer = vbDefault
End Sub

Private Sub TextC1_GotFocus()
  TextC1.Text = ""
End Sub

Private Sub TextC2_GotFocus()
  TextC2.Text = ""
End Sub

Private Sub TextC3_GotFocus()
  TextC3.Text = ""
End Sub

Private Sub TextC4_GotFocus()
  TextC4.Text = ""
End Sub

Private Sub TextC5_GotFocus()
  TextC5.Text = ""
End Sub

Private Sub TextC6_Change()
 If Len(TextC6.Text) >= TextC6.MaxLength Then CheckIVA.SetFocus
End Sub

Private Sub TextC6_GotFocus()
  TextC6.Text = ""
End Sub

Private Sub TextC6_LostFocus()
  Command2.SetFocus
End Sub

Private Sub Form_Load()
   'Centrar el formulario
   Left = (Screen.Width - Width) / 2
   Top = (Screen.Height - Height) / 2 - 600
   DataCodigos.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataUnidad.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataArt.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  Select Case Button.Key
    Case "Grabar"
        GrabarArticulo
        sSQL = "SELECT * FROM Materia_Prima ORDER BY Producto "
        DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
    Case "Eliminar"
        Codigos = LabelCodigo.Caption
        Mensajes = "Esta seguro que desea eliminar," & Chr$(13)
        Mensajes = Mensajes & "El Producto: (" & Codigos & ") '" & TextProducto.Text & "'."
        Titulo = "Eliminacion de Clientes"
        TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
        If J = 6 Then
           sSQL = "DELETE * FROM Materia_Prima "
           sSQL = sSQL & "WHERE Codigo = '" & Codigos & "';"
           DataArticulo.Database.Execute sSQL
           sSQL = "SELECT * FROM Materia_Prima ORDER BY Producto "
           DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
        End If
    Case "Nuevo"
        Nuevo = True
    Case "Primero"
        DataArticulo.Recordset.MoveFirst
    Case "Anterior"
        DataArticulo.Recordset.MovePrevious
        If DataArticulo.Recordset.BOF Then DataArticulo.Recordset.MoveFirst
    Case "Siguiente"
        DataArticulo.Recordset.MoveNext
        If DataArticulo.Recordset.EOF Then DataArticulo.Recordset.MoveLast
    Case "Ultimo"
        DataArticulo.Recordset.MoveLast
  End Select
  If DataArticulo.Recordset.RecordCount > 0 Then
     DBCArticulo.Text = DataArticulo.Recordset.Fields("Producto")
     LlenarArticulo DBCArticulo.Text
  End If
End Sub

