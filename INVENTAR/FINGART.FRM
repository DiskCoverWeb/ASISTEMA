VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FIngArt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Articulo"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
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
      MouseIcon       =   "FIngArt.frx":0000
   End
   Begin VB.TextBox TextProveedor 
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
      Left            =   5565
      TabIndex        =   19
      Top             =   1680
      Width           =   3585
   End
   Begin MSDBCtls.DBList DBLArt 
      Bindings        =   "FIngArt.frx":001C
      DataSource      =   "DataArticulo"
      Height          =   2985
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5265
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
   Begin VB.Data DataArtFact 
      Caption         =   "ArtFact"
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
      Top             =   3150
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   5565
      TabIndex        =   8
      Top             =   2520
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
      Left            =   8085
      TabIndex        =   15
      Top             =   2520
      Width           =   1065
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
      Left            =   8085
      TabIndex        =   9
      Top             =   2100
      Width           =   1065
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
      Left            =   5565
      TabIndex        =   10
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox TextArticulo 
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
      Left            =   5565
      TabIndex        =   7
      Top             =   1260
      Width           =   3480
   End
   Begin MSDBCtls.DBCombo DBCTipoInv 
      Bindings        =   "FIngArt.frx":0033
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   5565
      TabIndex        =   1
      Top             =   840
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327680
      Text            =   " "
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
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
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
      Top             =   3150
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
      Top             =   3465
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
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
      Top             =   2835
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
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
      Top             =   3465
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
      Height          =   435
      Left            =   7770
      TabIndex        =   17
      Top             =   3465
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
      Height          =   435
      Left            =   6405
      TabIndex        =   16
      Top             =   3465
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxCtaInv 
      Height          =   330
      Left            =   5565
      TabIndex        =   21
      Top             =   2940
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
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
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CTA. DE INV.:"
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
      TabIndex        =   22
      Top             =   2940
      Width           =   1380
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR"
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
      TabIndex        =   20
      Top             =   1680
      Width           =   1380
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
      Left            =   6930
      TabIndex        =   14
      Top             =   2520
      Width           =   1065
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
      Left            =   6930
      TabIndex        =   6
      Top             =   2100
      Width           =   1065
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
      Left            =   4095
      TabIndex        =   11
      Top             =   840
      Width           =   1380
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
      Left            =   4095
      TabIndex        =   5
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label Label3 
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
      Left            =   4095
      TabIndex        =   4
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARTICULO:"
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
      TabIndex        =   3
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE INVENTARIO:"
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
      Left            =   5565
      TabIndex        =   13
      Top             =   525
      Width           =   3480
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
            Picture         =   "FIngArt.frx":0045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0157
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0269
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":037B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":088D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0D9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":12B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO:"
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
      TabIndex        =   12
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARTICULO"
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
End
Attribute VB_Name = "FIngArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LlenarArticulo(CodArt As String)
  RatonReloj
  LabelCodigo.Visible = True
  DBLArt.Enabled = True
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArt, sSQL, False
  With DataArt.Recordset
   If .RecordCount > 0 Then
       TipoInventario = .Fields("TP")
       LabelCodigo.Caption = .Fields("Codigo")
       TextArticulo.Text = .Fields("Articulo")
       TextUnidad.Text = .Fields("Unidad")
       TextBodega.Text = .Fields("Bodega")
       TextMinimo.Text = .Fields("Minimo")
       TextMaximo.Text = .Fields("Maximo")
       TextProveedor.Text = .Fields("Proveedor")
       MBoxCtaInv.Text = FormatoCodigoCta(.Fields("Codigo_Cta"))
       sSQL = "SELECT * FROM Inventario "
       sSQL = sSQL & "WHERE TP = '" & TipoInventario & "' "
       SelectData DataArt, sSQL, False
       DBCTipoInv.Text = DataArt.Recordset.Fields("Concepto")
    End If
  End With
  Nuevo = False
  RatonNormal
End Sub

Public Sub NuevoArticulo()
  Nuevo = True
  LabelCodigo.Caption = ""
  TextArticulo.Text = ""
  TextUnidad.Text = ""
  TextBodega.Text = ""
  TextMinimo.Text = ""
  TextMaximo.Text = ""
  TextProveedor.Text = ""
  FormatoMaskCta MBoxCtaInv
End Sub

Public Sub GrabarArticulo(CodArt As String)
  RatonReloj
  DBLArt.Enabled = True
  Nuevo = False
  LabelCodigo.Caption = CodArt
  If CodArt = "" Then CodArt = Ninguno
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArtFact, sSQL, False
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArt, sSQL, False
  With DataArt.Recordset
       If .RecordCount > 0 Then
          .Edit
          Codigo = .Fields("Codigo")
       Else
          .AddNew
           Numero = ReadSetDataNum("Articulos", True, True)
           Codigo = FormatoCodigo(TextArticulo.Text, Numero)
       End If
      .Fields("Codigo") = Codigo
      .Fields("Articulo") = TextArticulo.Text
      .Fields("Unidad") = TextUnidad.Text
      .Fields("Bodega") = TextBodega.Text
      .Fields("Minimo") = TextMinimo.Text
      .Fields("Maximo") = TextMaximo.Text
      .Fields("Proveedor") = TextProveedor.Text
      .Fields("Codigo_Cta") = CambioCodigoCta(MBoxCtaInv.Text)
       sSQL = "SELECT * FROM Inventario "
       sSQL = sSQL & "WHERE Concepto = '" & DBCTipoInv.Text & "' "
       SelectData DataSQL, sSQL, False
      .Fields("TP") = DataSQL.Recordset.Fields("TP")
      .Update
       LabelCodigo.Caption = Codigo
  End With
  LabelCodigo.Visible = True
  DataArticulo.Refresh
  RatonNormal
End Sub

Private Sub Command1_Click()
  GrabarArticulo Codigo
End Sub

Private Sub Command2_Click()
  Unload FIngArt
End Sub

Private Sub DBCTipoInv_LostFocus()
   TipoInventario = MP
   sSQL = "SELECT * FROM Inventario "
   sSQL = sSQL & "WHERE Concepto = '" & DBCTipoInv.Text & "' "
   SelectData DataSQL, sSQL, False
   If DataSQL.Recordset.RecordCount > 0 Then TipoInventario = DataSQL.Recordset.Fields("TP")
End Sub

Private Sub DBLArt_DblClick()
  SiguienteControl
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub DBLArt_LostFocus()
 If (Nuevo = False) And (DataArticulo.Recordset.RecordCount > 0) Then
    Codigo = SinEspaciosIzq(DBLArt.Text)
    LlenarArticulo SinEspaciosIzq(Codigo)
 End If
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * FROM Articulo "
   SelectData DataArtFact, sSQL, False
   sSQL = "SELECT Codigo & '   ' & Articulo As NombArt FROM Articulo "
   sSQL = sSQL & "ORDER BY TP,Articulo "
   SelectDBList DBLArt, DataArticulo, sSQL, "NombArt"
   sSQL = "SELECT Concepto FROM Inventario "
   sSQL = sSQL & "ORDER BY Concepto "
   SelectDBCombo DBCTipoInv, DataInv, sSQL, "Concepto", False
   FormatoMaskCta MBoxCtaInv
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FIngArt
   DataArt.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
   DataInv.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
   DataSQL.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
   DataArtFact.DatabaseName = RutaEmpresa & "\INVENTAR.MDB"
End Sub

Private Sub TextArticulo_LostFocus()
 If TextArticulo.Text = "" Then TextArticulo.Text = Ninguno
End Sub

Private Sub TextBodega_LostFocus()
  If TextBodega.Text = "" Then TextBodega.Text = Ninguno
End Sub

Private Sub TextProveedor_LostFocus()
  If TextProveedor.Text = "" Then TextProveedor.Text = niguno
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 Select Case Button.Key
   Case "Eliminar"
      If DBLArt.Enabled Then
         sSQL = "SELECT Codigo FROM Kardex "
         sSQL = sSQL & "WHERE  Codigo = '" & Codigo & "' "
         SelectData DataArt, sSQL, False
         If DataArt.Recordset.RecordCount > 0 Then
            Mensajes = "No se puede eliminar este Producto (" & Codigo & ")," & Chr(13)
            Mensajes = Mensajes & "porque tiene valores procesables."
            MsgBox Mensajes
         Else
            Mensajes = "Esta seguro que desea eliminar" & Chr(13)
            Mensajes = Mensajes & "El Producto [" & Codigo & "]"
            Titulo = "Pregunta de Eliminacion"
            If BoxMensaje = 6 Then
               sSQL = "DELETE * FROM Articulo "
               sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
               DeleteData DataArt, sSQL
               sSQL = "DELETE * FROM Articulo "
               sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
               DeleteData DataArtFact, sSQL
               sSQL = "SELECT Codigo & ' ' & Articulo As NombArt FROM Articulo "
               sSQL = sSQL & "ORDER BY TP,Articulo "
               SelectData DataArticulo, sSQL, False
            End If
         End If
         DBLArt.SetFocus
      End If
   Case "Nuevo"
       NuevoArticulo
       DBLArt.Enabled = False
       LabelCodigo.Visible = False
       DBCTipoInv.SetFocus
       Codigo = Ninguno
   Case "Grabar"
       GrabarArticulo Codigo
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
 If (Nuevo = False) And (DataArticulo.Recordset.RecordCount > 0) Then
    DBLArt.Text = DataArticulo.Recordset.Fields(0)
    Codigo = SinEspaciosIzq(DBLArt.Text)
    LlenarArticulo Codigo
 End If
End Sub

