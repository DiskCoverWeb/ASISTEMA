VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form IngProdu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
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
            Key             =   "Inicio"
            Object.ToolTipText     =   "Ir al inicio de los Articulos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Atras"
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
            Key             =   "Final"
            Object.ToolTipText     =   "Ir ak final de los productos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
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
      Left            =   5565
      TabIndex        =   13
      Top             =   840
      Width           =   1380
   End
   Begin MSDBCtls.DBList DBLArt 
      Bindings        =   "IngProdu.frx":0000
      DataSource      =   "DataArt"
      Height          =   1425
      Left            =   105
      TabIndex        =   1
      Top             =   735
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2514
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
      Left            =   105
      MaxLength       =   35
      TabIndex        =   8
      Top             =   3150
      Width           =   3900
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2205
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   3780
      Width           =   2220
      _ExtentX        =   3916
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
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2940
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
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   5565
      TabIndex        =   14
      Top             =   1470
      Width           =   1380
   End
   Begin VB.CheckBox CheckIVA 
      Caption         =   "Producto con I.V.A."
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
      Left            =   2415
      TabIndex        =   12
      Top             =   3570
      Width           =   3060
   End
   Begin VB.TextBox TextPVP 
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
      Left            =   4095
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "IngProdu.frx":0012
      Top             =   3150
      Width           =   1380
   End
   Begin VB.TextBox TextLinea 
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
      MaxLength       =   35
      TabIndex        =   4
      Top             =   2520
      Width           =   3900
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DEL PRODUCTO"
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
      TabIndex        =   6
      Top             =   2940
      Width           =   3900
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
            Picture         =   "IngProdu.frx":0014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":0126
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":0238
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":085C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":0D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngProdu.frx":1280
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUENTA CONTABLE"
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
      TabIndex        =   10
      Top             =   3570
      Width           =   2220
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
      Left            =   4095
      TabIndex        =   5
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGO"
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
      Left            =   4095
      TabIndex        =   3
      Top             =   2310
      Width           =   1380
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.V.P."
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
      Left            =   4095
      TabIndex        =   7
      Top             =   2940
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DEL PRODUCTO"
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
      TabIndex        =   0
      Top             =   525
      Width           =   5370
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LINEA DEL PRODUCTO"
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
      TabIndex        =   2
      Top             =   2310
      Width           =   3900
   End
End
Attribute VB_Name = "IngProdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  GrabarArticulos
End Sub

Private Sub Command2_Click()
  Unload IngProdu
End Sub

Private Sub DBLArt_DblClick()
  SendKeys "{TAB}"
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = Key_Enter Then SendKeys "{TAB}"
End Sub

Private Sub DBLArt_LostFocus()
  Codigo = ""
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Articulo ='" & DBLArt.Text & "' "
  SelectData DataArticulo, sSQL, False
  With DataArticulo.Recordset
   If .RecordCount > 0 Then Codigo = .Fields("Codigo")
  End With
  LlenarArticulos Codigo
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * FROM Articulo ORDER BY Articulo "
   SelectDBList DBLArt, DataArt, sSQL, "Articulo"
   FormatoMaskCta MBoxCta
   RatonNormal MDIFact
End Sub

Private Sub Form_Load()
   CentrarForm IngProdu
   DataArt.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\CODIGOS.MDB"
End Sub

Private Sub TextLinea_LostFocus()
  If TextLinea.Text = "" Then TextLinea.Text = Ninguno
End Sub

Private Sub TextPVP_LostFocus()
  If TextPVP.Text = "" Then TextPVP.Text = "0"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  With DataArt.Recordset
  Select Case Button.Key
    Case "Grabar"
         GrabarArticulos
         If .RecordCount > 0 Then .MoveFirst
    Case "Eliminar"
         If LabelCodigo.Caption <> "" Then
            Mensajes = "Esta seguro de Eliminar este Producto."
            Titulo = "Pregunta de grabación"
            TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
            If J = 6 Then
               sSQL = "DELETE * FROM Articulo "
               sSQL = sSQL & "WHERE Codigo = '" & LabelCodigo.Caption & "' "
               DataArticulo.Database.Execute sSQL
               sSQL = "SELECT * FROM Articulo ORDER BY Articulo "
               SelectData DataArt, sSQL, False
            End If
         End If
         Nuevo = False
         If .RecordCount > 0 Then .MoveFirst
    Case "Nuevo"
         Nuevo = True
         TextLinea.Text = ""
         TextProducto.Text = ""
         LabelCodigo.Caption = ""
         TextPVP.Text = ""
         CheckIVA.Value = 0
         Codigos = ""
         MBoxCta.Text = CodigoMaskCta("")
    Case "Inicio"
         If .RecordCount > 0 Then .MoveFirst
         Nuevo = False
    Case "Atras"
         .MovePrevious
         If .BOF Then .MoveFirst
         Nuevo = False
    Case "Siguiente"
         .MoveNext
         If .EOF Then .MoveLast
         Nuevo = False
    Case "Final"
         If .RecordCount > 0 Then .MoveLast
         Nuevo = False
  End Select
  If .RecordCount > 0 And Nuevo = False Then
      DBLArt.Text = .Fields("Articulo")
      LlenarArticulos .Fields("Codigo")
  End If
  End With
  If Nuevo Then TextLinea.SetFocus
End Sub

Public Sub LlenarArticulos(CodigoArt As String)
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo ='" & CodigoArt & "' "
  DataArticulo.RecordSource = sSQL: DataArticulo.Refresh
  With DataArticulo.Recordset
  If .RecordCount > 0 Then
     TextLinea.Text = .Fields("Linea")
     TextProducto.Text = .Fields("Articulo")
     LabelCodigo.Caption = .Fields("Codigo")
     TextPVP.Text = .Fields("PVP")
     If .Fields("IVA") Then CheckIVA.Value = 1 Else CheckIVA.Value = 0
     MBoxCta.Text = CodigoMaskCta(.Fields("Cta"))
     TextProducto.SetFocus
  Else
     MsgBox "Este Articulo no exite."
     DBLArt.SetFocus
  End If
  End With
End Sub

Public Sub GrabarArticulos()
  Mensajes = "Esta seguro de Grabar este Producto."
  Titulo = "Pregunta de grabación"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  With DataArticulo.Recordset
  If J = 6 Then
     If Nuevo Then
        Numero = ReadSetDataNum(DataCodigos, "Articulos", True)
        LabelCodigo.Caption = FormatoCodigo(TextProducto.Text, Numero)
       .AddNew
       .Fields("Codigo") = LabelCodigo.Caption
       .Fields("Stock") = 0
     Else
       .Edit
     End If
    .Fields("Linea") = TextLinea.Text
    .Fields("Articulo") = TextProducto.Text
    .Fields("PVP") = TextPVP.Text
    .Fields("IVA") = CheckIVA.Value
    .Fields("Cta") = CambioCodigoCta(MBoxCta.Text)
    .Update
     sSQL = "SELECT * FROM Articulo ORDER BY Articulo "
     DataArt.RecordSource = sSQL: DataArt.Refresh
  End If
  End With
  Nuevo = False
End Sub
