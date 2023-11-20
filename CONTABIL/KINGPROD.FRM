VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form IngProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
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
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   15
      Top             =   5670
      Width           =   855
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
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   14
      Top             =   5250
      Width           =   855
   End
   Begin MSDBCtls.DBList DBLProd 
      Bindings        =   "KIngProd.frx":001C
      DataSource      =   "DataProd"
      Height          =   2400
      Left            =   105
      TabIndex        =   10
      Top             =   1050
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   4233
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
      Caption         =   "CANCELAR"
      Height          =   435
      Left            =   4200
      TabIndex        =   8
      Top             =   5460
      Width           =   1275
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
      Left            =   1785
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3990
      Width           =   3690
   End
   Begin VB.Data DataProd 
      Caption         =   "Prod"
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
   Begin VB.Data DataProceso 
      Caption         =   "Proceso"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3570
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GRABAR"
      Height          =   435
      Left            =   2835
      TabIndex        =   7
      Top             =   5460
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   1785
      TabIndex        =   9
      Top             =   4830
      Width           =   2325
      _ExtentX        =   4101
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
   Begin VB.Data DataProducto 
      Caption         =   "Producto"
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
      Bindings        =   "KIngProd.frx":002F
      DataSource      =   "DataProceso"
      Height          =   315
      Left            =   1785
      TabIndex        =   12
      Top             =   4410
      Width           =   3690
      _ExtentX        =   6509
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
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bodega No."
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
      TabIndex        =   16
      Top             =   5670
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Kardex"
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
      TabIndex        =   11
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
      Width           =   5370
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
            Picture         =   "KIngProd.frx":0045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":0157
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":0269
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":037B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":088D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":0D9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "KIngProd.frx":12B1
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
      TabIndex        =   6
      Top             =   4830
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
      Left            =   1785
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
      Top             =   5250
      Width           =   1695
   End
End
Attribute VB_Name = "IngProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LlenarArticulo(CodProd As String)
   If Nuevo Then
      LabelCodigo.Caption = ""
      TextProducto.Text = ""
      TextUnidad.Text = ""
      TextBodega.Text = ""
      MBoxCta.Text = FormatoCodigoCta(" ")
    Else
      sSQL = "SELECT Materia_Prima.*,Tipo_Kardex.* FROM Materia_Prima,Tipo_Kardex "
      sSQL = sSQL & "WHERE Codigo = '" & CodProd & "' "
      sSQL = sSQL & "AND TK = COD "
      DataProducto.RecordSource = sSQL: DataProducto.Refresh
      If DataProducto.Recordset.RecordCount > 0 Then
         LabelCodigo.Caption = DataProducto.Recordset.Fields("Codigo")
         TextProducto.Text = DataProducto.Recordset.Fields("Producto")
         TextUnidad.Text = DataProducto.Recordset.Fields("Unidad")
         MBoxCta.Text = DataProducto.Recordset.Fields("Cuenta")
         DBCTP.Text = DataProducto.Recordset.Fields("TK") & " " & DataProducto.Recordset.Fields("PROCESO")
         TextBodega.Text = DataProducto.Recordset.Fields("Bodega")
      End If
   End If
   Nuevo = False
   TextProducto.SetFocus
End Sub

Public Sub GrabarArticulo()
   Codigos = LabelCodigo.Caption
   Mensajes = "Esta seguro que desea grabar El Producto: " & Chr(13)
   Mensajes = Mensajes & TextProducto.Text
   Titulo = "Pregunta de Grabación"
   TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
   If J = 6 Then
      sSQL = "SELECT * FROM Materia_Prima "
      sSQL = sSQL & "WHERE Codigo = '" & Codigos & "' "
      DataProd.RecordSource = sSQL: DataProd.Refresh
      With DataProd.Recordset
       If .RecordCount > 0 Then
          .Edit
       Else
          .AddNew
           Numero = ReadSetDataNum(DataCodigos, "Materia Prima", True)
          .Fields("Codigo") = FormatoCodigo(TextProducto.Text, Numero)
       End If
      .Fields("Producto") = TextProducto.Text
      .Fields("Unidad") = TextUnidad.Text
      .Fields("TK") = Val(SinEspaciosIzq(DBCTP.Text))
      .Fields("Stock") = 0
      .Fields("Bodega") = Val(TextBodega.Text)
      .Fields("Cuenta") = FormatoCodigoCta(MBoxCta.Text)
      .Update
      End With
      Nuevo = False
   End If
End Sub

Private Sub Command1_Click()
   Unload IngProductos
End Sub

Private Sub Command2_Click()
  GrabarArticulo
  sSQL = "SELECT Codigo & '  ' & Producto As Productos FROM Materia_Prima ORDER BY Producto "
  DataProd.RecordSource = sSQL: DataProd.Refresh
End Sub

Private Sub DBLProd_DblClick()
  LlenarArticulo SinEspaciosIzq(DBLProd.Text)
End Sub

Private Sub Form_Activate()
   FormatoMaskCta MBoxCta
   Nuevo = False
   sSQL = "SELECT COD & '  ' & PROCESO As Procesos FROM Tipo_Kardex ORDER BY COD "
   DataProceso.RecordSource = sSQL: DataProceso.Refresh
   DBCTP.ListField = "Procesos"
   If DataProceso.Recordset.RecordCount > 0 Then DBCTP.Text = DataProceso.Recordset.Fields("Procesos")
   sSQL = "SELECT Codigo & '  ' & Producto As Productos FROM Materia_Prima ORDER BY Producto "
   DataProd.RecordSource = sSQL: DataProd.Refresh
   DBLProd.ListField = "Productos"
   If DataProd.Recordset.RecordCount > 0 Then
      DBLProd.Text = DataProd.Recordset.Fields("Productos")
   End If
   LlenarArticulo SinEspaciosIzq(DBLProd.Text)
   RatonNormal MDIKardex
End Sub

Private Sub Form_Load()
   CentrarForm IngProductos
   DataCodigos.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataProducto.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataProd.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
   DataProceso.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  Select Case Button.Key
    Case "Grabar"
        GrabarArticulo
        sSQL = "SELECT Codigo & '  ' & Producto As Productos FROM Materia_Prima ORDER BY Producto "
        DataProd.RecordSource = sSQL: DataProd.Refresh
    Case "Eliminar"
        Codigos = LabelCodigo.Caption
        Mensajes = "Esta seguro que desea eliminar," & Chr$(13)
        Mensajes = Mensajes & "El Producto: (" & Codigos & ") '" & TextProducto.Text & "'."
        Titulo = "Eliminacion de Productos"
        TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
        If J = 6 Then
           sSQL = "DELETE * FROM Materia_Prima "
           sSQL = sSQL & "WHERE Codigo = '" & Codigos & "' "
           DataProd.Database.Execute sSQL
           sSQL = "SELECT Codigo & '  ' & Producto As Productos FROM Materia_Prima ORDER BY Producto "
           DataProd.RecordSource = sSQL: DataProd.Refresh
        End If
    Case "Nuevo"
        Nuevo = True
    Case "Primero"
        DataProd.Recordset.MoveFirst
    Case "Anterior"
        DataProd.Recordset.MovePrevious
        If DataProd.Recordset.BOF Then DataProd.Recordset.MoveFirst
    Case "Siguiente"
        DataProd.Recordset.MoveNext
        If DataProd.Recordset.EOF Then DataProd.Recordset.MoveLast
    Case "Ultimo"
        DataProd.Recordset.MoveLast
  End Select
  If DataProd.Recordset.RecordCount > 0 Then
     DBLProd.Text = DataProd.Recordset.Fields("Productos")
     LlenarArticulo SinEspaciosIzq(DBLProd.Text)
  End If
End Sub

