VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form IngProdu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataList DLArt 
      Bindings        =   "IngProdu.frx":0000
      DataSource      =   "AdoArt"
      Height          =   1815
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3201
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
      Left            =   5250
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "IngProdu.frx":0015
      Top             =   3375
      Width           =   1380
   End
   Begin VB.CheckBox CheckIVA 
      Caption         =   "P&roducto con I.V.A."
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
      Left            =   5250
      TabIndex        =   10
      Top             =   3795
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
      Left            =   105
      MaxLength       =   40
      TabIndex        =   9
      Top             =   3375
      Width           =   5055
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
      Left            =   6720
      Picture         =   "IngProdu.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
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
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2640
      Width           =   2430
   End
   Begin VB.TextBox TextCodInv 
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
      Left            =   3045
      MaxLength       =   15
      TabIndex        =   6
      Top             =   3795
      Width           =   2115
   End
   Begin VB.TextBox TextCta 
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
      Left            =   3045
      MaxLength       =   16
      TabIndex        =   5
      Top             =   4110
      Width           =   2115
   End
   Begin VB.TextBox TextCosto 
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
      Left            =   3045
      MaxLength       =   16
      TabIndex        =   4
      Top             =   4425
      Width           =   2115
   End
   Begin VB.TextBox TextCtaIng 
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
      Left            =   3045
      MaxLength       =   16
      TabIndex        =   3
      Top             =   4740
      Width           =   2115
   End
   Begin VB.TextBox TextCxP 
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
      Left            =   3045
      MaxLength       =   16
      TabIndex        =   2
      Top             =   5055
      Width           =   2115
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
      Left            =   6720
      Picture         =   "IngProdu.frx":0459
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   210
      Top             =   1260
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Articulo"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   225
      Top             =   630
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   225
      Top             =   960
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "Art"
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
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "IngProdu.frx":06DB
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   2655
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C&ODIGO DE &INVENTARIO"
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
      TabIndex        =   22
      Top             =   3795
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&CUENTA POR COBRAR"
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
      Left            =   2625
      TabIndex        =   21
      Top             =   2325
      Width           =   4005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&NOMBRE DEL PRODUCTO"
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
      TabIndex        =   20
      Top             =   105
      Width           =   6525
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.&V.P."
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
      Left            =   5250
      TabIndex        =   19
      Top             =   3060
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&CODIGO"
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
      TabIndex        =   18
      Top             =   2310
      Width           =   2430
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DEL &PRODUCTO"
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
      TabIndex        =   17
      Top             =   3060
      Width           =   5055
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C&UENTA INVENTARIO"
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
      Top             =   4110
      Width           =   2955
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA DE COSTO DE &VENTA"
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
      TabIndex        =   15
      Top             =   4425
      Width           =   2955
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUEN&TA DE INGRESO"
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
      Top             =   4740
      Width           =   2955
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA DE VENTAS ANTIC."
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
      Top             =   5055
      Width           =   2955
   End
End
Attribute VB_Name = "IngProdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  TextoValido TextCodInv, , True
  TextoValido TextCta, , True
  TextoValido TextCosto, , True
  TextoValido TextCtaIng, , True
  TextoValido TextCxP, , True
  'TextoValido TextCodigoB, , True
  GrabarArticulos
End Sub

Private Sub Command2_Click()
  Unload IngProdu
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLArt_DblClick()
  SiguienteControl
End Sub

Private Sub DLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar el Producto: " _
              & DLArt.Text & "."
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then
        Codigo = SinEspaciosIzq(DLArt.Text)
        sSQL = "DELETE * " _
             & "FROM Catalogo_Productos " _
             & "WHERE Codigo_Inv ='" & Codigo & "' "
        ConectarAdoExecute sSQL
        DLArt.Enabled = True
        sSQL = "SELECT (Codigo_Inv & '     ' & Producto) As CodArt " _
             & "FROM Catalogo_Productos " _
             & "ORDER BY Codigo_Inv "
        SelectData AdoArt, sSQL
     End If
  End If
End Sub

Private Sub DLArt_LostFocus()
  Codigo = Ninguno
  If SinEspaciosIzq(DLArt.Text) <> "" Then Codigo = SinEspaciosIzq(DLArt.Text)
  LlenarArticulos Codigo
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT (Codigo_Inv & '      ' & Producto) As CodArt " _
        & "FROM Catalogo_Productos " _
        & "ORDER BY Codigo_Inv "
   SelectDBList DLArt, AdoArt, sSQL, "CodArt"
   sSQL = "SELECT (Codigo & '     ' & Linea) As CodLinea FROM Linea_Producto ORDER BY Codigo "
   SelectDBCombo DCLinea, AdoLinea, sSQL, "CodLinea", False
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm IngProdu
   ConectarAdodc AdoArt
   ConectarAdodc AdoLinea
   ConectarAdodc AdoArticulo
End Sub

Private Sub Label11_Click()

End Sub

Private Sub TextCodigo_GotFocus()
  MarcarTexto TextCodigo
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, , True
End Sub

Private Sub TextCodigoB_GotFocus()
  MarcarTexto TextCodigoB
End Sub

Private Sub TextCodigoB_LostFocus()
  TextoValido TextCodigoB
End Sub

Private Sub TextCodInv_GotFocus()
  MarcarTexto TextCodInv
End Sub

Private Sub TextCodInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCodInv_LostFocus()
  TextoValido TextCodInv
End Sub

Private Sub TextCosto_GotFocus()
  MarcarTexto TextCosto
End Sub

Private Sub TextCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCosto_LostFocus()
  TextoValido TextCosto
End Sub

Private Sub TextCta_GotFocus()
  MarcarTexto TextCta
End Sub

Private Sub TextCta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCta_LostFocus()
  TextoValido TextCta
End Sub

Private Sub TextCtaIng_GotFocus()
  MarcarTexto TextCtaIng
End Sub

Private Sub TextCtaIng_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCtaIng_LostFocus()
  TextoValido TextCtaIng, True
End Sub

Private Sub TextCxP_GotFocus()
  MarcarTexto TextCxP
End Sub

Private Sub TextCxP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCxP_LostFocus()
  TextoValido TextCxP, True
End Sub

Private Sub TextProducto_GotFocus()
  MarcarTexto TextProducto
End Sub

Private Sub TextProducto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPVP_GotFocus()
  MarcarTexto TextPVP
End Sub

Private Sub TextPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPVP_LostFocus()
  TextoValido TextPVP, True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim Nuevo As Boolean
  Nuevo = False
  DLArt.Enabled = True
  With AdoArt.Recordset
  Select Case Button.key
    Case "Nuevo"
         TextCodigo.Text = ""
         TextProducto.Text = ""
         TextPVP.Text = ""
         TextCodInv.Text = ""
         TextCta.Text = ""
         TextCosto.Text = ""
         Nuevo = True
    Case "Eliminar"
         
    Case "Grabar"
         GrabarArticulos
         If .RecordCount > 0 Then .MoveFirst
    Case "Inicio"
         If .RecordCount > 0 Then .MoveFirst
         Nuevo = False
    Case "Atras"
         If .RecordCount > 0 Then
            .MovePrevious
             If .BOF Then .MoveFirst
         End If
         Nuevo = False
    Case "Siguiente"
         If .RecordCount > 0 Then
            .MoveNext
             If .EOF Then .MoveLast
         End If
         Nuevo = False
    Case "Final"
         If .RecordCount > 0 Then .MoveLast
         Nuevo = False
  End Select
  If .RecordCount > 0 And Nuevo = False Then
      DLArt.Text = AdoArt.Recordset.Fields("CodArt")
      LlenarArticulos SinEspaciosIzq(DLArt.Text)
  End If
  End With
  If Nuevo Then
     DLArt.Enabled = False
     TextCodigo.SetFocus
  End If
End Sub

Public Sub LlenarArticulos(CodigoArt As String)
  sSQL = "SELECT * FROM Catalogo_Productos "
  sSQL = sSQL & "WHERE Codigo_Inv ='" & CodigoArt & "' "
  SelectData AdoArticulo, sSQL, False
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
       TextCodigo.Text = .Fields("Codigo_Inv")
       CodigoL = .Fields("CodigoL")
       TextProducto.Text = .Fields("Producto")
       TextPVP.Text = Format(.Fields("PVP"), "#,##0.00")
       TextCodInv.Text = .Fields("Codigo_Inv")
'       TextCodigoB.Text = .Fields("CodigoB")
       TextCta.Text = .Fields("Cta_Inv")
       TextCosto.Text = .Fields("Cta_Costo")
       TextCtaIng.Text = .Fields("Cta_Ingreso")
'       TextCxP.Text = .Fields("Cta_CxP")
       If .Fields("IVA") Then CheckIVA.Value = 1 Else CheckIVA.Value = 0
       sSQL = "SELECT * FROM Linea_Producto "
       sSQL = sSQL & "WHERE Codigo ='" & CodigoL & "' "
       SelectData AdoArticulo, sSQL, False
       If AdoArticulo.Recordset.RecordCount > 0 Then
          DCLinea.Text = AdoArticulo.Recordset.Fields("Codigo") & "  " & AdoArticulo.Recordset.Fields("Linea")
       End If
       TextProducto.SetFocus
   Else
       MsgBox "Este Articulo no exite."
   End If
  End With
End Sub

Public Sub GrabarArticulos()
  Codigo = TextCodigo.Text
  Mensajes = "Esta seguro de Grabar el Producto: " _
           & TextProducto.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "SELECT * FROM Catalogo_Productos "
     sSQL = sSQL & "WHERE Codigo_Inv = '" & Codigo & "' "
     SelectData AdoArticulo, sSQL, False
     With AdoArticulo.Recordset
          If .RecordCount > 0 Then
             '.Edit
          Else
             .AddNew
             .Fields("Codigo_Inv") = TextCodigo.Text
          End If
         .Fields("CodigoL") = SinEspaciosIzq(DCLinea.Text)
         .Fields("Producto") = TextProducto.Text
         .Fields("Cta_Inv") = TextCta.Text
         .Fields("Cta_Costo") = TextCosto.Text
         .Fields("Cta_Ingreso") = TextCtaIng.Text
      '   .Fields("Cta_CxP") = TextCxP.Text
        ' .Fields("Codigo_Inv") = TextCodInv.Text
      '   .Fields("CodigoB") = TextCodigoB.Text
         .Fields("PVP") = TextPVP.Text
         .Fields("IVA") = CheckIVA.Value
         .Update
          DLArt.Enabled = True
          sSQL = "SELECT (Codigo_Inv & '     ' & Producto) As CodArt FROM Catalogo_Productos ORDER BY Codigo_Inv "
          SelectData AdoArt, sSQL, False
     End With
  End If
  Nuevo = False
End Sub
