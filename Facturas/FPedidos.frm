VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PEDIDOS DE PRODUCTOS"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPVP 
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
      Left            =   3885
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "FPedidos.frx":0000
      Top             =   3255
      Width           =   1695
   End
   Begin VB.TextBox TxtCantidad 
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
      Left            =   1995
      TabIndex        =   6
      Top             =   3255
      Width           =   1695
   End
   Begin MSDataListLib.DataList DLProducto 
      Bindings        =   "FPedidos.frx":0007
      DataSource      =   "AdoProducto"
      Height          =   2400
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   4233
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   105
      Top             =   735
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Producto"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   7665
      Picture         =   "FPedidos.frx":0021
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Salir"
      DisabledPicture =   "FPedidos.frx":0463
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
      Left            =   7665
      MouseIcon       =   "FPedidos.frx":0EAD
      Picture         =   "FPedidos.frx":18F7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox TxtOrden 
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
      TabIndex        =   4
      Top             =   3255
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   105
      Top             =   1050
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "AdoAux"
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
   Begin VB.Label Label4 
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
      Left            =   1995
      TabIndex        =   5
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " P.V.P."
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
      Left            =   3885
      TabIndex        =   7
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " P.V.P."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5670
      TabIndex        =   1
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MESA No."
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
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE EL PRODUCTO"
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
      Width           =   5580
   End
End
Attribute VB_Name = "FPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
Dim Cod_Sub_Inv As String
  TextoValido TxtCantidad, True
  If Not IsNumeric(TxtOrden) Then TxtOrden = TrimStrg(MidStrg(TxtOrden, 1, 6))
  With AdoProducto.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo_Inv = '" & CodigoP & "' ")
       If Not .EOF Then
          Mensajes = "Esta seguro de Grabar El Pedido/Orden No."
          Titulo = "Pregunta de grabación"
          If BoxMensaje = vbYes And Val(TxtCantidad.Text) > 0 Then
             RatonReloj
             Cod_Sub_Inv = Ninguno
             Ln_No = 0
             sSQL = "SELECT * " _
                  & "FROM Trans_Pedidos " _
                  & "WHERE Item = '" & NumEmpresa & "' "
             If IsNumeric(TxtOrden) Then
                sSQL = sSQL & "AND Orden_No = " & Val(TxtOrden) & " "
             Else
                sSQL = sSQL & "AND No_Hab = '" & UCaseStrg(TxtOrden) & "' "
             End If
             sSQL = sSQL & "ORDER BY ID DESC "
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then Ln_No = AdoAux.Recordset.fields("ID") + 1
             Cod_Sub_Inv = CodigoCuentaSup(CodigoP)
             SetAdoAddNew "Trans_Pedidos"
             SetAdoFields "Fecha", FechaSistema
             SetAdoFields "Codigo", CodigoP
             SetAdoFields "Hora", Format$(Time, "HH:SS")
             SetAdoFields "Producto", .fields("Producto")
             SetAdoFields "Cantidad", Val(TxtCantidad)
             SetAdoFields "Precio", Val(TxtPVP)
             Total_IVA = 0
             If .fields("IVA") Then Total_IVA = Redondear((Val(TxtPVP.Text) * Val(TxtCantidad.Text)) * Porc_IVA, 2)
             SetAdoFields "Total_IVA", Total_IVA
             SetAdoFields "Total", (Val(TxtPVP.Text) * Val(TxtCantidad.Text))
             If IsNumeric(TxtOrden) Then
                SetAdoFields "Orden_No", Val(TxtOrden)
                SetAdoFields "TC", "OP"
             Else
                SetAdoFields "No_Hab", UCaseStrg(TxtOrden)
             End If
             SetAdoFields "Cta_Venta", .fields("Cta_Ventas")
             SetAdoFields "Codigo_Sup", Cod_Sub_Inv
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "ID", Ln_No
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
             TxtCantidad.Text = ""
          End If
       End If
   End If
  End With
  RatonNormal
  DLProducto.SetFocus
End Sub

Private Sub DLProducto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  TxtOrden = ""
  SQL2 = "SELECT (Producto & ' - ' & Codigo_Inv) As Productos ,P.* " _
       & "FROM Catalogo_Productos As P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDB_List DLProducto, AdoProducto, SQL2, "Productos"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FPedidos
  ConectarAdodc AdoAux
  ConectarAdodc AdoProducto
End Sub

Private Sub TxtCantidad_GotFocus()
  MarcarTexto TxtCantidad
End Sub

Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCantidad_LostFocus()
  TextoValido TxtCantidad, True
End Sub

Private Sub TxtOrden_GotFocus()
  TxtPVP = "0.00"
  CodigoP = SinEspaciosDer(DLProducto.Text)
  LblCodigo.Caption = CodigoP
  With AdoProducto.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo_Inv = '" & CodigoP & "' ")
       If Not .EOF Then
          TxtPVP = Format$(.fields("PVP"), "#,##0.00")
       End If
   End If
  End With
  MarcarTexto TxtOrden
End Sub

Private Sub TxtOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOrden_LostFocus()
  If IsNumeric(TxtOrden) Then
     TxtOrden = Format$(Val(TxtOrden), "0000000")
  End If
End Sub

Private Sub TxtPVP_GotFocus()
  MarcarTexto TxtPVP
End Sub

Private Sub TxtPVP_LostFocus()
  TextoValido TxtPVP, True, , 2
End Sub
