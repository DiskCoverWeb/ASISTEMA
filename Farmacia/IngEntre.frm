VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form IngEntrega 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Productos"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
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
      Height          =   960
      Left            =   9660
      Picture         =   "IngEntre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   525
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   11324
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.- INGRESOS DEL DIA"
      TabPicture(0)   =   "IngEntre.frx":0282
      Tab(0).ControlCount=   10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DBGEntrega"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MBoxFecha"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DBLArt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextPVP"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TextCant"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      TabCaption(1)   =   "&2.- CONTABILIZACION"
      TabPicture(1)   =   "IngEntre.frx":029E
      Tab(1).ControlCount=   7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LabelDebe"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LabelHaber"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DBGDatas2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TextConcepto"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(6).Enabled=   0   'False
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar &Asiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -67020
         Picture         =   "IngEntre.frx":02BA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   420
         Width           =   1485
      End
      Begin VB.TextBox TextConcepto 
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
         Left            =   -74895
         TabIndex        =   11
         Top             =   1050
         Width           =   7680
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
         Height          =   960
         Left            =   9555
         Picture         =   "IngEntre.frx":06FC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1470
         Width           =   1065
      End
      Begin VB.TextBox TextCant 
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
         Left            =   7875
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "IngEntre.frx":0B3E
         Top             =   1995
         Width           =   1380
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
         Left            =   7875
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "IngEntre.frx":0B40
         Top             =   1575
         Width           =   1380
      End
      Begin MSDBGrid.DBGrid DBGDatas2 
         Bindings        =   "IngEntre.frx":0B42
         Height          =   4425
         Left            =   -74895
         OleObjectBlob   =   "IngEntre.frx":0B58
         TabIndex        =   13
         Top             =   1470
         Width           =   10515
      End
      Begin MSDBCtls.DBList DBLArt 
         Bindings        =   "IngEntre.frx":150C
         DataSource      =   "DataArt"
         Height          =   1620
         Left            =   105
         TabIndex        =   2
         Top             =   735
         Width           =   6525
         _ExtentX        =   11509
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
      Begin MSMask.MaskEdBox MBoxFecha 
         Height          =   330
         Left            =   7875
         TabIndex        =   4
         Top             =   1155
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   327680
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
      End
      Begin MSDBGrid.DBGrid DBGEntrega 
         Bindings        =   "IngEntre.frx":151E
         Height          =   3795
         Left            =   105
         OleObjectBlob   =   "IngEntre.frx":1535
         TabIndex        =   18
         Top             =   2520
         Width           =   10515
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   10
         Top             =   735
         Width           =   7680
      End
      Begin VB.Label Label1 
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
         Left            =   6720
         TabIndex        =   7
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &FECHA"
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
         Left            =   6720
         TabIndex        =   3
         Top             =   1155
         Width           =   1170
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
         TabIndex        =   1
         Top             =   420
         Width           =   6525
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR"
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
         Left            =   6720
         TabIndex        =   5
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Label LabelHaber 
         Alignment       =   1  'Right Justify
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
         Left            =   -66705
         TabIndex        =   15
         Top             =   5985
         Width           =   1800
      End
      Begin VB.Label LabelDebe 
         Alignment       =   1  'Right Justify
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
         Left            =   -68490
         TabIndex        =   16
         Top             =   5985
         Width           =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T O T A L E S   "
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
         Left            =   -74895
         TabIndex        =   17
         Top             =   5985
         Width           =   6420
      End
   End
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3675
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataComp 
      Caption         =   "Comp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3990
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataBanco 
      Caption         =   "Banco"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3045
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataIngArt 
      Caption         =   "IngArt"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2100
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataRet 
      Caption         =   "Ret"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2730
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataAbonos2 
      Caption         =   "Abonos2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Data DataSubCtas 
      Caption         =   "SubCtas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2415
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataLinea 
      Caption         =   "Linea"
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
      Top             =   945
      Visible         =   0   'False
      Width           =   1905
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
      Top             =   1260
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
      Top             =   630
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "IngEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  GrabarArticulos
End Sub

Private Sub Command2_Click()
  Unload IngEntrega
End Sub

Private Sub Command3_Click()
  TextoValido TextConcepto
  NumComp = ReadSetDataNum("Diario", True, False)
  Mensajes = "Esta seguro de Grabar el Comprobante No. " & NumComp & "]"
  Titulo = "Pregunta de grabación"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then
     If DataAbonos2.Recordset.RecordCount > 0 Then
        RatonReloj
        FechaTexto = MBoxFecha.Text
        NumComp = ReadSetDataNum("Diario", True, True)
       'Grabacion del Comprobante
        Co.T = Normal
        Co.TP = CompDiario
        Co.Fecha = FechaTexto
        Co.Numero = NumComp
        Co.Concepto = TextConcepto.Text
        Co.Beneficiario = Ninguno
        Co.Efectivo = 0
        Co.Monto_Total = Debe
        GrabarComprobantes Co, DataAbonos2, DataSubCtas, , DataRet
        ImprimirComprobantesDe False, CompDiario, NumComp, NumEmpresa, DataComp, DataTrans, , DataRet
        IniciarAsientosDe DataAbonos2, DBGDatas2, DataSubCtas, DataBanco, DataRet, DataIngArt
        LabelDebe.Caption = Format(0, "#,##0.00")
        LabelHaber.Caption = Format(0, "#,##0.00")
        RatonNormal
        MBoxFecha.SetFocus
     End If
  End If
End Sub

Private Sub DBGEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     DBGEntrega.Col = 0: Codigo = DBGEntrega.Text
     Mensajes = "Esta seguro de Eliminar el Codigo: " & Codigo & "]"
     Titulo = "Pregunta de Eliminacion"
     TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If J = 6 Then
        sSQL = "DELETE * FROM Entrega "
        sSQL = sSQL & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
        sSQL = sSQL & "AND Codigo = '" & Codigo & "' "
        DeleteData DataArticulo, sSQL
        ProcesarAsientos
     End If
  End If
End Sub

Private Sub DBLArt_DblClick()
  SiguienteControl
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DBLArt_LostFocus()
  Codigo = Ninguno
  If SinEspaciosIzq(DBLArt.Text) <> "" Then Codigo = SinEspaciosIzq(DBLArt.Text)
  LlenarArticulos Codigo
End Sub

Private Sub Form_Activate()
   CTAsientoContable
   NuevoDiario = False
   IniciarAsientosDe DataAbonos2, DBGDatas2, DataSubCtas, DataBanco, DataRet, DataIngArt
   sSQL = "SELECT (Codigo & Space(5) & Articulo) As CodArt "
   sSQL = sSQL & "FROM Articulo ORDER BY Codigo "
   SelectDBList DBLArt, DataArt, sSQL, "CodArt"
   FechaValida MBoxFecha
   ProcesarAsientos
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm IngEntrega
   DataArt.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataLinea.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataIngArt.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSubCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComp.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataBanco.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAbonos2.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
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

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub

Public Sub LlenarArticulos(CodigoArt As String)
  DBGEntrega.Visible = False
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo ='" & CodigoArt & "' "
  SelectData DataArticulo, sSQL, False
  With DataArticulo.Recordset
   If .RecordCount > 0 Then
       Codigo = .Fields("Codigo")
       TextPVP.Text = "0.00"
       TextCant.Text = "0.00"
   Else
       MsgBox "Este Articulo no exite."
   End If
  End With
  ProcesarAsientos
  DBGEntrega.Visible = True
  TextPVP.SetFocus
End Sub

Public Sub GrabarArticulos()
  DBGEntrega.Visible = False
  Mensajes = "Esta seguro de Grabar el Producto: "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     FechaValida MBoxFecha
     sSQL = "SELECT * FROM Entrega "
     SelectData DataArticulo, sSQL, False
     With DataArticulo.Recordset
         .AddNew
         .Fields("Fecha") = MBoxFecha.Text
         .Fields("Codigo") = Codigo
         .Fields("Cantidad") = Val(TextCant.Text)
         .Fields("Precio") = Val(TextPVP.Text)
          If Val(TextCant.Text) > 0 And Val(TextPVP.Text) > 0 Then .Update
     End With
  End If
  ProcesarAsientos
  DBGEntrega.Visible = True
  RatonNormal
End Sub


Public Sub ProcesarAsientos()
  sSQL = "SELECT A.Codigo,A.Articulo,E.Precio,E.Cantidad,A.Cta_Ingreso,A.Cta_CxP "
  sSQL = sSQL & "FROM Entrega As E,Articulo As A "
  sSQL = sSQL & "WHERE E.Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
  sSQL = sSQL & "AND A.Codigo = E.Codigo "
  sSQL = sSQL & "ORDER BY A.Codigo "
  SelectDBGrid DBGEntrega, DataArticulo, sSQL
  RatonReloj
  With DataArticulo.Recordset
   If .RecordCount > 0 Then
       IniciarAsientosDe DataAbonos2, DBGDatas2, DataSubCtas, DataBanco, DataRet, DataIngArt
       TextConcepto.Text = "Cierre de entrega"
       Cta_Ingreso = .Fields("Cta_Ingreso")
       Cta_Aux = .Fields("Cta_CxP")
       Total = 0
       Do While Not .EOF
          If Cta_Ingreso <> .Fields("Cta_Ingreso") Or Cta_Aux <> .Fields("Cta_CxP") Then
             Total = Round(Total)
             InsertarAsientos DataAbonos2, Cta_Aux, 0, Total, 0
             InsertarAsientos DataAbonos2, Cta_Ingreso, 0, 0, Total
             Cta_Ingreso = .Fields("Cta_Ingreso")
             Cta_Aux = .Fields("Cta_CxP")
             Total = 0
          End If
          Precio = .Fields("Precio")
          Cantidad = .Fields("Cantidad")
          Total = Total + (Precio * Cantidad)
         .MoveNext
       Loop
       Total = Round(Total)
       InsertarAsientos DataAbonos2, Cta_Aux, 0, Total, 0
       InsertarAsientos DataAbonos2, Cta_Ingreso, 0, 0, Total
   End If
  End With
  Debe = 0: Haber = 0
  With DataAbonos2.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debe")
          Haber = Haber + .Fields("Haber")
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  LabelDebe.Caption = Format(0, "#,##0.00")
  LabelHaber.Caption = Format(0, "#,##0.00")
End Sub
