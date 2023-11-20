VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form AbonosProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PAGO CHEQUES"
   ClientHeight    =   3705
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   2835
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   3150
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   2115
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   7050
      Begin MSDataListLib.DataCombo DCClientes 
         Bindings        =   "CxP_Prov.frx":0000
         DataSource      =   "AdoClientes"
         Height          =   315
         Left            =   945
         TabIndex        =   8
         Top             =   945
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Cliente"
      End
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "CxP_Prov.frx":001A
         DataSource      =   "AdoBanco"
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   1680
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Banco"
      End
      Begin MSDataListLib.DataCombo DCCxP 
         Bindings        =   "CxP_Prov.frx":0031
         DataSource      =   "AdoCxP"
         Height          =   315
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Abonos"
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta Contable Egreso"
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   6840
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta Contable Egreso"
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   1365
         Width           =   6840
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cliente:"
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   945
         Width           =   855
      End
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1890
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   105
      Width           =   1800
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&CHEQUE No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   1995
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Factura"
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
   Begin VB.TextBox TextCajaMN 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   105
      MaxLength       =   14
      TabIndex        =   12
      Text            =   "0"
      Top             =   3045
      Width           =   2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6195
      Picture         =   "CxP_Prov.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2730
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   5145
      Picture         =   "CxP_Prov.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2730
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   5775
      TabIndex        =   3
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   105
      Top             =   2310
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "DetAcomp"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   1680
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   1365
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Banco"
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   2835
      Top             =   2730
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Factura"
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
   Begin MSAdodcLib.Adodc AdoCxP 
      Height          =   330
      Left            =   105
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
      Caption         =   "Factura"
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
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA"
      Height          =   330
      Left            =   4935
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MONTO"
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   2730
      Width           =   2640
   End
End
Attribute VB_Name = "AbonosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaValida MBFecha
  FechaTexto = MBFecha
  Fecha_Vence = MBFecha
  If IsNumeric(TxtRecibo) Then NoCheque = Format(Val(TxtRecibo), "00000000")
  
  Total = Val(TotalCajaMN)
  DetalleComp = Ninguno
  CodigoCliente = Ninguno
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCClientes & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Grupo = .Fields("Grupo")
       End If
   End If
  End With
  CodigoCli = CodigoCliente
    
  Trans_No = 201
  SubCtaGen = SinEspaciosIzq(DCCxP)
  Cta_Aux = SinEspaciosIzq(DCBanco)
  If Len(Cta_Aux) <= 1 Then Cta_Aux = Cta_CajaG
  
  BorrarAsientos True
  sSQL = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE TC = 'P' " _
       & "AND Cta = '" & SubCtaGen & "' " _
       & "AND DH = '1' " _
       & "AND TM = '1' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectAdodc AdoIngCaja, sSQL
  With AdoIngCaja.Recordset
      .AddNew
      .Fields("Prima") = 0
      .Fields("Fecha_V") = MBFecha
      .Fields("TC") = "P"
      .Fields("Prima") = 0
      .Fields("Factura") = 0
      .Fields("Codigo") = CodigoCliente
      .Fields("Beneficiario") = NombreCliente
      .Fields("Detalle_SubCta") = "Abono Anticipado"
      .Fields("Cta") = SubCtaGen
      .Fields("DH") = "1"
      .Fields("Valor") = Total
      .Fields("Valor_ME") = 0
      .Fields("TM") = "1"
      .Fields("Item") = NumEmpresa
      .Fields("T_No") = Trans_No
      .Fields("SC_No") = 1
      .Fields("CodigoU") = CodigoUsuario
      .Update
  End With
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectAdodc AdoIngCaja, sSQL
  
  InsertarAsientos AdoIngCaja, SubCtaGen, 0, Total, 0
  InsertarAsientos AdoIngCaja, Cta_Aux, 0, 0, Total
  
  Mensajes = "Esta Seguro que desea grabar Abono."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     FechaTexto = MBFecha ' FechaSistema
     RatonReloj
     NumComp = ReadSetDataNum("Egresos", True, True)
     Co.TP = CompEgreso
     Co.T = Normal
     Co.Fecha = FechaTexto
     Co.Numero = NumComp
     Co.Monto_Total = Total
     Co.Concepto = "Anticipado de: " & UCase(NombreCliente) & ", Grupo: " & Grupo
     Co.CodigoB = CodigoCliente
     Co.Efectivo = 0
     Co.Cotizacion = 0
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     Co.T_No = Trans_No
     GrabarComprobante Co
     Control_Procesos Normal, "Grabar Comprobante de: " & Co.TP & "No. " & NumComp
   ' Seteamos para el siguiente comprobante
     RatonNormal
     If TipoFactura <> "OP" Then ImprimirComprobantesDe False, Co
  End If
  Unload AbonosProveedores
End Sub

Private Sub Command2_Click()
   Unload AbonosProveedores
End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCxP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCxP_LostFocus()
Dim TipoCta As String
  TipoCta = Ninguno
  If AdoCxP.Recordset.RecordCount > 0 Then
     AdoCxP.Recordset.MoveFirst
     AdoCxP.Recordset.Find ("NomCuenta = '" & DCCxP & "' ")
     If Not AdoCxP.Recordset.EOF Then TipoCta = AdoCxP.Recordset.Fields("TC")
     'MsgBox TipoCta
     Select Case TipoCta
       Case "C", "P"
        sSQL = "SELECT C.Grupo,C.Codigo,C.Cliente,CP.Cta " _
             & "FROM Clientes As C, Catalogo_CxCxP As CP " _
             & "WHERE CP.Item = '" & NumEmpresa & "' " _
             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
             & "AND CP.Cta = '" & SinEspaciosIzq(DCCxP) & "' " _
             & "AND CP.Codigo = C.Codigo " _
             & "ORDER BY C.Cliente "
       Case Else
        sSQL = "SELECT Grupo,Codigo,Cliente " _
             & "FROM Clientes " _
             & "WHERE Codigo <> '.' " _
             & "ORDER BY Cliente "
     End Select
     SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
  End If
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta,TC " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN('C','P','PS','CS') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Mod_Gastos <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCxP, AdoCxP, sSQL, "NomCuenta"
  
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC,Codigo "
  SelectDBCombo DCBanco, AdoBanco, sSQL, "NomCuenta"
  Frame2.Visible = True
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format(DiarioCaja, "0000000") Else TxtRecibo = ""
  Mifecha = BuscarFecha(FechaTexto)
  MBFecha.Text = FechaSistema
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm AbonosProveedores
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCxP
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoClientes
End Sub

Private Sub TextCajaMN_GotFocus()
  TextCajaMN.Text = Saldo
  MarcarTexto TextCajaMN
End Sub

Private Sub TextCajaMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCajaMN_LostFocus()
  TextoValido TextCajaMN, True
  TotalCajaMN = Redondear(Val(CCur(TextCajaMN.Text)), 2)
  TextCajaMN.Text = Format(TotalCajaMN, "#,##0.00")
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
