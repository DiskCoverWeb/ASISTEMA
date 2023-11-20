VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AbonoBancos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   5370
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   6960
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
   ScaleHeight     =   5370
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OpcEli 
      Caption         =   "&Eliminar"
      Height          =   225
      Left            =   5670
      TabIndex        =   28
      Top             =   105
      Width           =   1170
   End
   Begin VB.OptionButton OpcAct 
      Caption         =   "&Actualizar"
      Height          =   225
      Left            =   4305
      TabIndex        =   27
      Top             =   105
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.CheckBox CheqInstitucion 
      Caption         =   "&ABONOS RECIBIDOS EN LA INSTITUCION"
      Height          =   225
      Left            =   105
      TabIndex        =   25
      Top             =   105
      Width           =   4110
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&RECIBO DE CAJA No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Value           =   1  'Checked
      Width           =   2220
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2415
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   420
      Width           =   1170
   End
   Begin VB.TextBox TextCheqNo 
      Height          =   330
      Left            =   3990
      MaxLength       =   8
      TabIndex        =   18
      Top             =   4515
      Width           =   1800
   End
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   105
      Top             =   2940
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "IngCaja"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   4095
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Height          =   330
      Left            =   3990
      MaxLength       =   14
      TabIndex        =   16
      Text            =   "0"
      Top             =   4095
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   5880
      Picture         =   "AbonoBan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1365
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   5880
      Picture         =   "AbonoBan.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   420
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "AbonoBan.frx":0D0C
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   3990
      TabIndex        =   7
      Top             =   840
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   4515
      TabIndex        =   3
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "AbonoBan.frx":0D25
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   1260
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   4410
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "AbonoBan.frx":0D3C
      DataSource      =   "AdoDetAcomp"
      Height          =   315
      Left            =   1365
      TabIndex        =   5
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   105
      Top             =   4725
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   26
      Top             =   2100
      Width           =   5685
   End
   Begin VB.Label LblNota 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota"
      ForeColor       =   &H00C000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   3675
      Width           =   5685
   End
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   3990
      TabIndex        =   20
      Top             =   4935
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
      Height          =   330
      Left            =   2100
      TabIndex        =   19
      Top             =   4935
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO FACT."
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2100
      TabIndex        =   24
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA DE EMISION"
      Height          =   330
      Left            =   105
      TabIndex        =   23
      Top             =   1680
      Width           =   2010
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
      Height          =   330
      Left            =   2100
      TabIndex        =   17
      Top             =   4515
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA"
      Height          =   330
      Left            =   3675
      TabIndex        =   2
      Top             =   420
      Width           =   855
   End
   Begin VB.Label LblObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      ForeColor       =   &H00C000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   3360
      Width           =   5685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3465
      TabIndex        =   9
      Top             =   1680
      Width           =   2325
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Abonado"
      Height          =   330
      Left            =   2100
      TabIndex        =   15
      Top             =   4095
      Width           =   1905
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3990
      TabIndex        =   12
      Top             =   2940
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F&actura No."
      Height          =   330
      Left            =   2205
      TabIndex        =   6
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   2100
      TabIndex        =   11
      Top             =   2940
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   2520
      Width           =   5685
   End
End
Attribute VB_Name = "AbonoBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Listar_Facturas_Pendientes()
  TipoFactura = DCTipo
  SQL1 = "SELECT F.TC,F.Factura,F.CodigoC,F.Fecha,F.Fecha_V,F.Saldo_MN,F.Cta_CxP,F.Nota," _
       & "F.Observacion,C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo,F.Autorizacion,F.Serie " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.TC,F.Factura "
  'MsgBox SQL1
  SelectDBCombo DCFactura, AdoFactura, SQL1, "Factura"
End Sub

Private Sub Command1_Click()
  FechaValida MBFecha
  If CheqInstitucion.Value = 1 Then
     If OpcAct.Value Then TipoDoc = "A" Else TipoDoc = "E"
  Else
     TipoDoc = Ninguno
  End If
  TextoValido TextCheqNo
  FechaTexto = MBFecha
  Mensajes = "Esta Seguro que desea grabar Abono."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     FechaTexto = MBFecha ' FechaSistema
     If CheqRecibo.Value = 1 Then
        DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
     Else
        DiarioCaja = Val(TxtRecibo)
     End If
     SaldoDisp = Saldo - TotalCajaMN
     LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
     Cta = SinEspaciosIzq(DCBanco)
     NivelNo = UCase(Trim(Mid(DCBanco, Len(Cta) + 1, Len(DCBanco))))
    'Abono de Factura
     TA.T = Normal
     TA.Fecha = MBFecha
     TA.Cta = Cta
     TA.Banco = UCase(Trim(Mid(DCBanco, Len(Cta) + 1, Len(DCBanco))))
     TA.Cheque = TextCheqNo
     TA.Abono = TotalCajaMN
     Grabar_Abonos TA
     T = "P"
     If SaldoDisp <= 0 Then
        T = "C"
        SaldoDisp = 0
     End If
     sSQL = "UPDATE Facturas " _
          & "SET Saldo_MN = " & SaldoDisp & ", T = '" & T & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Factura = " & Factura_No & " " _
          & "AND TC = '" & TipoFactura & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodigoC = '" & CodigoCliente & "' "
     ConectarAdoExecute sSQL
     RatonNormal
     'ImprimirReciboCaja AdoIngCaja, FechaTexto, NombreCliente
     Listar_Facturas_Pendientes
     MsgBox "Abono Realizado con éxito"
     DCFactura.SetFocus
  End If
  'Unload AbonoEfectivo
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de abonos de facturas"
   Unload Me
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBanco_LostFocus()
  TextCajaMN.SetFocus
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  Codigo1 = Ninguno
  Saldo = 0
  Total_IVA = 0
  Total_Ret = 0
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Cotizacion = 0
  TotalDolar = 0
  Saldo_ME = 0
  Label3.Caption = ""
  Label1.Caption = ""
  LabelPend.Caption = ""
  LabelSaldo.Caption = ""
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura) & " ")
       If Not .EOF Then
          Label8.Caption = " " & .Fields("Fecha")
          Label11.Caption = "C.I./R.U.C.: " & .Fields("CI_RUC")
          Grupo_No = " " & .Fields("Grupo")
          TextCheqNo = Grupo_No
          LblObs.Caption = " " & .Fields("Observacion")
          LblNota.Caption = " " & .Fields("Nota")
          CodigoCliente = .Fields("CodigoC")
          NombreCliente = .Fields("Cliente")
          DireccionCli = .Fields("Direccion")
          Factura_No = .Fields("Factura")
          Cta_Cobrar = .Fields("Cta_CxP")
          TipoFactura = .Fields("TC")
          Saldo = .Fields("Saldo_MN")
         'Datos del Abonos
          TA.Serie = .Fields("Serie")
          TA.Autorizacion = .Fields("Autorizacion")
          TA.TP = TipoFactura
          TA.Fecha = MBFecha
          TA.Cta_CxP = Cta_Cobrar
          TA.Factura = Factura_No
          TA.CodigoC = CodigoCliente
          
          LabelSaldo.Caption = Format(Saldo, "#,##0.00")
          Command1.Enabled = True
          Label3.Caption = NombreCliente
          Label1.Caption = "Factura No. " & Format(Factura_No, "0000000")
          SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
          LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
          AbonoBancos.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
          TextCajaMN.Text = LabelPend.Caption
          DCBanco.SetFocus
       Else
          MsgBox "Esta Factura no esta pendiente"
          Command1.Enabled = False
          DCFactura.SetFocus
       End If
    End If
  End With
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
   Listar_Facturas_Pendientes
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDBCombo DCTipo, AdoDetAcomp, sSQL, "TC"
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas "
  If Nuevo Then sSQL = sSQL & "WHERE TC = 'BA' " Else sSQL = sSQL & "WHERE TC = 'CJ' "
  sSQL = sSQL & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBanco, AdoBanco, sSQL, "NomCuenta"
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.Value = 1 Then TxtRecibo = Format(DiarioCaja, "0000000") Else TxtRecibo = ""
  Mifecha = BuscarFecha(FechaTexto)
  MBFecha.Text = FechaSistema
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm AbonoBancos
   ConectarAdodc AdoBanco
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoDetAcomp
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
  TotalCajaMN = Round(Val(CCur(TextCajaMN)), 2)
  TextCajaMN = Format(TotalCajaMN, "#,##0.00")
  SaldoDisp = Saldo - TotalCajaMN
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
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

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format(Val(TxtRecibo), "0000000")
End Sub
