VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form AbonosPV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   2655
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7800
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
   ScaleHeight     =   2655
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextTotalBaucher 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   4830
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "0"
      Top             =   1680
      Width           =   1800
   End
   Begin VB.TextBox TextBaucher 
      Height          =   330
      Left            =   1575
      MaxLength       =   8
      TabIndex        =   9
      Top             =   1680
      Width           =   1590
   End
   Begin VB.TextBox TextTarjeta 
      Height          =   330
      Left            =   1575
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "."
      Top             =   1365
      Width           =   3270
   End
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   105
      Top             =   945
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
      Top             =   2205
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
      Left            =   4830
      MaxLength       =   14
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Crédito"
      Height          =   855
      Left            =   6720
      Picture         =   "Abonos1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   6720
      Picture         =   "Abonos1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1260
      TabIndex        =   16
      Top             =   525
      Width           =   1905
   End
   Begin VB.Line Line4 
      X1              =   105
      X2              =   6615
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Abonado"
      Height          =   330
      Left            =   3255
      TabIndex        =   10
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Baucher No."
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Line Line3 
      X1              =   105
      X2              =   6615
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   4830
      TabIndex        =   13
      Top             =   2205
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
      Height          =   330
      Left            =   3255
      TabIndex        =   12
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tarjeta Crédito"
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   1485
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja MN"
      Height          =   330
      Left            =   3255
      TabIndex        =   4
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4830
      TabIndex        =   3
      Top             =   525
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F&actura No."
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   3255
      TabIndex        =   2
      Top             =   525
      Width           =   1590
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CLIENTE"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6525
   End
End
Attribute VB_Name = "AbonosPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Mensajes = "Esta Seguro que desea grabar estos pagos."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     FechaTexto = FechaSistema
     DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
     TA.Cta_CxP = Cta_Cobrar
     TA.CodigoC = CodigoCliente
    'Abono de Factura Caja MN
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Cta_CajaG
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "EFECTIVO MN"
     TA.Cheque = Grupo_No
     TA.Factura = Factura_No
     TA.Abono = TotalCajaMN
     Grabar_Abonos TA
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Cta_CajaBA
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = TextTarjeta.Text
     TA.Cheque = TextBaucher.Text
     TA.Factura = Factura_No
     TA.Abono = Total_Tarjeta
     Grabar_Abonos TA
     T = "P"
     If SaldoDisp <= 0 Then
        T = "C"
        SaldoDisp = 0
     End If
     sSQL = "UPDATE Facturas " _
          & "SET Saldo_MN = " & SaldoDisp & ",T = '" & T & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Factura = " & Factura_No & " " _
          & "AND TC = '" & TipoProc & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodigoC = '" & CodigoCliente & "' "
     ConectarAdoExecute sSQL
     RatonNormal
  End If
  Unload AbonosPV
End Sub

Private Sub Command2_Click()
   Unload AbonosPV
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  ControlEsNumerico TextTotalBaucher
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Total_IVA = 0
  Total_Ret = 0
  CodigoCliente = "9999999999"
  NombreCliente = "CONSUMIDOR FINAL"
  DireccionCli = "S/N"
  Label3.Caption = NombreCliente
  Label1.Caption = " " & Factura_No
  SQL1 = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE CodigoC = '" & CodigoCliente & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Factura = " & Factura_No & " " _
       & "AND NOT TC IN ('C','P') " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Factura "
  SelectAdodc AdoFactura, SQL1
  Mifecha = BuscarFecha(FechaTexto)
  Codigo1 = Ninguno
  Saldo = 0: Cotizacion = 0: TotalDolar = 0
  Saldo_ME = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Factura_No = .Fields("Factura")
       TipoProc = .Fields("TC")
       Cta_Cobrar = .Fields("Cta_CxP")
       Saldo = Redondear(.Fields("Saldo_MN"), 2)
       Saldo_ME = Redondear(.Fields("Saldo_ME"), 2)
       TotalDolar = .Fields("Total_ME")
       Cotizacion = .Fields("Cotizacion")
       If TotalDolar <> 0 Then LabelSaldoD.Caption = Format$(Saldo_ME, "#,##0.00")
       LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
       If TotalDolar <> 0 Then
          TextRet.SetFocus
       Else
          TextCajaMN.SetFocus
       End If
    End If
  End With
  SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
  TextCajaMN.Text = LabelPend.Caption
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm AbonosPV
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
End Sub

Private Sub TextBaucher_GotFocus()
  MarcarTexto TextBaucher
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
  TextCajaMN.Text = Format$(TotalCajaMN, "#,##0.00")
  SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextTarjeta_GotFocus()
  MarcarTexto TextTarjeta
End Sub

Private Sub TextTotalBaucher_GotFocus()
  MarcarTexto TextTotalBaucher
End Sub

Private Sub TextTotalBaucher_LostFocus()
  TextoValido TextTotalBaucher, True
  Total_Tarjeta = Redondear(Val(CCur(TextTotalBaucher.Text)), 2)
  TextTotalBaucher.Text = Format$(Total_Tarjeta, "#,##0.00")
  SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

