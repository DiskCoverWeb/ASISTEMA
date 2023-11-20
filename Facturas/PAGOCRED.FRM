VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PagoCredito 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORMA DE PAGO A CREDITO"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PagoCred.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGBancos 
      Bindings        =   "PagoCred.frx":000C
      Height          =   1800
      Left            =   105
      OleObjectBlob   =   "PagoCred.frx":001F
      TabIndex        =   9
      Top             =   1470
      Width           =   9990
   End
   Begin VB.TextBox TextCheque 
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
      Left            =   1365
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PagoCred.frx":09D3
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
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
      Left            =   6720
      TabIndex        =   8
      Top             =   945
      Width           =   1695
   End
   Begin VB.TextBox TextBanco 
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
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   3
      Top             =   525
      Width           =   2115
   End
   Begin VB.TextBox TextCheq_No 
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
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   5
      Top             =   945
      Width           =   1275
   End
   Begin VB.Data DataCaja 
      Caption         =   "Caja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MBoxFechaCheque 
      Height          =   330
      Left            =   3990
      TabIndex        =   7
      Top             =   945
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "##/##/##"
      PromptChar      =   "0"
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Cheq."
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
      Left            =   2730
      TabIndex        =   6
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   8400
      TabIndex        =   11
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Factura S/."
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
      TabIndex        =   10
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label LabelCheque 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   8400
      TabIndex        =   13
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Bancos S/."
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
      TabIndex        =   12
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CHEQUE S/."
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
      Width           =   1275
   End
   Begin VB.Label Label1x 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BANCO"
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
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
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
      Top             =   945
      Width           =   1275
   End
End
Attribute VB_Name = "PagoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If .Fields("Posf") Then
'               DataChequesPosf.Recordset.AddNew
'               DataChequesPosf.Recordset.Fields("T") = Normal
'               DataChequesPosf.Recordset.Fields("Fecha") = .Fields("Fecha")
'               DataChequesPosf.Recordset.Fields("Codigo_C") = CodigoCli
'               DataChequesPosf.Recordset.Fields("Factura") = Factura_No
'               DataChequesPosf.Recordset.Fields("TP") = Ventas
'               DataChequesPosf.Recordset.Fields("Monto") = .Fields("Valor")
'               DataChequesPosf.Recordset.Fields("Banco") = .Fields("Banco")
'               DataChequesPosf.Recordset.Fields("Cheq_No") = .Fields("Cheq_No")
'               DataChequesPosf.Recordset.Fields("Cheque") = .Fields("Cheque")
'               DataChequesPosf.Recordset.Update
            
Private Sub Command1_GotFocus()
  Saldo = Total_Factura - Efectivo - Retencion - Cheque
  LabelSaldo.Caption = Format(Saldo, "#,##0")
End Sub

Private Sub DBGBancos_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(DataCaja)
End Sub

Private Sub DBGBancos_Change()
  CalculosTotalCaja DataCaja
  Saldo = Total_Factura - Efectivo - Retencion - Cheque
  If Saldo < 0 Then
     Cadena = "El pago es mayor que el monto de la Factura, " & Chr(13)
     Cadena = Cadena & "Elimine o cambie alguno Pagos."
     MsgBox Cadena
     DBGBancos.SetFocus
  Else
     TextCheque.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "DELETE * FROM Ingreso_Caja "
  DeleteData DataCaja, sSQL
  SelectData DataCaja, "Ingreso_Caja", False
  DBAnchoMallas PagoCredito, DBGBancos, 0, DataCaja
  Saldo = Total_Factura
  LabelSaldo.Caption = Format(Saldo, "#,##0")
  TextCheque.Text = Str(Total_Factura)
  TextCheque.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm PagoCredito
   DataCaja.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub

Private Sub MBoxFechaCheque_LostFocus()
  If MBoxFechaCheque.Text = "00/00/00" Then MBoxFechaCheque.Text = FechaTexto1
  If TextBanco.Text = Ninguno Or TextCheq_No.Text = Ninguno Then Cheque = 0
  DataCaja.Recordset.AddNew
  DataCaja.Recordset.Fields("Banco") = TextBanco.Text
  DataCaja.Recordset.Fields("Cheq_No") = TextCheq_No.Text
  DataCaja.Recordset.Fields("Fecha") = MBoxFechaCheque.Text
  Si_No = False
  If MBoxFechaCheque.Text <> FechaTexto1 Then
     Efectivo = 0: Retencion = 0: Si_No = True
  End If
  Total_Abonos = Efectivo + Retencion + Cheque
  DataCaja.Recordset.Fields("Posf") = Si_No
  DataCaja.Recordset.Fields("Efectivo") = Efectivo
  DataCaja.Recordset.Fields("Retencion") = Retencion
  DataCaja.Recordset.Fields("Cheque") = Cheque
  DataCaja.Recordset.Fields("Total_Abono") = Total_Abonos
  DataCaja.Recordset.Fields("Valor") = 0
  DataCaja.Recordset.Fields("Saldo") = 0
  If Cheque <> 0 And MBoxFechaCheque.Text <> FechaTexto1 Then DataCaja.Recordset.Update
  Saldo = Total_Factura
  CalculosTotalCaja DataCaja
  CalcularSaldos DataCaja, Total_Factura
  CalculosTotalCaja DataCaja
  Saldo = Total_Factura
  LabelCheque.Caption = Format(Cheque, "#,##0")
  LabelSaldo.Caption = Format(Saldo, "#,##0")
  If Saldo < 0 Then
     Cadena = "El pago es mayor que el monto de la Factura, " & Chr(13)
     Cadena = Cadena & "Elimine o cambie alguno Pagos."
     MsgBox Cadena
     DBGBancos.SetFocus
  Else
     TextCheque.SetFocus
  End If
End Sub

Private Sub TextBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextCheque.SetFocus
End Sub

Private Sub TextCheq_No_GotFocus()
  TextCheq_No.Text = ""
End Sub

Private Sub TextCheq_No_LostFocus()
  If TextCheq_No.Text = "" Then TextCheq_No.Text = Ninguno
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then Command1.SetFocus
End Sub

Private Sub Command1_Click()
  DataCaja.Recordset.AddNew
  DataCaja.Recordset.Fields("Valor") = Total_Factura
  DataCaja.Recordset.Fields("Efectivo") = 0
  DataCaja.Recordset.Fields("Retencion") = 0
  DataCaja.Recordset.Fields("Cheque") = 0
  DataCaja.Recordset.Fields("Posf") = False
  DataCaja.Recordset.Fields("Total_Abono") = 0
  DataCaja.Recordset.Fields("Saldo") = Total_Factura
  DataCaja.Recordset.Fields("Banco") = Ninguno
  DataCaja.Recordset.Fields("Cheq_No") = Ninguno
  DataCaja.Recordset.Fields("Fecha") = FechaTexto1
  DataCaja.Recordset.Update
  CalculosTotalCaja DataCaja
  LabelSaldo.Caption = Format(Saldo, "#,##0")
  PagoCredito.Hide
End Sub

Private Sub TextBanco_GotFocus()
   TextBanco.Text = ""
End Sub

Private Sub TextBanco_LostFocus()
   If TextBanco.Text = "" Then TextBanco.Text = Ninguno
   TextBanco.Text = UCase(TextBanco.Text)
End Sub

Private Sub TextCheque_GotFocus()
  MarcarTexto TextCheque
End Sub

Private Sub TextCheque_LostFocus()
  If TextCheque.Text = "" Then TextCheque.Text = "0"
  Cheque = Val(TextCheque.Text)
End Sub

Private Sub TextEfectivo_GotFocus()
  TextEfectivo.Text = ""
End Sub

Private Sub TextEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then
     Efectivo = Total_Factura
     TextEfectivo.Text = Str(Efectivo)
     SendKeys "{TAB}"
  End If
End Sub

Private Sub TextEfectivo_LostFocus()
  If TextEfectivo.Text = "" Then TextEfectivo.Text = "0"
  Efectivo = Val(TextEfectivo.Text)
  Saldo = Total_Factura - Efectivo - Retencion - SumaBancos
  LabelSaldo.Caption = Format(Saldo, "#,##0")
End Sub

Private Sub TextRet_GotFocus()
  TextRet.Text = ""
End Sub

Private Sub TextRet_LostFocus()
  If TextRet.Text = "" Then TextRet.Text = "0"
  Retencion = Val(TextRet.Text)
  CalculosTotalCaja DataCaja
  Saldo = Total_Factura - Efectivo - Retencion - SumaBancos
  LabelSaldo.Caption = Format(Saldo, "#,##0")
End Sub
