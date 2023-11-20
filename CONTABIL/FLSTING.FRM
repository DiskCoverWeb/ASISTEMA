VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FListCompIng 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION: Comprobante de Ingreso."
   ClientHeight    =   6585
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGTrans 
      Bindings        =   "FLstIng.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "FLstIng.frx":001C
      TabIndex        =   25
      Top             =   3000
      Width           =   9855
   End
   Begin MSDBGrid.DBGrid DBGridFacturas 
      Bindings        =   "FLstIng.frx":09CF
      Height          =   1455
      Left            =   4680
      OleObjectBlob   =   "FLstIng.frx":09E5
      TabIndex        =   22
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox TextComp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data DataMovFact 
      Caption         =   "MovFact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data DataComprobantes 
      Caption         =   "Comprobantes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Imprimir Comprobante"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label LabelRecibi 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recibi"
      Height          =   255
      Left            =   1560
      TabIndex        =   24
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Recibi de:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LabelConcepto 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label LabelCta 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cta"
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label LabelBanco 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco"
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label LabelDeposito 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deposito"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label LabelFormaPago 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FormaPago"
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label LabelFecha 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad Total:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   492
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4092
   End
   Begin VB.Label LabelHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label LabelDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   10080
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Depósito No."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cta. (Cte./Ahrr.)"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Depositar en:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Por concepto de:"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Forma de pago:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ingreso No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha (DD/MM/AA):"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1692
   End
End
Attribute VB_Name = "FListCompIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CantFacturas, CantFactIng As Integer
Dim Saldo, Abono, TotalAbonos As Single
Dim Detectar As Boolean

Private Sub CmdCancelar_Click()
   Unload FListCompIng
End Sub

Private Sub CmdGrabar_Click()
  Mensajes = "Imprimir Comprobante."
  Titulo = "Pregunta de Impresion"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then ImprimirComprobante
  TextComp.SetFocus
End Sub

Private Sub Form_Activate()
   Label12.Caption = Empresa
   MDIFact.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   'Centrar el formulario
   Left = (Screen.Width - Width) / 2
   Top = (Screen.Height - Height) / 2 - 600
   'Abriendo bases relacionadas
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataMovFact.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub
Public Sub LlenarIngresoTexto(ByVal InicioX As Single, ByVal InicioY As Single)
Printer.Font.Bold = False: Printer.FontSize = 10
'-----------------------------------------------------------------
Printer.CurrentX = InicioX + 2.2: Printer.CurrentY = InicioY + 3.6
Printer.Print LabelFecha.Caption
Printer.CurrentX = InicioX + 14.2: Printer.CurrentY = InicioY + 3.6
Printer.Print Format(TextComp.Text, "000000")
Printer.CurrentX = InicioX + 2.7: Printer.CurrentY = InicioY + 4.7
Printer.Print LabelConcepto.Caption
Printer.CurrentX = InicioX + 13.2: Printer.CurrentY = InicioY + 4.7
Printer.Print FormatoCient(Val(LabelTotal.Caption), 13)
Printer.CurrentX = InicioX + 0.7: Printer.CurrentY = InicioY + 5.9
Printer.Print "Sucres."
Printer.CurrentX = InicioX + 11.5: Printer.CurrentY = InicioY + 5.4
Printer.Print LabelBanco.Caption
Printer.CurrentX = InicioX + 13: Printer.CurrentY = InicioY + 5.9
Printer.Print LabelCta.Caption
Printer.CurrentX = InicioX + 12.5: Printer.CurrentY = InicioY + 6.4
Printer.Print LabelDeposito.Caption
J = 10: PFil = 7.7: Tabular = 0.7
Printer.CurrentX = InicioX + 4: Printer.CurrentY = InicioY + 7.2
Printer.Print LabelConcepto.Caption
DataMovFact.Refresh
Do While Not DataMovFact.Recordset.EOF
  Printer.CurrentX = InicioX + Tabular: Printer.CurrentY = InicioY + PFil
  Printer.Print DataMovFact.Recordset.Fields("Factura")
  Tabular = Tabular + 1.5
  If Tabular > 16 Then
     PFil = PFil + 0.5: Tabular = 0.7
  End If
  DataMovFact.Recordset.MoveNext
Loop
PFil = 13.5
For I = 1 To 8
GridCtas.Col = 1: GridCtas.Row = I
If GridCtas.Text <> "" Then
  Printer.CurrentX = InicioX + 1: Printer.CurrentY = InicioY + PFil
  GridCtas.Col = 0: GridCtas.Row = I
  Printer.Print GridCtas.Text
  Printer.CurrentX = InicioX + 5: Printer.CurrentY = InicioY + PFil
  GridCtas.Col = 1: GridCtas.Row = I
  Printer.Print GridCtas.Text
  Printer.CurrentX = InicioX + 12.3: Printer.CurrentY = InicioY + PFil
  GridCtas.Col = 2: GridCtas.Row = I
  Printer.Print FormatoCient(Val(GridCtas.Text), 13)
  Printer.CurrentX = InicioX + 15.3: Printer.CurrentY = InicioY + PFil
  GridCtas.Col = 3: GridCtas.Row = I
  Printer.Print FormatoCient(Val(GridCtas.Text), 13)
  PFil = PFil + 0.5
End If
Next I
Printer.CurrentX = InicioX + 12.3: Printer.CurrentY = InicioY + 18.4
Printer.Print FormatoCient(LabelDebe.Caption, 13)
Printer.CurrentX = InicioX + 15.3: Printer.CurrentY = InicioY + 18.4
Printer.Print FormatoCient(LabelHaber.Caption, 13)
End Sub

Public Sub ImprimirComprobante()
Dim PosX, PosY As Single
Dim Msg   As String
MousePointer = vbHourglass        ' colocamos el mouse reloj
MargIzq = 1 ' Obtiene la mitad del ancho.
MargSup = 1 ' Obtiene la mitad del alto.
MargDer = 19.5
MargInf = 23
On Error GoTo Errorhandler  ' Configura controlador de errores.
Printer.ScaleMode = 7
Printer.DrawWidth = 6
Call FormatoIngresoLineas(MargIzq, MargDer, MargSup, MargInf)
Call FormatoIngresoTexto(MargIzq, MargSup)
Call LlenarIngresoTexto(MargIzq, MargSup)
'Inicio de impresión
Printer.EndDoc
MousePointer = vbDefault
Exit Sub
Errorhandler:
    MousePointer = vbDefault
    MsgBox "Error: Hubo un problema al imprimir en su impresora."
    Exit Sub
End Sub

Private Sub TextComp_GotFocus()
  TextComp.Text = ""
End Sub

Private Sub TextComp_LostFocus()
  MousePointer = vbHourglass
  SSQL = "SELECT Ingresos.*,Cuentas.Cuenta As Banco, Cuentas.No_Cta FROM Ingresos,Cuentas "
  SSQL = SSQL & "WHERE Ingreso = " & Val(TextComp.Text) & " "
  SSQL = SSQL & "AND Cta_Banco = Cuentas.Codigo "
  DataComprobantes.RecordSource = SSQL: DataComprobantes.Refresh
  If DataComprobantes.Recordset.RecordCount > 0 Then
     LabelFecha.Caption = DataComprobantes.Recordset.Fields("Fecha")
     LabelRecibi.Caption = DataComprobantes.Recordset.Fields("Recibi_de")
     LabelTotal.Caption = DataComprobantes.Recordset.Fields("Monto_Total")
     LabelConcepto.Caption = DataComprobantes.Recordset.Fields("Concepto")
     LabelFormaPago.Caption = DataComprobantes.Recordset.Fields("Forma_Pago")
     LabelBanco.Caption = DataComprobantes.Recordset.Fields("Banco")
     LabelCta.Caption = DataComprobantes.Recordset.Fields("No_Cta")
     LabelDeposito.Caption = DataComprobantes.Recordset.Fields("Deposito")
     SSQL = "SELECT Factura,Saldo,Abono,Retencion FROM Pago_Facturas "
     SSQL = SSQL & "WHERE Comprobante = " & Val(TextComp.Text)
     DataMovFact.RecordSource = SSQL: DataMovFact.Refresh
     'Llenar cuentas
     SSQL = "SELECT Cta,Cuentas.Cuenta,Debe,Haber FROM Transacciones,Cuentas "
     SSQL = SSQL & "WHERE TP = '" & CompIngreso & "' "
     SSQL = SSQL & "AND Comp = " & Val(TextComp.Text) & " "
     SSQL = SSQL & "AND Cuentas.Codigo = Cta "
     DataTransacciones.RecordSource = SSQL: DataTransacciones.Refresh
     SumaDebe = 0: SumaHaber = 0
     Do While Not DataTransacciones.Recordset.EOF
        SumaDebe = SumaDebe + DataTransacciones.Recordset.Fields("Debe")
        SumaHaber = SumaHaber + DataTransacciones.Recordset.Fields("Haber")
        DataTransacciones.Recordset.MoveNext
     Loop
     DBGTrans.Columns(0).Width = TextWidth(String(DataTransacciones.Recordset.Fields(0).Size, "X"))
     DBGTrans.Columns(1).Width = TextWidth(String(DataTransacciones.Recordset.Fields(1).Size, "X"))
     DBGTrans.Columns(DataTransacciones.Recordset.Fields.Count - 2).NumberFormat = "##,##0.00"
     DBGTrans.Columns(DataTransacciones.Recordset.Fields.Count - 1).NumberFormat = "##,##0.00"
     LabelDebe.Caption = Format(SumaDebe, "#,##0")
     LabelHaber.Caption = Format(SumaHaber, "#,##0")
  Else
     MsgBox "Este Comprobante no existe."
     TextComp.SetFocus
  End If
  MousePointer = vbDefault
End Sub

