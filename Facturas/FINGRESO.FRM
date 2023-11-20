VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"FINGRESO.frx":0000
   ClientHeight    =   6135
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   7140
      TabIndex        =   23
      Top             =   525
      Width           =   3375
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   1050
         TabIndex        =   30
         Text            =   "Text4"
         Top             =   1365
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1050
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   2205
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   210
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         Caption         =   "U.S.$"
         Height          =   225
         Left            =   105
         TabIndex        =   25
         Top             =   630
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sucres"
         Height          =   330
         Left            =   105
         TabIndex        =   24
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   2205
         TabIndex        =   32
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         Height          =   330
         Left            =   105
         TabIndex        =   31
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   2205
         TabIndex        =   29
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cotización"
         Height          =   330
         Left            =   1050
         TabIndex        =   28
         Top             =   1050
         Width           =   1065
      End
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   5565
      TabIndex        =   22
      Text            =   "0000000000000"
      Top             =   840
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   6300
      TabIndex        =   18
      Top             =   3255
      Width           =   2115
      Begin VB.OptionButton OpcH 
         Caption         =   "Haber"
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
         Left            =   1050
         TabIndex        =   8
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton OpcD 
         Caption         =   "Debe"
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
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSDBCtls.DBCombo DBCCtas 
      DataSource      =   "DataCuentas"
      Height          =   315
      Left            =   105
      TabIndex        =   6
      Top             =   3570
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4725
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TextRecibi 
      Height          =   330
      Left            =   1050
      MaxLength       =   30
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Data DataTransacciones 
      Caption         =   "Transacciones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2730
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   5595
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5565
      Width           =   1815
   End
   Begin VB.TextBox TextHaber 
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
      Left            =   8505
      TabIndex        =   10
      Top             =   3570
      Width           =   1905
   End
   Begin VB.TextBox TextConcepto 
      Height          =   645
      Left            =   1575
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2100
      Width           =   5370
   End
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10605
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   645
      Left            =   1575
      TabIndex        =   11
      Top             =   1260
      Width           =   5370
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " La cantidad de"
      Height          =   330
      Left            =   105
      TabIndex        =   33
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R.U.C./C.I."
      Height          =   330
      Left            =   4515
      TabIndex        =   21
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00/00/00"
      Height          =   330
      Left            =   9660
      TabIndex        =   20
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Recibi de:"
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Width           =   855
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
      TabIndex        =   17
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
      Height          =   435
      Left            =   8820
      TabIndex        =   16
      Top             =   5565
      Width           =   1695
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
      Height          =   435
      Left            =   7035
      TabIndex        =   15
      Top             =   5565
      Width           =   1695
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
      Height          =   435
      Left            =   5670
      TabIndex        =   14
      Top             =   5565
      Width           =   1275
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
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
      Left            =   8505
      TabIndex        =   9
      Top             =   3360
      Width           =   1905
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUENTA"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   6105
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Por concepto de:"
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha:"
      Height          =   330
      Left            =   8400
      TabIndex        =   0
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "FIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboFormaPago_LostFocus()
  If ComboFormaPago.Text = OpcEfectivo Then
    TextBanco.Enabled = False
    TextCtaCte.Enabled = False
    TextCheque.Enabled = False
    TextConcepto.SetFocus
  Else
    TextBanco.Enabled = True
    TextCtaCte.Enabled = True
    TextCheque.Enabled = True
    TextBanco.SetFocus
  End If
End Sub

Private Sub DBCCtas_GotFocus()
  sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cuenta FROM Cuentas "
  sSQL = sSQL & "WHERE TC <> 'G' AND Cuenta <> '" & Ninguno & "' ORDER BY Codigo "
  DataCuentas.RecordSource = sSQL: DataCuentas.Refresh
End Sub

Private Sub DBCCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then
     If (SumaDebe <> SumaHaber) Then DBCCtas.SetFocus Else CmdGrabar.SetFocus
  End If
End Sub

Private Sub DBCCtas_LostFocus()
  Cuenta = Ninguno: TipoCta = "G"
  Cadena = SinEspaciosIzq(DBCCtas.Text)
  sSQL = "SELECT * FROM Cuentas WHERE Codigo = '" & Cadena & "' AND TC <> 'G' "
  DataCtas.RecordSource = sSQL: DataCtas.Refresh
  If DataCtas.Recordset.RecordCount > 0 Then
    Cuenta = DataCtas.Recordset.Fields("Codigo")
    Cadena = DataCtas.Recordset.Fields("Cuenta")
    TipoCta = DataCtas.Recordset.Fields("TC")
    DBCCtas.Text = Cuenta & Space(20) & Cadena
  Else
    sSQL = "SELECT * FROM Cuentas WHERE Numero = " & Val(Cadena) & " AND TC <> 'G' "
    DataCtas.RecordSource = sSQL: DataCtas.Refresh
    If DataCtas.Recordset.RecordCount > 0 Then
       Cuenta = DataCtas.Recordset.Fields("Codigo")
       Cadena = DataCtas.Recordset.Fields("Cuenta")
       TipoCta = DataCtas.Recordset.Fields("TC")
       DBCCtas.Text = Cuenta & Space(20) & Cadena
    End If
  End If
  TextHaber.Text = ""
  If Cuenta = Ninguno Then DBCCtas.SetFocus
End Sub


Private Sub CmdCancelar_Click()
   Unload FIngreso
End Sub

Private Sub CmdGrabar_Click()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber,
  If (SumaDebe <> SumaHaber) Then
    Mensajes = "Las transacciones no cuadran correctamente" & Chr(13)
    Mensajes = Mensajes & "corrija los resultados de las cuentas"
    MsgBox Mensajes
    DBCCtas.SetFocus
  Else
    Mensajes = "Esta seguro de Grabar el Comp No. " & LabelComp.Caption & "]"
    Titulo = "Pregunta de grabación"
    TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
    If J = 6 Then
       If DataAsientos.Recordset.RecordCount > 0 Then
          MousePointer = vbHourglass
          FechaTexto = FormatoFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
          DataTransacciones.RecordSource = "Transacciones": DataTransacciones.Refresh
          DataComprobantes.RecordSource = "Ingresos ": DataComprobantes.Refresh
          sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Ingresos' "
          DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
          NumComp = DataCodigos.Recordset.Fields("Numero")
          sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Asiento' "
          DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
          Asiento = DataCodigos.Recordset.Fields("Numero")
          DataAsientos.Recordset.MoveFirst
          Do While Not DataAsientos.Recordset.EOF
             Codigos = DataAsientos.Recordset.Fields("CODIGO")
             TipoCta = DataAsientos.Recordset.Fields("TC")
             Debe = DataAsientos.Recordset.Fields("DEBE")
             Haber = DataAsientos.Recordset.Fields("HABER")
             DataTransacciones.Recordset.AddNew
             DataTransacciones.Recordset.Fields("Estado") = Normal
             DataTransacciones.Recordset.Fields("Asiento") = Asiento
             DataTransacciones.Recordset.Fields("Cheq_Dep") = TextCheque.Text
             DataTransacciones.Recordset.Fields("Fecha") = FechaTexto
             DataTransacciones.Recordset.Fields("TP") = CompIngreso
             DataTransacciones.Recordset.Fields("Comp") = NumComp
             DataTransacciones.Recordset.Fields("Concepto") = Mid(TextConcepto.Text, 1, 40)
             DataTransacciones.Recordset.Fields("Cta") = Codigos
             DataTransacciones.Recordset.Fields("TC") = TipoCta
             DataTransacciones.Recordset.Fields("Debe") = Debe
             DataTransacciones.Recordset.Fields("Haber") = Haber
             DataTransacciones.Recordset.Update
             DataAsientos.Recordset.MoveNext
          Loop
          'grabacion del Comp
          DataComprobantes.Recordset.AddNew
          DataComprobantes.Recordset.Fields("Estado") = Normal
          DataComprobantes.Recordset.Fields("Fecha") = FechaTexto
          DataComprobantes.Recordset.Fields("Recibi_de") = TextRecibi.Text
          DataComprobantes.Recordset.Fields("Ingreso") = NumComp
          DataComprobantes.Recordset.Fields("Monto_Total") = SumaDebe
          DataComprobantes.Recordset.Fields("Concepto") = TextConcepto.Text
          DataComprobantes.Recordset.Fields("Forma_Pago") = ComboFormaPago.Text
          DataComprobantes.Recordset.Fields("Banco") = TextBanco.Text
          DataComprobantes.Recordset.Fields("Cta_Cte") = TextCtaCte.Text
          DataComprobantes.Recordset.Fields("Deposito") = TextCheque.Text
          DataComprobantes.Recordset.Update
          'llenamos el encabezado del comprobante
          sSQL = "SELECT * FROM Ingresos "
          sSQL = sSQL & "WHERE Ingreso = " & NumComp & " "
          DataComprobantes.RecordSource = sSQL: DataComprobantes.Refresh
          'Llenar cuentas
          sSQL = "SELECT Cta,Cuentas.Cuenta,Debe,Haber FROM Transacciones,Cuentas "
          sSQL = sSQL & "WHERE TP = '" & CompIngreso & "' "
          sSQL = sSQL & "AND Comp = " & NumComp & " "
          sSQL = sSQL & "AND Cuentas.Codigo = Cta "
          DataTransacciones.RecordSource = sSQL: DataTransacciones.Refresh
          MousePointer = vbDefault
          'imprimir comprobante
          Mensajes = "Imprimir Comprobante"
          Titulo = "Pregunta de Impresion"
          TipoDeCaja = 4 + 32: ResultBox = MsgBox(Mensajes, TipoDeCaja, Titulo)
          If ResultBox = 6 Then ImprimirCompIngreso DataComprobantes, DataBanco, DataTransacciones
          sSQL = "UPDATE Codigos SET Numero = Numero + 1 WHERE Concepto = 'Ingresos' "
          DataCodigos.Database.Execute sSQL
          sSQL = "UPDATE Codigos SET Numero = Numero + 1 WHERE Concepto = 'Asiento' "
          DataCodigos.Database.Execute sSQL
          NumComp = NumComp + 1
          LabelComp.Caption = Format(NumComp, "000000")
          sSQL = "DELETE * FROM Asientos "
          DataAsientos.Database.Execute sSQL
          sSQL = "SELECT * FROM Asientos "
          DataAsientos.RecordSource = sSQL: DataAsientos.Refresh
          DataCodigos.Database.Close
          DBAnchoMallas FIngreso, DBGAsientos, DataAsientos
       End If
       TextFechaD.SetFocus
    End If
  End If
End Sub

Private Sub DBGAsientos_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar la Transaccion " & Chr(13)
  Mensajes = Mensajes & DataAsientos.Recordset.Fields("CODIGO") & "  "
  Mensajes = Mensajes & DataAsientos.Recordset.Fields("CUENTA") & "?"
  Titulo = "Confirmación de eliminación"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub DBGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then DBCCtas.SetFocus
End Sub

Private Sub DBGAsientos_LostFocus()
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
End Sub

Private Sub Form_Activate()
  sSQL = "DELETE * FROM Asientos "
  DataAsientos.Database.Execute sSQL
  sSQL = "SELECT * FROM Asientos "
  DataAsientos.RecordSource = sSQL: DataAsientos.Refresh
  'DBAnchoMallas FIngreso, DBGAsientos, DataAsientos
  
  sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cuenta FROM Cuentas "
  sSQL = sSQL & "WHERE TC <> 'G' AND Cuenta <> '" & Ninguno & "' ORDER BY Codigo "
  DataCuentas.RecordSource = sSQL: DataCuentas.Refresh
  DBCCtas.ListField = "Nombre_Cuenta"
  DBCCtas.Text = DataCuentas.Recordset.Fields("Nombre_Cuenta")
  
  sSQL = "SELECT * FROM Codigos WHERE Concepto = 'Ingresos' "
  DataCodigos.RecordSource = sSQL: DataCodigos.Refresh
  DataCodigos.Database.Close
  
  DataComprobantes.RecordSource = "Ingresos ": DataComprobantes.Refresh
  TextConcepto.MaxLength = DataComprobantes.Recordset.Fields("Concepto").Size
  TextConcepto.width = TextWidth(String(TextConcepto.MaxLength + 1, "X"))
  DataComprobantes.Database.Close
  
  Label12.Caption = Empresa
  MDIFact.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   CentrarForm FIngreso
   'Abriendo bases relacionadas
   DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTransacciones.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub TextBanco_GotFocus()
  TextBanco = ""
End Sub

Private Sub TextBanco_LostFocus()
  If TextBanco = "" Then TextBanco = Ninguno
End Sub

Private Sub TextCantidad_GotFocus()
   TextCantidad.Text = ""
End Sub

Private Sub TextCantidad_LostFocus()
  If TextCantidad.Text = "" Then TextCantidad.Text = "0"
  Abono = Val(TextCantidad.Text)
End Sub

Private Sub TextConcepto_Change()
  If Len(TextConcepto.Text) >= TextConcepto.MaxLength Then DBCCtas.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
   TextConcepto.Text = ""
End Sub

Private Sub TextConcepto_LostFocus()
  If TextConcepto.Text = "" Then TextConcepto.Text = Ninguno
End Sub

Private Sub TextCtaCte_GotFocus()
   TextCtaCte.Text = ""
End Sub

Private Sub TextCtaCte_LostFocus()
  If TextCtaCte.Text = "" Then TextCtaCte.Text = Ninguno
End Sub

Private Sub TextCheque_GotFocus()
   TextCheque.Text = ""
End Sub

Private Sub TextCheque_LostFocus()
  If TextCheque.Text = "" Then TextCheque.Text = Ninguno
End Sub

Private Sub TextFechaA_Change()
  If Len(TextFechaA.Text) >= TextFechaA.MaxLength Then TextRecibi.SetFocus
End Sub

Private Sub TextFechaA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaA_LostFocus()
  If TextFechaA.Text = "" Then TextFechaA = Anio
  FechaTexto = FormatoFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
End Sub

Private Sub TextFechaD_Change()
  If Len(TextFechaD.Text) >= TextFechaD.MaxLength Then TextFechaM.SetFocus
End Sub

Private Sub TextFechaD_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaD_LostFocus()
  If TextFechaD.Text = "" Then TextFechaD = Dia
End Sub

Private Sub TextFechaM_Change()
  If Len(TextFechaM.Text) >= TextFechaM.MaxLength Then TextFechaA.SetFocus
End Sub

Private Sub TextFechaM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaM_LostFocus()
If TextFechaM.Text = "" Then TextFechaM = Mes
End Sub

Private Sub TextHaber_LostFocus()
  If TextHaber.Text = "" Then TextHaber.Text = "0"
  If (TipoCta <> "G") And (Val(TextHaber.Text) <> 0) Then
      Debe = 0: Haber = 0
      If OpcD Then Debe = Val(TextHaber.Text) Else Haber = Val(TextHaber.Text)
      If DataAsientos.Recordset.RecordCount < 8 Then
         DataAsientos.Recordset.AddNew
         DataAsientos.Recordset.Fields("TC") = TipoCta
         DataAsientos.Recordset.Fields("Codigo") = Cuenta
         DataAsientos.Recordset.Fields("Cuenta") = Cadena
         DataAsientos.Recordset.Fields("Debe") = Debe
         DataAsientos.Recordset.Fields("Haber") = Haber
         DataAsientos.Recordset.Update
      End If
  End If
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
  DBCCtas.SetFocus
End Sub

Private Sub TextRecibi_Change()
  If Len(TextRecibi.Text) >= TextRecibi.MaxLength Then TextCantidad.SetFocus
End Sub

Private Sub TextRecibi_LostFocus()
  If TextRecibi.Text = "" Then TextRecibi.Text = Ninguno
End Sub

