VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CancelarCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA                                                                      Cancelacion Factura"
   ClientHeight    =   4455
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   10845
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "CancCheq.frx":0000
      Height          =   3900
      Left            =   105
      OleObjectBlob   =   "CancCheq.frx":0016
      TabIndex        =   1
      Top             =   420
      Width           =   10620
   End
   Begin VB.Data DataCheques 
      Caption         =   "Cheques"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8190
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   330
      Left            =   9240
      TabIndex        =   0
      Top             =   105
      Width           =   1485
   End
   Begin VB.Data DataDiarioCaja 
      Caption         =   "DiarioCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data DataClientes 
      Caption         =   "Clientes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataFactura 
      Caption         =   "Factura"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Coloquese en el cheque y presione la tecla <F2> para cancelar el pago"
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   9150
   End
End
Attribute VB_Name = "CancelarCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
   Unload CancelarCheque
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_F2 Then
     Total = DataCheques.Recordset.Fields("Monto")
     MiFecha = DataCheques.Recordset.Fields("Fecha")
     Cheque = DataCheques.Recordset.Fields("Valor_Cheq")
     Factura_No = DataCheques.Recordset.Fields("Factura")
     TextoBanco = DataCheques.Recordset.Fields("Banco")
     TextoCheque = DataCheques.Recordset.Fields("Cheq_No")
     Codigos = DataCheques.Recordset.Fields("Codigo")
     TipoProc = DataCheques.Recordset.Fields("TP")
     Mensajes = "Esta Seguro que desea pagar el Cheque de Fecha: " & MiFecha & "," & Chr(13) & Chr(13)
     Mensajes = Mensajes & "Banco: " & TextoBanco
     Mensajes = Mensajes & ", Cheque No. " & TextoCheque & Chr(13) & Chr(13)
     Mensajes = Mensajes & "Valor Cheque S/. " & Format(Cheque, "#,##0") & Chr(13)
     Titulo = "Formulario de Grabación"
     TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If J = 6 Then
        SelectDataHead DataDiarioCaja, "Diario_Caja"
        IngresoCaja = ReadSetDataNum("Ingreso Caja", True, True)
        DiarioCaja = 0
        Saldo = Total - Cheque
        If Saldo < 0 Then Saldo = 0
       With DataDiarioCaja.Recordset
           .AddNew
           .Fields("T") = Normal
           .Fields("Fecha") = FechaSistema
           .Fields("TP") = CxC
           .Fields("Diario_No") = DiarioCaja
           .Fields("Caja_No") = IngresoCaja
           .Fields("Factura") = Factura_No
           .Fields("Monto") = Total
           .Fields("Cheque") = Cheque
           .Fields("Efectivo") = 0
           .Fields("Retencion") = 0
           .Fields("Total_Abono") = Cheque
           .Fields("Saldo") = Saldo
           .Fields("Banco") = TextoBanco
           .Fields("Cheq_No") = TextoCheque
           .Fields("Codigo_C") = Codigos
           .Update
       End With
       sSQL = "UPDATE Facturas SET Saldo = " & Saldo & " "
       If Saldo <= 0 Then sSQL = sSQL & ", T = '" & Cancelado & "' "
       sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
       UpdateData DataFactura, sSQL
       sSQL = "UPDATE Cheques_Posf SET T = '" & Cancelado & "' "
       sSQL = sSQL & "WHERE TP = '" & TipoProc & "' "
       sSQL = sSQL & "AND Fecha = #" & MiFecha & "# "
       sSQL = sSQL & "AND Factura = " & Factura_No & " "
       sSQL = sSQL & "AND Codigo_C = '" & Codigos & "' "
       sSQL = sSQL & "AND Banco = '" & TextoBanco & "' "
       sSQL = sSQL & "AND Cheq_No = '" & TextoCheque & "' "
       sSQL = sSQL & "AND Cheque = " & Cheque & " "
       DataCheques.Database.Execute sSQL
       ListarCheques
     End If
  End If
End Sub

Private Sub Form_Activate()
   ListarCheques
   RatonNormal MDIFact
   If DataCheques.Recordset.RecordCount <= 0 Then Unload CancelarCheque
End Sub

Private Sub Form_Load()
   CentrarForm CancelarCheque
   DataCheques.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataFactura.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataDiarioCaja.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataClientes.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\SETEOS.MDB"
End Sub

Public Sub ListarCheques()
   sSQL = "SELECT Fecha_D,Fecha,TP,Factura,Monto,Banco,Cheq_No,Cheque As Valor_Cheq,Codigo "
   sSQL = sSQL & "FROM Cheques_Posf,Clientes "
   sSQL = sSQL & "WHERE Cheques_Posf.Codigo_C = Clientes.Codigo "
   sSQL = sSQL & "AND T = '" & Procesado & "' "
   sSQL = sSQL & "ORDER BY Fecha,TP,Factura "
   SelectDBGrid CancelarCheque, DBGrid1, 0, DataCheques, sSQL
End Sub
