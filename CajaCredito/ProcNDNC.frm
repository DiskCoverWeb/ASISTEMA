VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form ProcesarND_NC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COSTOS POR MANTENIMIENTO / FONDO MORTUORIO"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "ProcNDNC.frx":0000
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   2310
      TabIndex        =   7
      Top             =   1050
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.CommandButton Command4 
      Caption         =   "&Listar Proceso "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5460
      Picture         =   "ProcNDNC.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   105
      Width           =   1170
   End
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2310
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ProcNDNC.frx":0456
      Top             =   420
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DGLibreta 
      Bindings        =   "ProcNDNC.frx":045A
      Height          =   4740
      Left            =   105
      TabIndex        =   11
      Top             =   1470
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   1680
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
      Caption         =   "Aux"
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6720
      Picture         =   "ProcNDNC.frx":0472
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
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
      Height          =   1065
      Left            =   7980
      Picture         =   "ProcNDNC.frx":0CF4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Procesar Transaccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4200
      Picture         =   "ProcNDNC.frx":0FFE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1170
   End
   Begin VB.TextBox TextMontoMin 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2310
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "ProcNDNC.frx":1440
      Top             =   105
      Width           =   1800
   End
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   315
      Top             =   1995
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
      Caption         =   "Caja"
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
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   105
      Top             =   6195
      Width           =   9045
      _ExtentX        =   15954
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
      Caption         =   "Saldos"
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   2205
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
      Caption         =   "TP"
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
   Begin VB.Label LabelTotalInt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   330
      Left            =   2310
      TabIndex        =   5
      Top             =   735
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Transacción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1050
      Width           =   2220
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL PROCESADO"
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
      Top             =   735
      Width           =   2220
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto Transaccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   2220
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto Minimo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   2220
   End
End
Attribute VB_Name = "ProcesarND_NC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub InsertarND_NC(MBFecha As String, _
                         DtaCta As Adodc, _
                         CuentaNo As String, _
                         TDebe As Currency, _
                         THaber As Currency, _
                         MontoMinimo As Currency)
Dim Si_Debitar As Boolean
  Si_Debitar = True
  TiempoTexto = Format(Time, FormatoTimes)
  If (CuentaNo <> "00000000-0") Then
  TotalEncaje = 0: SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Libretas " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "AND TP = '" & TipoProc & "' " _
       & "AND Fecha = #" & BuscarFecha(FechaSistema) & "# "
  SelectAdodc DtaCta, sSQL
  If DtaCta.Recordset.RecordCount > 0 Then Si_Debitar = False
  If Si_Debitar Then
     sSQL = "SELECT TOP 1 * " _
          & "FROM Trans_Libretas " _
          & "WHERE Cuenta_No = '" & CuentaNo & "' " _
          & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
     SelectAdodc DtaCta, sSQL
     With DtaCta.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           NumeroLineas = .Fields("ID")
           ID_Trans = .Fields("IDT")
           If NumeroLineas <= 0 Then NumeroLineas = 1
           If NumeroLineas > 34 Then NumeroLineas = 1
       End If
      'Debite si el Disponible es mayor al Valor
       Valor = 0
       If TDebe > 0 Then Valor = TDebe
       If THaber > 0 Then Valor = THaber
       If 0 < SaldoDisp And SaldoDisp <= Valor Then Valor = SaldoDisp
       If TDebe > 0 Then TDebe = Valor
       If THaber > 0 Then THaber = Valor
       'MsgBox CuentaNo & vbCrLf & SaldoDisp
       If SaldoDisp <= MontoMinimo Then
          Valor = Round(Valor, 2)
          If SaldoDisp >= Valor And Valor > 0 Then
            .AddNew
            .Fields("Fecha") = MBFecha
            .Fields("Cuenta_No") = CuentaNo
            .Fields("TP") = TipoProc
            .Fields("Debitos") = Round(TDebe, 2)
            .Fields("Creditos") = Round(THaber, 2)
            .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
             If TipoGrupo Then
                If THaber <> 0 Then
                  .Fields("Saldo_Disp") = SaldoDisp
                  .Fields("T") = Pendiente
                Else
                  .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
                  .Fields("T") = Normal
                End If
             Else
               .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
               .Fields("T") = Normal
             End If
            .Fields("CodigoU") = CodigoUsuario
             NumeroLineas = NumeroLineas + 1
             If NumeroLineas > 34 Then NumeroLineas = 1
            .Fields("IDT") = ID_Trans + 1
            .Fields("Hora") = TiempoTexto
            .Fields("Item") = NumEmpresa
            .Fields("ME") = False
            .Fields("Cheque") = Ninguno
             'MsgBox "........"
             SetUpdate DtaCta
             If TDebe > 0 Then TotalAbonos = TotalAbonos + TDebe
             If THaber > 0 Then TotalAbonos = TotalAbonos + THaber
          End If
       End If
      End With
   End If
  End If
End Sub

Private Sub Command1_Click()
  MensajeEncabado = "LIBRETAS NO PROCESADAS"
  ImprimirAdodc AdoSaldos, True, 1, 9
End Sub

Private Sub Command2_Click()
Dim Dias_Aper As Long
Dim Fecha_Edad As Byte

  DGLibreta.Visible = False
  TotalAbonos = 0
  Monto_Total = Round(CCur(TextMontoMin), 2)
  Cadena = DCTP
  If Cadena = "" Then Cadena = Ninguno
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE TP = '" & Cadena & "' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
   End If
  End With
  Debe = 0: Haber = 0
  Select Case DC_Caja
    Case "C": Haber = Round(CCur(TxtMonto), 2)
    Case "D": Debe = Round(CCur(TxtMonto), 2)
  End Select
  MiTiempo = Time
  Titulo = "Pregunta de Grabacion"
  Mensajes = "Seguro de Grabar Transaccion" & vbCrLf _
           & "en Libretas"
  If BoxMensaje = vbYes Then
     sSQL = "SELECT CL.*,C.Fecha_N,C.Fecha As Fecha_Apert " _
          & "FROM Clientes_Datos_Extras As CL,Clientes As C " _
          & "WHERE CL.T <> 'A' " _
          & "AND CL.Codigo = C.Codigo " _
          & "ORDER BY CL.Cuenta_No "
     SelectAdodc AdoSaldos, sSQL
     With AdoSaldos.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          Contador = 1
          Do While Not .EOF
             Cuenta = .Fields("Cuenta_No")
             Fecha_Edad = Round((CFechaLong(FechaSistema) - CFechaLong(.Fields("Fecha_N"))) / 365)
             Dias_Aper = (CFechaLong(FechaSistema) - CFechaLong(.Fields("Fecha_Apert")))
             If Fecha_Edad >= 18 Then
                Debe = 0: Haber = 0
                Select Case DC_Caja
                  Case "C": Haber = Round(CCur(TxtMonto), 2)
                  Case "D": Debe = Round(CCur(TxtMonto), 2)
                End Select
                If Dias_Aper >= 30 And TipoProc = "NDMT" Then
                   InsertarND_NC FechaSistema, AdoCaja, Cuenta, Debe, Haber, Monto_Total
                   ProcesarND_NC.Caption = Format(MiTiempo - Time, "hh:mm:ss") & " " & Contador & "/" & .RecordCount & " Cuenta_No. " & Cuenta & "       Total Debitos Realizados: " & Format(TotalAbonos, "#,##0.00")
                   Contador = Contador + 1
                Else
                   InsertarND_NC FechaSistema, AdoCaja, Cuenta, Debe, Haber, Monto_Total
                   ProcesarND_NC.Caption = Format(MiTiempo - Time, "hh:mm:ss") & " " & Contador & "/" & .RecordCount & " Cuenta_No. " & Cuenta & "       Total Debitos Realizados: " & Format(TotalAbonos, "#,##0.00")
                   Contador = Contador + 1
                End If
             End If
            .MoveNext
          Loop
      End If
     End With
  End If
  ListarMantenimientos
End Sub

Private Sub Command3_Click()
  Unload ProcesarND_NC
End Sub

Private Sub DGLibreta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto ProcesarND_NC, AdoSaldos
End Sub

Private Sub Command4_Click()
  ListarMantenimientos
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT TP " _
       & "FROM Catalogo_Proceso " _
       & "WHERE Nivel = 2 " _
       & "ORDER BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ProcesarND_NC
  ConectarAdodc AdoTP
  ConectarAdodc AdoAux
  ConectarAdodc AdoCaja
  ConectarAdodc AdoSaldos
End Sub

Private Sub TextMontoMin_GotFocus()
  MarcarTexto TextMontoMin
End Sub

Private Sub TextMontoMin_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMontoMin_LostFocus()
  TextoValido TextMontoMin, True
  TextMontoMin.Text = Format(TextMontoMin.Text, "#,##0.00")
End Sub

Private Sub TxtMonto_GotFocus()
  MarcarTexto TxtMonto
End Sub

Private Sub TxtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonto_LostFocus()
  TextoValido TxtMonto, True
  TxtMonto.Text = Format(TxtMonto.Text, "#,##0.00")
End Sub

Public Sub ListarMantenimientos()
  DGLibreta.Visible = False
  TotalAbonos = 0
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No <> '.' " _
       & "ORDER BY Cuenta_No "
  SelectAdodc AdoAux, sSQL
  
  sSQL = "SELECT Cta.T,Cta.Cuenta_No,Cliente,TL.Debitos,TL.Saldo_Disp,TL.Saldo_Cont " _
       & "FROM Trans_Libretas As TL,Clientes As C,Clientes_Datos_Extras As Cta " _
       & "WHERE TL.TP = '" & DCTP & "' " _
       & "AND TL.Fecha = #" & BuscarFecha(FechaSistema) & "# " _
       & "AND C.Codigo = Cta.Codigo " _
       & "AND Cta.Cuenta_No = TL.Cuenta_No " _
       & "ORDER BY Cliente,Cta.Cuenta_No "
  'MsgBox sSQL
  SelectDataGrid DGLibreta, AdoSaldos, sSQL
  'MsgBox sSQL
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TotalAbonos = TotalAbonos + .Fields("Debitos")
          Cuenta_No = .Fields("Cuenta_No")
          SaldoDisp = .Fields("Saldo_Disp")
          If SaldoDisp <= 0 Then
             AdoAux.Recordset.MoveFirst
             AdoAux.Recordset.Find ("Cuenta_No like '" & Cuenta_No & "' ")
             If Not AdoAux.Recordset.EOF Then
                AdoAux.Recordset.Fields("Fecha_A") = FechaSistema
                AdoAux.Recordset.Fields("T") = Anulado
                AdoAux.Recordset.Update
             End If
          End If
         .MoveNext
       Loop
  End If
  End With
  DGLibreta.Visible = True
  RatonNormal
  LabelTotalInt.Caption = Format(TotalAbonos, "#,##0.00")
End Sub
