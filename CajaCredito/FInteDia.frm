VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FIntLibretas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ACREDITAR INTERESES MENSUALES EN LIBRETAS"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCTipoLibreta 
      Bindings        =   "FInteDia.frx":0000
      DataSource      =   "AdoTipoLibreta"
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   105
      Width           =   7260
      _ExtentX        =   12806
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Listar Intereses"
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
      Left            =   105
      Picture         =   "FInteDia.frx":001D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3675
      Width           =   1800
   End
   Begin VB.TextBox TxTAño 
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
      Left            =   1890
      MaxLength       =   4
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FInteDia.frx":0327
      Top             =   105
      Width           =   750
   End
   Begin VB.ListBox LMes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DGInt 
      Bindings        =   "FInteDia.frx":032E
      Height          =   5895
      Left            =   1995
      TabIndex        =   10
      Top             =   525
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   2835
      Top             =   1470
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Height          =   960
      Left            =   105
      Picture         =   "FInteDia.frx":0343
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5775
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Acreditar &Intereses"
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
      Left            =   105
      Picture         =   "FInteDia.frx":0C0D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4725
      Width           =   1800
   End
   Begin MSAdodcLib.Adodc AdoInt 
      Height          =   330
      Left            =   1995
      Top             =   6405
      Width           =   5265
      _ExtentX        =   9287
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
      Caption         =   "Int"
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
   Begin MSAdodcLib.Adodc AdoTipoLibreta 
      Height          =   330
      Left            =   2730
      Top             =   1155
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Año"
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
      Width           =   1800
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
      Left            =   8295
      TabIndex        =   7
      Top             =   6405
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTALES"
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
      Left            =   7245
      TabIndex        =   8
      Top             =   6405
      Width           =   1065
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Mes de Interes"
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
      Top             =   525
      Width           =   1800
   End
End
Attribute VB_Name = "FIntLibretas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub InsertarMontosInt(MBFecha As String, _
                             DtaCta As Adodc, _
                             CuentaNo As String, _
                             Valor As Currency)
  TiempoTexto = Format(Time, FormatoTimes)
  If (CuentaNo <> "00000000-0") Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  sSQL = "SELECT TOP 1 * FROM Trans_Libretas " _
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
       End If
      Valor = Redondear(Valor, 2)
      If Valor > 0 Then
      .AddNew
      .Fields("Fecha") = MBFecha
      .Fields("Cuenta_No") = CuentaNo
      .Fields("TP") = "INT"
      .Fields("Debitos") = 0
      .Fields("Creditos") = Valor
      .Fields("Saldo_Cont") = SaldoCont + Valor
       If TipoGrupo Then
          If THaber <> 0 Then
            .Fields("Saldo_Disp") = SaldoDisp
            .Fields("T") = Pendiente
          Else
            .Fields("Saldo_Disp") = SaldoDisp + Valor
            .Fields("T") = Normal
          End If
       Else
         .Fields("Saldo_Disp") = SaldoDisp + Valor
         .Fields("T") = Normal
       End If
      .Fields("CodigoU") = CodigoUsuario
       If NumeroLineas < 0 Then NumeroLineas = 0
       If NumeroLineas > 36 Then NumeroLineas = 0
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
       SetUpdate DtaCta
      End If
      If Valor > 0 Then
         sSQL = "UPDATE Trans_Saldo_Libretas " _
              & "SET P = True " _
              & "WHERE Fecha <= #" & BuscarFecha(Mifecha) & "# " _
              & "AND Cuenta_No = '" & CuentaNo & "' "
         ConectarAdoExecute sSQL
      End If
  End With
  End If
End Sub

Private Sub Command1_Click()
  RatonReloj
  Si_No = True
  Dia = 1
  Mes = Val(SinEspaciosIzq(LMes.Text))
  Anio = Val(TxTAño.Text)
  Select Case Mes
    Case 1, 3, 5, 7, 8, 10, 12
         Dia = 31
    Case 4, 6, 9, 11
         Dia = 30
    Case 2
         If (Anio Mod 4 <> 0) Then Dia = 28
         If (Anio Mod 4 = 0) Then Dia = 29
  End Select
  FechaIni = "01/01/" _
           & Format(Anio, "0000")
           
  FechaFin = Format(Dia, "00") & "/" _
           & Format(Mes, "00") & "/" _
           & Format(Anio, "0000")
           
  Mifecha = FechaFin
  FechaIni = BuscarFecha(FechaIni)
  FechaFin = BuscarFecha(FechaFin)
  Total = 0
  DGInt.Visible = False
  RatonReloj
  TipoCta = DCTipoLibreta
  If TipoCta = "" Then TipoCta = Ninguno
  
'''  sSQL = "UPDATE Trans_Intereses " _
'''       & "SET Interes = ROUND(Interes,2) " _
'''       & "WHERE Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND P = " & Val(adFalse) & " "
'''  ConectarAdoExecute sSQL
'''
'''  sSQL = "DELETE * " _
'''       & "FROM Trans_Intereses " _
'''       & "WHERE Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND Interes <= 0 " _
'''       & "AND P = " & Val(adFalse) & " "
'''  ConectarAdoExecute sSQL
  
  sSQL = "SELECT Ct.Tipo,Cliente,I.Cuenta_No,SUM(I.Interes) As InteresAcum " _
       & "FROM Trans_Intereses As I,Clientes As C,Clientes_Datos_Extras As Ct " _
       & "WHERE I.Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Ct.Tipo_Dato = 'LIBRETAS' " _
       & "AND I.P = " & Val(adFalse) & " " _
       & "AND Ct.Tipo = '" & TipoCta & "' " _
       & "AND I.Cuenta_No = Ct.Cuenta_No " _
       & "AND Ct.Codigo = C.Codigo " _
       & "GROUP BY Ct.Tipo,Cliente,I.Cuenta_No "
  SQLDec = "InteresAcum 4|."
  
  SelectDataGrid DGInt, AdoInt, sSQL, SQLDec
  DGInt.Visible = False
  With AdoInt.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Do While Not .EOF
          Total = Total + .Fields("InteresAcum")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGInt.Visible = True
  If Si_No Then
     DGInt.Caption = "MES DE " & UCase(LMes.Text) & " SIN PROCESAR"
  Else
     DGInt.Caption = "MES DE " & UCase(LMes.Text) & " PROCESADO"
  End If
  LabelTotalInt.Caption = Format(Total, "#,##0.0000")
  RatonNormal
End Sub

Private Sub Command2_Click()
  DGInt.Visible = False
  MiTiempo = Time
  RatonReloj
  With AdoInt.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Contador = 1
       Do While Not .EOF
          FIntLibretas.Caption = Format(MiTiempo - Time, "hh:mm:ss") & " " & Contador & "/" & .RecordCount & " Cuenta_No. " & Cuenta
          Cuenta = .Fields("Cuenta_No")
          Valor = .Fields("InteresAcum")
          If Valor > 0 Then InsertarMontosInt FechaSistema, AdoCaja, Cuenta, Valor
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
 'Actualizar los intereses del tipo de cuenta
  TipoCta = DCTipoLibreta
  If TipoCta = "" Then TipoCta = Ninguno
  If SQL_Server Then
     sSQL = "UPDATE Trans_Intereses " _
          & "SET P = " & Val(adTrue) & " " _
          & "FROM Trans_Intereses As I,Clientes As C,Clientes_Datos_Extras As Ct "
  Else
     sSQL = "FROM Trans_Intereses As I,Clientes As C,Clientes_Datos_Extras As Ct " _
          & "SET I.P = " & Val(adTrue) & " "
  End If
  sSQL = sSQL _
       & "WHERE I.Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND I.P = " & Val(adFalse) & " " _
       & "AND Ct.Tipo = '" & TipoCta & "' " _
       & "AND I.Cuenta_No = Ct.Cuenta_No " _
       & "AND Ct.Codigo = C.Codigo "
  ConectarAdoExecute sSQL
  RatonNormal
  Unload FIntLibretas
End Sub

Private Sub Command3_Click()
  Unload FIntLibretas
End Sub

Private Sub DGInt_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGInt.Visible = False
     GenerarDataTexto FIntLibretas, AdoInt
     DGInt.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     Mensajes = "Imprimir Listado"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
        DGInt.Visible = False
        MensajeEncabData = "REPORTE DE INTERES POR: " & DCTipoLibreta
        SQLMsg1 = DGInt.Caption
        ImprimirAdodc AdoInt, 1, 8
        DGInt.Visible = True
     End If
  End If
  
End Sub

Private Sub Form_Activate()
  RatonNormal
  sSQL = "SELECT Tipo,Acreditacion " _
       & "FROM Catalogo_Interes " _
       & "WHERE TP = 'C' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Tipo,Acreditacion " _
       & "ORDER BY Tipo "
  SelectDBCombo DCTipoLibreta, AdoTipoLibreta, sSQL, "Tipo"
  Si_No = False
  Mifecha = FechaSistema
  LMes.Clear
  For NoMeses = 1 To 12
      LMes.AddItem Format(NoMeses, "00") & " - " & MesesLetras(NoMeses)
  Next NoMeses
  TxTAño.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FIntLibretas
  ConectarAdodc AdoInt
  ConectarAdodc AdoCaja
  ConectarAdodc AdoTipoLibreta
End Sub

Private Sub LMes_DblClick()
  SiguienteControl
End Sub

Private Sub LMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxTAño_GotFocus()
  MarcarTexto TxTAño
End Sub

Private Sub TxTAño_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxTAño_LostFocus()
   TextoValido TxTAño
End Sub
