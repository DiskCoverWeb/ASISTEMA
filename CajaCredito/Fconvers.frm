VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FConversion 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   7350
      Picture         =   "Fconvers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1785
      Width           =   2325
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Orden Trans"
      Height          =   750
      Left            =   7350
      Picture         =   "Fconvers.frx":09F6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   945
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reprocesar intereses"
      Height          =   750
      Left            =   4935
      Picture         =   "Fconvers.frx":0E38
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1785
      Width           =   2325
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Procesar Cambio de Credito"
      Height          =   750
      Left            =   4935
      Picture         =   "Fconvers.frx":127A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   945
      Width           =   2325
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar Cambio de Libretas"
      Height          =   750
      Left            =   4935
      Picture         =   "Fconvers.frx":16BC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   2325
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4530
      Left            =   105
      TabIndex        =   9
      Top             =   2730
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7990
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "CAJA"
      TabPicture(0)   =   "Fconvers.frx":1AFE
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data1"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "CLICTA"
      TabPicture(1)   =   "Fconvers.frx":1B1A
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Data2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DBGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "CLIENTE"
      TabPicture(2)   =   "Fconvers.frx":1B36
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Data3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DBGrid3"
      Tab(2).Control(1).Enabled=   0   'False
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "Fconvers.frx":1B52
      Tab(3).ControlCount=   2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Data5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "DBGrid4"
      Tab(3).Control(1).Enabled=   0   'False
      Begin VB.Data Data5 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4095
         Width           =   11040
      End
      Begin VB.Data Data3 
         Caption         =   "Data1"
         Connect         =   "dBASE IV;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4095
         Width           =   11040
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "dBASE IV;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74895
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4095
         Width           =   11040
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "dBASE IV;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   105
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   ""
         Top             =   4095
         Width           =   11040
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Fconvers.frx":1B6E
         Height          =   3690
         Left            =   105
         OleObjectBlob   =   "Fconvers.frx":1B7E
         TabIndex        =   10
         Top             =   420
         Width           =   11040
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "Fconvers.frx":2530
         Height          =   3690
         Left            =   -74895
         OleObjectBlob   =   "Fconvers.frx":2540
         TabIndex        =   11
         Top             =   420
         Width           =   11040
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "Fconvers.frx":2EF2
         Height          =   3690
         Left            =   -74895
         OleObjectBlob   =   "Fconvers.frx":2F02
         TabIndex        =   12
         Top             =   420
         Width           =   11040
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "Fconvers.frx":38B4
         Height          =   3690
         Left            =   -74895
         OleObjectBlob   =   "Fconvers.frx":38C4
         TabIndex        =   14
         Top             =   420
         Width           =   11040
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar la base de datos"
      Height          =   750
      Left            =   7350
      Picture         =   "Fconvers.frx":4276
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   5
      Top             =   0
      Width           =   2430
      Begin VB.OptionButton Option2 
         Caption         =   "*.DBF"
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
         Left            =   1155
         TabIndex        =   7
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "*.MDB"
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
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   2430
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   105
      TabIndex        =   1
      Top             =   1155
      Width           =   2430
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   2625
      TabIndex        =   0
      Top             =   315
      Width           =   2220
   End
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   525
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2940
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataCtaNo 
      Caption         =   "CtaNo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   525
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2835
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataInt 
      Caption         =   "Int"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3150
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataSaldos 
      Caption         =   "Saldos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3150
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ARCHIVOS"
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
      Left            =   2625
      TabIndex        =   3
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIDAD Y DIRECTORIO"
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
      TabIndex        =   2
      Top             =   630
      Width           =   2430
   End
End
Attribute VB_Name = "FConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub GenerarTablaPrestamoConv(DtaTabla As Data, TxtInt As Single, TxTMeses As Single, TxtMonto As Double, Meses_Dias As Boolean)
  'Interes = Round_ME4(TxtInt / 100)
  F = Contador
  If 0 <= Plazo And Plazo <= 6 Then
     Interes = Round_ME4(12 / 100)
  ElseIf 6 < Plazo And Plazo <= 12 Then
     Interes = Round_ME4(11.17 / 100)
  ElseIf 12 < Plazo And Plazo <= 24 Then
     Interes = Round_ME4(9.5 / 100)
  ElseIf 24 < Plazo And Plazo <= 36 Then
     Interes = Round_ME4(11.65 / 100)
  Else
     Interes = Round_ME4(12 / 100)
  End If
  Numero = Round_ME(TxTMeses)
  Total = Round_ME(TxtMonto)
  Saldo = Total:  Valor_ME = 0
  Total_ME = 0:  Valor = 0
  sSQL = "SELECT * FROM Tabla_Prestamo "
  SelectData DtaTabla, sSQL
  sSQL = "DELETE * FROM Tabla_Prestamo "
  DeleteData DtaTabla, sSQL
  With DtaTabla.Recordset
    If Meses_Dias Then   'Si_No = True Dias else Meses
       For I = 0 To 1
          .AddNew
          .Fields("Mes_No") = I
          .Fields("Dia") = DiasLetras(WeekDay(MiFecha))
           NoDias = Numero
           If I = 0 Then
              Valor_ME = Round_ME(((Total * Interes) / 360) * (NoDias + 3))
              Total_ME = Round_ME(Total)
              Valor = Round_ME(Total + Valor_ME)
             .Fields("Fecha") = MiFecha
             .Fields("Capital") = Total_ME
             .Fields("Interes") = Valor_ME
             .Fields("Pagos") = 0
             .Fields("Saldo") = Round_ME(Total + Valor_ME)
              'MiFecha = SiguienteMes(MiFecha)
              MiFecha = CLongFecha(CFechaLong(MiFecha) + Numero)
           Else
             .Fields("Fecha") = MiFecha
             .Fields("Capital") = 0
             .Fields("Interes") = 0
             .Fields("Pagos") = Total
             .Fields("Saldo") = 0
           End If
          .Update
       Next I
    Else
       Total = Round_ME(Total + (Total * ((Numero / 12) * Interes)))
       Tasa = 0
       Do
         Tasa = Round_ME4(Tasa + 0.0001)
         Cuota = Round_ME(((Saldo * Tasa) / 12) / (1 - (1 + (Tasa / 12)) ^ -Numero))
       Loop Until (Cuota * Numero) >= Total
       Contador = 1: Total = Saldo
       Valor = Round_ME(((12 * Total) + (Total * Interes * Numero)) / (12 * Numero))
       Valor_ME = 0: Total_ME = 0: Comision = 0
       For I = 0 To Numero
          .AddNew
          .Fields("Mes_No") = I
          .Fields("Dia") = DiasLetras(WeekDay(MiFecha))
           If I = 0 Then
             .Fields("Fecha") = MiFecha
             .Fields("Capital") = 0
             .Fields("Interes") = 0
             .Fields("Comision") = 0
             .Fields("Pagos") = 0
              MiFecha = SiguienteMes(MiFecha)
              'MiFecha = CLongFecha(CFechaLong(MiFecha) + 30)
           Else
             .Fields("Fecha") = MiFecha
             .Fields("Capital") = Total_ME
             .Fields("Interes") = Valor_ME
             .Fields("Comision") = Comision
             .Fields("Pagos") = Valor
              MiFecha = SiguienteMes(MiFecha)
              'MiFecha = CLongFecha(CFechaLong(MiFecha) + 30)
           End If
          .Fields("Saldo") = Total
          .Update
          'Comision del 1%
           Comision = Round_ME(Total * 0.012)
          'Interes Inicial
           Valor_ME = Round_ME(Total * (Tasa / 12))
          'Amortizacion o Capital
           Total_ME = Round_ME(Valor - Valor_ME)
          'Saldo Pendiente
           Total = Round_ME(Total - Total_ME)
          'Interes Final
           Valor_ME = Round(Valor - Total_ME - Comision)
          'Valor_ME = Round_ME(Total * (Tasa / 12))
          'Total_ME = Round_ME(Valor - Valor_ME)
          'Total = Round_ME(Total - Total_ME)
           Contador = Contador + 1
       Next I
    End If
  End With
  sSQL = "SELECT * FROM Tabla_Prestamo "
  sSQL = sSQL & "ORDER BY Mes_No "
  SelectData DtaTabla, sSQL
  If Meses_Dias = False Then
     With DtaTabla.Recordset
      If .RecordCount > 0 Then
         .MoveLast
         .Edit
         'Comision del 1%
          Valor = Round_ME(.Fields("Interes"))
          Total = Round_ME(.Fields("Capital"))
          Abono = Round_ME(.Fields("Pagos"))
          Saldo = Round_ME(.Fields("Saldo"))
          'MsgBox Total
          Comision = Round_ME(Total * 0.012)
         .Fields("Interes") = Abono - Total - Saldo - Comision
         .Fields("Comision") = Comision
         .Fields("Capital") = Total + Saldo
         .Fields("Saldo") = 0
         .Update
      End If
     End With
  End If
  Contador = F
End Sub

Public Sub InsertarMontosInt(MBFecha As String, DtaCta As Data, CuentaNo As String, Valor As Double)
  If ((CuentaNo <> "00000000-0") And (Valor > 0)) Then
  TiempoTexto = Format(Time, FormatoTimes)
  SaldoDisp = 0: SaldoCont = 0
  sSQL = "SELECT * FROM Transacciones " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "AND CL = True " _
       & "ORDER BY Fecha,IDT,Hora,ID "
  SelectData DtaCta, sSQL, False
  With DtaCta.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           NumeroLineas = .Fields("ID")
           ID_Trans = .Fields("IDT")
           If NumeroLineas <= 0 Then NumeroLineas = 1
       End If
      .AddNew
      .Fields("CL") = True
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
      .Fields("Usuario") = NombreUsuario
      .Fields("ID") = NumeroLineas + 1
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
      .Update
  End With
  Saldo = 0
  sSQL = "SELECT * FROM Cuentas " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' "
  SelectData DtaCta, sSQL, False
  With DtaCta.Recordset
   If .RecordCount > 0 Then
      .Edit
      .Fields("Saldo_Cont") = SaldoCont + Valor
      .Fields("Saldo_Disp") = SaldoDisp + Valor
       Saldo = Round_ME(SaldoDisp + Valor)
       If Saldo < 0 Then Saldo = 0
      .Fields("Fecha_T") = FechaSistema
      .Fields("Interes") = Saldo
      .Update
   End If
  End With
  End If
End Sub

Public Sub InsertarMontosConv(MBFecha As String, DtaCta As Data, DtaBanc As Data, CuentaNo As String, TDebe As Double, THaber As Double, NoCheque As String, NomBanco As String)
  SaldoDisp = 0: SaldoCont = 0
  sSQL = "SELECT * FROM Transacciones " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "AND CL = True " _
       & "ORDER BY Fecha,ID,Hora "
  SelectData DtaCta, sSQL, False
  With DtaCta.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
       End If
      .AddNew
      .Fields("CL") = True
      .Fields("Fecha") = MBFecha
      .Fields("Cuenta_No") = CuentaNo
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
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
      .Fields("Usuario") = NombreUsuario
      .Fields("ID") = NumeroLineas + 1
      .Fields("Hora") = Format(Time, FormatoTimes)
      .Fields("Item") = NumEmpresa
      .Fields("ME") = Moneda_US
      .Fields("Cheque") = NoCheque
      .Update
      'Caja
       Select Case TipoProc
         Case "APER", "DEP", "RET", "DEPS", "RETS", "BOVE", "DEPC", "DDA", "DDAC"
             .AddNew
             .Fields("CL") = False
             .Fields("Fecha") = MBFecha
             .Fields("Cuenta_No") = CuentaNo
             .Fields("TP") = TipoProc
             .Fields("Debitos") = THaber
             .Fields("Creditos") = TDebe
             .Fields("Saldo_Cont") = 0
             .Fields("Saldo_Disp") = 0
             .Fields("T") = Normal
             .Fields("Usuario") = NombreUsuario
             .Fields("ID") = 0
             .Fields("Hora") = Format(Time, FormatoTimes)
             .Fields("Item") = NumEmpresa
             .Fields("ME") = Moneda_US
             .Fields("Cheque") = NoCheque
             .Update
       End Select
  End With
  If TipoGrupo Then
     Select Case TipoProc
       Case "DEPC", "DEAC"
            sSQL = "SELECT * FROM Bancos " _
                 & "WHERE Cuenta_No = '" & CuentaNo & "' "
            SelectData DataCtaNo, sSQL, False
            With DtaBanc.Recordset
                .AddNew
                .Fields("T") = Pendiente
                .Fields("ME") = Moneda_US
                .Fields("Fecha") = MBFecha
                .Fields("Cuenta_No") = CuentaNo
                .Fields("TP") = TipoProc
                .Fields("Banco") = NomBanco
                .Fields("Cheque") = NoCheque
                .Fields("Valor") = THaber - TDebe
                .Fields("Dias") = 4
                .Fields("Item") = NumEmpresa
                .Update
            End With
     End Select
  End If
End Sub

Private Sub Command1_Click()
   Data1.DatabaseName = Dir1.Path & "\"
   Data1.RecordSource = "CAJA"
   Data1.Refresh
   Data2.DatabaseName = Dir1.Path & "\"
   Data2.RecordSource = "CLICTA"
   Data2.Refresh
   Data3.DatabaseName = Dir1.Path & "\"
   Data3.RecordSource = "CLIENTE"
   Data3.Refresh
End Sub

Private Sub Command2_Click()
Dim Contador As Long
  RatonReloj
  Suma_MN = 0
  Contador = 0
  MiTiempo = Time
  TipoProc = "APER"
    
  sSQL = "DELETE * FROM Cuentas "
  DeleteData DataCuentas, sSQL
  sSQL = "DELETE * FROM Trans_Libretas "
  DeleteData DataSaldos, sSQL
  sSQL = "DELETE * FROM Bloqueos "
  DeleteData DataCtaNo, sSQL
  
  sSQL = "SELECT * FROM Cuentas "
  SelectData DataCuentas, sSQL
  sSQL = "SELECT * FROM Trans_Libretas "
  SelectData DataSaldos, sSQL
  sSQL = "SELECT * FROM Bloqueos "
  SelectData DataCtaNo, sSQL
  
  sSQL = "SELECT * FROM Libretas "
  sSQL = sSQL & "ORDER BY CM2CUENTA "
  SelectData Data5, sSQL, False
  With Data5.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Nombres = Ninguno
          NoCheque = Ninguno
          Haber = Valor_Decimal(.Fields("CM2SCEFEC"))
          SaldoDisp = Valor_Decimal(.Fields("CM2SCEFEC"))
          SaldoCont = Valor_Decimal(.Fields("CM2SCTOTA"))
          Valor = Valor_Decimal(.Fields("CM2SCBLOQ"))
          Apellidos = .Fields("Socios")
          Cuenta_No = Mid(.Fields("CM2CUENTA"), 1, 8) & "-" & Mid(.Fields("CM2CUENTA"), 9, 1)
          If Apellidos = "" Then Apellidos = Ninguno
          DataCuentas.Recordset.AddNew
          SetField DataCuentas, "T", "N"
          SetField DataCuentas, "TC", "CJ"
          SetField DataCuentas, "PJ", True
          SetField DataCuentas, "ME", False
          SetField DataCuentas, "Fecha", FechaSistema
          SetField DataCuentas, "Cuenta_No", Cuenta_No
          SetField DataCuentas, "Representante", Apellidos
          SetField DataCuentas, "Nombres", Nombres
          SetField DataCuentas, "Apellidos", Apellidos
          SetField DataCuentas, "Profesion", Ninguno
          SetField DataCuentas, "Actividad", Ninguno
          SetField DataCuentas, "Casilla_Postal", Ninguno
          SetField DataCuentas, "RUC_CI", FormatoCodigoRUC_CI(.Fields("CM2DCCEDU"), True)
          SetField DataCuentas, "CI", FormatoCodigoRUC_CI(.Fields("CM2DCCEDU"), False)
          SetField DataCuentas, "Est_Civil", Ninguno
          SetField DataCuentas, "No_Soc", 1
          SetField DataCuentas, "No_Dep", 1
          SetField DataCuentas, "Ciudad", "Lago Agrio"
          SetField DataCuentas, "LugarTrabajo", Ninguno
          SetField DataCuentas, "FAX", Ninguno
          SetField DataCuentas, "Telefono", Ninguno
          SetField DataCuentas, "TelefonoT", Ninguno
          SetField DataCuentas, "Direccion", Ninguno
          SetField DataCuentas, "DireccionT", Ninguno
          SetField DataCuentas, "Usuario", NombreUsuario
          SetField DataCuentas, "Sector", Ninguno
          SetField DataCuentas, "Area", Ninguno
          SetField DataCuentas, "Item", NumEmpresa
          DataCuentas.Recordset.Update
         'Libretas
          Debe = 0:
          DataSaldos.Recordset.AddNew
          SetField DataSaldos, "T", Normal
          SetField DataSaldos, "Fecha", FechaSistema
          SetField DataSaldos, "Cuenta_No", Cuenta_No
          SetField DataSaldos, "TP", TipoProc
          SetField DataSaldos, "Debitos", Debe
          SetField DataSaldos, "Creditos", Haber
          SetField DataSaldos, "Saldo_Cont", SaldoCont
          SetField DataSaldos, "Saldo_Disp", SaldoDisp + Valor
          SetField DataSaldos, "Usuario", NombreUsuario
          SetField DataSaldos, "ID", 1
          SetField DataSaldos, "IDT", 1
          SetField DataSaldos, "Hora", Format(Time, FormatoTimes)
          SetField DataSaldos, "Item", NumEmpresa
          SetField DataSaldos, "ME", False
          SetField DataSaldos, "Cheque", NoCheque
          DataSaldos.Recordset.Update
          If Round(Valor) > 0 Then
             DataCtaNo.Recordset.AddNew
             SetField DataCtaNo, "T", Normal
             SetField DataCtaNo, "Fecha", FechaSistema
             SetField DataCtaNo, "Cuenta_No", Cuenta_No
             SetField DataCtaNo, "Valor", Round(Valor)
             SetField DataCtaNo, "Item", NumEmpresa
             DataCtaNo.Recordset.Update
          End If
          Contador = Contador + 1
          FConversion.Caption = Cuenta_No & " Registros actualizados: " & Contador & "/" & .RecordCount & ", Tiempo: " & Format(Time - MiTiempo, "hh:mm:ss")
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub Command3_Click()
Dim Contador As Long
  RatonReloj
  Suma_MN = 0
  Contador = 0
  MiTiempo = Time

  sSQL = "SELECT * FROM Garantes "
  SelectData DataCuentas, sSQL
  sSQL = "SELECT * FROM Prestamos "
  SelectData DataSaldos, sSQL
  sSQL = "SELECT * FROM Trans_Prestamos "
  SelectData DataSaldos, sSQL
  
  sSQL = "DELETE * FROM Garantes "
  DeleteData DataCuentas, sSQL
  sSQL = "DELETE * FROM Prestamos "
  DeleteData DataSaldos, sSQL
  sSQL = "DELETE * FROM Trans_Prestamos "
  DeleteData DataSaldos, sSQL
  
  sSQL = "SELECT * FROM PrestamosVigentes "
  sSQL = sSQL & "ORDER BY NoPrestamo "
  SelectData Data5, sSQL, False
  With Data5.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TipoProc = "S/F"
          Apellidos = Ninguno
          Nombres = Ninguno
          MiFecha = .Fields("Fecha_Inic")
          Numero = .Fields("NoPrestamo")
          Cuenta_No = Mid(.Fields("Cuenta"), 1, 8) & "-" & Mid(.Fields("Cuenta"), 9, 1)
          DataCuentas.Recordset.AddNew
          SetField DataCuentas, "TP", TipoProc
          SetField DataCuentas, "Credito_No", Numero
          SetField DataCuentas, "Cuenta_No", Cuenta_No
          SetField DataCuentas, "Nombres", Nombres
          SetField DataCuentas, "Apellidos", Apellidos
          SetField DataCuentas, "CI", "000000000-0"
          SetField DataCuentas, "LugarTrabajo", Ninguno
          SetField DataCuentas, "Telefono", Ninguno
          SetField DataCuentas, "Direccion", Ninguno
          DataCuentas.Recordset.Update
          Contador = Contador + 1
          FConversion.Caption = Numero & " Registros actualizados: " & Contador & "/" & .RecordCount & ", Tiempo: " & Format(Time - MiTiempo, "hh:mm:ss")
         .MoveNext
       Loop
   End If
  End With
    
  sSQL = "SELECT * FROM Prestamos "
  SelectData DataCuentas, sSQL
  
  sSQL = "SELECT * FROM Trans_Prestamos "
  SelectData DataSaldos, sSQL
  
  sSQL = "SELECT * FROM PrestamosVigentes "
  sSQL = sSQL & "ORDER BY NoPrestamo "
  SelectData Data5, sSQL, False
  ProcesarPrestamos Data5
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub Command4_Click()
  'Reprocesar intereses
  sSQL = "SELECT T.* FROM Transacciones As T,Cuentas As C " _
       & "WHERE T.Cuenta_No = C.Cuenta_No " _
       & "AND T.CL = True " _
       & "ORDER BY T.Cuenta_No,T.Fecha,T.IDT,T.Hora,T.ID "
  SelectData DataCuentas, sSQL
  Contador = 1
  MiTiempo = Time
  With DataCuentas.Recordset
   If .RecordCount > 0 Then
       MiFecha = .Fields("Fecha")
       SaldoAnterior = 0
       Do While Not .EOF
          FechaIni = MiFecha
          MiFecha = .Fields("Fecha")
          Cuenta = .Fields("Cuenta_No")
          Saldo = .Fields("Saldo_Disp")
          Contador = Contador + 1
          FConversion.Caption = "Registros actualizados: " & Contador & "/" & .RecordCount & ", Tiempo: " & Format(Time - MiTiempo, "hh:mm:ss")
          TotalInteres = 0: SaldoFinal = 0
          If Saldo < 0 Then Saldo = 0
          sSQL = "SELECT * FROM Tasa_Interes " _
               & "WHERE ME = " & Val(Moneda_US) & " " _
               & "ORDER BY Desde "
          SelectData DataInt, sSQL, False
          If DataInt.Recordset.RecordCount > 0 Then
             Do While Not DataInt.Recordset.EOF
                SaldoInic = DataInt.Recordset.Fields("Desde")
                SaldoFinal = DataInt.Recordset.Fields("Hasta")
                If (SaldoInic <= Saldo) And (Saldo < SaldoFinal) Then
                   TotalInteres = DataInt.Recordset.Fields("Interes")
                End If
                DataInt.Recordset.MoveNext
             Loop
          End If
          Saldo = Round_ME(((Saldo * TotalInteres) / 100) / 360)
          If Saldo < 0 Then Saldo = 0
          sSQL = "SELECT * FROM Intereses "
          SelectData DataSaldos, sSQL, False
          FechaIniN = CFechaLong(FechaIni)
          FechaFinN = CFechaLong(MiFecha)
          NumDias = (FechaFinN - FechaIniN)
          If NumDias > 0 Then
             For I = 1 To NumDias
                 'MsgBox FechaIni
                 sSQL = "DELETE * FROM Intereses " _
                      & "WHERE Fecha = #" & BuscarFecha(FechaIni) & "# " _
                      & "AND Cuenta_No = '" & Cuenta & "' "
                 DeleteData DataSaldos, sSQL
                 DataSaldos.Recordset.AddNew
                 DataSaldos.Recordset.Fields("Fecha") = FechaIni
                 DataSaldos.Recordset.Fields("Cuenta_No") = Cuenta
                 DataSaldos.Recordset.Fields("Interes") = SaldoAnterior
                 DataSaldos.Recordset.Update
                 FechaIni = CLongFecha(CFechaLong(FechaIni) + 1)
             Next I
          End If
         'MsgBox "Fecha: " & MiFecha
          sSQL = "DELETE * FROM Intereses " _
               & "WHERE Fecha = #" & BuscarFecha(MiFecha) & "# " _
               & "AND Cuenta_No = '" & Cuenta & "' "
          DeleteData DataSaldos, sSQL
          DataSaldos.Recordset.AddNew
          DataSaldos.Recordset.Fields("Fecha") = MiFecha
          DataSaldos.Recordset.Fields("Cuenta_No") = Cuenta
          DataSaldos.Recordset.Fields("Interes") = Saldo
          SaldoAnterior = Saldo
          DataSaldos.Recordset.Update
         .MoveNext
       Loop
    End If
  End With
End Sub

Private Sub Command5_Click()
  sSQL = "UPDATE Transacciones SET IDT = 0 "
  UpdateData DataCtaNo, sSQL
  
  sSQL = "SELECT * FROM Transacciones " _
       & "WHERE CL = True " _
       & "ORDER BY Cuenta_No,Fecha,Hora,ID "
  SelectData DataCtaNo, sSQL, False
  With DataCtaNo.Recordset
   If .RecordCount > 0 Then
       ID_Trans = 0: Contador = 0
       Cuenta = .Fields("Cuenta_No")
       Do While Not .EOF
          Contador = Contador + 1
          ID_Trans = ID_Trans + 1
          FConversion.Caption = Contador & "/" & .RecordCount
         .Edit
          If Cuenta <> .Fields("Cuenta_No") Then
             Cuenta = .Fields("Cuenta_No")
             ID_Trans = 0
          End If
         .Fields("IDT") = ID_Trans
         .Update
         .MoveNext
       Loop
   End If
  End With
End Sub

Private Sub Command6_Click()
  Unload FConversion
End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
   RutaBackup = Mid(Drive1.Drive, 1, 2) & "\SYSBASES"
   DataInt.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
   DataCtaNo.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
   DataSaldos.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
   DataCuentas.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
   Data4.DatabaseName = RutaBackup & "\AUXILIAR.MDB"
   Data5.DatabaseName = RutaBackup & "\AUXILIAR.MDB"
End Sub

Private Sub Option1_Click()
  File1.filename = Dir1.Path & "\*.MDB"
End Sub

Private Sub Option2_Click()
  File1.filename = Dir1.Path & "\*.DBF"
End Sub

Public Sub ProcesarPrestamos(Dta5 As Data)
With Dta5.Recordset
   Contador = 0
   If .RecordCount > 0 Then
      .MoveFirst
       TipoProc = "S/F"
       Tasa = 20
       MiFecha = .Fields("Fecha_Inic")
       Numero = .Fields("NoPrestamo")
       NumComp = .Fields("NoPrestamo")
       Monto_Total = .Fields("Monto")
       Saldo = .Fields("Saldo_Capital")
       Saldo_ME = .Fields("Saldo_Capital")
       Plazo = .Fields("Cuotas")
       NoDias = .Fields("No_Div")
       NoCert = Plazo - NoDias
       Cuenta_No = Mid(.Fields("Cuenta"), 1, 8) & "-" & Mid(.Fields("Cuenta"), 9, 1)
       Do While Not .EOF
          If NumComp <> .Fields("NoPrestamo") Then
             'MsgBox NumComp
             DataCuentas.Recordset.AddNew
             SetField DataCuentas, "T", Procesado
             SetField DataCuentas, "TP", TipoProc
             SetField DataCuentas, "ME", False
             SetField DataCuentas, "Tasa", Tasa
             SetField DataCuentas, "Plazo", Plazo
             SetField DataCuentas, "Credito_No", NumComp
             SetField DataCuentas, "Cuenta_No", Cuenta_No
             SetField DataCuentas, "Meses", Plazo
             SetField DataCuentas, "Dia", DiasLetras(WeekDay(MiFecha))
             SetField DataCuentas, "Fecha", MiFecha
             SetField DataCuentas, "Interes", Round_ME((Monto_Total * Tasa) / 100)
             SetField DataCuentas, "Capital", Monto_Total
             SetField DataCuentas, "Pagos", Saldo_ME
             SetField DataCuentas, "Saldo_Pendiente", Saldo_ME
             SetField DataCuentas, "Patrimonio", 0
             SetField DataCuentas, "Encaje", 0
             DataCuentas.Recordset.Update
             'MsgBox Plazo & Chr(13) & NoDias & Chr(13) & NoCert
             GenerarTablaPrestamoConv Data4, Tasa, Plazo, Monto_Total, False
             
             sSQL = "DELETE * FROM Tabla_Prestamo "
             sSQL = sSQL & "WHERE Mes_No <= " & NoCert & " "
             DeleteData Data4, sSQL
             
             sSQL = "SELECT * FROM Tabla_Prestamo "
             SelectData Data4, sSQL
             If Data4.Recordset.RecordCount > 0 Then
                'MsgBox NumComp
                ProcesarPrestamosMeses Data4
             End If
             TipoProc = "S/F"
             Tasa = 20
             MiFecha = .Fields("Fecha_Inic")
             Numero = .Fields("NoPrestamo")
             NumComp = .Fields("NoPrestamo")
             Monto_Total = .Fields("Monto")
             Saldo = .Fields("Saldo_Capital")
             Saldo_ME = .Fields("Saldo_Capital")
             Plazo = .Fields("Cuotas")
             NoDias = .Fields("No_Div")
             NoCert = Plazo - NoDias
             Cuenta_No = Mid(.Fields("Cuenta"), 1, 8) & "-" & Mid(.Fields("Cuenta"), 9, 1)
          End If
          Contador = Contador + 1
          FConversion.Caption = "Registros actualizados: " & Contador & "/" & .RecordCount & ", Tiempo: " & Format(Time - MiTiempo, "hh:mm:ss")
         .MoveNext
       Loop
       DataCuentas.Recordset.AddNew
       SetField DataCuentas, "T", Procesado
       SetField DataCuentas, "TP", TipoProc
       SetField DataCuentas, "ME", False
       SetField DataCuentas, "Tasa", Tasa
       SetField DataCuentas, "Plazo", Plazo
       SetField DataCuentas, "Credito_No", NumComp
       SetField DataCuentas, "Cuenta_No", Cuenta_No
       SetField DataCuentas, "Meses", Plazo
       SetField DataCuentas, "Dia", DiasLetras(WeekDay(MiFecha))
       SetField DataCuentas, "Fecha", MiFecha
       SetField DataCuentas, "Interes", Round_ME((Monto_Total * Tasa) / 100)
       SetField DataCuentas, "Capital", Monto_Total
       SetField DataCuentas, "Pagos", SaldoTotal
       SetField DataCuentas, "Saldo_Pendiente", Saldo
       SetField DataCuentas, "Patrimonio", 0
       SetField DataCuentas, "Encaje", 0
       DataCuentas.Recordset.Update
       GenerarTablaPrestamoConv Data4, Tasa, Plazo, Monto_Total, False
       sSQL = "DELETE * FROM Tabla_Prestamo "
       sSQL = sSQL & "WHERE Mes_No <= " & NoCert & " "
       DeleteData Data4, sSQL
             
       sSQL = "SELECT * FROM Tabla_Prestamo "
       SelectData Data4, sSQL
       If Data4.Recordset.RecordCount > 0 Then
          ProcesarPrestamosMeses Data4
       End If
   End If
  End With
End Sub

Public Sub ProcesarPrestamosMeses(Dta5 As Data)
With Dta5.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          FechaTexto = .Fields("Fecha")
          DataSaldos.Recordset.AddNew
          SetField DataSaldos, "T", Procesado
          SetField DataSaldos, "TP", TipoProc
          SetField DataSaldos, "ME", False
          SetField DataSaldos, "Credito_No", NumComp
          SetField DataSaldos, "Cuenta_No", Cuenta_No
          SetField DataSaldos, "Mes_No", .Fields("Mes_No")
          SetField DataSaldos, "Dia", DiasLetras(WeekDay(FechaTexto))
          SetField DataSaldos, "Fecha", FechaTexto
          SetField DataSaldos, "Interes", .Fields("Interes")
          SetField DataSaldos, "Comision", .Fields("Comision")
          SetField DataSaldos, "Capital", .Fields("Capital")
          SetField DataSaldos, "Pagos", .Fields("Pagos")
          SetField DataSaldos, "Saldo", .Fields("Saldo")
          DataSaldos.Recordset.Update
          'FConversion.Caption = "(" & .Fields("Mes_No") & "), Registros actualizados: " & Contador & "/" & .RecordCount & ", Tiempo: " & Format(Time - MiTiempo, "hh:mm:ss")
         .MoveNext
       Loop
   End If
  End With
End Sub

Public Function Valor_Decimal(ValorD As Double) As Double
  Cadena = Format(ValorD, "000000000000")
  Cadena1 = Mid(Cadena, 1, 10) & "." & Mid(Cadena, 11, 2)
  'MsgBox "C 123456789012" & Chr(13)
  '     & "C " & Cadena & Chr(13)
  '     & "N " & Cadena1
  Valor_Decimal = Cadena1
End Function
