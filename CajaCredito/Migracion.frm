VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FMigracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDOS DIARIOS"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   11475
   Begin MSDataGridLib.DataGrid DGSaldos 
      Bindings        =   "Migracion.frx":0000
      Height          =   2535
      Left            =   105
      TabIndex        =   5
      Top             =   3675
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   4471
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
      Left            =   105
      Top             =   630
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   2100
      TabIndex        =   1
      Top             =   105
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
   Begin VB.CommandButton Command4 
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
      Left            =   6405
      Picture         =   "Migracion.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar Interes Diario"
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
      Left            =   4725
      Picture         =   "Migracion.frx":08E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1590
   End
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   105
      Top             =   6195
      Width           =   11250
      _ExtentX        =   19844
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
   Begin MSAdodcLib.Adodc AdoInt 
      Height          =   330
      Left            =   1890
      Top             =   630
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   3675
      Top             =   630
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Trans"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   105
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
   Begin MSDataGridLib.DataGrid DGClientes 
      Bindings        =   "Migracion.frx":1024
      Height          =   2115
      Left            =   105
      TabIndex        =   6
      Top             =   1155
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   3731
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   3255
      Width           =   11250
      _ExtentX        =   19844
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
   Begin MSDataGridLib.DataGrid DGPrestamos 
      Bindings        =   "Migracion.frx":103E
      Height          =   2115
      Left            =   105
      TabIndex        =   7
      Top             =   6615
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   3731
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
   Begin MSAdodcLib.Adodc AdoPrestamos 
      Height          =   330
      Left            =   105
      Top             =   8715
      Width           =   11250
      _ExtentX        =   19844
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA DE PROCESO"
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
      Width           =   2010
   End
End
Attribute VB_Name = "FMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Inic, Final As Double
  FechaValida MBFecha
  FechaValida MBFechaF
  If NombreUsuario <> "Administrador de Red" Then MBFechaF.Text = MBFecha.Text
  Total = 0
  MiTiempo = Time
  FechaIni = BuscarFecha(MBFecha.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  
  sSQL = "DELETE * " _
       & "FROM Trans_Intereses " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT Cuenta_No " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE T <> 'A' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY Cuenta_No "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       DGSaldos.Visible = False
       Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
          FechaIni = CLongFecha(CFechaLong(MBFecha.Text) - 35)
          FechaIni = BuscarFecha(MBFecha.Text)
          FechaFin = BuscarFecha(MBFechaF.Text)
          For I = CFechaLong(MBFecha) To CFechaLong(MBFechaF)
              FechaTexto = CLongFecha(I)
              Saldo = 0: SaldoAnterior = 0
              TotalInteres = 0: SaldoFinal = 0
              SaldoDisp = 0: SaldoCont = 0
              Cuenta_No = .Fields("Cuenta_No")
             'Consulta de Saldos de Transacciones
              Si_No = False
              sSQL = "SELECT TOP 1 * " _
                   & "FROM Trans_Libretas " _
                   & "WHERE Fecha <= #" & BuscarFecha(FechaTexto) & "# " _
                   & "AND Cuenta_No = '" & Cuenta_No & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND NOT TP = 'DEFR' " _
                   & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
              SelectAdodc AdoTrans, sSQL
              If AdoTrans.Recordset.RecordCount > 0 Then
                 Moneda_US = AdoTrans.Recordset.Fields("ME")
                 SaldoDisp = AdoTrans.Recordset.Fields("Saldo_Disp")
                 SaldoCont = AdoTrans.Recordset.Fields("Saldo_Cont")
              End If
             'Sacar tasa de intereses segun saldo disponible
              FIntereses.Caption = Cuenta_No & " - " & FechaTexto & ": (Total Interes = " & Format(Total, "#,##0.0000") & ") " & Mifecha & " - " & Contador & "/" & .RecordCount & " - " & Format(Time - MiTiempo, "hh:mm:ss")
              Saldo = Interes_Caja(SaldoDisp, "..")
             'MsgBox SaldoDisp
              Total = Total + Saldo
              If Saldo > 0 Then
                 sSQL = "INSERT INTO Trans_Intereses " _
                      & "(AC,P,Fecha,Cuenta_No,Interes,Saldo_Disp,Saldo_Cont,Item) " _
                      & "VALUES " _
                      & "(0,0,#" & BuscarFecha(FechaTexto) & "#,'" & Cuenta_No & "'," & Saldo & "," & SaldoDisp & "," & SaldoCont & ",'" & NumEmpresa & "') "
                 ConectarAdoExecute sSQL
              End If
          Next I
         .MoveNext
       Loop
       DGSaldos.Visible = True
       FMigracion.Caption = "(Total Interes = " & Format(Total, "#,##0.0000") & ") " & Contador & "/" & .RecordCount & " - " & Format(Time - MiTiempo, "hh:mm:ss")
       RatonNormal
   End If
  End With
  SaldoCont = 0: SaldoDisp = 0: TotalInteres = 0
  FechaValida MBFecha
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFecha)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "SELECT C.T,Cta.Cuenta_No,Cliente,SL.Fecha,SL.Saldo_Disp,SL.Saldo_Cont,SL.Interes " _
       & "FROM Clientes_Datos_Extras As Cta,Trans_Intereses As SL,Clientes C " _
       & "WHERE SL.Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Cta.Cuenta_No = SL.Cuenta_No " _
       & "AND Cta.Codigo = C.Codigo " _
       & "ORDER BY Cta.Cuenta_No "
  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  
  RatonNormal
  MBFecha.SetFocus
End Sub

Private Sub Command4_Click()
  Unload FMigracion
End Sub

Private Sub Form_Activate()
  RatonNormal
  If NombreUsuario = "Administrador de Red" Then MBFechaF.Visible = True
  MBFecha.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FMigracion
  ConectarAdodc AdoInt
  ConectarAdodc AdoAux
  ConectarAdodc AdoTrans
  ConectarAdodc AdoSaldos
  ConectarAdodc AdoClientes
  ConectarAdodc AdoPrestamos
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

Public Sub Listar_Intereses_Ganados()
  RatonReloj
  Total = 0
  sSQL = "SELECT C.Cuenta_No,Cl.Cliente As Nombre_Socio,I.Interes,Cl.Representante " _
       & "FROM Clientes_Datos_Extras As C,Clientes As Cl,Trans_Intereses As I " _
       & "WHERE I.Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# " _
       & "AND C.Cuenta_No = I.Cuenta_No " _
       & "AND C.Codigo = Cl.Codigo " _
       & "ORDER BY C.Cuenta_No "
  SQLDec = "Interes 4|."
  SelectDataGrid DGSaldos, AdoSaldos, sSQL, SQLDec
  DGSaldos.Visible = False
  RatonReloj
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Interes")
         .MoveNext
       Loop
   End If
  End With
  LabelPromedio.Caption = Format(Total, "#,##0.0000")
  DGSaldos.Visible = True
  RatonNormal
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
    FechaValida MBFechaF
End Sub
