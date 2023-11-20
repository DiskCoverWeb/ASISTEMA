VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form SaldoBancos 
   Caption         =   "SALDO DE CAJA BANCOS"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGCta 
      Bindings        =   "Saldoban.frx":0000
      Height          =   855
      Left            =   8505
      TabIndex        =   17
      Top             =   525
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin MSDataGridLib.DataGrid DGBanco 
      Bindings        =   "Saldoban.frx":0016
      Height          =   5160
      Left            =   105
      TabIndex        =   15
      ToolTipText     =   "abc"
      Top             =   1785
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9102
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
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   210
      TabIndex        =   16
      Top             =   420
      Width           =   4320
      Begin VB.OptionButton OpcFlujoEfec 
         Caption         =   "Flujo de Efectivo"
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
         Left            =   2205
         TabIndex        =   18
         Top             =   840
         Width           =   2010
      End
      Begin VB.OptionButton OpcEspec 
         Caption         =   "Cuentas Especiales"
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
         Left            =   2205
         TabIndex        =   6
         Top             =   525
         Width           =   2010
      End
      Begin VB.OptionButton OpcCJBA 
         Caption         =   "Flujo Caja Bancos"
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
         Left            =   2205
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1905
      End
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         OLEDragMode     =   1
         OLEDropMode     =   2
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
         OLEDragMode     =   1
         OLEDropMode     =   2
      End
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   840
         TabIndex        =   2
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         OLEDragMode     =   1
         OLEDropMode     =   2
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
         OLEDragMode     =   1
         OLEDropMode     =   2
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Hasta"
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
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Desde"
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
         TabIndex        =   1
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   7455
      Picture         =   "Saldoban.frx":002D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   525
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1590
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   2805
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "SALDO &CAJA/BANCOS"
      TabPicture(0)   =   "Saldoban.frx":08F7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&DETALLE DE CHEQ. GIRADOS"
      TabPicture(1)   =   "Saldoban.frx":0913
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DETALLE DE C&HEQ. POSFECHADOS"
      TabPicture(2)   =   "Saldoban.frx":092F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command7 
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
         Height          =   855
         Left            =   -69540
         Picture         =   "Saldoban.frx":094B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command5 
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
         Height          =   855
         Left            =   -69540
         Picture         =   "Saldoban.frx":1215
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command2 
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
         Height          =   855
         Left            =   6405
         Picture         =   "Saldoban.frx":1ADF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Listar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5460
         Picture         =   "Saldoban.frx":23A9
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Listar Posfech."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -70485
         Picture         =   "Saldoban.frx":27EB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Listar Cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -70485
         Picture         =   "Saldoban.frx":2C2D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4515
         Picture         =   "Saldoban.frx":306F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   960
      End
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   1785
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   6930
      Width           =   11355
      _ExtentX        =   20029
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
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   315
      Top             =   2100
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Ctas1"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2415
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
End
Attribute VB_Name = "SaldoBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload SaldoBancos
End Sub

Private Sub Command2_Click()
  DGBanco.Visible = False
  SQLMsg3 = ""
  SQLMsg2 = "Desde:  " & MBFechaI.Text & "   al   " & MBFechaF.Text
  Mensajes = "Imprimir Resumido"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = vbYes Then
     SQLMsg1 = "TABLERO DE BORDE"
     Imprimir_Saldos_Flujo AdoBanco, 1
  Else
     SQLMsg1 = Cadena1
     ImprimirSaldosBancos AdoBanco, 1
  End If
  DGBanco.Visible = True
End Sub

Private Sub Command3_Click()
Dim TotPresupuesto As Currency
Dim TotDiferencia As Currency
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  DGBanco.Visible = False
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Fecha = #" & FechaIni & "# " _
       & "AND Fecha_Venc = #" & FechaFin & "# " _
       & "AND TP = 'CJBA' "
  Ejecutar_SQL_SP sSQL
  If OpcCJBA.value Then
     Cadena1 = UCaseStrg(OpcCJBA.Caption)
     sSQL = "SELECT ME,TC,Codigo,Cuenta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC IN ('BA','CJ') " _
          & "ORDER BY ME,TC,Codigo "
  ElseIf OpcEspec.value Then
     Cadena1 = UCaseStrg(OpcEspec.Caption)
     sSQL = "SELECT ME,TC,Codigo,Cuenta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC IN ('BA','CJ','C','P','I','G','RF','RI') " _
          & "ORDER BY ME,TC,Codigo "
  Else
     Cadena1 = UCaseStrg(OpcCJBA.Caption)
     sSQL = "SELECT ME,TC,Codigo,Cuenta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'BA' " _
          & "ORDER BY ME,TC,Codigo "
  End If
  Contador = 0
  Select_Adodc AdoCtas1, sSQL
  With AdoCtas1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          TipoProc = .Fields("TC")
          Codigo = .Fields("Codigo")
          Cuenta = .Fields("Cuenta")
          Moneda_US = .Fields("ME")
          Saldo = 0: Saldo_ME = 0
          Contador = Contador + 1
          SaldoBancos.Caption = Cadena1 & " - " & Format(Contador / .RecordCount, "00%")
          TotPresupuesto = 0
          If OpcCJBA.value Or OpcEspec.value Then Procesar_Flujo_de_Caja
         .MoveNext
       Loop
   End If
  End With
  If OpcFlujoEfec.value Then Procesar_Flujo_de_Efectivo
  If OpcCJBA.value Or OpcEspec.value Then ListarSaldosDiarios
  SaldoBancos.Caption = Cadena1
  DGCta.Visible = True
  RatonNormal
End Sub

Private Sub Command4_Click()
  ListarSaldosDiarios
End Sub

Private Sub Command5_Click()
  DGBanco.Visible = False
  Mensajes = "Imprimir los cheques girados"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = vbYes Then
     SQLMsg2 = "CHEQUE GIRADOS DE BANCOS"
     ImprimirCheques AdoBanco, 1
  End If
  DGBanco.Visible = True
End Sub

Private Sub Command6_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "SELECT C.Cuenta,Cliente,T.Fecha,T.TP,T.Numero,T.Cheq_Dep As Cheque_No,T.Haber As Valor " _
       & "FROM Catalogo_Cuentas As C,Comprobantes As Co,Transacciones As T,Clientes As Cl " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Co.T = '" & Normal & "' " _
       & "AND T.TP = '" & CompEgreso & "' " _
       & "AND T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.TP = Co.TP " _
       & "AND T.Numero = Co.Numero " _
       & "AND T.Item = Co.Item " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = Co.Periodo " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Cta = C.Codigo " _
       & "AND Co.Codigo_B = Cl.Codigo " _
       & "ORDER BY C.Codigo,T.Fecha,T.TP,T.Numero "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  DGCta.Visible = False
  RatonNormal
End Sub

Private Sub Command7_Click()
DGBanco.Visible = False
  Mensajes = "Imprimir los cheques Posfechados"
  Titulo = "Pregunta de Impresion"
  If BoxMensaje = vbYes Then
     SQLMsg2 = "CHEQUES POSFECHADOS"
     ImprimirAdodc AdoBanco, True, 1, 8
  End If
  DGBanco.Visible = True
End Sub

Private Sub Command8_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "SELECT C.Cuenta,Cliente,T.Fecha_Efec,T.TP,T.Numero,T.Cheq_Dep As Cheque_No,T.Haber As Valor " _
       & "FROM Catalogo_Cuentas As C,Comprobantes As Co,Transacciones As T,Clientes As Cl " _
       & "WHERE T.Fecha_Efec > #" & FechaFin & "# " _
       & "AND T.T = '" & Normal & "' " _
       & "AND T.TP = '" & CompEgreso & "' " _
       & "AND T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.TP = Co.TP " _
       & "AND T.Numero = Co.Numero " _
       & "AND T.Item = Co.Item " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = Co.Periodo " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Cta = C.Codigo " _
       & "AND Co.Codigo_B = Cl.Codigo " _
       & "ORDER BY C.Codigo,T.Fecha,T.TP,T.Numero "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  DGCta.Visible = False
  RatonNormal
End Sub

Private Sub DGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto SaldoBancos, AdoBanco
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGBanco.AllowUpdate = True
     MsgBox "Proceso Aceptado, puede Modificar"
     DGBanco.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm SaldoBancos
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCtas1
  ConectarAdodc AdoBanco
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF = UltimoDiaMes(MBFechaI)
End Sub

Public Sub ListarSaldosDiarios()
  RatonReloj
  DGBanco.Visible = False
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "SELECT SD.TC,CC.Cuenta,SD.ME,SD.Saldo_Anterior," _
       & "SD.Ingresos As Debitos,SD.Egresos As Creditos,SD.Saldo_Actual,SD.Presupuesto,SD.Diferencia " _
       & "FROM Saldo_Diarios As SD,Catalogo_Cuentas As CC " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.Fecha = #" & FechaIni & "# " _
       & "AND SD.Fecha_Venc = #" & FechaFin & "# " _
       & "AND SD.TP = 'CJBA' " _
       & "AND SD.Cta = CC.Codigo " _
       & "AND SD.Item = CC.Item " _
       & "ORDER BY SD.ME,CC.Codigo,SD.TC,SD.Cta "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  DGBanco.Visible = True
  sSQL = "SELECT ME,TC,SUM(Saldo_Actual) As Total_Saldos " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Fecha = #" & FechaIni & "# " _
       & "AND Fecha_Venc = #" & FechaFin & "# " _
       & "AND TP = 'CJBA' " _
       & "GROUP BY ME,TC "
  Select_Adodc_Grid DGCta, AdoCtas, sSQL
  RatonNormal
End Sub

Public Sub Procesar_Flujo_de_Caja()
    TotPresupuesto = 0
    sSQL = "SELECT Cta,SUM(Presupuesto) AS Presupuestos " _
           & "FROM Trans_Presupuestos " _
           & "WHERE Mes_No BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
           & "AND Cta = '" & Codigo & "' " _
           & "AND Codigo = '" & Ninguno & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "GROUP BY Cta "
    Select_Adodc AdoCtas, sSQL, False
    If AdoCtas.Recordset.RecordCount > 0 Then TotPresupuesto = AdoCtas.Recordset.Fields("Presupuestos")
    sSQL = "SELECT * " _
         & "FROM Transacciones " _
         & "WHERE Fecha <= #" & FechaFin & "# " _
         & "AND Cta = '" & Codigo & "' " _
         & "AND T = '" & Normal & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "ORDER BY Cta,Fecha,TP,Numero,Debe DESC,Haber,ID "
    Select_Adodc AdoCtas, sSQL, False
    If AdoCtas.Recordset.RecordCount > 0 Then
       AdoCtas.Recordset.MoveLast
       Saldo = Round(AdoCtas.Recordset.Fields("Saldo"), 2)
       If Moneda_US Then Saldo = Round(AdoCtas.Recordset.Fields("Saldo_ME"), 2)
    End If
    Debe = 0: Haber = 0
    sSQL = "SELECT Cta,SUM(Debe) As Debe1,SUM(Haber) As Haber1,SUM(Parcial_ME) As Parcial_ME1 " _
         & "FROM Transacciones " _
         & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Cta = '" & Codigo & "' " _
         & "AND T = '" & Normal & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "GROUP BY Cta "
    Select_Adodc AdoCtas, sSQL
    If AdoCtas.Recordset.RecordCount > 0 Then
       Debe = Round(AdoCtas.Recordset.Fields("Debe1"), 2)
       Haber = Round(AdoCtas.Recordset.Fields("Haber1"), 2)
    End If
  If Saldo <> 0 Then
     TotDiferencia = 0
     If TotPresupuesto <> 0 Then TotDiferencia = TotPresupuesto - Saldo
     SetAdoAddNew "Saldo_Diarios"
     SetAdoFields "Fecha", MBFechaI.Text
     SetAdoFields "Fecha_Venc", MBFechaF.Text
     SetAdoFields "TC", TipoProc
     SetAdoFields "ME", 0
     SetAdoFields "Cta", Codigo
     SetAdoFields "TP", "CJBA"
     SetAdoFields "Saldo_Anterior", Saldo + Haber - Debe
     SetAdoFields "Ingresos", Debe
     SetAdoFields "Egresos", Haber
     SetAdoFields "Saldo_Actual", Saldo
     SetAdoFields "Presupuesto", TotPresupuesto
     SetAdoFields "Diferencia", TotDiferencia
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoUpdate
  End If
End Sub

Public Sub Procesar_Flujo_de_Efectivo()
Dim Fecha_Ini1 As String
Dim Fecha_Fin1 As String
Dim Reprocesar As Boolean
Dim TipoProy(5) As TipoProyectar
    Reprocesar = False
     Mensajes = "Reprocesar Saldos Anteriores" & vbCrLf
     Titulo = "Pregunta de grabación"
     If BoxMensaje = vbYes Then Reprocesar = True

    NoMes = Month(MBFechaI) - 1
    NoAnio = Year(MBFechaI)
    If NoMes < 1 Then
       NoMes = 12
       NoAnio = NoAnio - 1
    End If
    Fecha_Ini1 = "01/" & Format(NoMes, "00") & "/" & Format(NoAnio, "0000")
    Fecha_Fin1 = UltimoDiaMes(Fecha_Ini1)
    Fecha_Ini1 = BuscarFecha(Fecha_Ini1)
    Fecha_Fin1 = BuscarFecha(Fecha_Fin1)

   Codigo1 = SaldoBancos.Caption
   DGCta.Visible = False
   Contador = 0
   If Reprocesar Then
     sSQL = "UPDATE Catalogo_Cuentas " _
          & "SET Proyectado=0,Procesado=0,Diferencia=0 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
   Else
     sSQL = "UPDATE Catalogo_Cuentas " _
          & "SET Procesado=0,Diferencia=0 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
   End If
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Catalogo_Cuentas " _
        & "SET Listar = " & Val(adFalse) & " " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND DG = 'G' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "DELETE * " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Codigo IN ('A','B','C','D','E') "
   Ejecutar_SQL_SP sSQL

   SetAdoAddNew "Catalogo_Cuentas"
   SetAdoFields "Codigo", "A"
   SetAdoFields "TC", "N"
   SetAdoFields "DG", "G"
   SetAdoFields "Listar", True
   SetAdoFields "Cuenta", "SALDO CAJA/BANCOS"
   SetAdoUpdate
  
   SetAdoAddNew "Catalogo_Cuentas"
   SetAdoFields "Codigo", "B"
   SetAdoFields "TC", "N"
   SetAdoFields "DG", "G"
   SetAdoFields "Listar", True
   SetAdoFields "Cuenta", "TOTAL INGRESOS"
   SetAdoUpdate
  
   SetAdoAddNew "Catalogo_Cuentas"
   SetAdoFields "Codigo", "C"
   SetAdoFields "TC", "N"
   SetAdoFields "DG", "G"
   SetAdoFields "Listar", True
   SetAdoFields "Cuenta", "TOTAL CAJA/BANCOS + INGRESOS"
   SetAdoUpdate
  
   SetAdoAddNew "Catalogo_Cuentas"
   SetAdoFields "Codigo", "D"
   SetAdoFields "TC", "N"
   SetAdoFields "DG", "G"
   SetAdoFields "Listar", True
   SetAdoFields "Cuenta", "TOTAL COSTOS,GASTOS E INVERSIONES"
   SetAdoUpdate
   
   SetAdoAddNew "Catalogo_Cuentas"
   SetAdoFields "Codigo", "E"
   SetAdoFields "TC", "N"
   SetAdoFields "DG", "G"
   SetAdoFields "Listar", True
   SetAdoFields "Cuenta", "S A L D O S"
   SetAdoUpdate
   
   For I = 0 To 5
      TipoProy(I).Diferencia = 0
      TipoProy(I).Procesado = 0
      TipoProy(I).Proyectado = 0
   Next I
   sSQL = "SELECT * " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE Listar <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND DG = 'D' " _
        & "ORDER BY Codigo "
   Select_Adodc AdoCtas, sSQL
   With AdoCtas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Cta = .Fields("Codigo")
           Codigo = .Fields("Codigo")
           TipoDoc = .Fields("TC")
           TipoCta = Normal
           Contador = Contador + 1
           SaldoBancos.Caption = Format(Contador / .RecordCount, "00%") & ", Cta: " & Cta & ", SubModulo: " & Codigo & "..."
           SaldoActual = 0
           SaldoAnterior = 0
           sSQL = "SELECT Cta,SUM(Debe) As TDebe,SUM(Haber) As THaber " _
                & "FROM Transacciones " _
                & "WHERE Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
                & "AND T <> '" & Anulado & "' " _
                & "AND Cta = '" & Cta & "' " _
                & "AND TP IN ('CD','CE','CI','ND','NC') " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           sSQL = sSQL & "GROUP BY Cta "
           Select_Adodc AdoAux, sSQL
           'MsgBox sSQL
           If AdoAux.Recordset.RecordCount > 0 Then
              Select Case MidStrg(Cta, 1, 1)
                Case "1": SaldoActual = AdoAux.Recordset.Fields("THaber")
                Case "2": SaldoActual = AdoAux.Recordset.Fields("TDebe")
                Case "4": SaldoActual = AdoAux.Recordset.Fields("THaber")
                Case "5": SaldoActual = AdoAux.Recordset.Fields("TDebe")
              End Select
           End If
           If TipoDoc = "BA" Then
              sSQL = "SELECT Cta,SUM(Debe) As TDebe,SUM(Haber) As THaber " _
                   & "FROM Transacciones " _
                   & "WHERE Fecha < #" & FechaIni & "# " _
                   & "AND T <> '" & Anulado & "' " _
                   & "AND Cta = '" & Cta & "' " _
                   & "AND TP IN ('CD','CE','CI','ND','NC') " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              sSQL = sSQL & "GROUP BY Cta "
              Select_Adodc AdoAux, sSQL
              
              If AdoAux.Recordset.RecordCount > 0 Then
                 Select Case MidStrg(Cta, 1, 1)
                   Case "1": SaldoAnterior = AdoAux.Recordset.Fields("TDebe") - AdoAux.Recordset.Fields("THaber")
                 End Select
              End If
           Else
              sSQL = "SELECT Cta,SUM(Debe) As TDebe,SUM(Haber) As THaber " _
                   & "FROM Transacciones " _
                   & "WHERE Fecha Between #" & Fecha_Ini1 & "# and #" & Fecha_Fin1 & "# " _
                   & "AND T <> '" & Anulado & "' " _
                   & "AND Cta = '" & Cta & "' " _
                   & "AND TP IN ('CD','CE','CI','ND','NC') " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              sSQL = sSQL & "GROUP BY Cta "
              Select_Adodc AdoAux, sSQL
              If AdoAux.Recordset.RecordCount > 0 Then
                 Select Case MidStrg(Cta, 1, 1)
                   Case "1": SaldoAnterior = AdoAux.Recordset.Fields("THaber")
                   Case "2": SaldoAnterior = AdoAux.Recordset.Fields("TDebe")
                   Case "4": SaldoAnterior = AdoAux.Recordset.Fields("THaber")
                   Case "5": SaldoAnterior = AdoAux.Recordset.Fields("TDebe")
                 End Select
              End If
           End If
          If Not Reprocesar Then SaldoAnterior = .Fields("Proyectado")
          If TipoDoc = "BA" Then
             SaldoActual = 0
             Diferencia = 0
          End If
          'MsgBox "====>: " & sSQL
          Diferencia = SaldoActual - SaldoAnterior
          
          If (SaldoAnterior + SaldoActual) <> 0 Then
              Cta_Sup = CodigoCuentaSup(Cta)
              'MsgBox Cta & vbCrLf & Cta_Sup
                sSQL = "UPDATE Catalogo_Cuentas " _
                     & "SET Proyectado = " & SaldoAnterior _
                     & ",Procesado = " & SaldoActual _
                     & ",Diferencia = " & Diferencia & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo = '" & Cta & "' "
                Ejecutar_SQL_SP sSQL
                sSQL = "UPDATE Catalogo_Cuentas " _
                     & "SET Listar = " & Val(adTrue) & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo = '" & Cta_Sup & "' "
                Ejecutar_SQL_SP sSQL
          End If
          Select Case MidStrg(Cta, 1, 1)
            Case "1"
                     If TipoDoc = "BA" Then
                        TipoProy(0).Proyectado = TipoProy(0).Proyectado + SaldoAnterior
                        TipoProy(0).Procesado = TipoProy(0).Procesado + SaldoActual
                        TipoProy(0).Diferencia = TipoProy(0).Diferencia + Diferencia
                        
                        TipoProy(2).Proyectado = TipoProy(2).Proyectado + SaldoAnterior
                        TipoProy(2).Procesado = TipoProy(2).Procesado + SaldoActual
                        TipoProy(2).Diferencia = TipoProy(2).Diferencia + Diferencia
                            
                        TipoProy(4).Proyectado = TipoProy(4).Proyectado + SaldoAnterior
                        TipoProy(4).Procesado = TipoProy(4).Procesado + SaldoActual
                        TipoProy(4).Diferencia = TipoProy(4).Diferencia + Diferencia
                     End If
            Case "4"
                     TipoProy(1).Proyectado = TipoProy(1).Proyectado + SaldoAnterior
                     TipoProy(1).Procesado = TipoProy(1).Procesado + SaldoActual
                     TipoProy(1).Diferencia = TipoProy(1).Diferencia + Diferencia
                     
                     TipoProy(2).Proyectado = TipoProy(2).Proyectado + SaldoAnterior
                     TipoProy(2).Procesado = TipoProy(2).Procesado + SaldoActual
                     TipoProy(2).Diferencia = TipoProy(2).Diferencia + Diferencia
                     
                     TipoProy(4).Proyectado = TipoProy(4).Proyectado + SaldoAnterior
                     TipoProy(4).Procesado = TipoProy(4).Procesado + SaldoActual
                     TipoProy(4).Diferencia = TipoProy(4).Diferencia + Diferencia
            Case "5"
                     TipoProy(3).Proyectado = TipoProy(3).Proyectado + SaldoAnterior
                     TipoProy(3).Procesado = TipoProy(3).Procesado + SaldoActual
                     TipoProy(3).Diferencia = TipoProy(3).Diferencia + Diferencia
                     
                     TipoProy(4).Proyectado = TipoProy(4).Proyectado + SaldoAnterior
                     TipoProy(4).Procesado = TipoProy(4).Procesado + SaldoActual
                     TipoProy(4).Diferencia = TipoProy(4).Diferencia + Diferencia
          End Select
         .MoveNext
        Loop
    End If
   End With
  'Sacamos Totales
   For I = 0 To 4
       sSQL = "UPDATE Catalogo_Cuentas " _
            & "SET Proyectado = " & TipoProy(I).Proyectado _
            & ",Procesado = " & TipoProy(I).Procesado _
            & ",Diferencia = " & TipoProy(I).Diferencia & " " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND TC = 'N' " _
            & "AND DG = 'G' " _
            & "AND Listar <> " & Val(adFalse) & " "
        Select Case I
          Case 0: sSQL = sSQL & "AND Codigo = 'A' "
          Case 1: sSQL = sSQL & "AND Codigo = 'B' "
          Case 2: sSQL = sSQL & "AND Codigo = 'C' "
          Case 3: sSQL = sSQL & "AND Codigo = 'D' "
          Case 4: sSQL = sSQL & "AND Codigo = 'E' "
        End Select
        Ejecutar_SQL_SP sSQL
   Next I
   sSQL = "SELECT Codigo,Cuenta,Proyectado,Procesado,Diferencia,TC,Item,Periodo " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Listar <> " & Val(adFalse) & " " _
        & "ORDER BY Codigo "
   Select_Adodc_Grid DGBanco, AdoBanco, sSQL
   DGCta.Visible = True
   DGBanco.Visible = True
   SaldoBancos.Caption = Codigo1
End Sub
