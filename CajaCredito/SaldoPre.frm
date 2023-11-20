VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form SaldoPrestamo 
   Caption         =   "SALDO DE CAJA BANCOS"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
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
      Height          =   750
      Left            =   10395
      Picture         =   "SaldoPre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   525
      Width           =   855
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
      Height          =   750
      Left            =   9450
      Picture         =   "SaldoPre.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   525
      Width           =   855
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   840
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1275
      Left            =   105
      TabIndex        =   8
      Top             =   105
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   2249
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Saldos Diarios"
      TabPicture(0)   =   "SaldoPre.frx":0BD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CheqPrest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DCTipo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tabla Temporizada"
      TabPicture(1)   =   "SaldoPre.frx":0BF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(1)=   "Command5"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command6 
         Caption         =   "&Calcular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   -67545
         Picture         =   "SaldoPre.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Listar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   -66600
         Picture         =   "SaldoPre.frx":104E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   420
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DCTipo 
         Bindings        =   "SaldoPre.frx":1358
         DataSource      =   "AdoDetCheq"
         Height          =   315
         Left            =   1470
         TabIndex        =   10
         Top             =   735
         Width           =   5895
         _ExtentX        =   10398
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
         Height          =   750
         Left            =   7455
         Picture         =   "SaldoPre.frx":1371
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   435
         Width           =   855
      End
      Begin VB.CheckBox CheqPrest 
         Caption         =   "&Tipo Prestamo"
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
         Left            =   1470
         TabIndex        =   11
         Top             =   420
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8400
         Picture         =   "SaldoPre.frx":17B3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   -74895
         TabIndex        =   14
         Top             =   420
         Width           =   1275
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
         TabIndex        =   13
         Top             =   420
         Width           =   1275
      End
   End
   Begin MSDataGridLib.DataGrid DGBanco 
      Bindings        =   "SaldoPre.frx":1ABD
      Height          =   2115
      Left            =   105
      TabIndex        =   3
      Top             =   1470
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
      Caption         =   "SALDO DE PRESTAMOS"
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   210
      Top             =   1575
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
      Top             =   3570
      Width           =   11145
      _ExtentX        =   19659
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
   Begin MSAdodcLib.Adodc AdoDetCheq 
      Height          =   330
      Left            =   210
      Top             =   2205
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
      Caption         =   "DetCheq"
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
   Begin MSAdodcLib.Adodc AdoDetCheqPosf 
      Height          =   330
      Left            =   210
      Top             =   1890
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
      Caption         =   "DetCheqPosf"
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   105
      Top             =   7140
      Width           =   11145
      _ExtentX        =   19659
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
      Caption         =   "Detalle"
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
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "SaldoPre.frx":1AD4
      Height          =   2745
      Left            =   105
      TabIndex        =   17
      Top             =   4410
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   4842
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
      Caption         =   "DETALLE DE SALDO DE PRESTAMOS"
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
   Begin VB.Label LblCapital 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5775
      TabIndex        =   6
      Top             =   3990
      Width           =   1800
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAPITAL"
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
      Left            =   4410
      TabIndex        =   7
      Top             =   3990
      Width           =   1380
   End
   Begin VB.Label LblPendiente 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9555
      TabIndex        =   4
      Top             =   3990
      Width           =   1800
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO PENDIENTE"
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
      Left            =   7665
      TabIndex        =   5
      Top             =   3990
      Width           =   1905
   End
End
Attribute VB_Name = "SaldoPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqPrest_Click()
  If CheqPrest.value = 1 Then DCTipo.Visible = True Else DCTipo.Visible = False
End Sub

Private Sub Command1_Click()
  Unload SaldoPrestamo
End Sub

Private Sub Command2_Click()
  DGBanco.Visible = False
  SQLMsg3 = ""
  SQLMsg1 = DGBanco.Caption
  SQLMsg2 = "Al " & MBoxFechaF
  ImprimirAdo AdoBanco, True, 1, 8
  ImprimirAdo AdoDetalle, True, 1, 8
  DGBanco.Visible = True
End Sub

Private Sub Command3_Click()
  RatonReloj
  If CheqPrest.value = 1 Then
    DGBanco.Caption = UCase(DCTipo.Text)
  Else
    DGBanco.Caption = "SALDO DE PRESTAMOS"
  End If
  Contador = 0
  FechaValida MBoxFechaF
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  DGBanco.Visible = False
  Codigo = SinEspaciosIzq(DCTipo.Text)
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectAdodc AdoBanco, sSQL
  
  sSQL = "SELECT Fecha,TP,Tasa,Credito_No,Cuenta_No " _
       & "FROM Prestamos " _
       & "WHERE Fecha <= #" & FechaFin & "# " _
       & "AND T <> 'N' " _
       & "AND Item = '" & NumEmpresa & "' "
  If CheqPrest.value = 1 Then sSQL = sSQL & "AND TP = '" & Codigo & "' "
  sSQL = sSQL & "ORDER BY TP,Cuenta_No,Credito_No "
  'MsgBox "TP = " & vbCrLf & sSQL
  SelectAdodc AdoDetCheqPosf, sSQL
  With AdoDetCheqPosf.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          RatonReloj
          Contador = Contador + 1
          Debe = .Fields("Tasa")
          Mifecha = .Fields("Fecha")
          TipoProc = .Fields("TP")
          Credito_No = .Fields("Credito_No")
          Cuenta_No = .Fields("Cuenta_No")
          Saldo = 0: Haber = 0
          CodigoCli = Ninguno
          SaldoPrestamo.Caption = TipoProc & ": " & Contador & "/" & .RecordCount & ": " & Cuenta_No & " -> " & Contrato_No
          sSQL = "SELECT Codigo " _
               & "FROM Clientes_Datos_Extras " _
               & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
               & "AND Tipo_Dato = 'LIBRETAS' "
          SelectAdodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount > 0 Then
             CodigoCli = AdoCtas.Recordset.Fields("Codigo")
          End If
          sSQL = "SELECT TP,Credito_No,SUM(Capital) As TotCap " _
               & "FROM Trans_Prestamos " _
               & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
               & "AND Credito_No = '" & Credito_No & "' " _
               & "AND TP = '" & TipoProc & "' " _
               & "GROUP BY TP,Credito_No "
          SelectAdodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount > 0 Then
             Total = AdoCtas.Recordset.Fields("TotCap")
          End If
          Saldo = 0
          sSQL = "SELECT TP,Credito_No,SUM(Capital) As TotCap " _
               & "FROM Trans_Prestamos " _
               & "WHERE Fecha_C <= #" & FechaFin & "# " _
               & "AND T = 'C' " _
               & "AND Cuenta_No = '" & Cuenta_No & "' " _
               & "AND Credito_No = '" & Credito_No & "' " _
               & "AND TP = '" & TipoProc & "' " _
               & "GROUP BY TP,Credito_No "
          'MsgBox sSQL
          SelectAdodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount > 0 Then
             Do While Not AdoCtas.Recordset.EOF
                Saldo = Saldo + AdoCtas.Recordset.Fields("TotCap")
                AdoCtas.Recordset.MoveNext
             Loop
          End If
          'If Cuenta_No = "00100231-0" Then MsgBox Total
          If (Total - Saldo) <> 0 Then
             AdoBanco.Recordset.AddNew
             AdoBanco.Recordset.Fields("TP") = TipoProc
             AdoBanco.Recordset.Fields("Fecha") = Mifecha
             AdoBanco.Recordset.Fields("Cta") = Credito_No
             AdoBanco.Recordset.Fields("Grupo_No") = Cuenta_No
             AdoBanco.Recordset.Fields("CodigoC") = CodigoCli
             AdoBanco.Recordset.Fields("Numero") = 0
             AdoBanco.Recordset.Fields("Ingresos") = Debe
             AdoBanco.Recordset.Fields("Egresos") = Total
             AdoBanco.Recordset.Fields("Total") = Total - Saldo
             AdoBanco.Recordset.Fields("Item") = NumEmpresa
             AdoBanco.Recordset.Fields("CodigoU") = CodigoUsuario
             AdoBanco.Recordset.Update
          End If
          RatonNormal
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Cliente,CR.Fecha,CR.TP,CT.Cuenta_No,CR.Cta As Credito,Ingresos As Tasa,Egresos As Creditos,Total As Saldo_Pend " _
       & "FROM Saldo_Diarios As CR,Clientes As C,Clientes_Datos_Extras As CT " _
       & "WHERE CR.CodigoC = C.Codigo " _
       & "AND C.Codigo = Ct.Codigo " _
       & "AND CR.Grupo_No = CT.Cuenta_No " _
       & "AND CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY CR.Fecha,CR.TP,CR.Cta,Cliente "
  SelectDataGrid DGBanco, AdoBanco, sSQL
  Total = 0: Saldo = 0
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Creditos")
          Saldo = Saldo + .Fields("Saldo_Pend")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LblCapital.Caption = Format(Total, "#,##0.00")
  LblPendiente.Caption = Format(Saldo, "#,##0.00")
  DGBanco.Visible = True
  RatonNormal
End Sub

Private Sub Command4_Click()
  sSQL = "SELECT Cliente,CR.Fecha,TP,Ct.Cuenta_No,Cuenta As Credito,Debitos As Tasa,Creditos,Valor_ME As Saldo_Pend " _
       & "FROM Saldo_Diarios As CR,Clientes As C,Clientes_Datos_Extras As Ct " _
       & "WHERE CR.Beneficiario = C.Codigo " _
       & "AND C.Codigo = Ct.Codigo " _
       & "AND CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY CR.Fecha,TP,Cuenta,Cliente "
  SelectDataGrid DGBanco, AdoBanco, sSQL
  Total = 0: Saldo = 0
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Creditos")
          Saldo = Saldo + .Fields("Saldo_Pend")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LblCapital.Caption = Format(Total, "#,##0.00")
  LblPendiente.Caption = Format(Saldo, "#,##0.00")
  DGBanco.Visible = True
End Sub

Private Sub Command5_Click()
  FechaValida MBoxFechaF
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  DGBanco.Visible = True
'''  sSQL = "SELECT Cta,TP, " _
'''       & "SUM(De1a30) As SDe1a30," _
'''       & "SUM(De31a90) As SDe31a90," _
'''       & "SUM(De91a180) As SDe91a180," _
'''       & "SUM(De181a360) As SDe181a360," _
'''       & "SUM(MasDe360) As SMasDe360," _
'''       & "SUM(Total) As STotal " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Fecha = #" & FechaFin & "# " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "GROUP BY Cta,TP "
'''  SelectDataGrid DGBanco, AdoBanco, sSQL
'''  Sumatoria = 0
'''  With AdoBanco.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Sumatoria = Sumatoria + .Fields("STotal")
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  LblCapital.Caption = Format(Sumatoria, "#,##0.00")
  FechaStr = "31/12/" & Format(Year(FechaSistema), "0000")
  sSQL = "SELECT T,TP,Credito_No,Cuenta_No " _
       & "FROM Prestamos " _
       & "WHERE Fecha_C = #" & BuscarFecha(FechaStr) & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP,Credito_No "
  SelectDataGrid DGBanco, AdoBanco, sSQL

End Sub

Private Sub Command6_Click()
'Totales de Vencidos
Dim VenDe1a30 As Currency
Dim VenDe31a90 As Currency
Dim VenDe91a180 As Currency
Dim VenDe181a360 As Currency
Dim VenMasDe360 As Currency
'Totales de Vigentes
Dim VigDe1a30 As Currency
Dim VigDe31a90 As Currency
Dim VigDe91a180 As Currency
Dim VigDe181a360 As Currency
Dim VigMasDe360 As Currency
  FechaStr = "31/12/" & Format(Year(FechaSistema), "0000")
  FechaValida MBoxFechaF
  Mifecha = MBoxFechaF.Text
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  FechaN = CFechaLong(Mifecha)
'''  sSQL = "DELETE * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Fecha = #" & FechaFin & "# " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  ConectarAdoExecute sSQL
'''  sSQL = "SELECT * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Fecha = #" & FechaFin & "# " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  SelectAdodc AdoBanco, sSQL
  DGBanco.Visible = False
  sSQL = "UPDATE Prestamos " _
       & "SET No_Venc=0," _
       & "Fecha_C=#" & BuscarFecha(FechaStr) & "# " _
       & "WHERE Item <> '.' "
  ConectarAdoExecute sSQL
' Verificamos Vencidos
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP,Credito_No,Fecha "
  SelectAdodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Credito_No = .Fields("Credito_No")
       TipoDoc = .Fields("TP")
       NoMeses = 0
       Do While Not .EOF
          Contador = Contador + 1
          If Credito_No <> .Fields("Credito_No") _
             Or TipoDoc <> .Fields("TP") Then
             SaldoPrestamo.Caption = Contador & "/" & .RecordCount & " => " & TipoDoc & " = " & Credito_No
             sSQL = "UPDATE Prestamos " _
                  & "SET No_Venc = " & NoMeses & "," _
                  & "Fecha_C = #" & FechaTexto & "# " _
                  & "WHERE Credito_No = '" & Credito_No & "' " _
                  & "AND TP = '" & TipoDoc & "' " _
                  & "AND Item = '" & NumEmpresa & "' "
             ConectarAdoExecute sSQL
             Credito_No = .Fields("Credito_No")
             TipoDoc = .Fields("TP")
             NoMeses = 0
          End If
          FechaIniN = CFechaLong(.Fields("Fecha"))
          FechaFinN = CFechaLong(.Fields("Fecha_C"))
          FechaTexto = BuscarFecha(.Fields("Fecha_C"))
         'Preguntamos Fecha de Pago con Canc
          If (FechaIniN <> FechaFinN) And (FechaIniN < CFechaLong(Mifecha)) Then
             NoMeses = NoMeses + 1
          End If
         .MoveNext
       Loop
   End If
  End With
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP,Credito_No,Fecha "
  SelectAdodc AdoBanco, sSQL
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          T = .Fields("T")
          TipoDoc = .Fields("TP")
          Credito_No = .Fields("Credito_No")
          Cuenta_No = .Fields("Cuenta_No")
          SaldoPrestamo.Caption = Contador & "/" & .RecordCount & " => " & TipoDoc & " = " & Credito_No
          If T <> "N" Then
          sSQL = "SELECT * " _
               & "FROM Trans_Prestamos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND TP = '" & TipoDoc & "' " _
               & "AND Credito_No = '" & Credito_No & "' "
          SelectAdodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount <= 0 Then
             SetAddNew AdoBanco
             SetFields AdoBanco, "T", T
             SetFields AdoBanco, "Fecha", FechaStr
             SetFields AdoBanco, "Fecha_C", FechaStr
             SetFields AdoBanco, "TP", TipoDoc
             SetFields AdoBanco, "Encaje", 0
             SetFields AdoBanco, "Credito_No", Credito_No
             SetFields AdoBanco, "Cuenta_No", Cuenta_No
             SetFields AdoBanco, "Tasa", 0
             SetFields AdoBanco, "Meses", 0
             SetFields AdoBanco, "Patrimonio", -1
             SetFields AdoBanco, "No_Venc", 0
             SetFields AdoBanco, "Item", NumEmpresa
             SetUpdate AdoBanco
          End If
          End If
         .MoveNext
       Loop
   End If
  End With
' Empezamos a procesar tabla temporizada
  Contador = 0
  sSQL = "SELECT P.TP,P.Credito_No,P.Cuenta_No,P.No_Venc,P.Fecha_C,C.Codigo " _
       & "FROM Prestamos As P,Clientes_Datos_Extras As Ct,Clientes As C " _
       & "WHERE P.Cuenta_No = Ct.Cuenta_No " _
       & "AND Ct.Codigo = C.Codigo " _
       & "AND P.Item = '" & NumEmpresa & "' " _
       & "ORDER BY P.TP,P.Credito_No,P.Fecha "
  SelectAdodc AdoDetCheq, sSQL
  With AdoDetCheq.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TipoDoc = .Fields("TP")
          Codigo = .Fields("Codigo")
          Numero = .Fields("No_Venc")
          FechaTexto = .Fields("Fecha_C")
          Credito_No = .Fields("Credito_No")
          VigDe1a30 = 0: VigDe31a90 = 0: VigDe91a180 = 0
          VigDe181a360 = 0: VigMasDe360 = 0
          VenDe1a30 = 0: VenDe31a90 = 0: VenDe91a180 = 0
          VenDe181a360 = 0: VenMasDe360 = 0
          Contador = Contador + 1
          SaldoPrestamo.Caption = Contador & "/" & .RecordCount & " => " & TipoDoc & " - " & Credito_No
          sSQL = "SELECT * " _
               & "FROM Trans_Prestamos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND TP = '" & TipoDoc & "' " _
               & "AND Credito_No = '" & Credito_No & "' " _
               & "ORDER BY TP,Credito_No,Fecha "
          SelectAdodc AdoCtas, sSQL
          If AdoCtas.Recordset.RecordCount > 0 Then
             Do While Not AdoCtas.Recordset.EOF
                I = CFechaLong(AdoCtas.Recordset.Fields("Fecha"))
                J = CFechaLong(AdoCtas.Recordset.Fields("Fecha_C"))
                NoDias = FechaN - I
                If (I < J) And (FechaN < I) Then
                  'Vencidos
                   If FechaN >= NumDias Then
                      Select Case NoDias
                        Case 0 To 30: VenDe1a30 = VenDe1a30 + AdoCtas.Recordset.Fields("Capital")
                        Case 31 To 90: VenDe31a90 = VenDe31a90 + AdoCtas.Recordset.Fields("Capital")
                        Case 91 To 180: VenDe91a180 = VenDe91a180 + AdoCtas.Recordset.Fields("Capital")
                        Case 181 To 360: VenDe181a360 = VenDe181a360 + AdoCtas.Recordset.Fields("Capital")
                        Case Is > 360: VenMasDe360 = VenMasDe360 + AdoCtas.Recordset.Fields("Capital")
                      End Select
                   End If
                Else
                  'Vigentes
                   If FechaN >= NumDias Then
                      Select Case NoDias
                        Case 0 To 30: VigDe1a30 = VigDe1a30 + AdoCtas.Recordset.Fields("Capital")
                        Case 31 To 90: VigDe31a90 = VigDe31a90 + AdoCtas.Recordset.Fields("Capital")
                        Case 91 To 180: VigDe91a180 = VigDe91a180 + AdoCtas.Recordset.Fields("Capital")
                        Case 181 To 360: VigDe181a360 = VigDe181a360 + AdoCtas.Recordset.Fields("Capital")
                        Case Is > 360: VigMasDe360 = VigMasDe360 + AdoCtas.Recordset.Fields("Capital")
                      End Select
                End If
                End If
                AdoCtas.Recordset.MoveNext
             Loop
          End If
          Cta = "142"
          If Numero > 1 Then Cta = "141"
          If Numero > 2 And TipoDoc = "C/C" Then Cta = "141"
          Select Case TipoDoc
            Case "C/C", "P/E", "P/H", "S/C": Cta = Cta & "3"
            Case "P/A", "P/C": Cta = Cta & "2"
            Case "S/F": Cta = Cta & "1"
          End Select

          Total = De0 + De1a30 + De31a90 + De91a180 + De181a360 + MasDe360
             SaldoPrestamo.Caption = "Vigentes: " & Contador & "/" & .RecordCount & " => " & TipoDoc & " - " & Credito_No
             Cta = "140"
             Select Case TipoDoc
               Case "C/C", "P/E", "P/H", "S/C": Cta = Cta & "3"
               Case "P/A", "P/C": Cta = Cta & "2"
               Case "S/F": Cta = Cta & "1"
             End Select
             IngSaldoDiarios

          IngSaldoDiarios
         .MoveNext
       Loop
   End If
  End With
  DGBanco.Visible = True
  sSQL = "SELECT Cta,TP, " _
       & "SUM(De1a30) As SDe1a30," _
       & "SUM(De31a90) As SDe31a90," _
       & "SUM(De91a180) As SDe91a180," _
       & "SUM(De181a360) As SDe181a360," _
       & "SUM(MasDe360) As SMasDe360," _
       & "SUM(Total) As STotal " _
       & "FROM Saldo_Diarios " _
       & "WHERE Fecha = #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Cta,TP "
  sSQL = "SELECT T,Credito_No,Cuenta_No " _
       & "FROM Prestamos " _
       & "WHERE Fecha_C = #" & BuscarFecha(FechaStr) & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP,Credito_No "
  SelectDataGrid DGBanco, AdoBanco, sSQL
End Sub

Private Sub DGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGBanco.Visible = False
     GenerarDataTexto SaldoPrestamo, AdoBanco
     DGBanco.Visible = True
  End If
  
  If CtrlDown And KeyCode = vbKeyB Then
     BuscarDatos DGBanco, AdoBanco
  End If
  If CtrlDown And KeyCode = vbKeyL Then
     FechaValida MBoxFechaF
     Cta = DGBanco.Columns(0).Text
     TipoDoc = DGBanco.Columns(1).Text
     FechaFin = BuscarFecha(MBoxFechaF.Text)
     sSQL = "SELECT SD.Cta,SD.TP,C.Cliente,SD.Credito_No," _
          & "SUM(De1a30) As SDe1a30," _
          & "SUM(De31a90) As SDe31a90," _
          & "SUM(De91a180) As SDe91a180," _
          & "SUM(De181a360) As SDe181a360," _
          & "SUM(MasDe360) As SMasDe360," _
          & "SUM(Total) As STotal " _
          & "FROM Saldo_Diarios As SD,Clientes As C " _
          & "WHERE SD.Fecha = #" & FechaFin & "# " _
          & "AND SD.TP = '" & TipoDoc & "' " _
          & "AND SD.Cta = '" & Cta & "' " _
          & "AND SD.Item = '" & NumEmpresa & "' " _
          & "AND SD.CodigoC = C.Codigo " _
          & "GROUP BY SD.Cta,SD.TP,C.Cliente,SD.Credito_No "
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     DGBanco.Visible = False
     SQLMsg3 = ""
     SQLMsg1 = DGBanco.Caption
     SQLMsg2 = "Al " & MBoxFechaF.Text
     ImprimirAdodc AdoBanco, 1, 2, 8
     DGBanco.Visible = True
  End If
End Sub

Private Sub DGDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGDetalle.Visible = False
     GenerarDataTexto SaldoPrestamo, AdoDetalle
     DGDetalle.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyB Then
     BuscarDatos DGDetalle, AdoDetalle
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     DGDetalle.Visible = False
     SQLMsg3 = ""
     SQLMsg1 = DGDetalle.Caption
     SQLMsg2 = "Al " & MBoxFechaF.Text
     ImprimirAdodc AdoDetalle, 1, 2, 8
     DGDetalle.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT (CTP & '   ' & Descripcion) As TipoPrest " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE TC <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY CTP,Descripcion "
  SelectDBCombo DCTipo, AdoDetCheq, sSQL, "TipoPrest"
  RatonNormal
End Sub

Private Sub Form_Load()
  'CentrarForm SaldoPrestamo
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoDetCheq
  ConectarAdodc AdoDetCheqPosf
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Public Sub IngSaldoDiarios()
  If Total <> 0 Then
     SetAddNew AdoBanco
     SetFields AdoBanco, "Fecha", Mifecha
     SetFields AdoBanco, "TP", TipoDoc
     SetFields AdoBanco, "CodigoC", Codigo
     SetFields AdoBanco, "Credito_No", Credito_No
     SetFields AdoBanco, "De1a30", De1a30
     SetFields AdoBanco, "De31a90", De31a90
     SetFields AdoBanco, "De91a180", De91a180
     SetFields AdoBanco, "De181a360", De181a360
     SetFields AdoBanco, "MasDe360", MasDe360
     SetFields AdoBanco, "Total", Total
     SetFields AdoBanco, "Item", NumEmpresa
     SetFields AdoBanco, "Cta", Cta
     SetUpdate AdoBanco
  End If
End Sub
