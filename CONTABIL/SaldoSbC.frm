VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BalanceSubCtas 
   Caption         =   "SALDO DE CAJA BANCOS"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10290
      Top             =   6615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SaldoSbC.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SaldoSbC.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SaldoSbC.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SaldoSbC.frx":1046
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CheqNivel 
      Caption         =   "Por Cuenta Contable"
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
      Left            =   4305
      TabIndex        =   10
      Top             =   735
      Width           =   2115
   End
   Begin MSDataGridLib.DataGrid DGBanco 
      Bindings        =   "SaldoSbC.frx":1920
      Height          =   5370
      Left            =   105
      TabIndex        =   13
      Top             =   1155
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   9472
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
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
      Left            =   11235
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   5790
      _ExtentX        =   10213
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
      Height          =   645
      Left            =   2415
      TabIndex        =   1
      Top             =   0
      Width           =   7050
      Begin VB.OptionButton OpcC 
         Caption         =   "Ctas x Cobrar"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "Ctas x Pagar"
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
         Left            =   1785
         TabIndex        =   3
         Top             =   210
         Width           =   1485
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "Ctas Gastos"
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
         Left            =   3465
         TabIndex        =   4
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "Ctas de Ingreso"
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
         Left            =   5145
         TabIndex        =   5
         Top             =   210
         Width           =   1695
      End
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2940
      TabIndex        =   9
      Top             =   735
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
      TabIndex        =   7
      Top             =   735
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
   Begin MSAdodcLib.Adodc AdoDetCheq 
      Height          =   330
      Left            =   315
      Top             =   1680
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "SaldoSbC.frx":1937
      DataSource      =   "AdoCtas1"
      Height          =   315
      Left            =   6510
      TabIndex        =   11
      Top             =   735
      Visible         =   0   'False
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Ctas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   315
      Top             =   2310
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Procesar"
            Object.ToolTipText     =   "Procesar Submodulos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Listar"
            Object.ToolTipText     =   "Listar Submodulos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelTotHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   13860
      TabIndex        =   14
      Top             =   735
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Total"
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
      Left            =   12705
      TabIndex        =   15
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      Left            =   2205
      TabIndex        =   8
      Top             =   735
      Width           =   750
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      TabIndex        =   6
      Top             =   735
      Width           =   750
   End
End
Attribute VB_Name = "BalanceSubCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListarBalanceDeSubModulos()
  RatonReloj
  'MBFechaI.Text = FechaSistema
  'MBFechaF.Text = FechaSistema
  DGBanco.Visible = False
  TipoDoc = "T"
  If OpcC.value Then
     TipoDoc = "C"
     DGBanco.Caption = "BALANCE DE SUBMODULO CUENTAS POR COBRAR"
  End If
  If OpcP.value Then
     TipoDoc = "P"
     DGBanco.Caption = "BALANCE DE SUBMODULO CUENTAS POR PAGAR"
  End If
  If OpcI.value Then
     TipoDoc = "I"
     DGBanco.Caption = "BALANCE DE SUBMODULO DE INGRESOS"
  End If
  If OpcG.value Then
     TipoDoc = "G"
     DGBanco.Caption = "BALANCE DE SUBMODULO DE GASTOS"
  End If
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC = '" & TipoDoc & "' " _
       & "AND TP = 'SUBM' "
  Select_Adodc AdoBanco, sSQL
  If AdoBanco.Recordset.RecordCount > 0 Then
     MBFechaI.Text = AdoBanco.Recordset.fields("Fecha")
     MBFechaF.Text = AdoBanco.Recordset.fields("Fecha_Venc")
  End If
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  DGBanco.Caption = "BALANCE DE SUBMODULOS"
  If OpcC.value Or OpcP.value Then
     sSQL = "SELECT SD.TC,C.Cliente As Detalles,SD.Cta,SD.Comprobante As Cuentas,SD.Saldo_Anterior,SD.Ingresos,SD.Egresos,SD.Saldo_Actual " _
          & "FROM Saldo_Diarios As SD,Clientes As C " _
          & "WHERE SD.CodigoU = '" & CodigoUsuario & "' " _
          & "AND SD.Item = '" & NumEmpresa & "' " _
          & "AND SD.TC = '" & TipoDoc & "' " _
          & "AND SD.TP = 'SUBM' " _
          & "AND SD.CodigoC = C.Codigo "
     If TipoDoc <> "T" Then sSQL = sSQL & "AND SD.TC = '" & TipoDoc & "' "
     sSQL = sSQL & "ORDER BY SD.TC,C.Cliente,SD.Cta "
  End If
  If OpcI.value Or OpcG.value Then
     sSQL = "SELECT SD.TC,C.Detalle As Detalles,SD.Cta,SD.Comprobante As Cuentas,SD.Saldo_Anterior,SD.Ingresos,SD.Egresos,SD.Saldo_Actual " _
          & "FROM Saldo_Diarios As SD,Catalogo_SubCtas As C " _
          & "WHERE SD.CodigoU = '" & CodigoUsuario & "' " _
          & "AND C.Periodo = '" & Periodo_Contable & "' " _
          & "AND SD.Item = '" & NumEmpresa & "' " _
          & "AND SD.TC = '" & TipoDoc & "' " _
          & "AND SD.TP = 'SUBM' " _
          & "AND SD.Item = C.Item " _
          & "AND SD.CodigoC = C.Codigo "
     If TipoDoc <> "T" Then sSQL = sSQL & "AND SD.TC = '" & TipoDoc & "' "
     sSQL = sSQL & "ORDER BY SD.TC,C.Detalle,SD.Cta "
  End If
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  SaldoTotal = 0
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If CheqNivel.value = 1 Then
             SaldoTotal = SaldoTotal + .fields("Saldo_Actual")
          Else
             If .fields("Cta") = "I==============>" Then SaldoTotal = SaldoTotal + .fields("Saldo_Actual")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelTotHaber.Caption = Format(SaldoTotal, "#,##0.00")
  BalanceSubCtas.Caption = "BALANCE DE SUBMODULOS"
  DGBanco.Visible = True
  RatonNormal
End Sub

Private Sub CheqNivel_Click()
  If CheqNivel.value = 1 Then DCCtas.Visible = True Else DCCtas.Visible = False
End Sub

Private Sub Command1_Click()
  Unload BalanceSubCtas
End Sub

Private Sub Imprimir()
  DGBanco.Visible = False
  If OpcC.value Then
     SQLMsg1 = "SALDO DE CUENTAS POR COBRAR"
  ElseIf OpcP.value Then
     SQLMsg1 = "SALDO DE CUENTAS POR PAGAR"
  ElseIf OpcI.value Then
     SQLMsg1 = "SALDO DE CUENTAS DE INGRESO"
  ElseIf OpcG.value Then
     SQLMsg1 = "SALDO DE CUENTAS DE GASTOS"
  End If
  SQLMsg2 = "Desde: " & MBFechaI.Text & " al " & MBFechaF.Text
  SQLMsg3 = ""
  ImprimirBalanceSubCta AdoBanco, SaldoTotal, 1, 7.5
  DGBanco.Visible = True
End Sub

Private Sub Procesar()
  DGBanco.Visible = False
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Cta = SinEspaciosIzq(DCCtas)
  If Cta = "" Then Cta = Ninguno
  DGBanco.Caption = "BALANCE DE SUBMODULOS"
  TipoDoc = "T"
  If OpcC.value Then
     TipoDoc = "C"
     DGBanco.Caption = "BALANCE DE SUBMODULO CUENTAS POR COBRAR"
  End If
  If OpcP.value Then
     TipoDoc = "P"
     DGBanco.Caption = "BALANCE DE SUBMODULO CUENTAS POR PAGAR"
  End If
  If OpcI.value Then
     TipoDoc = "I"
     DGBanco.Caption = "BALANCE DE SUBMODULO DE INGRESOS"
  End If
  If OpcG.value Then
     TipoDoc = "G"
     DGBanco.Caption = "BALANCE DE SUBMODULO DE GASTOS"
  End If
 'Consultamos el SubModulo
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC = '" & TipoDoc & "' " _
       & "AND TP = 'SUBM' "
  Ejecutar_SQL_SP sSQL

  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC = '" & TipoDoc & "' " _
       & "AND TP = 'SUBM' "
  Select_Adodc AdoBanco, sSQL

  sSQL = "SELECT TC,Codigo,Cta " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If TipoDoc <> "T" Then sSQL = sSQL & "AND TC = '" & TipoDoc & "' "
  If CheqNivel.value = 1 Then sSQL = sSQL & "AND Cta = '" & Cta & "' "
  sSQL = sSQL & "GROUP BY TC,Codigo,Cta " _
       & "ORDER BY TC,Codigo,Cta "
  Contador = 0
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Progreso_Barra.Mensaje_Box = "SUBMODULOS: " & Codigo
       Progreso_Iniciar
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       SubTotal = 0
       Codigo1 = .fields("TC")
       Codigo = .fields("Codigo")
       Cuenta = .fields("Cta")
       
       Do While Not .EOF
          Contador = Contador + 1
          'BalanceSubCtas.Caption = "Procesando: " & Format(Contador / .RecordCount, "00%") & " - " & Codigo
          If Codigo <> .fields("Codigo") Then
             If SubTotal <> 0 And CheqNivel.value <> 1 Then
                SetAdoAddNew "Saldo_Diarios"
                SetAdoFields "Fecha", MBFechaI.Text
                SetAdoFields "Fecha_Venc", MBFechaF.Text
                SetAdoFields "TC", Codigo1
                SetAdoFields "Cta", "I" & String(14, "=") & ">"
                SetAdoFields "CodigoC", Codigo
                SetAdoFields "Comprobante", " * SubTotal:"
                SetAdoFields "TP", "SUBM"
                SetAdoFields "Saldo_Actual", SubTotal
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoUpdate
             End If
             Codigo1 = .fields("TC")
             Codigo = .fields("Codigo")
             Cuenta = .fields("Cta")
             Progreso_Barra.Mensaje_Box = "SUBMODULOS: " & Codigo
             Progreso_Esperar
             SubTotal = 0
          End If
          Saldo = 0
          sSQL = "SELECT * " _
               & "FROM Trans_SubCtas " _
               & "WHERE Fecha <= #" & FechaFin & "# " _
               & "AND Cta = '" & .fields("Cta") & "' " _
               & "AND Codigo = '" & .fields("Codigo") & "' " _
               & "AND T = '" & Normal & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY Codigo,Cta,Fecha,TP,Numero,Factura,Debitos DESC,Creditos,ID "
''               & "ORDER BY Fecha DESC,TP DESC,Numero DESC,Factura DESC,Debitos ASC,Creditos DESC,ID DESC "
          Select_Adodc AdoDetCheq, sSQL
          If AdoDetCheq.Recordset.RecordCount > 0 Then
             AdoDetCheq.Recordset.MoveLast
             Saldo = AdoDetCheq.Recordset.fields("Saldo_MN")
          End If
          Debe = 0: Haber = 0
          sSQL = "SELECT Cta,SUM(Debitos) As Debe1,SUM(Creditos) As Haber1 " _
               & "FROM Trans_SubCtas " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
               & "AND Cta = '" & .fields("Cta") & "' " _
               & "AND Codigo = '" & .fields("Codigo") & "' " _
               & "AND T = '" & Normal & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "GROUP BY Cta "
          Select_Adodc AdoDetCheq, sSQL
          If AdoDetCheq.Recordset.RecordCount > 0 Then
             Debe = Round(AdoDetCheq.Recordset.fields("Debe1"), 2)
             Haber = Round(AdoDetCheq.Recordset.fields("Haber1"), 2)
          End If
          If Saldo <> 0 Then
             If OpcCoop Then
                Select Case MidStrg(.fields("Cta"), 1, 1)
                  Case "1", "4", "6", "8"
                       Total = Saldo - Debe + Haber
                  Case "2", "3", "5", "7", "9"
                       Total = Saldo - Haber + Debe
                End Select
             Else
                Select Case MidStrg(.fields("Cta"), 1, 1)
                  Case "1", "5", "6", "8"
                       Total = Saldo - Debe + Haber
                  Case "2", "3", "4", "7", "9"
                       Total = Saldo - Haber + Debe
                End Select
             End If
             SetAdoAddNew "Saldo_Diarios"
             SetAdoFields "Fecha", MBFechaI.Text
             SetAdoFields "Fecha_Venc", MBFechaF.Text
             SetAdoFields "TC", .fields("TC")
             SetAdoFields "ME", False
             SetAdoFields "Cta", .fields("Cta")
             SetAdoFields "CodigoC", .fields("Codigo")
             SetAdoFields "Comprobante", "-"
             SetAdoFields "TP", "SUBM"
             SetAdoFields "Saldo_Anterior", Total
             SetAdoFields "Ingresos", Debe
             SetAdoFields "Egresos", Haber
             SetAdoFields "Saldo_Actual", Saldo
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
             SubTotal = SubTotal + Saldo
          End If
          Progreso_Esperar
         .MoveNext
       Loop
       If SubTotal <> 0 And CheqNivel.value <> 1 Then
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Fecha", MBFechaI.Text
          SetAdoFields "Fecha_Venc", MBFechaF.Text
          SetAdoFields "TC", Codigo1
          SetAdoFields "Cta", "I" & String(14, "=") & ">"
          SetAdoFields "CodigoC", Codigo
          SetAdoFields "Comprobante", " * SubTotal:"
          SetAdoFields "TP", "SUBM"
          SetAdoFields "Saldo_Actual", SubTotal
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
       End If
      .MoveFirst
   End If
  End With
 'ListarSaldosDiarios
  If SQL_Server Then
     sSQL = "UPDATE Saldo_Diarios " _
          & "SET Comprobante = CC.Cuenta " _
          & "FROM Saldo_Diarios As SD,Catalogo_Cuentas AS CC "
  Else
     sSQL = "UPDATE Saldo_Diarios As SD,Catalogo_Cuentas AS CC " _
          & "SET SD.Comprobante = CC.Cuenta "
  End If
  sSQL = sSQL & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.TC = '" & TipoDoc & "' " _
       & "AND SD.Item = CC.Item " _
       & "AND SD.Cta = CC.Codigo "
  Ejecutar_SQL_SP sSQL
  RatonNormal
  Progreso_Final
  ListarBalanceDeSubModulos
End Sub

Private Sub Listar()
  FechaValida MBFechaI
  FechaValida MBFechaF
  ListarBalanceDeSubModulos
End Sub

Private Sub DGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGBanco.Visible = False
     GenerarDataTexto BalanceSubCtas, AdoBanco
     DGBanco.Visible = True
  End If
  'If CtrlDown And KeyCode = vbKeyB Then BuscarDatos DGBanco, AdoBanco
End Sub

Private Sub Form_Activate()
  Tipo_de_SubCtas "C"
  If Bloquear_Control Then
     Toolbar1.buttons("Procesar").Enabled = False
     Toolbar1.buttons("Listar").Enabled = False
     Toolbar1.buttons("Imprimir").Enabled = False
  End If
End Sub

Private Sub Form_Load()
 'CentrarForm BalanceSubCtas
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCtas1
  ConectarAdodc AdoBanco
  ConectarAdodc AdoDetCheq
  
  DGBanco.Height = MDI_Y_Max - DGBanco.Top - 300
  DGBanco.width = MDI_X_Max - DGBanco.Left
  AdoBanco.Top = DGBanco.Top + DGBanco.Height + 30
  AdoBanco.width = MDI_X_Max - AdoBanco.Left
  
'  Label3.Top = DGBalance.Top + DGBalance.Height + 30
'  Label6.Top = DGBalance.Top + DGBalance.Height + 30
'  Label9.Top = DGBalance.Top + DGBalance.Height + 30
'  Label11.Top = DGBalance.Top + DGBalance.Height + 30
'  LabelTotSaldo.Top = DGBalance.Top + DGBalance.Height + 30
'  LabelTotDebe.Top = DGBalance.Top + DGBalance.Height + 30
'  LabelTotHaber.Top = DGBalance.Top + DGBalance.Height + 30
  
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub OpcC_Click()
  Tipo_de_SubCtas "C"
End Sub

Private Sub OpcG_Click()
  Tipo_de_SubCtas "G"
End Sub

Private Sub OpcI_Click()
  Tipo_de_SubCtas "I"
End Sub

Private Sub OpcP_Click()
  Tipo_de_SubCtas "P"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' MsgBox Button.key
    Select Case Button.key
      Case "Salir": Unload BalanceSubCtas
      Case "Procesar": Procesar
      Case "Listar": Listar
      Case "Imprimir": Imprimir
    End Select
End Sub

Public Sub Tipo_de_SubCtas(TipoCta As String)
  sSQL = "SELECT TSC.Cta & Space(20) & CC.Cuenta As Nombre_Cta,COUNT(TSC.Cta) As TotTrans " _
       & "FROM Catalogo_Cuentas As CC,Trans_SubCtas As TSC " _
       & "WHERE TSC.TC = '" & TipoCta & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Codigo = TSC.Cta " _
       & "AND CC.Item = TSC.Item " _
       & "AND CC.Periodo = TSC.Periodo " _
       & "GROUP BY TSC.Cta, CC.Cuenta " _
       & "ORDER BY TSC.Cta "
  SelectDB_Combo DCCtas, AdoCtas1, sSQL, "Nombre_Cta"
  RatonNormal
End Sub
