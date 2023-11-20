VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form RPrestamo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDOS DIARIOS"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11040
   Begin VB.CommandButton Command6 
      Caption         =   "&Consultar Detalle"
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
      Left            =   9975
      Picture         =   "RPretamo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1680
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "RPretamo.frx":0442
      DataSource      =   "AdoTipo"
      Height          =   315
      Left            =   2415
      TabIndex        =   24
      Top             =   1050
      Width           =   7470
      _ExtentX        =   13176
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
   Begin MSDataGridLib.DataGrid DGSaldos 
      Bindings        =   "RPretamo.frx":0458
      Height          =   4635
      Left            =   105
      TabIndex        =   23
      Top             =   1470
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   8176
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   420
      Top             =   2415
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
      Caption         =   "Tipo"
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
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   945
      TabIndex        =   2
      Top             =   1050
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.CommandButton Command5 
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
      Left            =   10500
      Picture         =   "RPretamo.frx":0470
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   105
      Width           =   435
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9975
      Picture         =   "RPretamo.frx":0572
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   435
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
      Left            =   105
      TabIndex        =   10
      Top             =   0
      Width           =   4215
      Begin VB.OptionButton OpcC 
         Caption         =   "C&ancelados"
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
         Left            =   2730
         TabIndex        =   6
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton OpcV 
         Caption         =   "&Vigentes"
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
         TabIndex        =   12
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "&Prestamos"
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
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   4410
      TabIndex        =   14
      Top             =   0
      Width           =   5475
      Begin VB.ComboBox CBPrestamos 
         Height          =   315
         Left            =   3465
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   210
         Width           =   1905
      End
      Begin VB.CheckBox CheqLib_No 
         Caption         =   "&Número de Libreta"
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
         TabIndex        =   17
         Top             =   210
         Width           =   1905
      End
      Begin MSMask.MaskEdBox MBoxCuenta 
         Height          =   330
         Left            =   1995
         TabIndex        =   19
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   210
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   192
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "CCCCCCCC-C"
         Mask            =   "########-#"
         PromptChar      =   "0"
      End
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   735
      Width           =   1380
      _ExtentX        =   2434
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
      Height          =   855
      Left            =   9975
      Picture         =   "RPretamo.frx":0674
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3570
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Consultar"
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
      Left            =   9975
      Picture         =   "RPretamo.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   735
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir Listado"
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
      Left            =   9975
      Picture         =   "RPretamo.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2625
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   105
      Top             =   6090
      Width           =   9780
      _ExtentX        =   17251
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   2730
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
   Begin VB.Label LabelCuota 
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
      Left            =   7980
      TabIndex        =   21
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label LabelCapital 
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
      Left            =   5145
      TabIndex        =   18
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUOTA"
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
      Left            =   7140
      TabIndex        =   22
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   3150
      TabIndex        =   20
      Top             =   6510
      Width           =   2010
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &TIPO DE PRESTAMO"
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
      Left            =   2415
      TabIndex        =   13
      Top             =   735
      Width           =   7470
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &HASTA:"
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
      TabIndex        =   9
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label LabelInteres 
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
      Left            =   1155
      TabIndex        =   7
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INTERES"
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
      TabIndex        =   8
      Top             =   6510
      Width           =   1065
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &DESDE:"
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
      Top             =   735
      Width           =   855
   End
End
Attribute VB_Name = "RPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaValida MBoxFecha, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFecha.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  If OpcP.value Then
     sSQL = "SELECT C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Tasa,P.Meses,P.Fecha,P.Pagos,P.Saldo_Pendiente,P.Encaje " _
          & "FROM Prestamos As P, Clientes_Datos_Extras As C "
  ElseIf OpcV.value Then
     sSQL = "SELECT P.V,C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Mes_No,P.Fecha,P.Capital,P.Pagos,P.Saldo " _
          & "FROM Trans_Prestamos As P, Clientes_Datos_Extras As C "
  Else
     sSQL = "SELECT P.V,C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Mes_No,P.Fecha,P.Capital,P.Pagos,P.Saldo " _
          & "FROM Trans_Prestamos As P, Clientes_Datos_Extras As C "
  End If
  sSQL = sSQL & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND P.Cuenta_No = C.Cuenta_No " _
       & "AND P.TP = '" & Codigo & "' "
  If OpcP.value Then
     sSQL = sSQL & "AND P.T = 'P' "
  ElseIf OpcV.value Then
     sSQL = sSQL & "AND P.T = 'P' "
  Else
     sSQL = sSQL & "AND P.T = 'C' "
  End If
  If CheqLib_No.value = 1 Then
     sSQL = sSQL & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' "
  End If
  If NombreCampo = "Fecha" Then NombreCampo = "P." & NombreCampo
  sSQL = sSQL & "ORDER BY " & NombreCampo & " ASC "
  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  DGSaldos.Visible = False
  RatonReloj
  Total = 0: Total_ME = 0: Saldo = 0
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcP.value Then
             Total_ME = Total_ME + .Fields("Saldo_Pendiente")
          Else
             Total = Total + .Fields("Pagos")
             Total_ME = Total_ME + .Fields("Capital")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  If OpcP.value Then Label5.Caption = " SALDO PENDIENTE"
  If OpcV.value Then Label5.Caption = " CAPITAL"
  If OpcC.value Then Label5.Caption = " CAPITAL"
  LabelCuota.Caption = Format(Total, "#,##0.00")
  LabelCapital.Caption = Format(Total_ME, "#,##0.00")
  LabelInteres.Caption = Format(Saldo, "#,##0.00")
  DGSaldos.Visible = True
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Command2_Click()
  SQLMsg2 = "": SQLMsg3 = ""
  If Opcion = 99 Then
     SQLMsg1 = "LISTADO DE HISTORIAL DE PRESTAMOS"
     ImprimirPrestamosVencidos AdoSaldos, True, 2, 9, False, CheqLib_No.value, True
  Else
  SQLMsg1 = "LISTADO DE " & UCase(DCTipo.Text) & " ["
  If OpcP.value Then
     SQLMsg1 = SQLMsg1 & UCase(OpcP.Caption) & "]"
  ElseIf OpcV.value Then
     SQLMsg1 = SQLMsg1 & UCase(OpcV.Caption) & "]"
  Else
     SQLMsg1 = SQLMsg1 & UCase(OpcC.Caption) & "]"
  End If
  ImprimirPrestamosVencidos AdoSaldos, True, 2, 9, OpcP.value, CheqLib_No.value
  End If
End Sub

Private Sub Command3_Click()
  Opcion = 1
  FechaValida MBoxFecha, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFecha.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = SinEspaciosIzq(DCTipo.Text)
  CodigoCli = Ninguno
'''  sSQL = "SELECT * FROM Clientes_Datos_Extras " _
'''       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
'''  SelectAdodc AdoSaldos, sSQL
'''  If AdoSaldos.Recordset.RecordCount > 0 Then
'''     CodigoCli = AdoSaldos.Recordset.Fields("Codigo")
'''  End If
  If OpcP.value Then
     sSQL = "SELECT P.T,C.Cliente,P.Cuenta_No,P.Credito_No," _
          & "P.Capital,P.Tasa,P.Meses,P.Fecha,P.Pagos," _
          & "P.Saldo_Pendiente,P.Encaje,Cta.Tipo " _
          & "FROM Prestamos As P, Clientes As C,Clientes_Datos_Extras As Cta " _
          & "WHERE P.T <> 'A' " _
          & "AND P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ElseIf OpcV.value Then
     sSQL = "SELECT P.V,C.Cliente,P.Cuenta_No,P.Credito_No," _
          & "P.Cuota_No,P.Fecha,P.Fecha_C,P.Capital,P.Pagos,P.Saldo,Cta.Tipo " _
          & "FROM Trans_Prestamos As P,Clientes As C,Clientes_Datos_Extras As Cta " _
          & "WHERE P.T = 'P' " _
          & "AND P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT P.V,C.Cliente,P.Cuenta_No,P.Credito_No," _
          & "P.Cuota_No,P.Fecha,P.Fecha_C,P.Capital,P.Pagos,P.Saldo,Cta.Tipo " _
          & "FROM Trans_Prestamos As P,Clientes As C,Clientes_Datos_Extras As Cta " _
          & "WHERE P.T = 'C' " _
          & "AND P.Fecha_C BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If CheqLib_No.value = 1 Then
     sSQL = sSQL & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' "
  End If
  sSQL = sSQL & "AND P.Cuenta_No = Cta.Cuenta_No " _
       & "AND Cta.Codigo = C.Codigo " _
       & "AND P.TP = '" & Codigo & "' " _
       & "AND P.Item = '" & NumEmpresa & "' "
  If OpcP.value Then
     sSQL = sSQL & "ORDER BY C.Cliente,P.Credito_No,Meses "
  Else
     sSQL = sSQL & "ORDER BY C.Cliente,P.Credito_No,Cuota_No "
  End If

  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  DGSaldos.Visible = False
  RatonReloj
  Total = 0: Total_ME = 0: Saldo = 0
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcP.value Then
             Total_ME = Total_ME + .Fields("Saldo_Pendiente")
          Else
             Total = Total + .Fields("Pagos")
             Total_ME = Total_ME + .Fields("Capital")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  If OpcP.value Then Label5.Caption = " SALDO PENDIENTE"
  If OpcV.value Then Label5.Caption = " CAPITAL"
  If OpcC.value Then Label5.Caption = " CAPITAL"
  LabelCuota.Caption = Format(Total, "#,##0.00")
  LabelCapital.Caption = Format(Total_ME, "#,##0.00")
  LabelInteres.Caption = Format(Saldo, "#,##0.00")
  DGSaldos.Visible = True
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Command4_Click()
  Unload RPrestamo
End Sub

Private Sub Command5_Click()
  FechaValida MBoxFecha, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFecha.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  If OpcP.value Then
     sSQL = "SELECT C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Tasa,P.Meses,P.Fecha,P.Pagos,P.Saldo_Pendiente,P.Encaje " _
          & "FROM Prestamos As P, Clientes_Datos_Extras As C "
  ElseIf OpcV.value Then
     sSQL = "SELECT P.V,C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Mes_No,P.Fecha,P.Capital,P.Pagos,P.Saldo " _
          & "FROM Trans_Prestamos As P, Clientes_Datos_Extras As C "
  Else
     sSQL = "SELECT P.V,C.Apellidos,C.Nombres,P.Cuenta_No,P.Credito_No," _
          & "P.Mes_No,P.Fecha,P.Capital,P.Pagos,P.Saldo " _
          & "FROM Trans_Prestamos As P, Clientes_Datos_Extras As C "
  End If
  sSQL = sSQL & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND P.Cuenta_No = C.Cuenta_No " _
       & "AND P.TP = '" & Codigo & "' "
  If OpcP.value Then
     sSQL = sSQL & "AND P.T = 'P' "
  ElseIf OpcV.value Then
     sSQL = sSQL & "AND P.T = 'P' "
  Else
     sSQL = sSQL & "AND P.T = 'C' "
  End If
  If CheqLib_No.value = 1 Then
     sSQL = sSQL & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' "
  End If
  If NombreCampo = "Fecha" Then NombreCampo = "P." & NombreCampo
  sSQL = sSQL & "ORDER BY " & NombreCampo & " DESC "
  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  DGSaldos.Visible = False
  RatonReloj
  Total = 0: Total_ME = 0: Saldo = 0
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcP.value Then
             Total_ME = Total_ME + .Fields("Saldo_Pendiente")
          Else
             Total = Total + .Fields("Pagos")
             Total_ME = Total_ME + .Fields("Capital")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  If OpcP.value Then Label5.Caption = " SALDO PENDIENTE"
  If OpcV.value Then Label5.Caption = " CAPITAL"
  If OpcC.value Then Label5.Caption = " CAPITAL"
  LabelCuota.Caption = Format(Total, "#,##0.00")
  LabelCapital.Caption = Format(Total_ME, "#,##0.00")
  LabelInteres.Caption = Format(Saldo, "#,##0.00")
  DGSaldos.Visible = True
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub DBGSaldos_DblClick()
  NombreCampo = DBGSaldos.Columns(I).adoField
End Sub

Private Sub DBGSaldos_HeadClick(ByVal ColIndex As Integer)
  I = ColIndex
End Sub

Private Sub Command6_Click()
  FechaValida MBoxFecha, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFecha.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = SinEspaciosIzq(DCTipo.Text)
  CodigoCli = Ninguno
  If CheqLib_No.value = 1 Then
     sSQL = "SELECT P.T,C.Cliente,P.Cuenta_No,P.Credito_No," _
          & "P.Cuota_No,P.Fecha,P.Fecha_C,P.Capital,P.Pagos,P.Saldo " _
          & "FROM Trans_Prestamos As P,Clientes As C,Clientes_Datos_Extras As Cta " _
          & "WHERE P.T <> 'A' " _
          & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' " _
          & "AND P.TP = '" & Codigo & "' "
     If CBPrestamos.Text <> "Todos" Then sSQL = sSQL & "AND P.Credito_No = '" & CBPrestamos.Text & "' "
     sSQL = sSQL & "AND P.Item = '" & NumEmpresa & "' " _
          & "AND P.Cuenta_No = Cta.Cuenta_No " _
          & "AND Cta.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente,P.Credito_No,Cuota_No "
     SelectDataGrid DGSaldos, AdoSaldos, sSQL
  End If
  Opcion = 99
End Sub

Private Sub DGSaldos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     DGSaldos.Visible = False
     GenerarDataTexto RPrestamo, AdoSaldos
     DGSaldos.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel(3) Or CNivel(6) Then
        Command2.Enabled = False
     End If
  End If
  NombreCampo = "Apellidos"
  sSQL = "SELECT CTP & ' ' & Descripcion As TipoPrest " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE TC = True " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY CTP,Descripcion "
  SelectDBCombo DCTipo, AdoTipo, sSQL, "TipoPrest"
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm RPrestamo
  ConectarAdodc AdoAux
  ConectarAdodc AdoTipo
  ConectarAdodc AdoSaldos
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_LostFocus()
   CBPrestamos.Clear
   CBPrestamos.AddItem "Todos"
   sSQL = "SELECT * " _
        & "FROM Prestamos " _
        & "WHERE Cuenta_No = '" & MBoxCuenta.Text & " ' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "ORDER BY TP,Credito_No "
   SelectAdodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      Do While Not AdoAux.Recordset.EOF
         CBPrestamos.AddItem AdoAux.Recordset.Fields("Credito_No")
         AdoAux.Recordset.MoveNext
      Loop
   End If
   CBPrestamos.Text = "Todos"
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub
