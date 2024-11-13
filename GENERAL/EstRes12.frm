VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form EstadoResult12Meses 
   Caption         =   "ESTADO DE RESULTADOS POR MES"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGBalanceG 
      Bindings        =   "EstRes12.frx":0000
      Height          =   4425
      Left            =   105
      TabIndex        =   10
      Top             =   1470
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   7805
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "EstRes12.frx":001A
      DataSource      =   "AdoAgencias"
      Height          =   375
      Left            =   2310
      TabIndex        =   8
      Top             =   945
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   661
      _Version        =   393216
      ListField       =   "|"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   420
      Top             =   3045
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Agencias"
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
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "No&mbre de Agencia:"
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
      TabIndex        =   7
      Top             =   945
      Width           =   2115
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   420
      Top             =   3360
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   420
      Top             =   3675
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin MSAdodcLib.Adodc AdoResultado 
      Height          =   330
      Left            =   420
      Top             =   3990
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Resultado"
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
   Begin MSAdodcLib.Adodc AdoPres 
      Height          =   330
      Left            =   420
      Top             =   4305
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Pres"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11340
      Top             =   1680
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
            Picture         =   "EstRes12.frx":0034
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EstRes12.frx":090E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EstRes12.frx":0D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EstRes12.frx":19B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1588
      ButtonWidth     =   1376
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del modulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar"
            Key             =   "Analitico"
            Object.ToolTipText     =   "Procesar Analitico Mensual"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ER"
                  Text            =   "Estado de Resultado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ES"
                  Text            =   "Estado de Situacion"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "Excel"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "PDF"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3465
         TabIndex        =   0
         Top             =   0
         Width           =   12300
         Begin VB.OptionButton OpcSM 
            Caption         =   "P&rocesar sin SubCuentas de Bloque"
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
            Left            =   8715
            TabIndex        =   6
            Top             =   210
            Width           =   3480
         End
         Begin VB.OptionButton OpcT 
            Caption         =   "Pr&ocesar Mayores y Subcuentas de Bloque"
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
            Left            =   4410
            TabIndex        =   5
            Top             =   210
            Value           =   -1  'True
            Width           =   4110
         End
         Begin MSMask.MaskEdBox MBFechaF 
            Height          =   330
            Left            =   2940
            TabIndex        =   4
            Top             =   315
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
            Top             =   315
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
            Top             =   315
            Width           =   750
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
            Left            =   2205
            TabIndex        =   3
            Top             =   315
            Width           =   750
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoBalanceG 
      Height          =   330
      Left            =   105
      Top             =   6195
      Width           =   9885
      _ExtentX        =   17436
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
      Caption         =   "BalanceG"
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
Attribute VB_Name = "EstadoResult12Meses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ProcesoBAM As Boolean

Private Sub CheckAgencia_Click()
  If CheckAgencia.value = 1 Then
     DCAgencia.Visible = True
  Else
     DCAgencia.Visible = False
  End If
End Sub

Private Sub Form_Activate()
  If CNivel(7) Then
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload EstadoResult12Meses
     RatonNormal
  Else
     Eliminar_Duplicados_SP "Catalogo_Cuentas", "Codigo"
     Fechas_Balances "Balance_Analitico", MBFechaI, MBFechaF
     DGBalanceG.Caption = "ESTADO DE RESULTADOS ANALITICOS MENSUALES"
     sSQL = "SELECT (Item & '  ' & Empresa) As NombreEmpresa " _
          & "FROM Empresas " _
          & "WHERE Empresa <> '" & Ninguno & "' " _
          & "ORDER BY Item "
     SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NombreEmpresa", False
     RatonNormal
  End If
End Sub

Private Sub Form_Load()
 'CentrarForm EstadoResult12Meses
  ConectarAdodc AdoPres
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoResultado
  ConectarAdodc AdoAgencias
  ConectarAdodc AdoBalanceG
     
  DGBalanceG.width = MDI_X_Max - 100
  DGBalanceG.Height = MDI_Y_Max - DGBalanceG.Top - 350
  AdoBalanceG.Top = DGBalanceG.Top + DGBalanceG.Height + 5
  AdoBalanceG.width = MDI_X_Max - 100
  ProcesoBAM = False
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF, False
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI, False
End Sub

Public Sub Procesar_Balance_12_Meses(TipoBalance As String)
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
  
  Control_Procesos Normal, "(" & TipoBalance & ") Analitico Mensual del " & MBFechaI & "  al " & MBFechaF
  Update_Fechas "Balance_Analitico", MBFechaI, MBFechaF
  DGBalanceG.Visible = False
  RatonReloj
  NumItemTemp = 0
  If CheckAgencia.value = 1 Then
     NumItemTemp = SinEspaciosIzq(DCAgencia.Text)
  Else
     NumItemTemp = NumEmpresa
  End If
  FechaValida MBFechaI
  FechaValida MBFechaF
 
 'Empezamos a sacar los saldos
  Select Case TipoBalance
    Case "ER": DGBalanceG.Caption = "RESUMEN ANALITICO MENSUAL DE ESTADO DE RESULTADO"
    Case "ES": DGBalanceG.Caption = "RESUMEN ANALITICO MENSUAL DE ESTADO DE SITUACION"
  End Select
  'DGBalanceG.Caption = DGBalanceG.Caption & "DESDE EL: " & MBFechaI.Text & "  AL " & MBFechaF.Text
  
 'Insertamos las Cuentas y SubModulos del Balance Analitico Mensual
  Procesar_Balance_Analitico_Mensual_SP TipoBalance, OpcT.value, MBFechaI, MBFechaF, sSQL
  sSQL = sSQL _
       & "FROM Reporte_Analitico_Mensual " _
       & "WHERE Item = '" & NumItemTemp & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TB = '" & TipoBalance & "' " _
       & "ORDER BY Codigo_Aux "
  Select_Adodc_Grid DGBalanceG, AdoBalanceG, sSQL
  With AdoBalanceG.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total_Presupuesto = Total_Presupuesto + .fields("Presupuesto")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGBalanceG.Visible = True
  ProcesoBAM = True
  RatonNormal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "Salir"
         Unload EstadoResult12Meses
    Case "Excel"
         If ProcesoBAM Then
            DGBalanceG.Visible = False
            GenerarDataTexto EstadoResult12Meses, AdoBalanceG
            DGBalanceG.Visible = True
         Else
            MsgBox "No se puede generar el informe," & vbCrLf & vbCrLf & "porque no ha realizado el proceso"
         End If
    Case "PDF"
         If ProcesoBAM Then
            DGBalanceG.Visible = False
            SQLMsg1 = "DEL " & MBFechaI & " AL " & MBFechaF
            PDF_Analitico_Mensual AdoBalanceG, DGBalanceG.Caption, MBFechaF
            DGBalanceG.Visible = True
         Else
            MsgBox "No se puede generar el informe," & vbCrLf & vbCrLf & "porque no ha realizado el proceso"
         End If
  End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 'MsgBox ButtonMenu.key
  Procesar_Balance_12_Meses ButtonMenu.key
End Sub
