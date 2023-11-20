VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LibroRetenciones 
   Caption         =   "LISTAR COMPROBANTE DE RETENCION DE EGRESOS"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGRet 
      Bindings        =   "RepoRete.frx":0000
      Height          =   5370
      Left            =   105
      TabIndex        =   17
      Top             =   1680
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   9472
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   105
      Top             =   7035
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
      Caption         =   "DetRet"
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
      Height          =   1590
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   11250
      Begin MSDataListLib.DataCombo DCCtas 
         Bindings        =   "RepoRete.frx":0018
         DataSource      =   "AdoBanco"
         Height          =   315
         Left            =   105
         TabIndex        =   16
         Top             =   1155
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Ctas"
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
      Begin VB.OptionButton OpcDEP 
         Caption         =   "Por Relación Dependencia"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   945
         Width           =   2745
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Modificar"
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
         Left            =   8085
         Picture         =   "RepoRete.frx":002F
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   210
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   1365
         TabIndex        =   4
         Top             =   525
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
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   1365
         TabIndex        =   2
         Top             =   210
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
         Height          =   855
         Left            =   10185
         Picture         =   "RepoRete.frx":0471
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
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
         Left            =   9135
         Picture         =   "RepoRete.frx":0D3B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command4 
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
         Left            =   7035
         Picture         =   "RepoRete.frx":1605
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcTodos 
         Caption         =   "Todas las Retención"
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
         TabIndex        =   9
         Top             =   945
         Value           =   -1  'True
         Width           =   2220
      End
      Begin VB.OptionButton OpcIndiv 
         Caption         =   "Una Cta. de Ret."
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
         Left            =   2310
         TabIndex        =   10
         Top             =   945
         Width           =   1905
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Consulta"
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
         Left            =   2835
         TabIndex        =   5
         Top             =   210
         Width           =   4110
         Begin VB.OptionButton OpcT 
            Caption         =   "Todas"
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
            Left            =   3045
            TabIndex        =   8
            Top             =   210
            Width           =   960
         End
         Begin VB.OptionButton OpcA 
            Caption         =   "Anuladas"
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
            Left            =   1680
            TabIndex        =   7
            Top             =   210
            Width           =   1170
         End
         Begin VB.OptionButton OpcP 
            Caption         =   "Pendientes"
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
            Left            =   210
            TabIndex        =   6
            Top             =   210
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Final"
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
         Width           =   1275
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Inicial"
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
         Width           =   1275
      End
      Begin VB.Label LabelTotSaldo 
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
         Left            =   9240
         TabIndex        =   15
         Top             =   1155
         Width           =   1905
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Retenciones"
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
         Left            =   7455
         TabIndex        =   14
         Top             =   1155
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   210
      Top             =   2310
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   210
      Top             =   2625
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Asientos"
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
Attribute VB_Name = "LibroRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  RatonReloj
  Si_No = False
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = SinEspaciosIzq(DCCtas.Text)
  sSQL = "SELECT * " _
       & "FROM Trans_Retenciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If OpcIndiv.Value Then sSQL = sSQL & "AND Cta = '" & Codigo & "' "
  If OpcP.Value Then
     sSQL = sSQL & "AND T = 'N' "
  ElseIf OpcA.Value Then
     sSQL = sSQL & "AND T = 'A' "
  End If
  sSQL = sSQL & "ORDER BY Cta,Porc,Fecha "
  SelectDataGrid DGRet, AdoDetRet, sSQL
  Sumatoria = 0
  With AdoDetRet.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Sumatoria = Sumatoria + .Fields("Valor_Ret")
         .MoveNext
       Loop
   End If
  End With
  DGRet.AllowDelete = False
  DGRet.AllowUpdate = True
  LabelTotSaldo.Caption = Format(Sumatoria, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command2_Click()
  DGRet.Visible = False
  If OpcDEP.Value Then
     MensajeEncabData = "RETENCIONES POR RELACION DE DEPENDENCIA"
     SQLMsg1 = "Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
     Imprimir_Libro_Dependencia AdoDetRet, 8
  Else
     MensajeEncabData = "REPORTE DE RETENCIONES"
     SQLMsg1 = "Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
     Imprimir_Libro_Retenciones AdoDetRet
  End If
  DGRet.Visible = True
End Sub

Private Sub Command3_Click()
  Unload LibroRetenciones
End Sub

Private Sub Command4_Click()
  RatonReloj
  Si_No = True
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = SinEspaciosIzq(DCCtas.Text)
  If OpcDEP.Value Then
     sSQL = "SELECT R.Fecha,C.Cliente,C.CI_RUC As Cedula,R.SN, R.AporteP As Aporte_Per, R.PorAp As Porc_Aporte, R.InLiqui As Ingresos_Liq, R.Salario As Base_Imponible, R.Retenido " _
          & "FROM Trans_Rol_Pagos As R, Clientes As C " _
          & "WHERE R.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND R.Item = '" & NumEmpresa & " ' " _
          & "AND R.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = R.Codigo " _
          & "ORDER BY R.Fecha,C.Cliente "
  Else
     sSQL = "SELECT Cat.Cuenta,C.T,C.Fecha,R.Autorizacion,R.Secuencial,Cl.Cliente,Cl.CI_RUC," _
          & "C.TP,C.Numero,R.Valor_Fact,(MontoIVA1 + MontoIVA2) As IVA,R.Valor_Ret,R.Porc " _
          & "FROM Comprobantes As C,Trans_Retenciones As R,Clientes As Cl,Catalogo_Cuentas As Cat " _
          & "WHERE R.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND R.Item = '" & NumEmpresa & "' " _
          & "AND R.Periodo = '" & Periodo_Contable & "' " _
          & "AND R.Numero = C.Numero " _
          & "AND R.TP = C.TP " _
          & "AND R.Item = C.Item " _
          & "AND R.Item = Cat.Item " _
          & "AND R.Periodo = C.Periodo " _
          & "AND R.Periodo = Cat.Periodo " _
          & "AND R.Cta = Cat.Codigo " _
          & "AND C.Codigo_B = Cl.Codigo "
     If OpcIndiv.Value Then sSQL = sSQL & "AND R.Cta = '" & Codigo & "' "
     If OpcP.Value Then
        sSQL = sSQL & "AND R.T = 'N' "
     ElseIf OpcA.Value Then
        sSQL = sSQL & "AND R.T = 'A' "
     End If
     sSQL = sSQL & "ORDER BY Cta,R.Porc,R.Fecha "
  End If
  SelectDataGrid DGRet, AdoDetRet, sSQL
  Sumatoria = 0
  With AdoDetRet.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If OpcDEP.Value Then
             Sumatoria = Sumatoria + .Fields("Aporte_Per")
          Else
             Sumatoria = Sumatoria + .Fields("Valor_Ret")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGRet.AllowUpdate = False
  DGRet.AllowDelete = True
  LabelTotSaldo.Caption = Format(Sumatoria, "#,##0.00")
  RatonNormal
End Sub

Private Sub DGRet_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then GenerarDataTexto LibroRetenciones, AdoDetRet
  If CtrlDown And KeyCode = vbKeyDelete Then
     If Si_No Then
     Mensajes = "Seguro de Elimnar Retencion de: " & vbCrLf _
              & AdoDetRet.Recordset.Fields("Cliente") & vbCrLf _
              & "Comprobante de " & AdoDetRet.Recordset.Fields("TP") _
              & ", Numero No. " & AdoDetRet.Recordset.Fields("Numero")
     Titulo = "ELIMINACION"
     If BoxMensaje = vbYes Then
     SQL1 = "DELETE * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = '" & AdoDetRet.Recordset.Fields("TP") & "' " _
          & "AND Numero = " & AdoDetRet.Recordset.Fields("Numero") & " " _
          & "AND Secuencial = '" & AdoDetRet.Recordset.Fields("Secuencial") & "' "
     ConectarAdoExecute SQL1
     Command4_Click
     End If
     Else
       MsgBox "No se puede eliminar"
     End If
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(10) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('RF','IV') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCtas, AdoBanco, sSQL, "Nombre_Cta"
  RatonNormal
End Sub

Private Sub Form_Load()
' CentrarForm LibroRetenciones
  ConectarAdodc AdoBanco
  ConectarAdodc AdoDetRet
  ConectarAdodc AdoAsientos
End Sub

Private Sub MBoxFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub
