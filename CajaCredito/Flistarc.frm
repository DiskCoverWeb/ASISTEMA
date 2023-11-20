VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FListarCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTADO DE CUENTAS"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGCuentas 
      Bindings        =   "Flistarc.frx":0000
      Height          =   4950
      Left            =   105
      TabIndex        =   14
      ToolTipText     =   "<Ctrl + P> Imprimir <Ctrl + B> Buscar Datos, <F1> Generar Archivo"
      Top             =   1050
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   8731
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   420
      Top             =   3780
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
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por:"
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
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      Begin VB.OptionButton OpcT 
         Caption         =   "&Todas"
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
         TabIndex        =   4
         Top             =   525
         Width           =   1065
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
         Left            =   1680
         TabIndex        =   3
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton OpcA 
         Caption         =   "A&nuladas"
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
         TabIndex        =   2
         Top             =   525
         Width           =   1380
      End
      Begin VB.OptionButton OpcAp 
         Caption         =   "A&perturadas"
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
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&A->Z"
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
         Left            =   8400
         Picture         =   "Flistarc.frx":0019
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Z->A"
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
         Left            =   9030
         Picture         =   "Flistarc.frx":06AF
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   645
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
         Height          =   645
         Left            =   9660
         Picture         =   "Flistarc.frx":0D45
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   210
         Width           =   645
      End
      Begin VB.ComboBox CmbLibretas 
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
         Left            =   5565
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   525
         Width           =   2745
      End
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   4200
         TabIndex        =   8
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   525
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
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   4200
         TabIndex        =   7
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   210
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
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha &hasta"
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
         Left            =   2940
         TabIndex        =   6
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha &desde"
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
         Left            =   2940
         TabIndex        =   5
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Patron de &Ordenacion"
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
         Left            =   5565
         TabIndex        =   9
         Top             =   210
         Width           =   2745
      End
   End
   Begin MSAdodcLib.Adodc AdoListBusq 
      Height          =   330
      Left            =   420
      Top             =   4095
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
      Caption         =   "ListBusq"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   105
      Top             =   5985
      Width           =   10410
      _ExtentX        =   18362
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
      Caption         =   "Cuentas"
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
Attribute VB_Name = "FListarCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbLibretas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then DBGCuentas.SetFocus
End Sub

Private Sub Command1_Click()
  Listar_Ctas_Ahorros
  DGCuentas.SetFocus
End Sub

Private Sub Command2_Click()
  Listar_Ctas_Ahorros
  DGCuentas.SetFocus
End Sub

Private Sub Command3_Click()
   Unload FListarCtas
End Sub

Private Sub DGCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyB Then BuscarDatos DGCuentas, AdoCuentas
  If KeyCode = vbKeyF1 Then
     DGCuentas.Visible = False
     GenerarDataTexto FListarCtas, AdoCuentas
     DGCuentas.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     Mensajes = "Imprimir Listado"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
        DGCuentas.Visible = False
        SQLMsg1 = "REPORTE DE CLIENTES"
        ImprimirAdodc AdoCuentas, 1, 8
        DGCuentas.Visible = True
     End If
  End If
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT C.T,C.Fecha_Registro,C.Cuenta_No,Cliente,Representante,Cl.Ciudad,Cl.CI_RUC," _
        & "Cl.Telefono,Cl.FAX,Cl.Lugar_Trabajo,Cl.Direccion " _
        & "FROM Clientes_Datos_Extras As C,Clientes As Cl " _
        & "WHERE C.Item = '" & NumEmpresa & "' " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Cl.Codigo " _
        & "ORDER BY C.Fecha_Registro,Cuenta_No "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   CmbLibretas.Clear
   With AdoCuentas
      For I = 0 To .Recordset.Fields.Count - 1
          CmbLibretas.AddItem .Recordset.Fields(I).Name
      Next I
   End With
   CmbLibretas.Text = CmbLibretas.List(0)
   Listar_Ctas_Ahorros
   RatonNormal
   CmbLibretas.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FListarCtas
   ConectarAdodc AdoCtas
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListBusq
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
End Sub

Public Sub Listar_Ctas_Ahorros()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "SELECT C.T,C.Fecha_Registro,C.Cuenta_No,Cliente,Representante,Cl.Ciudad,Cl.CI_RUC," _
       & "Cl.Telefono,Cl.FAX,Cl.TelefonoT,Cl.Lugar_Trabajo,Cl.Direccion,Cl.Profesion " _
       & "FROM Clientes_Datos_Extras As C,Clientes As Cl "
  If OpcAp.value Then
     sSQL = sSQL & "WHERE C.Fecha_A BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = sSQL & "WHERE C.Fecha_Registro BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  sSQL = sSQL _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND C.Codigo = Cl.Codigo "
  If OpcA.value Then sSQL = sSQL & "AND C.T = 'A' "
  If OpcV.value Then sSQL = sSQL & "AND C.T = 'N' "
  'If OpcAp.Value Then sSQL = sSQL & "AND C.T = 'P' "
  Codigo = CmbLibretas.Text
  Select Case UCase(Codigo)
    Case "T", "FECHA_REGISTRO": Codigo = "C." & Codigo
  End Select
  sSQL = sSQL & "AND C.ITem = '" & NumEmpresa & "' " _
       & "ORDER BY " & Codigo & " "
  'MsgBox sSQL
  SelectDataGrid DGCuentas, AdoCuentas, sSQL
End Sub
