VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form NotaDebCred 
   Caption         =   "CONCILIACION DE BANCOS"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "Nota_DC.frx":0000
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   105
      TabIndex        =   20
      Top             =   105
      Width           =   7785
      _ExtentX        =   13732
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
   Begin MSDataGridLib.DataGrid DGAsientos 
      Bindings        =   "Nota_DC.frx":0017
      Height          =   4845
      Left            =   105
      TabIndex        =   19
      Top             =   1890
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   8546
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   2100
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
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   3255
      TabIndex        =   3
      Top             =   525
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   525
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   10290
      Picture         =   "Nota_DC.frx":0031
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   1065
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
      Picture         =   "Nota_DC.frx":08FB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1065
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
      Left            =   7980
      Picture         =   "Nota_DC.frx":11C5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   105
      TabIndex        =   13
      Top             =   945
      Width           =   11250
      Begin VB.TextBox TextDC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9135
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Nota_DC.frx":1607
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox TextValor 
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
         Left            =   9555
         MaxLength       =   13
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Nota_DC.frx":1609
         Top             =   420
         Width           =   1590
      End
      Begin VB.TextBox TextConcepto 
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
         Left            =   2835
         MaxLength       =   50
         TabIndex        =   9
         Top             =   420
         Width           =   5370
      End
      Begin VB.TextBox TextDocumento 
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
         MaxLength       =   8
         TabIndex        =   8
         Top             =   420
         Width           =   1380
      End
      Begin MSMask.MaskEdBox MBoxFechaDC 
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "N/C = 2 "
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
         Left            =   8295
         TabIndex        =   12
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "N/D = 1 "
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
         Left            =   8295
         TabIndex        =   18
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA:"
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
         TabIndex        =   17
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DOCUMENTO:"
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
         Left            =   1470
         TabIndex        =   16
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO:"
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
         Left            =   2835
         TabIndex        =   15
         Top             =   210
         Width           =   5370
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR"
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
         Left            =   9555
         TabIndex        =   14
         Top             =   210
         Width           =   1590
      End
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   2415
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   105
      Top             =   6720
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
      Left            =   2415
      TabIndex        =   2
      Top             =   525
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
      TabIndex        =   0
      Top             =   525
      Width           =   750
   End
End
Attribute VB_Name = "NotaDebCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload NotaDebCred
End Sub

Private Sub Command2_Click()
  With Heads
      .MsgTitulo = "REPORTE DE NOTAS DE DEBITOS Y CREDITOS"
      .MsgObjetivo = "Banco: "
      .MsgConcepto = "Fechas: "
      .TextoObjetivo = DCCtas.Text
      .TextoConcepto = "desde: " & MBoxFechaI.Text & " hasta: " & MBoxFechaF.Text
  End With
  ImprimirDataObjetivo AdoAsientos, True, 1, 8
End Sub

Private Sub Command3_Click()
  RatonReloj
  Frame1.Visible = False
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  sSQL = "SELECT Fecha,Concepto,TP,Documento,Debitos,Creditos,Parcial_ME,Cta,ME," _
       & "CodigoU,Item,Periodo " _
       & "FROM Trans_Conciliacion " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo1 & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Fecha "
  SelectDataGrid DGAsientos, AdoAsientos, sSQL
  Frame1.Visible = True
  RatonNormal
  MBoxFechaDC.SetFocus
End Sub

Private Sub DCCtas_LostFocus()
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  LeerCta Codigo1
  Label5.Caption = " VALOR M/N"
  If Moneda_US Then Label5.Caption = " VALOR M/E"
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCtas, AdoBanco, sSQL, "Nombre_Cta"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm NotaDebCred
  Frame1.Visible = False
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoAsientos
End Sub

Private Sub MBoxFechaDC_GotFocus()
  MBoxFechaDC.Text = LimpiarFechas
End Sub

Private Sub MBoxFechaDC_LostFocus()
  FechaValida MBoxFechaDC
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub TextConcepto_GotFocus()
  TextConcepto.Text = ""
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextDC_GotFocus()
  MarcarTexto TextDC
End Sub

Private Sub TextDC_LostFocus()
  If Val(TextDC.Text) < 1 Or Val(TextDC.Text) > 2 Then
     TextDC.Text = "": TextDC.SetFocus
  End If
End Sub

Private Sub TextDocumento_GotFocus()
  TextDocumento.Text = ""
End Sub

Private Sub TextDocumento_LostFocus()
  TextoValido TextDocumento, True
End Sub

Private Sub TextValor_GotFocus()
  TextValor.Text = ""
End Sub

Private Sub TextValor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValor_LostFocus()
  TextoValido TextValor, True
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  If Val(CCur(TextValor.Text)) > 0 Then
     SetAddNew AdoAsientos
     SetFields AdoAsientos, "Cta", Codigo1
     SetFields AdoAsientos, "Fecha", MBoxFechaDC.Text
     SetFields AdoAsientos, "Concepto", TextConcepto.Text
     SetFields AdoAsientos, "Documento", Val(CCur(TextDocumento.Text))
     If Val(TextDC.Text) = 1 Then
        SetFields AdoAsientos, "TP", "N/D"
        SetFields AdoAsientos, "Debitos", Val(CCur(TextValor.Text))
     Else
        SetFields AdoAsientos, "TP", "N/C"
        SetFields AdoAsientos, "Creditos", Val(CCur(TextValor.Text))
     End If
     SetFields AdoAsientos, "ME", False
     SetFields AdoAsientos, "Parcial_ME", 0
     SetFields AdoAsientos, "Item", NumEmpresa
     SetFields AdoAsientos, "Periodo", Periodo_Contable
     SetFields AdoAsientos, "CodigoU", CodigoUsuario
     SetUpdate AdoAsientos
  End If
  MBoxFechaDC.SetFocus
End Sub
