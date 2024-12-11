VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form BuscarContabil 
   Caption         =   "Presentacion de una base de datos:"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Consultar"
            Object.ToolTipText     =   "Consultar Datos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Envia a Excel los resultados"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modificar Comprobantes"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   2415
         TabIndex        =   8
         Top             =   0
         Width           =   5895
         Begin VB.CheckBox CheqPorFecha 
            Caption         =   "Por Fecha"
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
            Top             =   210
            Width           =   1275
         End
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   4410
            TabIndex        =   10
            Top             =   210
            Visible         =   0   'False
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
            Left            =   2205
            TabIndex        =   11
            Top             =   210
            Visible         =   0   'False
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
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hasta"
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
            Left            =   3675
            TabIndex        =   13
            Top             =   210
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Desde"
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
            Visible         =   0   'False
            Width           =   750
         End
      End
   End
   Begin VB.ComboBox CmbCampos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   105
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "CmbCampos"
      Top             =   1050
      Width           =   1485
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "BuscarC.frx":0000
      Height          =   3165
      Left            =   1680
      TabIndex        =   5
      Top             =   1155
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5583
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
   Begin MSAdodcLib.Adodc AdoListar 
      Height          =   330
      Left            =   2100
      Top             =   2100
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Listar"
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
   Begin VB.TextBox TextPatron 
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
      Left            =   4095
      TabIndex        =   3
      Top             =   735
      Width           =   5685
   End
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   9765
      Picture         =   "BuscarC.frx":0017
      TabIndex        =   4
      Top             =   735
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   1680
      Top             =   4305
      Width           =   8415
      _ExtentX        =   14843
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
      Caption         =   "Query"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   105
      Top             =   7455
      Width           =   9990
      _ExtentX        =   17621
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
      Caption         =   "Query1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DGQuery1 
      Bindings        =   "BuscarC.frx":08E1
      Height          =   2220
      Left            =   105
      TabIndex        =   6
      Top             =   5250
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   3916
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10290
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuscarC.frx":08F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuscarC.frx":0C13
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuscarC.frx":0F2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BuscarC.frx":1B7F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PATRON DE BUSQUEDA"
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
      TabIndex        =   1
      Top             =   735
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAMPO:"
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
      Width           =   1485
   End
End
Attribute VB_Name = "BuscarContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqPorFecha_Click()
 If CheqPorFecha.value = 0 Then
    Label8.Visible = False
    Label3.Visible = False
    MBoxFechaI.Visible = False
    MBoxFechaF.Visible = False
 Else
    Label8.Visible = True
    Label3.Visible = True
    MBoxFechaI.Visible = True
    MBoxFechaF.Visible = True
    MBoxFechaI.SetFocus
 End If
End Sub

Private Sub Command2_Click()
  Unload BuscarContabil
End Sub

Private Sub DGQuery_DblClick()
    Co.Fecha = DGQuery.Columns(1).Text
    Co.TP = DGQuery.Columns(2).Text
    Co.Numero = DGQuery.Columns(3).Text
    Co.Beneficiario = DGQuery.Columns(4).Text
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
'  Keys_Especiales Shift
'  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto BuscarContabil, AdoQuery
End Sub

Private Sub DGQuery1_DblClick()
    Co.Fecha = DGQuery1.Columns(1).Text
    Co.TP = DGQuery1.Columns(2).Text
    Co.Numero = DGQuery1.Columns(3).Text
    Co.Beneficiario = DGQuery1.Columns(4).Text
End Sub

Private Sub Form_Activate()
   BuscarContabil.WindowState = vbMaximized
   BuscarContabil.Refresh
   DGQuery.width = BuscarContabil.width - DGQuery.Left - 350
   DGQuery.Height = (BuscarContabil.Height / 2) - DGQuery.Top - 100
   AdoQuery.Top = DGQuery.Top + DGQuery.Height
   AdoQuery.width = DGQuery.width
   CmbCampos.Height = (BuscarContabil.Height / 2) - DGQuery.Top + 270
   
   DGQuery1.Top = AdoQuery.Top + AdoQuery.Height
   DGQuery1.width = BuscarContabil.width - DGQuery1.Left - 380
   DGQuery1.Height = BuscarContabil.Height - AdoQuery.Top - 1300
    
   AdoQuery1.Top = DGQuery1.Top + DGQuery1.Height
   AdoQuery1.width = DGQuery1.width
   
   sSQL = "SELECT MIN(Fecha) AS MinFecha, MAX(Fecha) AS MaxFecha " _
       & "FROM Comprobantes " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
   Select_Adodc AdoListar, sSQL
   If AdoListar.Recordset.RecordCount > 0 Then
      MBoxFechaI = AdoListar.Recordset.fields("MinFecha")
      MBoxFechaF = AdoListar.Recordset.fields("MaxFecha")
   End If
  
   CmbCampos.Clear
   CmbCampos.AddItem "T"
   CmbCampos.AddItem "Fecha"
   CmbCampos.AddItem "TP"
   CmbCampos.AddItem "Numero"
   CmbCampos.AddItem "Factura"
   CmbCampos.AddItem "Cheq_Dep"
   CmbCampos.AddItem "Cliente"
   CmbCampos.AddItem "SubModulo"
   CmbCampos.AddItem "CodigoU"
   CmbCampos.AddItem "Concepto"
   CmbCampos.AddItem "Detalle_SubCta"
   CmbCampos.AddItem "Cotizacion"
   CmbCampos.AddItem "Cta"
   CmbCampos.AddItem "Debe"
   CmbCampos.AddItem "Haber"
   CmbCampos.AddItem "Debitos"
   CmbCampos.AddItem "Creditos"
   CmbCampos.AddItem "Parcial_ME"
   CmbCampos.Text = CmbCampos.List(0)
   CmbCampos.SetFocus
   RatonNormal
End Sub

Private Sub Form_Deactivate()
  BuscarContabil.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
 'CentrarForm BuscarContabil
  ConectarAdodc AdoQuery
  ConectarAdodc AdoQuery1
  ConectarAdodc AdoListar
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub TextPatron_GotFocus()
   MarcarTexto TextPatron
End Sub

Private Sub TextPatron_LostFocus()
  TextoValido TextPatron
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
      Case "Salir"
           Unload BuscarContabil
      Case "Consultar"
           Consultar
      Case "Excel"
           If AdoQuery.Recordset.RecordCount > 0 Then GenerarDataTexto BuscarContabil, AdoQuery
           If AdoQuery1.Recordset.RecordCount > 0 Then GenerarDataTexto BuscarContabil, AdoQuery1
      Case "Modificar"
           If ClaveContador Then
              FechaComp = Co.Fecha
              NumeroComp = Co.Numero
              Co.Item = NumEmpresa
              Mensajes = "Seguro que quiere Modificar" & vbCrLf & "El Comprobante: " & Co.TP & " No. " & Co.Numero & " Con Fecha: " & Co.Fecha & vbCrLf
              If Len(Co.Beneficiario) > 1 Then Mensajes = Mensajes & vbCrLf & "De " & Co.Beneficiario
              Titulo = "Pregunta de Modificacion"
              If BoxMensaje = vbYes Then
                 ModificarComp = True
                 CopiarComp = False
                 NuevoComp = False
                 Trans_No = 1
                 FechaComp = Co.Fecha
                 Unload BuscarContabil
                 Eliminar_Asientos_SP True
                 FComprobantes.Show
              End If
           End If
    End Select
End Sub

Public Sub Consultar()
Dim LenStrg As Byte
Dim CampoBuscar As String

  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  LenStrg = Len(TextPatron.Text)
 'Seteos de Transacciones
 '& "AND 1 = 0 "
  sSQL = "SELECT T.T,T.Fecha,T.TP,T.Numero,T.Cheq_Dep,Cl.Cliente,C.CodigoU,C.Concepto,C.Cotizacion,T.Cta,T.Debe,T.Haber,T.Parcial_ME " _
       & "FROM Comprobantes AS C,Transacciones As T,Clientes AS Cl " _
       & "WHERE C.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.TP = T.TP " _
       & "AND C.Fecha = T.Fecha " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND C.Periodo = T.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo "
  Select_Adodc AdoListar, sSQL
 
  Cadena = ""
  Select Case CmbCampos.Text
     Case "Cta", "TP", "Numero", "Cheq_Dep", "Fecha", "T", "Debe", "Haber", "Parcial_ME"
          Cadena = "T."
          CampoBuscar = CmbCampos.Text
     Case "CodigoU", "Concepto", "Cotizacion"
          Cadena = "C."
          CampoBuscar = CmbCampos.Text
     Case "Cliente"
          Cadena = "Cl."
          CampoBuscar = CmbCampos.Text
     Case Else
          Cadena = "T."
          CampoBuscar = "T"
  End Select
  
  sSQL = "SELECT T.T,T.Fecha,T.TP,T.Numero,Cl.Cliente,C.Concepto,T.Cheq_Dep,C.Cotizacion,T.Cta,T.Debe,T.Haber,T.Parcial_ME,C.CodigoU " _
       & "FROM Comprobantes AS C,Transacciones As T,Clientes AS Cl " _
       & "WHERE C.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' "
  With AdoListar.Recordset.fields(CampoBuscar)
    Select Case .Type
      Case TadBoolean
           sSQL = sSQL & "AND " & Cadena & CampoBuscar & " = " & TextPatron.Text & " "
      Case TadDate, TadDate1
           If IsDate(TextPatron.Text) And CheqPorFecha.value = 0 Then
              sSQL = sSQL & "AND " & Cadena & CampoBuscar & " = #" & BuscarFecha(TextPatron.Text) & "# "
           Else
              MsgBox "Formato de Fecha invalida"
           End If
      Case TadByte, TadInteger, TadLong
           sSQL = sSQL & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadSingle, TadCurrency
           sSQL = sSQL & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadText
           sSQL = sSQL & "AND " & Cadena & CampoBuscar & " Like '%" & TextPatron.Text & "%' "
      Case Else
           sSQL = sSQL & "AND 1 = 0 "
    End Select
  End With
  If CheqPorFecha.value = 1 Then sSQL = sSQL & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
  sSQL = sSQL _
       & "AND C.TP = T.TP " _
       & "AND C.Fecha = T.Fecha " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND C.Periodo = T.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "ORDER BY T.Fecha,T.TP,T.Numero,Cl.Cliente "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL
  
 'Seteos de Trans_SubCtas
 '& "AND 1 = 0 "
  SQL2 = "SELECT T.T,T.Fecha,T.TP,T.Numero,B.Detalle,T.Cta,T.Detalle_SubCta,T.Factura,T.Debitos,T.Creditos,T.Parcial_ME,T.CodigoU,T.Item " _
       & "FROM Catalogo_SubCtas As B, Trans_SubCtas As T " _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND B.Periodo = T.Periodo " _
       & "AND B.Item = T.Item " _
       & "AND B.Codigo = T.Codigo "
  Select_Adodc AdoListar, SQL2
  
  Cadena = ""
  Select Case CmbCampos.Text
     Case "Cta", "TP", "Numero", "Factura", "Fecha", "T", "Detalle_SubCta", "Debitos", "Creditos", "Parcial_ME", "CodigoU"
          Cadena = "T."
          CampoBuscar = CmbCampos.Text
     Case "SubModulo"
          Cadena = "B."
          CampoBuscar = "Detalle"
     Case Else
          Cadena = "T."
          CampoBuscar = "T"
  End Select
  
  '----------------------------------------------------------------------------------------------
  SQL1 = "SELECT T.T,T.Fecha,T.TP,T.Numero,B.Detalle As SubModulo,T.Cta,T.Detalle_SubCta,T.Factura,T.Debitos,T.Creditos,T.Parcial_ME,T.CodigoU,T.Item " _
       & "FROM Catalogo_SubCtas As B, Trans_SubCtas As T " _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  With AdoListar.Recordset.fields(CampoBuscar)
    Select Case .Type
      Case TadBoolean
           SQL1 = SQL1 & "AND " & Cadena & CampoBuscar & " = " & TextPatron.Text & " "
      Case TadDate, TadDate1
           If IsDate(TextPatron.Text) And CheqPorFecha.value = 0 Then
              SQL1 = SQL1 & "AND " & Cadena & CampoBuscar & " = #" & BuscarFecha(TextPatron.Text) & "# "
           Else
              MsgBox "Formato de Fecha invalida"
           End If
      Case TadByte, TadInteger, TadLong
           SQL1 = SQL1 & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadSingle, TadCurrency
           SQL1 = SQL1 & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadText
           SQL1 = SQL1 & "AND " & Cadena & CampoBuscar & " Like '%" & TextPatron.Text & "%' "
      Case Else
           MsgBox "SubModulo"
           SQL1 = SQL1 & "AND 1 = 0 "
    End Select
  End With
  If CheqPorFecha.value = 1 Then SQL1 = SQL1 & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
     
  SQL1 = SQL1 _
       & "AND B.Periodo = T.Periodo " _
       & "AND B.Item = T.Item " _
       & "AND B.Codigo = T.Codigo "
  
 'Seteos de Trans_SubCtas
  SQL2 = "SELECT T.T,T.Fecha,T.TP,T.Numero, C.Cliente,T.Cta,T.Detalle_SubCta,T.Factura,T.Debitos,T.Creditos,T.Parcial_ME,T.CodigoU,T.Item " _
       & "FROM Catalogo_CxCxP As B, Trans_SubCtas As T, Clientes As C " _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND 1 = 0 " _
       & "AND B.Periodo = T.Periodo " _
       & "AND B.Item = T.Item " _
       & "AND C.Codigo = T.Codigo "
  Select_Adodc AdoListar, SQL2
 
 'Trans_SubCtas de: CxC y CxP
  Cadena = ""
  Select Case CmbCampos.Text
     Case "Cta", "TP", "Numero", "Factura", "Fecha", "T", "Detalle_SubCta", "Debitos", "Creditos", "Parcial_ME", "CodigoU"
          Cadena = "T."
          CampoBuscar = CmbCampos.Text
     Case "SubModulo"
          Cadena = "C."
          CampoBuscar = "Cliente"
     Case Else
          Cadena = "T."
          CampoBuscar = "T"
  End Select
  
  SQL2 = "SELECT T.T,T.Fecha,T.TP,T.Numero, C.Cliente As SubModulo,T.Cta,T.Detalle_SubCta,T.Factura,T.Debitos,T.Creditos,T.Parcial_ME,T.CodigoU,T.Item " _
       & "FROM Catalogo_CxCxP As B, Trans_SubCtas As T, Clientes As C " _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  With AdoListar.Recordset.fields(CampoBuscar)
    Select Case .Type
      Case TadBoolean
           SQL2 = SQL2 & "AND " & Cadena & CampoBuscar & " = " & TextPatron.Text & " "
      Case TadDate, TadDate1
           If IsDate(TextPatron.Text) And CheqPorFecha.value = 0 Then
              SQL2 = SQL2 & "AND " & Cadena & CampoBuscar & " = #" & BuscarFecha(TextPatron.Text) & "# "
           Else
              MsgBox "Formato de Fecha invalida"
           End If
      Case TadByte, TadInteger, TadLong
           SQL2 = SQL2 & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadSingle, TadCurrency
           SQL2 = SQL2 & "AND " & Cadena & CampoBuscar & " = " & Val(TextPatron.Text) & " "
      Case TadText
           SQL2 = SQL2 & "AND " & Cadena & CampoBuscar & " Like '%" & TextPatron.Text & "%' "
      Case Else
           SQL2 = SQL2 & "AND 1 = 0 "
    End Select
  End With
  If CheqPorFecha.value = 1 Then SQL2 = SQL2 & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
     
  SQL2 = SQL2 _
       & "AND B.Periodo = T.Periodo " _
       & "AND B.Item = T.Item " _
       & "AND B.Codigo = T.Codigo " _
       & "AND B.Codigo = C.Codigo "
  SQL1 = SQL1 & " UNION " & SQL2 & "ORDER BY T.Fecha,T.TP,T.Numero "
  Select_Adodc_Grid DGQuery1, AdoQuery1, SQL1
  CmbCampos.SetFocus
End Sub
