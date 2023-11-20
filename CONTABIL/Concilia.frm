VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form Conciliacion 
   Caption         =   "CONCILIACION DE BANCOS DE DEPOSITOS Y RETIROS"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   14325
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Procesar"
            Object.ToolTipText     =   "Procesar Conciliacion"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Transito"
            Object.ToolTipText     =   "Procesar Debitos/Creditos en transito"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Insertar"
            Object.ToolTipText     =   "Ingresar Transito"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Transito"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Resultados"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.Frame Frame3 
         Caption         =   "&Fechas Desde - Hasta"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   0
         Width           =   12405
         Begin MSMask.MaskEdBox MBoxFechaI 
            Height          =   330
            Left            =   105
            TabIndex        =   2
            Top             =   210
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
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   1575
            TabIndex        =   3
            Top             =   210
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
         Begin MSDataListLib.DataCombo DCCtas 
            Bindings        =   "Concilia.frx":0000
            DataSource      =   "AdoBanco"
            Height          =   315
            Left            =   4725
            TabIndex        =   5
            Top             =   210
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
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
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta &Bancaria:"
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
            TabIndex        =   4
            Top             =   210
            Width           =   1695
         End
      End
   End
   Begin VB.Frame FrmInsTransaccion 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CREAR UNA TRANSACCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6420
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
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
         Left            =   5250
         Picture         =   "Concilia.frx":0017
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
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
         Left            =   5250
         Picture         =   "Concilia.frx":0321
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1155
         Width           =   1065
      End
      Begin VB.TextBox TxtHaber 
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
         Left            =   1995
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1995
         Width           =   1695
      End
      Begin VB.TextBox TxtDebe 
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
         MaxLength       =   12
         TabIndex        =   17
         Top             =   1995
         Width           =   1695
      End
      Begin VB.TextBox TxtConcepto 
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
         Left            =   1785
         MaxLength       =   60
         TabIndex        =   15
         Top             =   1155
         Width           =   3375
      End
      Begin VB.TextBox TxtDocumento 
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
         Left            =   3465
         MaxLength       =   16
         TabIndex        =   13
         Top             =   735
         Width           =   1695
      End
      Begin VB.TextBox TxtTP 
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
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   11
         Top             =   735
         Width           =   435
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   1785
         TabIndex        =   9
         Top             =   315
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
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Credto"
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
         Left            =   1995
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Debito"
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
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Detalle"
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
         TabIndex        =   14
         Top             =   1155
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Documento"
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
         Left            =   2310
         TabIndex        =   12
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
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
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Documento"
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
         TabIndex        =   10
         Top             =   735
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DGBalance 
      Bindings        =   "Concilia.frx":062B
      Height          =   5790
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10213
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
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
      Left            =   420
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   420
      Top             =   2730
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
      Top             =   6615
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11865
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":0645
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":095F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":0C79
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":0F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":12AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":15C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Concilia.frx":18E1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Conciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Imprimir()
  DGBalance.Visible = False
  sSQL = "SELECT C,FECHA,BENEFICIARIO,TP,NUMERO,CHEQ_DEP,DEBE,HABER " _
       & "FROM Asiento_C " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY FECHA,TP,NUMERO,DEBE DESC,HABER "
  Select_Adodc AdoAsientos, sSQL
  
  MensajeEncabado = "LISTADO DE CONCILICIACION BANCARIA DE TRANSACCIONES"
  ImprimirAdodc AdoAsientos, 1, 8
  
  sSQL = "SELECT * " _
       & "FROM Asiento_C " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY FECHA,TP,NUMERO,DEBE DESC,HABER "
  Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
  DGBalance.Visible = True
End Sub

Private Sub Procesar()
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  RatonReloj
  Contador = 0
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  sSQL = "DELETE * " _
       & "FROM Asiento_C " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Asiento_C " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  sSQL = "SELECT C,Cliente As Beneficiario,T.Fecha,T.TP,T.Numero,Cheq_Dep,Debe,Haber,Parcial_ME,T.Item,T.ID " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes AS Cl " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.Cta = '" & Codigo1 & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "UNION " _
       & "SELECT C,Cliente As Beneficiario,T.Fecha,T.TP,T.Numero,Cheq_Dep,Debe,Haber,Parcial_ME,T.Item,T.ID " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes AS Cl " _
       & "WHERE T.Fecha < #" & FechaIni & "# " _
       & "AND T.Cta = '" & Codigo1 & "' " _
       & "AND C = " & Val(adFalse) & " "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "ORDER BY T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
  Select_Adodc AdoCtas, sSQL
  RatonReloj
  DGBalance.Visible = False
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
         Contador = Contador + 1
         Conciliacion.Caption = "Conciliando la fecha: " & .fields("Fecha")
         SetAddNew AdoAsientos
         SetFields AdoAsientos, "C", .fields("C")
         SetFields AdoAsientos, "FECHA", .fields("Fecha")
         SetFields AdoAsientos, "BENEFICIARIO", .fields("Beneficiario")
         SetFields AdoAsientos, "TP", .fields("TP")
         SetFields AdoAsientos, "NUMERO", .fields("Numero")
         SetFields AdoAsientos, "CHEQ_DEP", .fields("Cheq_Dep")
         SetFields AdoAsientos, "DEBE", .fields("Debe")
         SetFields AdoAsientos, "HABER", .fields("Haber")
         SetFields AdoAsientos, "ME", Moneda_US
         SetFields AdoAsientos, "CodigoU", CodigoUsuario
         SetFields AdoAsientos, "Item", .fields("Item")
         SetFields AdoAsientos, "T_No", Trans_No
         SetFields AdoAsientos, "IdTrans", Format(Contador, "0000")
         SetUpdate AdoAsientos
        .MoveNext
       Loop
   End If
  End With
  RatonNormal
  sSQL = "SELECT * " _
       & "FROM Asiento_C " _
       & "WHERE CodigoU='" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " " _
       & "ORDER BY FECHA DESC,TP,NUMERO,DEBE DESC,HABER "
  Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
  SumaDebe = 0: SumaHaber = 0
  DGBalance.Visible = True
  Cadena = "Registros: " & Format(AdoCtas.Recordset.RecordCount, "#,##0") & ".   Páginas: " _
         & Format((AdoCtas.Recordset.RecordCount / 45) + 1, "#,##0") & "."
  Conciliacion.Caption = "CONCILIACION DE BANCOS"
  AdoCtas.Caption = Cadena
  Opcion = 1
End Sub

Private Sub Transito()
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  RatonReloj
  Contador = 0
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  DGBalance.Visible = False
  sSQL = "SELECT C, Cta, Fecha, TP, Numero, Concepto, Documento, Debe, Haber, ID " _
       & "FROM Trans_Transito " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo1 & "' " _
       & "ORDER BY Fecha, TP, Numero "
  Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
  Cadena = "Registros: " & Format(AdoAsientos.Recordset.RecordCount, "#,##0") & ".   Páginas: " _
         & Format((AdoAsientos.Recordset.RecordCount / 45) + 1, "#,##0") & "."
  Conciliacion.Caption = "CONCILIACION DE BANCOS DEBITOS/CREDITOS EN TRANSITO"
  AdoAsientos.Caption = Cadena
  DGBalance.Visible = True
  Opcion = 2
  RatonNormal
End Sub

Private Sub Insertar_Transito()
  FrmInsTransaccion.Visible = True
  MBFecha.SetFocus
End Sub

Private Sub Borrar_Transito()
   If Opcion = 2 And ID_Trans <> 0 Then
      Titulo = "CONFIRMACION DE ELIMINACION DE TRANSACCION"
      Mensajes = "Eliminar esta transaccion"
      If BoxMensaje = vbYes Then
         sSQL = "DELETE * " _
              & "FROM Trans_Transito " _
              & "WHERE Periodo = '" & Periodo_Contable & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND ID = " & ID_Trans & " " _
              & "AND Cta = '" & Codigo1 & "' "
         Ejecutar_SQL_SP sSQL
        'MsgBox sSQL
         Transito
      End If
   End If
End Sub

Private Sub Grabar()
  RatonReloj
  DGBalance.Visible = False
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       If SQL_Server Then
          sSQL = "UPDATE Transacciones " _
               & "SET C = AC.C " _
               & "FROM Transacciones As T,Asiento_C As AC "
       Else
          sSQL = "UPDATE Transacciones As T,Asiento_C As AC " _
               & "SET T.C = AC.C "
       End If
       sSQL = sSQL & "WHERE AC.TP = T.TP " _
            & "AND AC.Numero = T.Numero " _
            & "AND AC.DEBE = T.Debe " _
            & "AND AC.HABER = T.Haber " _
            & "AND AC.FECHA = T.Fecha " _
            & "AND AC.CHEQ_DEP = T.Cheq_Dep " _
            & "AND AC.Item = T.Item " _
            & "AND AC.CodigoU = '" & CodigoUsuario & "' " _
            & "AND AC.T_No = " & Trans_No & " " _
            & "AND T.Periodo = '" & Periodo_Contable & "' " _
            & "AND T.Cta = '" & Codigo1 & "' "
       Ejecutar_SQL_SP sSQL
       MsgBox "Proceso Grabado"
   End If
  End With
  SumaDebe = 0: SumaHaber = 0
  DGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command1_Click()
    Titulo = "CONFIRMACION DE INSERCION DE TRANSACCION"
    Mensajes = "Insertar esta transaccion"
    If BoxMensaje = vbYes Then
       SetAdoAddNew "Trans_Transito"
       SetAdoFields "Item", NumEmpresa
       SetAdoFields "Periodo", Periodo_Contable
       SetAdoFields "Cta", Codigo1
       SetAdoFields "Fecha", MBFecha
       SetAdoFields "TP", TxtTP
       SetAdoFields "Concepto", TxtConcepto
       SetAdoFields "Documento", TxtDocumento
       SetAdoFields "Debe", Val(TxtDebe)
       SetAdoFields "Haber", Val(TxtHaber)
       SetAdoUpdate
       Transito
       MsgBox "Transaccion Insertada con exito"
    End If
    FrmInsTransaccion.Visible = False
End Sub

Private Sub Command2_Click()
  FrmInsTransaccion.Visible = False
End Sub

Private Sub DCCtas_LostFocus()
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  Codigo = Leer_Cta_Catalogo(Codigo1)
End Sub

Private Sub DGBalance_Click()
   If Opcion = 2 Then ID_Trans = DGBalance.Columns(9)
End Sub

Private Sub DGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto Conciliacion, AdoAsientos
End Sub

Private Sub DGBalance_KeyPress(KeyAscii As Integer)
   If Opcion = 1 Then
      Codigo1 = DGBalance.Columns(13)
      Select Case Chr(KeyAscii)
        Case "s", "S", "y", "Y": ' Si
             NivelCta = 1
        Case "n", "N"  ' No
             NivelCta = 0
      End Select
      Select Case Chr(KeyAscii)
        Case "s", "S", "y", "Y", "n", "N":
             If AdoAsientos.Recordset.RecordCount > 0 Then
                sSQL = "UPDATE Asiento_C " _
                     & "SET C = " & NivelCta & " " _
                     & "WHERE IdTrans = '" & Codigo1 & "' "
                If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                sSQL = sSQL & "AND CodigoU = '" & CodigoUsuario & "' " _
                     & "AND T_No = " & Trans_No & " "
                Ejecutar_SQL_SP sSQL
                sSQL = "SELECT * " _
                     & "FROM Asiento_C " _
                     & "WHERE CodigoU = '" & CodigoUsuario & "' "
                If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                sSQL = sSQL & "AND T_No = " & Trans_No & " " _
                     & "ORDER BY IdTrans "
                Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
                AdoAsientos.Recordset.MoveFirst
                AdoAsientos.Recordset.Find ("IdTrans = '" & Codigo1 & "' ")
                DGBalance.SetFocus
             End If
      End Select
   End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT (Codigo & Space(20) & Cuenta) As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoBanco, sSQL, "Nombre_Cta", False
  DGBalance.Visible = False
  Opcion = 0
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoAsientos
  Trans_No = 30
  DGBalance.Height = MDI_Y_Max - DGBalance.Top - 300
  DGBalance.width = MDI_X_Max - DGBalance.Left - 100
  AdoAsientos.Top = DGBalance.Top + DGBalance.Height + 10
  AdoAsientos.width = MDI_X_Max - AdoAsientos.Left - 100
End Sub

Private Sub MBFecha_GotFocus()
   MarcarTexto MBFecha
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  MBoxFechaF = UltimoDiaMes(MBoxFechaI)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir":  Unload Conciliacion
      Case "Procesar": Procesar
      Case "Transito": Transito
      Case "Insertar": Insertar_Transito
      Case "Borrar": Borrar_Transito
      Case "Grabar": Grabar
      Case "Imprimir": Imprimir
    End Select
End Sub

Private Sub TxtConcepto_GotFocus()
  MarcarTexto TxtConcepto
End Sub

Private Sub TxtConcepto_LostFocus()
   TxtConcepto = UCase(TxtConcepto)
End Sub

Private Sub TxtDebe_GotFocus()
  MarcarTexto TxtDebe
End Sub

Private Sub TxtDebe_LostFocus()
   TextoValido TxtDebe, True, , 2
End Sub

Private Sub TxtDocumento_GotFocus()
  MarcarTexto TxtDocumento
End Sub

Private Sub TxtDocumento_LostFocus()
  TxtDocumento = UCase(TxtDocumento)
End Sub

Private Sub TxtHaber_GotFocus()
  MarcarTexto TxtHaber
End Sub

Private Sub TxtHaber_LostFocus()
  TextoValido TxtHaber, True, , 2
End Sub

Private Sub TxtTP_GotFocus()
  MarcarTexto TxtTP
End Sub

Private Sub TxtTP_LostFocus()
  TxtTP = UCase(TxtTP)
End Sub
