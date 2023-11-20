VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EstadoResultMes 
   Caption         =   "ESTADO DE RESULTADOS POR MES"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictSubCtas 
      Height          =   330
      Left            =   105
      ScaleHeight     =   270
      ScaleWidth      =   11505
      TabIndex        =   22
      Top             =   6930
      Width           =   11565
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
      Left            =   10500
      Picture         =   "EstResuM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3150
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Estado de Resultado"
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
      Left            =   10500
      Picture         =   "EstResuM.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   105
      TabIndex        =   17
      Top             =   840
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10583
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ESTADO DE RESULTADO MENSUAL"
      TabPicture(0)   =   "EstResuM.frx":0D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "AdoResultado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DGBalance"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ESTADO DE RESULTADO MENSUAL POR SEMANA"
      TabPicture(1)   =   "EstResuM.frx":0D28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGSemana"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "AdoSemana"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DGSemana 
         Bindings        =   "EstResuM.frx":0D44
         Height          =   5160
         Left            =   -74895
         TabIndex        =   21
         Top             =   420
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   9102
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
      Begin MSDataGridLib.DataGrid DGBalance 
         Bindings        =   "EstResuM.frx":0D5C
         Height          =   5160
         Left            =   105
         TabIndex        =   20
         Top             =   420
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   9102
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
      Begin VB.CommandButton Command1 
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
         Left            =   -64605
         Picture         =   "EstResuM.frx":0D77
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1365
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
         Left            =   10395
         Picture         =   "EstResuM.frx":1081
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1365
         Width           =   1065
      End
      Begin MSAdodcLib.Adodc AdoResultado 
         Height          =   330
         Left            =   105
         Top             =   5565
         Width           =   10200
         _ExtentX        =   17992
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
      Begin MSAdodcLib.Adodc AdoSemana 
         Height          =   330
         Left            =   -74895
         Top             =   5565
         Width           =   10200
         _ExtentX        =   17992
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
         Caption         =   "Semana"
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
   Begin VB.Frame Frame1 
      Caption         =   "TIPO DE PRESENTACION"
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
      Left            =   2205
      TabIndex        =   13
      Top             =   0
      Width           =   8100
      Begin VB.OptionButton OpcD 
         Caption         =   "Solo Cuentas de Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5355
         TabIndex        =   16
         Top             =   210
         Width           =   2430
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "Solo Cuentas de Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2940
         TabIndex        =   15
         Top             =   210
         Width           =   2325
      End
      Begin VB.OptionButton OpcDG 
         Caption         =   "Cuentas de Grupo y Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   2745
      End
   End
   Begin VB.TextBox TextCotiza 
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
      Left            =   10395
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   420
      Width           =   1170
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   420
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   210
      Top             =   1365
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoFechaBal 
      Height          =   330
      Left            =   210
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
      Caption         =   "FechaBal"
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
   Begin VB.Label LabelHaber 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8400
      TabIndex        =   6
      Top             =   7350
      Width           =   1905
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MN"
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
      Left            =   7770
      TabIndex        =   9
      Top             =   7350
      Width           =   645
   End
   Begin VB.Label LabelDebe 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5985
      TabIndex        =   7
      Top             =   7350
      Width           =   1800
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ME"
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
      Left            =   5355
      TabIndex        =   8
      Top             =   7350
      Width           =   645
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cotizacion:"
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
      Left            =   10395
      TabIndex        =   12
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Utilidad/Perdida"
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
      Left            =   3780
      TabIndex        =   10
      Top             =   7350
      Width           =   1590
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
      TabIndex        =   2
      Top             =   420
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
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "EstadoResultMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Proceso_de_Barras As Progreso_Barras

Private Sub Command1_Click()
  MensajeEncabData = DGSemana.Caption
  ImprimirResultSemana AdoSemana, True, 1, 8
End Sub

Private Sub Command2_Click()
  DGBalance.Visible = False
  sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TB = 'MER' "
  If OpcG.Value Then sSQL = sSQL & "AND DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Ln "
  SelectDataGrid DGBalance, AdoResultado, sSQL
  DGBalance.Visible = False
  SQLMsg1 = "ESTADO DE RESULTADOS MENSUALES"
  SQLMsg2 = "DESDE EL: " & MBoxFechaI.Text & "  AL " & MBoxFechaF.Text
  If OpcCoop Then
     Imprimir_General_Con AdoResultado, 1, False
  Else
     Imprimir_General AdoResultado, 1
  End If
  DGBalance.Visible = True
End Sub

Private Sub Command4_Click()
  Unload EstadoResultMes
End Sub

Private Sub Command5_Click()
Dim Domingos(4) As Integer
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
  DGBalance.Visible = False
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  Control_Procesos Normal, "Balance Mensual del " & MesesLetras(Month(MBoxFechaI.Text))
  sSQL = "DELETE * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumItemTemp & "' " _
       & "AND Mid(TB,1,1) = 'M' "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumItemTemp & "' " _
       & "AND Mid(TB,1,1) = 'M' "
  SelectAdodc AdoResultado, sSQL
  'MsgBox "..."
  ProcesarBalanceMes EstadoResultMes, MBoxFechaI, MBoxFechaF, AdoCtas, AdoTrans, AdoSemana, False
  'MsgBox "......"
  'sSQL = "SELECT TC,DG,Codigo,Cuenta, Saldo_Anterior, Debitos, Creditos, Saldo_Total, Saldo_Total_ME "
  

       
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       'SetProgBar ProgBar, .RecordCount
      .MoveFirst
       TotalActivo = 0
       TotalPasivo = 0
       TotalCapital = 0
       Cta = Mid(.Fields("Codigo"), 1, 1)
       TipoDoc = .Fields("TC")
       Do While Not .EOF
          If Cta <> Mid(.Fields("Codigo"), 1, 1) Then
             Select Case Cta
               Case "1"
                    InsertarTotales "MES", " ", "G", " - TOTAL ACTIVO.", TotalActivo, TotalActivo_ME
               Case "2"
                    InsertarTotales "MES", " ", "G", " - TOTAL PASIVO", TotalPasivo, TotalPasivo_ME
               Case "3"
                    Saldo = TotalActivo - TotalPasivo - TotalCapital
                    Saldo_ME = TotalActivo_ME - TotalPasivo_ME - TotalCapital_ME
                    InsertarTotales "MES", " ", "G", " - TOTAL PATRIMONIO", TotalCapital, TotalCapital_ME
                    InsertarTotales "MES", " ", "G", " - TOTAL PASIVO", TotalPasivo, TotalPasivo_ME
                    InsertarTotales "MES", "+/-", "G", " - UTILIDAD(Pérdida) DEL PERIODO", Saldo, Saldo_ME
                    InsertarTotales "MES", "+", "G", " - TOTAL PASIVO Y PATRIMONIO.", TotalPasivo + TotalCapital, TotalPasivo_ME + TotalCapital_ME
               Case "4"
                    If OpcCoop Then
                       InsertarTotales "MER", " ", "G", " - TOTAL EGRESO", TotalIngreso, TotalIngreso_ME
                    Else
                       InsertarTotales "MER", " ", "G", " - TOTAL INGRESO", TotalIngreso, TotalIngreso_ME
                    End If
             End Select
             Cta = Mid(.Fields("Codigo"), 1, 1)
          End If
          TipoDoc = .Fields("TC")
          TipoCta = .Fields("DG")
          Codigo = .Fields("Codigo")
          Cuenta = .Fields("Cuenta")
          Total = .Fields("Saldo_Total")
          Total_ME = .Fields("Saldo_Total_ME")
          Select Case Codigo
            Case "1": TotalActivo = Total: TotalActivo_ME = Total_ME
            Case "2": TotalPasivo = Total: TotalPasivo_ME = Total_ME
            Case "3": TotalCapital = Total: TotalCapital_ME = Total_ME
            Case "4": TotalIngreso = Total: TotalIngreso_ME = Total_ME
            Case "5": TotalEgreso = Total: TotalEgreso_ME = Total_ME
         End Select
         If Len(Codigo) = 1 Then
            Total = 0: Total_ME = 0
         End If
         Select Case Mid(Codigo, 1, 1)
           Case "1", "2", "3": InsertarTotales "MES", Codigo, TipoCta, Cuenta, Total, Total_ME
           Case "4", "5":      InsertarTotales "MER", Codigo, TipoCta, Cuenta, Total, Total_ME
         End Select
         'IncProgBar ProgBar
        .MoveNext
      Loop
  End If
  End With
  If OpcCoop Then
     InsertarTotales "MER", " ", "G", " - TOTAL CUENTAS DE RESULTADOS ACREEDORAS", TotalEgreso, TotalEgreso_ME
     Saldo = TotalEgreso - TotalIngreso
     Saldo_ME = TotalEgreso_ME - TotalIngreso_ME
  Else
     InsertarTotales "MER", " ", "G", " - TOTAL EGRESOS", TotalEgreso, TotalEgreso_ME
     Saldo = TotalIngreso - TotalEgreso
     Saldo_ME = TotalIngreso_ME - TotalEgreso_ME
  End If
  InsertarTotales "MER", "+/-", "G", " - UTILIDAD(Pérdida) DEL PERIODO.", Saldo, Saldo_ME
  DGBalance.Visible = False
  TextCotiza.Text = Format(Dolar, "#,##0.00")
  Fecha_Procesos "Balance Mes", MBoxFechaI, MBoxFechaF
' Actualizamos las cuentas Principales
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET Debitos = 0.00,Creditos = 0.00 " _
       & "WHERE DG = 'G' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
' Listar Balance de Comprobacion
  TextCotiza.Text = Format(Dolar, "#,##0.00")
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Detalle = 'Balance Mes' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCtas, sSQL
  If AdoCtas.Recordset.RecordCount > 0 Then
     MBoxFechaI.Text = AdoCtas.Recordset.Fields("Fecha_Inicial")
     MBoxFechaF.Text = AdoCtas.Recordset.Fields("Fecha_Final")
  End If
  sSQL = "DELETE * " _
       & "FROM Balances_Mes " _
       & "WHERE (Total_N1 + Total_N2 + Total_N3 + Total_N4 + Total_N5 + Total_N6) = 0 " _
       & "AND Len(Codigo)>1 "
  ConectarAdoExecute sSQL
  
  DGBalance.Caption = "BALANCE DE COMPROBACION: " & FechaStrgCorta(MBoxFechaI.Text) & "  -  " & FechaStrgCorta(MBoxFechaF.Text)
  sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TB = 'MER' "
  If OpcG.Value Then sSQL = sSQL & "AND DG = 'G' "
  If OpcD.Value Then sSQL = sSQL & "AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Ln "
  SelectDataGrid DGBalance, AdoResultado, sSQL
  DGBalance.Visible = False
  SumaDebe = 0: SumaHaber = 0
  DGBalance.Visible = True
  'Diferencia = SumaDebe - SumaHaber
  'MsgBox Diferencia
  'LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  'LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  'LabelTotSaldo.Caption = Format(Diferencia, "#,##0.00")
  For I = 0 To 4: Domingos(I) = 0: Next I
  J = 4
  Mifecha = MBoxFechaF.Text
  For I = Day(MBoxFechaF.Text) To Day(MBoxFechaI.Text) Step -1
      If Weekday(Mifecha) = 1 Then
         Domingos(J) = I
         J = J - 1
      End If
      Mifecha = CLongFecha(CFechaLong(Mifecha) - 1)
  Next I
  'ProgBar.Value = ProgBar.Max
  EstadoResultMes.Caption = "ESTADO DE RESULTADO MENSUAL"
  RatonNormal
  sSQL = "SELECT Codigo,Cuenta " _
       & ",Semana_01 As [Del 1 al " & Domingos(1) - 1 & "] " _
       & ",Semana_02 As [Del " & Domingos(1) & " al " & Domingos(2) - 1 & "] " _
       & ",Semana_03 As [Del " & Domingos(2) & " al " & Domingos(3) - 1 & "] " _
       & ",Semana_04 As [Del " & Domingos(3) & " al " & Domingos(4) - 1 & "] " _
       & ",Semana_05 As [Del " & Domingos(4) & " al " & Day(MBoxFechaF.Text) & "] " _
       & ",TOTAL " _
       & "FROM Balances_Mes " _
       & "ORDER BY Codigo "
  'MsgBox sSQL
  SelectDataGrid DGSemana, AdoSemana, sSQL
  DGBalance.Caption = "ESTADO DE RESULTADO DEL MES DE " & UCase(MesesLetras(Month(MBoxFechaI.Text)))
  DGSemana.Caption = "ESTADO DE RESULTADO DEL MES DE " & UCase(MesesLetras(Month(MBoxFechaI.Text))) & " POR SEMANAS"
  RatonNormal
End Sub

Private Sub DGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto EstadoResultMes, AdoResultado
End Sub

Private Sub Form_Activate()
  If CNivel_7 Then
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload EstadoResult12Meses
     RatonNormal
  Else
     Label11.Visible = OpcCoop
     TextCotiza.Visible = OpcCoop
     RatonNormal
  End If
End Sub

Private Sub Form_Load()
  CentrarForm EstadoResultMes
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoSemana
  ConectarAdodc AdoFechaBal
  ConectarAdodc AdoResultado
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  MBoxFechaF = UltimoDiaMes(MBoxFechaI)
End Sub

Private Sub TextCotiza_GotFocus()
  TextCotiza.Text = Dolar
End Sub

Private Sub TextCotiza_LostFocus()
  TextoValido TextCotiza, True
  If Val(TextCotiza.Text) = 0 Then Dolar = Val(TextCotiza.Text)
End Sub
