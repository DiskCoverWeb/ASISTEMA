VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Historial 
   Caption         =   "LISTADO DE SALDOS"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   13920
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar ToolbarHistorial 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Libreta"
            Object.ToolTipText     =   "Imprimir Libreta"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MovimientoLibreta"
            Object.ToolTipText     =   "Movimiento de Libreta"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Mov_x_Tipo_Proc"
            Object.ToolTipText     =   "Movimiento por tipo de Proceso"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cuentas Abiertas"
            Object.ToolTipText     =   "Listar Cuentas Abiertas"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Depositos Excedidos"
            Object.ToolTipText     =   "Depositos Excedidos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificados"
            Object.ToolTipText     =   "Listar Certificados"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Saldos Libretas"
            Object.ToolTipText     =   "Presenta Saldo de Libretas"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Saldos Real Libretas"
            Object.ToolTipText     =   "Saldos Real Libretas con transacciones"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGCuentas 
      Bindings        =   "Historia.frx":0000
      Height          =   2535
      Left            =   105
      TabIndex        =   13
      Top             =   1575
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   4471
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   10185
      TabIndex        =   17
      Top             =   1575
      Width           =   330
   End
   Begin VB.TextBox TextLinea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "0"
      Top             =   1155
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3990
      TabIndex        =   3
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   735
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
      Left            =   1365
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   735
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   525
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
      Caption         =   "TP"
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
   Begin MSMask.MaskEdBox MBCuenta 
      Height          =   330
      Left            =   6930
      TabIndex        =   5
      ToolTipText     =   "Ingrese el número de Cuenta"
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
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "Historia.frx":0019
      DataSource      =   "AdoTP"
      Height          =   360
      Left            =   6510
      TabIndex        =   10
      Top             =   1155
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   105
      Top             =   5040
      Width           =   9045
      _ExtentX        =   15954
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
   Begin MSDataGridLib.DataGrid DGCartilla 
      Bindings        =   "Historia.frx":002D
      Height          =   2010
      Left            =   105
      TabIndex        =   16
      Top             =   5775
      Visible         =   0   'False
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   3545
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
      Caption         =   "HISTORIAL DE CARTILLAS EMITIDAS"
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
   Begin MSAdodcLib.Adodc AdoCartilla 
      Height          =   330
      Left            =   105
      Top             =   7770
      Visible         =   0   'False
      Width           =   11985
      _ExtentX        =   21140
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
   Begin VB.Label LblPromedio 
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
      Left            =   10290
      TabIndex        =   14
      Top             =   5025
      Width           =   1800
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROMEDIO"
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
      Left            =   9135
      TabIndex        =   15
      Top             =   5025
      Width           =   1170
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &POR TIPO DE TRANSACCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2730
      TabIndex        =   11
      Top             =   1155
      Width           =   3795
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Linea No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label LabelSocio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9345
      TabIndex        =   6
      Top             =   735
      Width           =   5790
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SOCIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   7
      Top             =   735
      Width           =   960
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Cuenta No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5355
      TabIndex        =   4
      Top             =   735
      Width           =   1590
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":0047
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":0361
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":067B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":0995
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":0CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":0FC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":12E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":15FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":17D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Historia.frx":1AF1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   1275
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
      Left            =   2730
      TabIndex        =   2
      Top             =   735
      Width           =   1275
   End
End
Attribute VB_Name = "Historial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Movimiento_x_Tipo_de_Proceso()
   Presentar_Malla False
   DGCuentas.Caption = "MOVIMIENTO POR TIPO DE PROCESO"
   sSQL = "SELECT C.Cliente,TL.ME,TL.Fecha,TL.TP,TL.Cuenta_No,TL.Debitos,TL.Creditos,TL.CodigoU,C.Fecha_N,C.Fecha As Fecha_Apert " _
        & "FROM Trans_Libretas As TL, Clientes_Datos_Extras As CL, Clientes As C " _
        & "WHERE TL.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND CL.Tipo_Dato = 'LIBRETAS' " _
        & "AND TL.TP = '" & DCTP.Text & "' " _
        & "AND C.Codigo = CL.Codigo " _
        & "AND CL.Cuenta_No = TL.Cuenta_No " _
        & "ORDER BY TL.ME,TL.Cuenta_No,TL.Fecha,TL.TP "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   RatonReloj
   DGCuentas.Visible = False
   Debe = 0: Haber = 0
   With AdoCuentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Debe = Debe + .Fields("Debitos")
           Haber = Haber + .Fields("Creditos")
          .MoveNext
        Loop
    End If
   End With
   DGCuentas.Visible = True
   RatonNormal
   Opcion = 3
   LblPromedio.Caption = Format(Abs(Debe - Haber), "#,##0.00")
End Sub

Public Sub Cuentas_Abiertas()
   Presentar_Malla False
   DGCuentas.Caption = "CUENTAS ABIERTAS"
   sSQL = "SELECT T.ME,T.Fecha,T.TP,T.Cuenta_No,Cl.Cliente,Cl.CI_RUC " _
        & "FROM Trans_Libretas As T,Clientes_Datos_Extras As C,Clientes As Cl " _
        & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND T.TP = 'APER' " _
        & "AND T.Cuenta_No = C.Cuenta_No " _
        & "AND C.Codigo = Cl.Codigo " _
        & "ORDER BY T.ME,T.Fecha,T.TP,T.Cuenta_No "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   Opcion = 4
   RatonNormal
End Sub

Public Sub Depositos_Excedidos()
   Presentar_Malla False
   DGCuentas.Caption = "DEPOSITOS EXCEDIDOS"
   sSQL = "SELECT TD As Tipo, CI_RUC, C.Cliente, C.Direccion, Ct.Cuenta_No," _
        & "SUM(Creditos) As Valor,'D' As Moneda,'03' As Tipo_Trans," _
        & "COUNT(Creditos) As Transacciones " _
        & "FROM Trans_Libretas As TL, Clientes_Datos_Extras As Ct, Clientes As C " _
        & "WHERE Ct.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Ct.Codigo " _
        & "AND Ct.Cuenta_No = TL.Cuenta_No " _
        & "AND TL.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TL.Item = '" & NumEmpresa & "' " _
        & "AND Creditos > 0 " _
        & "GROUP BY TD,CI_RUC,C.Cliente,C.Direccion,Ct.Cuenta_No " _
        & "HAVING SUM(Creditos) >= 2000 "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   Opcion = 6
End Sub

Public Sub Reimpresion_Libreta()
  Presentar_Malla False
  DGCuentas.Caption = "REIMPRESION DE LIBRETAS"
  Mensajes = "Seguro de Reimprimir Libreta"
  Titulo = "REIMPRESION DE LIBRETAS"
  If BoxMensaje = vbYes Then
     sSQL = "UPDATE Trans_Libretas " _
          & "SET IP = " & Val(adFalse) & " " _
          & "WHERE Cuenta_No = '" & MBCuenta & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     ConectarAdoExecute sSQL
     Imprimir_Libreta MBCuenta, AdoCuentas, 1, 8, CByte(TextLinea)
     Opcion = 0
  End If
End Sub

Private Sub Command1_Click()
  Unload Historial
End Sub

Private Sub DGCartilla_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   If CtrlDown And KeyCode = vbKeyF5 Then
      DGCartilla.AllowUpdate = True
      MsgBox "Proceso Aceptado, puede Modificar"
      DGCartilla.SetFocus
   End If
End Sub

Private Sub DGCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGCuentas.Visible = False
     GenerarDataTexto Historial, AdoCuentas
     DGCuentas.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGCuentas.AllowUpdate = True
     MsgBox "Proceso Aceptado, puede Modificar"
     DGCuentas.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel(2) Or CNivel(4) Or CNivel(5) Or CNivel(6) Then
        ToolbarHistorial.Buttons("Imprimir").Enabled = False
        ToolbarHistorial.Buttons("Libreta").Enabled = False
        ToolbarHistorial.Buttons("MovimientoLibreta").Enabled = False
        ToolbarHistorial.Buttons("Mov_x_Tipo_Proc").Enabled = False
        ToolbarHistorial.Buttons("Cuentas Abiertas").Enabled = False
        ToolbarHistorial.Buttons("Depositos Excedidos").Enabled = False
     End If
  End If
  sSQL = "SELECT TP " _
       & "FROM Trans_Libretas " _
       & "WHERE TP <> '.' " _
       & "GROUP BY TP " _
       & "ORDER BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  
  Presentar_Malla False
  Opcion = 0
  RatonNormal
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoTP
   ConectarAdodc AdoAux
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoCartilla
End Sub

Private Sub MBCuenta_GotFocus()
  MarcarTexto MBCuenta
End Sub

Private Sub MBCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyR Then
     sSQL = "SELECT * " _
          & "FROM Trans_Certificados " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "ORDER BY Cuenta_No,Fecha,IDT,Hora,ID "
     SelectAdodc AdoAux, sSQL, False
     SaldoAnterior = 0
     Contador = 0
     With AdoAux.Recordset
      If .RecordCount > 0 Then
         'MsgBox LineaNo
          Cuenta_No = .Fields("Cuenta_No")
          SaldoAnterior = 0
          RatonReloj
          Do While Not .EOF
             Historial.Caption = Format(Contador / .RecordCount, "00%")
             Contador = Contador + 1
            'MsgBox PosLinea
             If Cuenta_No <> .Fields("Cuenta_No") Then
                Cuenta_No = .Fields("Cuenta_No")
                SaldoAnterior = 0
             End If
             SaldoAnterior = SaldoAnterior + .Fields("Creditos") - .Fields("Debitos")
            .Fields("Saldo_Disp") = SaldoAnterior
            .Fields("Saldo_Cont") = SaldoAnterior
            .Update
            .MoveNext
         Loop
         RatonNormal
      End If
     End With
  End If
  PresionoEnter KeyCode
End Sub

Public Sub Movimiento_Libreta()
   Presentar_Malla True
   DGCuentas.Caption = "MOVIMIENTO DE LIBRETA"
   CodigoCliente = Ninguno
   
   sSQL = "SELECT Cl.*,C.Tipo_Dato " _
        & "FROM Clientes_Datos_Extras As C,Clientes As Cl " _
        & "WHERE Cuenta_No = '" & MBCuenta.Text & "' " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Cl.Codigo "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        LabelSocio.Caption = " " & .Fields("Cliente")
        CodigoCliente = .Fields("Codigo")
    Else
        LabelSocio.Caption = "NO EXISTE"
    End If
   End With
   Sumatoria = 0
   sSQL = "SELECT T,Fecha,TP,Cheque,Debitos,Creditos,Saldo_Cont,Saldo_Disp,(Saldo_Lib-Saldo_Cont) As Diferencia,Hora,Cartilla_No,Papeleta_No " _
        & "FROM Trans_Libretas " _
        & "WHERE Cuenta_No = '" & MBCuenta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "ORDER BY Fecha,IDT,Hora,ID "
   SQLDec = "Debitos 4|Creditos 4|Saldo_Cont 4|Saldo_Disp 4|."
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   DGCuentas.Visible = False
   With AdoCuentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Sumatoria = Sumatoria + .Fields("Saldo_Cont")
          .MoveNext
        Loop
        Sumatoria = Sumatoria / .RecordCount
    End If
   End With
   DGCuentas.Visible = True
   sSQL = "SELECT * " _
        & "FROM Trans_Cartillas " _
        & "WHERE Cuenta_No = '" & MBCuenta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "ORDER BY Fecha,Cartilla_No,Detalle "
   SelectDataGrid DGCartilla, AdoCartilla, sSQL
   LblPromedio.Caption = Format(Sumatoria, "#,##0.00")
   Opcion = 1
End Sub

Public Sub Listar_Certificados()
   Presentar_Malla False
   DGCuentas.Caption = "LISTAR CERTIFICADOS"
   sSQL = "SELECT T, Fecha, TP, Cheque, Debitos, Creditos, Saldo_Disp, Hora " _
        & "FROM Trans_Certificados " _
        & "WHERE Cuenta_No = '" & MBCuenta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "ORDER BY Fecha, IDT, Hora, ID "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   Opcion = 7
End Sub

Private Sub MBCuenta_LostFocus()
   CodigoCliente = Ninguno
   sSQL = "SELECT Cl.*,C.Tipo_Dato " _
        & "FROM Clientes_Datos_Extras As C,Clientes As Cl " _
        & "WHERE Cuenta_No = '" & MBCuenta & "' " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Cl.Codigo "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        LabelSocio.Caption = " " & .Fields("Cliente")
        CodigoCliente = .Fields("Codigo")
    Else
        LabelSocio.Caption = "NO EXISTE"
    End If
   End With
End Sub

Private Sub MBFechaF_GotFocus()
   MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

''Private Sub MBoxTarjeta_LostFocus()
''   FechaValida MBoxFechaI
''   FechaValida MBoxFechaF
''   FechaIni = BuscarFecha(MBoxFechaI.Text)
''   FechaFin = BuscarFecha(MBoxFechaF.Text)
''   sSQL = "SELECT C.* FROM Tarjetas As T,Clientes_Datos_Extras As C " _
''        & "WHERE Tarjeta_No = '" & MBoxTarjeta.Text & "' " _
''        & "AND T.Cuenta_No = C.Cuenta_No "
''   SelectAdodc AdoAux, sSQL
''   With AdoAux.Recordset
''    If .RecordCount > 0 Then
''        LabelSocio1.Caption = " " & .Fields("Nombres") & " " & .Fields("Apellidos")
''    Else
''        LabelSocio1.Caption = "NO EXISTE"
''    End If
''   End With
''   sSQL = "SELECT T,Fecha,TP,Ticket,Debitos,Creditos,Saldo_Disp,Hora " _
''        & "FROM Trans_Tarjetas " _
''        & "WHERE Tarjeta_No = '" & MBoxTarjeta.Text & "' " _
''        & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''        & "ORDER BY Fecha,IDT,Hora,ID "
''   SelectDataGrid DGTarjeta, AdoTarjeta, sSQL
''End Sub

Private Sub TextLinea_GotFocus()
  MarcarTexto TextLinea
End Sub

Private Sub TextLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextLinea_LostFocus()
  TextoValido TextLinea
  If Not IsNumeric(TextLinea) Then TextLinea = "1"
End Sub

Public Sub Impresiones()
    MsgBox "Tipo de Impresion: " & Opcion
    Select Case Opcion
      Case 1
           SQLMsg1 = "LISTADO DE MOVIMIENTO"
           SQLMsg2 = "Cuenta No. " & MBCuenta
           ImprimirAdo AdoCuentas, 1, 10, True
      Case 2
           MensajeEncabData = "REPORTE DE FLUJO DE CAJA"
           ImprimirAdo AdoCuentas, True, 1, 7
      Case 3
           MensajeEncabData = "MOVIMIENTO POR TIPO DE PROCESO"
           SQLMsg1 = "PROCESO: " & DCTP
           ImprimirAdo AdoCuentas, True, 1, 7
      Case 4
           SQLMsg1 = "LIBRETAS ABIERTAS"
           SQLMsg2 = "Desde " & MBFechaI & " al " & MBFechaF
           ImprimirAdodc AdoCuentas, 1, 8
      Case 5
           SQLMsg1 = "LISTADO DE MOVIMIENTO DE TARJETA"
           SQLMsg2 = "Tarjeta No. " '& MBTarjeta
           ImprimirAdodc AdoCuentas, 1, 8
      Case 6
           SQLMsg1 = "DEPOSITOS EXCEDIDOS"
           SQLMsg2 = "Desde " & MBFechaI & " al " & MBFechaF
           ImprimirAdodc AdoCuentas, 1, 8
      Case 7
           Mensajes = "Seguro de Reimprimir Certificado"
           Titulo = "REIMPRESION DE CERTIFICADO"
           If BoxMensaje = vbYes Then
              sSQL = "UPDATE Trans_Certificados " _
                   & "SET IP = " & Val(adFalse) & " " _
                   & "WHERE Cuenta_No = '" & MBCuenta & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
              ConectarAdoExecute sSQL
              Imprimir_Certificados MBCuenta, AdoCuentas, 1, 8, CByte(TextLinea)
           End If
      Case 8
           SQLMsg1 = "LISTADO DE SALDO DE LIBRETAS"
           SQLMsg2 = "Corte al " & MBFechaF
           ImprimirAdodc AdoCuentas, 1, 10
    End Select
End Sub

Private Sub ToolbarHistorial_ButtonClick(ByVal Button As ComctlLib.Button)
    FechaValida MBFechaI
    FechaValida MBFechaF
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload Historial
      Case "Imprimir"
           Impresiones
      Case "Libreta"
           Reimpresion_Libreta
      Case "MovimientoLibreta"
           Movimiento_Libreta
      Case "Mov_x_Tipo_Proc"
           Movimiento_x_Tipo_de_Proceso
      Case "Cuentas Abiertas"
           Cuentas_Abiertas
      Case "Depositos Excedidos"
           Depositos_Excedidos
      Case "Certificados"
           Listar_Certificados
      Case "Saldos Libretas"
           Saldos_Libretas
      Case "Saldos Real Libretas"
           Saldos_Real_Libretas
    End Select
End Sub

Public Sub Presentar_Malla(MediaMalla As Boolean)
  DGCuentas.width = MDI_X_Max - 100
  If MediaMalla Then
     DGCartilla.Visible = True
     AdoCartilla.Visible = True
     
     DGCuentas.Height = ((MDI_Y_Max - 1000) / 2)
     AdoCuentas.width = MDI_X_Max - 3100
     
     AdoCuentas.Top = DGCuentas.Top + DGCuentas.Height + 10
     Label29.Top = AdoCuentas.Top
     LblPromedio.Top = AdoCuentas.Top
     Label29.Left = AdoCuentas.Left + AdoCuentas.width + 10
     LblPromedio.Left = Label29.Left + Label29.width
     
     DGCartilla.width = MDI_X_Max - 100
     AdoCartilla.width = MDI_X_Max - 100
     DGCartilla.Top = AdoCuentas.Top + AdoCuentas.Height + 20
     DGCartilla.Height = MDI_Y_Max - AdoCuentas.Top - AdoCuentas.Height - AdoCartilla.Height
     AdoCartilla.Top = DGCartilla.Top + DGCartilla.Height + 10
     Label29.Caption = " PROMEDIO"
  Else
     AdoCartilla.Visible = False
     DGCartilla.Visible = False
     
     DGCuentas.Height = MDI_Y_Max - 1800
     AdoCuentas.width = MDI_X_Max - 3100
     AdoCuentas.Top = DGCuentas.Top + DGCuentas.Height + 10
     Label29.Top = AdoCuentas.Top
     LblPromedio.Top = AdoCuentas.Top
     Label29.Left = AdoCuentas.Left + AdoCuentas.width + 10
     LblPromedio.Left = Label29.Left + Label29.width
     Label29.Caption = " TOTAL"
  End If
End Sub

Public Sub Saldos_Libretas()
   Presentar_Malla False
   DGCuentas.Caption = "SALDOS DE LIBRETAS"
   sSQL = "SELECT C.Fecha As Fecha_Aper,C.Cliente,TL.Cuenta_No,(SUM(TL.Creditos)-SUM(TL.Debitos)) As Saldo_Libreta " _
        & "FROM Trans_Libretas As TL, Clientes_Datos_Extras As CL, Clientes As C " _
        & "WHERE TL.Fecha <= #" & FechaFin & "# " _
        & "AND TL.Item = '" & NumEmpresa & "' " _
        & "AND CL.Tipo_Dato = 'LIBRETAS' and TL.T <> '.' " _
        & "AND C.Codigo = CL.Codigo " _
        & "AND CL.Cuenta_No = TL.Cuenta_No " _
        & "GROUP BY C.Fecha,C.Cliente,TL.Cuenta_No " _
        & "ORDER BY C.Cliente,TL.Cuenta_No,C.Fecha "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   RatonReloj
   DGCuentas.Visible = False
   Saldo = 0
   With AdoCuentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Saldo = Saldo + .Fields("Saldo_Libreta")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   DGCuentas.Visible = True
   RatonNormal
   Opcion = 8
   LblPromedio.Caption = Format(Saldo, "#,##0.00")
End Sub

Public Sub Saldos_Real_Libretas()
   Presentar_Malla False
   DGCuentas.Visible = False
  'Fecha_Saldo
   Progreso_Barra.Mensaje_Box = "Procesando Saldos al " & MBFechaF & ", Espere un momento."
   Progreso_Iniciar

   DGCuentas.Caption = "SALDOS DE LIBRETAS"
   sSQL = "SELECT * " _
        & "FROM Clientes_Datos_Extras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Tipo_Dato = 'LIBRETAS' " _
        & "ORDER BY Cuenta_No "
   SelectAdodc AdoCuentas, sSQL
   With AdoCuentas.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
        Do While Not .EOF
           Saldo = 0
           Total = 0
           sSQL = "SELECT TOP 1 * " _
                & "FROM Trans_Libretas " _
                & "WHERE Cuenta_No = '" & .Fields("Cuenta_No") & "' " _
                & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
           SelectAdodc AdoAux, sSQL
           If AdoAux.Recordset.RecordCount > 0 Then
              Saldo = AdoAux.Recordset.Fields("Saldo_Cont")
              Total = AdoAux.Recordset.Fields("Saldo_Lib")
           End If
          .Fields("Fecha_Saldo") = MBFechaF
          .Fields("Saldo_Libreta") = Saldo
          .Fields("Saldo_Real") = Total
          .Update
           Progreso_Esperar
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   
   sSQL = "SELECT CL.Fecha_Saldo,C.Cliente,CL.Cuenta_No,CL.Saldo_Libreta,CL.Saldo_Real,(CL.Saldo_Real-CL.Saldo_Libreta) As Diferencias " _
        & "FROM Clientes_Datos_Extras As CL, Clientes As C " _
        & "WHERE CL.Fecha_Saldo = #" & FechaFin & "# " _
        & "AND CL.Item = '" & NumEmpresa & "' " _
        & "AND CL.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = CL.Codigo " _
        & "ORDER BY C.Cliente, CL.Cuenta_No "
   SelectDataGrid DGCuentas, AdoCuentas, sSQL
   DGCuentas.Visible = False
   RatonReloj
   Saldo = 0
   With AdoCuentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Saldo = Saldo + .Fields("Saldo_Libreta")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   DGCuentas.Visible = True
   Progreso_Final
   RatonNormal
   Opcion = 8
   LblPromedio.Caption = Format(Saldo, "#,##0.00")
End Sub


