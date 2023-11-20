VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FVencidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDITOS VENCIDOS"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11145
   Begin MSDataGridLib.DataGrid DGSaldos 
      Bindings        =   "FVencido.frx":0000
      Height          =   4530
      Left            =   105
      TabIndex        =   23
      Top             =   1995
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   7990
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2205
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
      Height          =   1905
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   2745
      Begin VB.OptionButton OpcPV 
         Caption         =   "Por Vencer"
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
         TabIndex        =   19
         Top             =   1575
         Width           =   2010
      End
      Begin VB.OptionButton OpcVD 
         Caption         =   "Vencidos del &día"
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
         TabIndex        =   18
         Top             =   1365
         Width           =   2010
      End
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   1365
         TabIndex        =   22
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
      Begin VB.OptionButton OpcV 
         Caption         =   "C&uotas Vigentes"
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
         Top             =   1155
         Width           =   2010
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "&Vencidos"
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
         TabIndex        =   16
         Top             =   945
         Value           =   -1  'True
         Width           =   1170
      End
      Begin MSMask.MaskEdBox MBoxFecha 
         Height          =   330
         Left            =   105
         TabIndex        =   21
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
         Left            =   1365
         TabIndex        =   3
         Top             =   210
         Width           =   1275
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
         TabIndex        =   20
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "I&mprimir Listado"
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
      Left            =   9975
      Picture         =   "FVencido.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4620
      Width           =   1065
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
      Height          =   435
      Left            =   9975
      Picture         =   "FVencido.frx":089A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1995
      Width           =   435
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
      Height          =   435
      Left            =   10605
      Picture         =   "FVencido.frx":099C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1995
      Width           =   435
   End
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   2940
      TabIndex        =   4
      Top             =   0
      Width           =   8100
      Begin MSDataListLib.DataCombo DCArea 
         Bindings        =   "FVencido.frx":0A9E
         DataSource      =   "AdoArea"
         Height          =   315
         Left            =   2100
         TabIndex        =   26
         Top             =   1050
         Width           =   1275
         _ExtentX        =   2249
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
      Begin MSDataListLib.DataCombo DCSector 
         Bindings        =   "FVencido.frx":0AB4
         DataSource      =   "AdoSector"
         Height          =   315
         Left            =   4515
         TabIndex        =   25
         Top             =   945
         Width           =   3480
         _ExtentX        =   6138
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
      Begin MSDataListLib.DataCombo DCTipo 
         Bindings        =   "FVencido.frx":0ACC
         DataSource      =   "AdoTrans"
         Height          =   315
         Left            =   105
         TabIndex        =   24
         Top             =   525
         Width           =   7890
         _ExtentX        =   13917
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
      Begin MSMask.MaskEdBox MBoxCuenta 
         Height          =   330
         Left            =   2100
         TabIndex        =   5
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1365
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   192
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
         Format          =   "CCCCCCCC-C"
         Mask            =   "########-#"
         PromptChar      =   "0"
      End
      Begin VB.CheckBox CheqArea 
         Caption         =   "Area Geográfica"
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
         Top             =   1050
         Width           =   1800
      End
      Begin VB.CheckBox CheqSector 
         Caption         =   "Sector"
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
         TabIndex        =   8
         Top             =   1050
         Width           =   960
      End
      Begin VB.CheckBox CheqLib_No 
         Caption         =   "Número de Libreta"
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
         TabIndex        =   6
         Top             =   1470
         Width           =   1905
      End
      Begin VB.Label LabelPromedio 
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
         Left            =   5880
         TabIndex        =   10
         Top             =   1365
         Width           =   2115
      End
      Begin VB.Label Label29 
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
         Left            =   4515
         TabIndex        =   11
         Top             =   1365
         Width           =   1380
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
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   7890
      End
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
      Height          =   960
      Left            =   9975
      Picture         =   "FVencido.frx":0AE3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5670
      Width           =   1065
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
      Height          =   960
      Left            =   9975
      Picture         =   "FVencido.frx":14D9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1065
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
      Height          =   960
      Left            =   9975
      Picture         =   "FVencido.frx":191B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3570
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoInt 
      Height          =   330
      Left            =   315
      Top             =   2520
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
      Caption         =   "Int"
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
      Left            =   315
      Top             =   2835
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
   Begin MSAdodcLib.Adodc AdoSector 
      Height          =   330
      Left            =   315
      Top             =   3150
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
      Caption         =   "Sector"
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
   Begin MSAdodcLib.Adodc AdoArea 
      Height          =   330
      Left            =   315
      Top             =   3465
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
      Caption         =   "Area"
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
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   105
      Top             =   6510
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
End
Attribute VB_Name = "FVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  If NombreCampo = "Fecha" Then NombreCampo = "P." & NombreCampo
  sSQL = sSQL & "ORDER BY " & NombreCampo & " ASC "
End Sub

Private Sub Command2_Click()
   Opcion = 1
   SQLMsg1 = UCase(DCTipo.Text)
   If OpcP.value Then
      SQLMsg2 = "Vencidos: Desde " & MBoxFecha.Text & " al " & MBoxFechaF.Text
      Opcion = 2
   ElseIf OpcV.value Then
      SQLMsg2 = "Cuotas Vigentes: Desde " & MBoxFecha.Text & " al " & MBoxFechaF.Text
   ElseIf OpcPV.value Then
      SQLMsg2 = "Por Vencer del " & MBoxFecha.Text
      Opcion = 2
   Else
      SQLMsg2 = "Vencidos del " & MBoxFecha.Text
      Opcion = 2
   End If
   Imprimir_Vencidos AdoSaldos, True, 1, 8, Opcion
End Sub

Private Sub Command3_Click()
  FechaValida MBoxFecha, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFecha.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = UCase(SinEspaciosIzq(DCTipo.Text))
  Codigo1 = UCase(DCArea.Text)
  Codigo2 = UCase(DCSector.Text)
  Codigo3 = MBoxCuenta.Text
  sSQL = "SELECT P.Credito_No,C.Cliente,P.Cuenta_No,C.Direccion," _
       & "Cta.Ciudad As Sector,Cta.Area,C.Telefono,C.TelefonoT," _
       & "P.Fecha,P.Cuota_No,P.Pagos,P.Capital " _
       & "FROM Trans_Prestamos As P,Clientes As C,Clientes_Datos_Extras As Cta "
  If OpcVD.value Then
     sSQL = sSQL & "WHERE P.Fecha_V = #" & BuscarFecha(MBoxFecha) & "# "
  ElseIf OpcPV.value Then
     sSQL = sSQL & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = sSQL & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If OpcP.value Then
     sSQL = sSQL & "AND P.T = 'P' "
     ' "AND P.V = " & Val(adTrue) & " "
  ElseIf OpcV.value Then
     sSQL = sSQL & "AND P.V = " & Val(adFalse) & "  " _
          & "AND P.T = 'P' "
  ElseIf OpcPV.value Then
     sSQL = sSQL & "AND P.V = " & Val(adFalse) & "  " _
          & "AND P.T = 'P' "
  Else
     sSQL = sSQL & "AND P.V <> " & Val(adFalse) & " "
  End If
  sSQL = sSQL _
       & "AND Cta.Tipo_Dato = 'LIBRETAS' " _
       & "AND P.Cuenta_No = Cta.Cuenta_No " _
       & "AND Cta.Codigo = C.Codigo " _
       & "AND P.TP = '" & Codigo & "' "
  If CheqArea.value = 1 Then sSQL = sSQL & "AND UCase(Cta.Area) = '" & Codigo1 & "' "
  If CheqSector.value = 1 Then sSQL = sSQL & "AND UCase(Cta.Ciudad) = '" & Codigo2 & "' "
  If CheqLib_No.value = 1 Then sSQL = sSQL & "AND P.Cuenta_No = '" & Codigo3 & "' "
  sSQL = sSQL & "AND Cta.Item = '" & NumEmpresa & "' " _
       & "ORDER BY P.Fecha,Cuota_No,C.Cliente,P.Credito_No "
  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  DGSaldos.Visible = False
  RatonReloj
  Total = 0
  With AdoSaldos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcP.value Then
             Total = Total + .Fields("Capital")
          ElseIf OpcVD.value Then
             Total = Total + .Fields("Capital")
          Else
             Total = Total + .Fields("Pagos")
          End If
         .MoveNext
       Loop
   End If
  End With
  LabelPromedio.Caption = Format(Total, "#,##0.00")
  DGSaldos.Visible = True
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Command4_Click()
  Unload FVencidos
End Sub

Private Sub Command6_Click()
   SQLMsg1 = UCase(DCTipo.Text)
   Opcion = 1
   If OpcP.value Then
      SQLMsg2 = "Vencidos: Desde " & MBoxFecha.Text & " al " & MBoxFechaF.Text
      Opcion = 2
   ElseIf OpcV.value Then
      SQLMsg2 = "Cuotas Vigentes: Desde " & MBoxFecha.Text & " al " & MBoxFechaF.Text
   ElseIf OpcPV.value Then
      SQLMsg2 = "Por Vencer del " & MBoxFecha.Text
      Opcion = 2
   Else
      SQLMsg2 = "Vencidos del " & MBoxFecha.Text
      Opcion = 2
   End If
   Imprimir_Vencidos AdoSaldos, True, 2, 8, Opcion
End Sub

Private Sub DGSaldos_DblClick()
  NombreCampo = DGSaldos.Columns(I)
End Sub

Private Sub DGSaldos_HeadClick(ByVal ColIndex As Integer)
  I = ColIndex
End Sub

Private Sub DGSaldos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     DGSaldos.Visible = False
     GenerarDataTexto FVencidos, AdoSaldos
     DGSaldos.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel(3) Or CNivel(6) Then
     Command2.Enabled = False
     Command6.Enabled = False
     End If
  End If
  NombreCampo = Ninguno
  sSQL = "SELECT CTP & '   ' & Descripcion As TipoPrest " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE TC = True " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "ORDER BY CTP,Descripcion "
  SelectDBCombo DCTipo, AdoTrans, sSQL, "TipoPrest", False
  sSQL = "SELECT Area " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Tipo_Dato = 'LIBRETAS' " _
       & "GROUP BY Area "
  SelectDBCombo DCArea, AdoArea, sSQL, "Area"
  sSQL = "SELECT Ciudad " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Tipo_Dato = 'LIBRETAS' " _
       & "GROUP BY Ciudad "
  SelectDBCombo DCSector, AdoSector, sSQL, "Ciudad", False
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FVencidos
  ConectarAdodc AdoInt
  ConectarAdodc AdoAux
  ConectarAdodc AdoTrans
  ConectarAdodc AdoArea
  ConectarAdodc AdoSector
  ConectarAdodc AdoSaldos
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub
