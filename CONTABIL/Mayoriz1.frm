VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form ListMayorizacion1 
   Caption         =   "Mayor Analitico"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
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
            Key             =   "Excel"
            Object.ToolTipText     =   "Bajar a Excel el Reporte"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnMayor"
            Object.ToolTipText     =   "Consultar un Mayor Auxiliar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "VariosMayores"
            Object.ToolTipText     =   "Consultar Varios Mayor Auxiliar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Patron de Busqueda"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   3570
         TabIndex        =   0
         Top             =   0
         Width           =   16920
         Begin VB.CheckBox CheqTC 
            Caption         =   "TC"
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
            Left            =   7035
            TabIndex        =   26
            Top             =   210
            Width           =   645
         End
         Begin VB.CheckBox CheqImpSubCtas 
            Caption         =   "Con Submodulos/Centro de Costos"
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
            Left            =   13440
            TabIndex        =   10
            Top             =   210
            Width           =   3375
         End
         Begin MSMask.MaskEdBox MBoxCtaI 
            Height          =   330
            Left            =   1785
            TabIndex        =   2
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBoxCtaF 
            Height          =   330
            Left            =   5250
            TabIndex        =   4
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBoxFechaI 
            Height          =   330
            Left            =   9660
            TabIndex        =   7
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
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   12075
            TabIndex        =   9
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
         Begin MSDataListLib.DataCombo DCTC 
            Bindings        =   "Mayoriz1.frx":0000
            DataSource      =   "AdoTC"
            Height          =   345
            Left            =   7770
            TabIndex        =   5
            Top             =   210
            Visible         =   0   'False
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   609
            _Version        =   393216
            Text            =   "XXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &4.- Hasta"
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
            Left            =   11025
            TabIndex        =   8
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &3.- Desde"
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
            Left            =   8610
            TabIndex        =   6
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &2.- Cuenta Final"
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
            Left            =   3570
            TabIndex        =   3
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &1.- Cuenta Inicial"
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
            Width           =   1695
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7665
      Width           =   330
   End
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "Mayoriz1.frx":0014
      DataSource      =   "AdoAgencias"
      Height          =   345
      Left            =   7770
      TabIndex        =   14
      Top             =   735
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "Agencia:"
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
      Left            =   6615
      TabIndex        =   13
      Top             =   735
      Width           =   1170
   End
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por &Usuario:"
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
      Top             =   735
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid DGMayor 
      Bindings        =   "Mayoriz1.frx":002E
      Height          =   3795
      Left            =   105
      TabIndex        =   22
      Top             =   1575
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   6694
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   525
      Top             =   3465
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
      Left            =   525
      Top             =   4095
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "Mayoriz1.frx":0045
      DataSource      =   "AdoCtas"
      Height          =   345
      Left            =   105
      TabIndex        =   15
      Top             =   1155
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Ctas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoMayor 
      Height          =   330
      Left            =   420
      Top             =   7665
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   "Mayor"
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
      Left            =   525
      Top             =   3780
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
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "Mayoriz1.frx":005B
      DataSource      =   "AdoUsuario"
      Height          =   345
      Left            =   1575
      TabIndex        =   12
      Top             =   735
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   2730
      Top             =   3465
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   2730
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
      Caption         =   "Usuario"
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
   Begin MSAdodcLib.Adodc AdoTC 
      Height          =   330
      Left            =   2730
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
      Caption         =   "TC"
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
   Begin VB.Label LblPatron 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Patron"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   14910
      TabIndex        =   27
      Top             =   735
      Width           =   5580
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   15540
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":0074
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":038E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":09C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":0CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mayoriz1.frx":2ADE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelTotSaldoAnt 
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
      Left            =   13125
      TabIndex        =   23
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Anterior MN"
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
      Left            =   11340
      TabIndex        =   24
      Top             =   1155
      Width           =   1800
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
      Left            =   9870
      TabIndex        =   21
      Top             =   7665
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual MN"
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
      Left            =   8190
      TabIndex        =   20
      Top             =   7665
      Width           =   1695
   End
   Begin VB.Label LabelTotHaber 
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
      Left            =   6510
      TabIndex        =   19
      Top             =   7665
      Width           =   1695
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Haber:"
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
      Left            =   5775
      TabIndex        =   18
      Top             =   7665
      Width           =   750
   End
   Begin VB.Label LabelTotDebe 
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
      Left            =   4095
      TabIndex        =   17
      Top             =   7665
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe:"
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
      Left            =   3360
      TabIndex        =   16
      Top             =   7665
      Width           =   750
   End
End
Attribute VB_Name = "ListMayorizacion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''Private Sub Command6_Click()
''''  RatonReloj
''''  DGMayor.Visible = False
''''  FechaValida MBoxFechaI
''''  FechaValida MBoxFechaF
''''  Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
''''  Codigo2 = CambioCodigoCta(MBoxCtaF.Text)
''''  If Codigo1 = " " Then Codigo1 = "1"
''''  If Codigo2 = " " Then Codigo2 = "9"
''''  DGMayor.Caption = " Mayor de: " & DCCtas.Text & "."
''''  FechaIni = BuscarFecha(MBoxFechaI.Text)
''''  FechaFin = BuscarFecha(MBoxFechaF.Text)
''''  SumaDebe = 0: SumaHaber = 0: Suma_ME = 0
''''  sSQL = "SELECT CC.TC,Cl.Codigo,Cl.Cliente,T.Fecha,T.TP,T.Numero,C.Concepto," _
''''       & "Debe,Haber,Saldo,Parcial_ME,Saldo_ME,ID,Cta " _
''''       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl,Catalogo_Cuentas As CC " _
''''       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
''''       & "AND T.T = '" & Normal & "' " _
''''       & "AND T.Cta BETWEEN '" & Codigo1 & "' AND '" & Codigo2 & "' " _
''''       & "AND T.Item = '" & NumEmpresa & "' " _
''''       & "AND CC.TC <> 'N' " _
''''       & "AND CC.TC <> 'CJ' " _
''''       & "AND CC.TC <> 'BA' " _
''''       & "AND C.Codigo_B = Cl.Codigo " _
''''       & "AND T.Cta = CC.Codigo " _
''''       & "AND T.TP = C.TP " _
''''       & "AND T.Numero = C.Numero " _
''''       & "AND T.Fecha = C.Fecha " _
''''       & "AND T.Item = C.Item " _
''''       & "AND T.Item = CC.Item " _
''''       & "ORDER BY T.Cta,T.Fecha,T.TP,T.Numero,Debe DESC,Haber,ID "
''''  Select_Adodc_Grid DGMayor, AdoMayor, sSQL
''''  DGMayor.Visible = False
''''      With AdoMayor.Recordset
''''     If .RecordCount > 0 Then
''''         ProgBar.Min = 0: ProgBar.Max = .RecordCount: I = 0
''''         Do While Not .EOF
''''            ProgBar.Value = I
''''            Suma_ME = Suma_ME + .Fields("Parcial_ME")
''''            SumaDebe = SumaDebe + .Fields("Debe")
''''            SumaHaber = SumaHaber + .Fields("Haber")
''''            SaldoTotal = .Fields("Saldo")
''''            I = I + 1
''''           .MoveNext
''''         Loop
''''        .MoveFirst
''''     End If
''''    End With
''''    DGMayor.Visible = True
''''    LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
''''    LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
''''    LabelTotSaldo.Caption = Format(SaldoTotal, "#,##0.00")
''''    SaldoAnterior = CalculosSaldoAnt(DCCtas.Text, SumaDebe, SumaHaber, SaldoTotal)
''''    LabelTotSaldoAnt.Caption = Format(SaldoAnterior, "#,##0.00")
''''    ProgBar.Value = ProgBar.Max
''''    ListMayorizacion1.Caption = "MAYOR ANALITICO"
''''    RatonNormal
''''    DCCtas.SetFocus
''''End Sub

Private Sub CheqTC_Click()
  If CheqTC.value = 0 Then DCTC.Visible = False Else DCTC.Visible = True
End Sub

Private Sub Command1_Click()
    RatonNormal
    Unload ListMayorizacion1
End Sub

Private Sub DCCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTC_LostFocus()
   sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND DG = 'D' " _
        & "AND TC = '" & DCTC.Text & "' " _
        & "AND Codigo IN (SELECT Cta " _
        & "               FROM Transacciones " _
        & "               WHERE Item = '" & NumEmpresa & "' " _
        & "               AND Periodo = '" & Periodo_Contable & "' " _
        & "               GROUP BY Cta) " _
        & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoCtas, sSQL, "Nombre_Cta"
End Sub

Private Sub DGMayor_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGMayor.AllowUpdate = True
     MsgBox "Puede Empezar a modificar los Conceptos"
  End If
  
  If KeyCode = vbKeyF10 Then
    If ClaveContador Then
       Co.Fecha = DGMayor.Columns(0).Text
       Co.TP = DGMayor.Columns(1).Text
       Co.Numero = DGMayor.Columns(2).Text
       Co.Beneficiario = DGMayor.Columns(3).Text
       FechaComp = Co.Fecha
       NumeroComp = Co.Numero
       Mensajes = "Seguro que quiere Modificar" & vbCrLf & "El Comprobante: " & Co.TP & " No. " & Co.Numero & " Con Fecha: " & Co.Fecha & vbCrLf
       If Len(Co.Beneficiario) > 1 Then Mensajes = Mensajes & "De " & ULCase(Co.Beneficiario)
       Titulo = "Pregunta de Modificacion"
       If BoxMensaje = vbYes Then
          ModificarComp = True
          CopiarComp = False
          NuevoComp = False
          Trans_No = 1
          Unload ListMayorizacion1
          FComprobantes.Show
       End If
    End If
  End If
End Sub

Private Sub Form_Activate()
  SQLPatron = ""
  LblPatron.Caption = "Patron Busqueda:"
  Co.Item = NumEmpresa
  If CNivel(7) Then
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload ListMayorizacion1
  Else
     If Supervisor = False Then
        If CNivel(3) Or CNivel(6) Then
           Command2.Enabled = False
        End If
     End If
     FormatoMaskCta MBoxCtaI
     FormatoMaskCta MBoxCtaF
     sSQL = "SELECT TC " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND DG = 'D' " _
          & "AND Codigo IN (SELECT Cta " _
          & "               FROM Transacciones " _
          & "               WHERE Item = '" & NumEmpresa & "' " _
          & "               AND Periodo = '" & Periodo_Contable & "' " _
          & "               GROUP BY Cta) " _
          & "GROUP BY TC " _
          & "ORDER BY TC "
     SelectDB_Combo DCTC, AdoTC, sSQL, "TC"
     
     sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND DG = 'D' " _
          & "AND Codigo IN (SELECT Cta " _
          & "               FROM Transacciones " _
          & "               WHERE Item = '" & NumEmpresa & "' " _
          & "               AND Periodo = '" & Periodo_Contable & "' " _
          & "               GROUP BY Cta) " _
          & "ORDER BY Codigo "
     SelectDB_Combo DCCtas, AdoCtas, sSQL, "Nombre_Cta"
     ListMayorizacion1.Caption = "MAYOR ANALITICO"
     If Individual Then
        DGMayor.ToolTipText = "<F10>:Modificar Comprobante  <Ctrl>+<F5>:Modificar Conceptos Contables"
     Else
        DGMayor.ToolTipText = "<F10>:Modificar Comprobante"
     End If
     sSQL = "SELECT (Nombre_Completo & '  ' & Codigo) As CodUsuario " _
          & "FROM Accesos " _
          & "WHERE Codigo <> '*' " _
          & "ORDER BY Nombre_Completo "
     SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "CodUsuario", False
  End If
  If ConSucursal Then
     sSQL = "SELECT (Item & '  ' & Empresa) As NomEmpresa " _
          & "FROM Empresas " _
          & "WHERE Item IN (" & ListSucursales & ") " _
          & "ORDER BY Item,Empresa "
     SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NomEmpresa"
     CheckAgencia.value = 0
     DCAgencia.Visible = True
     CheckAgencia.Visible = True
  Else
     DCAgencia.Visible = False
     CheckAgencia.Visible = False
  End If
  MBoxFechaI.Text = FechaSistema
  MBoxFechaF.Text = FechaSistema
  MBoxCtaI.SelText = "1"
  MBoxCtaF.SelText = "9"
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  ListarMayoresAux True, Individual
  Obtener_Campos_Patron_Busqueda AdoMayor
  ListMayorizacion1.Caption = "MAYOR ANALITICO"
  Opcion = 1
  MBoxCtaI.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
 'CentrarForm ListMayorizacion1
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTC
  ConectarAdodc AdoTrans
  ConectarAdodc AdoMayor
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoAgencias
  
  DGMayor.Height = MDI_Y_Max - DGMayor.Top - 400
  DGMayor.width = MDI_X_Max - DGMayor.Left
  AdoMayor.Top = DGMayor.Top + DGMayor.Height + 10
  Command1.Top = AdoMayor.Top
  LblPatron.width = MDI_X_Max - LblPatron.Left - 10
  
  Label6.Top = DGMayor.Top + DGMayor.Height + 10
  Label9.Top = DGMayor.Top + DGMayor.Height + 10
  Label11.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotSaldo.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotDebe.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotHaber.Top = DGMayor.Top + DGMayor.Height + 10
  Opcion = 0
  SQLPatron = ""
End Sub

Private Sub MBoxCtaF_GotFocus()
  MarcarTexto MBoxCtaF
End Sub

Private Sub MBoxCtaF_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxCtaF_LostFocus()
   Codigo1 = CambioCodigoCta(MBoxCtaI)
   Codigo2 = CambioCodigoCta(MBoxCtaF)
   If Codigo1 <> "0" And Codigo2 <> "0" Then
      sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
           & "FROM Catalogo_Cuentas " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Codigo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' " _
           & "AND DG = 'D' " _
           & "ORDER BY Codigo "
      SelectDB_Combo DCCtas, AdoCtas, sSQL, "Nombre_Cta"
   End If
End Sub

Private Sub MBoxCtaI_GotFocus()
  MarcarTexto MBoxCtaI
End Sub

Private Sub MBoxCtaI_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxCtaI_LostFocus()
  Codigo1 = CambioCodigoCta(MBoxCtaI)
'''  If Codigo1 = "0" Then
'''     Codigo1 = "1"
'''     MBoxCtaI = FormatoCodigoCta(Codigo1)
'''  End If
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Public Sub ListarMayoresAux(OpcUno As Boolean, PorConceptos As Boolean)
  DGMayor.Visible = False
  Progreso_Iniciar
  
  If PorConceptos Then
     sSQL = "SELECT T.Fecha,T.TP,T.Numero,Cl.Cliente,T.Detalle As Concepto,T.Cheq_Dep,T.Debe,T.Haber,T.Saldo," _
          & "T.Parcial_ME,T.Saldo_ME,T.ID,T.Cta,T.Item " _
          & "FROM Transacciones As T, Clientes As Cl "
  Else
     sSQL = "SELECT T.Fecha,T.TP,T.Numero,Cl.Cliente,Co.Concepto,T.Cheq_Dep,Debe,Haber,Saldo," _
          & "Parcial_ME,Saldo_ME,T.ID,T.Cta,CC.TC,CC.Cuenta,T.Item " _
          & "FROM Catalogo_Cuentas As CC, Transacciones As T, Comprobantes As Co, Clientes As Cl "
  End If
  sSQL = sSQL & "WHERE T.Periodo = '" & Periodo_Contable & "' "
  If CheckAgencia.value = 1 Then
     sSQL = sSQL & "AND T.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     If Not ConSucursal Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  End If
  sSQL = sSQL & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.T = '" & Normal & "' "
  If CheqTC.value = 0 Then
     If OpcUno Then
        sSQL = sSQL & "AND T.Cta = '" & Codigo1 & "' "
     Else
        sSQL = sSQL & "AND T.Cta BETWEEN '" & Codigo1 & "' AND '" & Codigo2 & "' "
     End If
  Else
     sSQL = sSQL & "AND CC.TC = '" & DCTC.Text & "' "
  End If
  
  sSQL = sSQL & SQLPatron
  
  If PorConceptos Then
     sSQL = sSQL & "AND T.Codigo_C = Cl.Codigo "
  Else
     If CheckUsuario.value = 1 Then sSQL = sSQL & "AND C.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
     sSQL = sSQL _
          & "AND T.Item = Co.Item " _
          & "AND T.Item = CC.Item " _
          & "AND T.Periodo = Co.Periodo " _
          & "AND T.Periodo = CC.Periodo " _
          & "AND T.Cta = CC.Codigo " _
          & "AND T.Fecha = Co.Fecha " _
          & "AND T.TP = Co.TP " _
          & "AND T.Numero = Co.Numero " _
          & "AND Co.Codigo_B = Cl.Codigo "
  End If
  sSQL = sSQL & "ORDER BY T.Cta,T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
  Select_Adodc_Grid DGMayor, AdoMayor, sSQL
  'MsgBox "..."
 'Consulta de Totales
  sSQL = "SELECT T.Cta,SUM(T.Debe) As TDebe, SUM(T.Haber) As THaber, SUM(T.Parcial_ME) As TParcial_ME "
  If PorConceptos Then
     sSQL = sSQL & "FROM Transacciones As T,Clientes As Cl "
  Else
     sSQL = sSQL & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl "
  End If
  sSQL = sSQL _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.T = '" & Normal & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  If OpcUno Then
     sSQL = sSQL & "AND T.Cta = '" & Codigo1 & "' "
  Else
     sSQL = sSQL & "AND T.Cta BETWEEN '" & Codigo1 & "' AND '" & Codigo2 & "' "
  End If
  If CheckAgencia.value = 1 Then
     sSQL = sSQL & "AND T.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
  Else
     If Not ConSucursal Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  End If
  If PorConceptos Then
     sSQL = sSQL & "AND T.Codigo_C = Cl.Codigo "
  Else
     If CheckUsuario.value = 1 Then sSQL = sSQL & "AND C.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
     sSQL = sSQL _
          & "AND C.Codigo_B = Cl.Codigo " _
          & "AND T.TP = C.TP " _
          & "AND T.Numero = C.Numero " _
          & "AND T.Periodo = C.Periodo " _
          & "AND T.Fecha = C.Fecha " _
          & "AND T.Item = C.Item "
  End If
  sSQL = sSQL _
       & "GROUP BY T.Cta " _
       & "ORDER BY T.Cta "
  Select_Adodc AdoTrans, sSQL
  
 'Totalizamos mayores
  SumaDebe = 0: SumaHaber = 0: Suma_ME = 0: SaldoTotal = 0
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount + 1
       If .RecordCount = 1 Then
           Do While Not .EOF
              Cta = .Fields("Cta")
              Progreso_Barra.Mensaje_Box = "Totalozando Cta: " & Cta
              Progreso_Esperar
              Suma_ME = Suma_ME + .Fields("TParcial_ME")
              SumaDebe = SumaDebe + .Fields("TDebe")
              SumaHaber = SumaHaber + .Fields("THaber")
             .MoveNext
           Loop
          'Obtenemos el ultimo Saldo de la Cta
           If AdoMayor.Recordset.RecordCount > 0 Then
              AdoMayor.Recordset.MoveLast
              SaldoTotal = AdoMayor.Recordset.Fields("Saldo")
              AdoMayor.Recordset.MoveFirst
           End If
       End If
   End If
  End With
 'LabelSaldo_ME.Caption = Format(Suma_ME, "#,##0.00")
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SaldoTotal, "#,##0.00")
  
  SaldoAnterior = CalculosSaldoAnt(DCCtas.Text, SumaDebe, SumaHaber, SaldoTotal)
  LabelTotSaldoAnt.Caption = Format(SaldoAnterior, "#,##0.00")
  
  DGMayor.AllowUpdate = False
  DGMayor.Visible = True
  Progreso_Final
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key & " - " & BalanceCC
    RatonReloj
    DGMayor.Visible = False
    FechaValida MBoxFechaI
    FechaValida MBoxFechaF
    FechaIni = BuscarFecha(MBoxFechaI)
    FechaFin = BuscarFecha(MBoxFechaF)
    Select Case Button.key
      Case "Salir"
            RatonNormal
            Unload ListMayorizacion1
      Case "Imprimir"
            FechaCorte = "Desde " & MBoxFechaI & " al " & MBoxFechaF
            Select Case Opcion
              Case 1
                   Codigo1 = SinEspaciosIzq(DCCtas.Text)
                   ListarMayoresAux True, Individual
              Case 2
                   Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
                   Codigo2 = CambioCodigoCta(MBoxCtaF.Text)
                   If Codigo1 = " " Then Codigo1 = "1"
                   If Codigo2 = " " Then Codigo2 = "9"
                   ListarMayoresAux False, Individual
              Case Else: ListarMayoresAux True, Individual
            End Select
            Imprimir_Mayor AdoMayor, CheqImpSubCtas.value
            DGMayor.Caption = " Mayor de: " & DCCtas.Text & "."
      Case "Excel"
            DGMayor.Visible = False
            Exportar_AdoDB_Excel AdoMayor.Recordset, "Mayores " & BuscarFecha(MBoxFechaI) & " al " & BuscarFecha(MBoxFechaF)
           'GenerarDataTexto ListMayorizacion1, AdoMayor
            DGMayor.Visible = True
      Case "UnMayor"
            Codigo1 = SinEspaciosIzq(DCCtas.Text)
            ListarMayoresAux True, Individual
            ListMayorizacion1.Caption = "MAYOR ANALITICO"
            Opcion = 1
      Case "VariosMayores"
            RatonReloj
            Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
            Codigo2 = CambioCodigoCta(MBoxCtaF.Text)
            If Codigo1 = " " Then Codigo1 = "1"
            If Codigo2 = " " Then Codigo2 = "9"
            ListarMayoresAux False, Individual
            ListMayorizacion1.Caption = "MAYOR ANALITICO"
            Opcion = 2
      Case "Buscar"
            FPatronBusqueda.Show 1
            If SQLPatron <> "" Then
               LblPatron.Caption = "Patron Busqueda: " & SQLPatron
               LblPatron.Refresh
               Codigo1 = SinEspaciosIzq(DCCtas.Text)
               ListarMayoresAux True, Individual
               ListMayorizacion1.Caption = "MAYOR ANALITICO"
               Opcion = 1
            End If
    End Select
    If Button.key <> "Salir" Then
        DGMayor.Visible = True
        DGMayor.Caption = " Mayor de: " & DCCtas.Text & "."
        RatonNormal
        LblPatron.Caption = "Patron Busqueda: " & SQLPatron
        LblPatron.Refresh
        DCCtas.SetFocus
    End If
End Sub
