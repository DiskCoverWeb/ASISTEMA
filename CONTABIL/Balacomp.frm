VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form BalanceComp 
   Caption         =   "BALANCE DE COMPROBACION"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "Balacomp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
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
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Procesar_Balance"
            Object.ToolTipText     =   "Procesar Balance de Comprobación"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Procesar_Balance_Mensual"
            Object.ToolTipText     =   "Procesa Balance Mensual"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Procesar_Balance_Consolidado"
            Object.ToolTipText     =   "Procesa Balance Consolidado de Varias Sucursales/Agencias"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Presentar_Balance_Comprobacion"
            Object.ToolTipText     =   "Presenta Balance de Comprobación"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Presenta_Estado_Situacion"
            Object.ToolTipText     =   "Presenta Estado de Situación (General)"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Presenta_Estado_Resultado"
            Object.ToolTipText     =   "Presenta Estado de Resultado"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Presenta_Balance_Semanal"
            Object.ToolTipText     =   "Presenta Balance Mensual por semanas"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SBSB11"
            Object.ToolTipText     =   "SBS B11"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   6510
         TabIndex        =   1
         Top             =   0
         Width           =   6735
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
            Left            =   5460
            MaxLength       =   11
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   210
            Width           =   1170
         End
         Begin MSMask.MaskEdBox MBFechaI 
            Height          =   330
            Left            =   840
            TabIndex        =   3
            Top             =   210
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
         Begin MSMask.MaskEdBox MBFechaF 
            Height          =   330
            Left            =   2940
            TabIndex        =   5
            Top             =   210
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
         Begin VB.Label Label5 
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
            Left            =   4305
            TabIndex        =   6
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Hasta:"
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
            TabIndex        =   4
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Desde:"
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
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Presentacion de Cuentas"
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
         Left            =   13335
         TabIndex        =   8
         Top             =   0
         Width           =   3270
         Begin VB.OptionButton OpcDG 
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
            Height          =   225
            Left            =   105
            TabIndex        =   9
            Top             =   315
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OpcG 
            Caption         =   "Grupo"
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
            Left            =   1155
            TabIndex        =   10
            Top             =   315
            Width           =   855
         End
         Begin VB.OptionButton OpcD 
            Caption         =   "Detalle"
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
            Left            =   2205
            TabIndex        =   11
            Top             =   315
            Width           =   960
         End
      End
   End
   Begin MSDataGridLib.DataGrid DGBalance 
      Bindings        =   "Balacomp.frx":0442
      Height          =   4635
      Left            =   105
      TabIndex        =   20
      Top             =   1470
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   8176
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   210
      TabIndex        =   19
      Top             =   7665
      Width           =   225
   End
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   315
      Top             =   2940
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   315
      Top             =   3255
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Ctas1"
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
      Left            =   315
      Top             =   3570
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSAdodcLib.Adodc AdoBalance 
      Height          =   330
      Left            =   0
      Top             =   7140
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Balance"
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
   Begin MSAdodcLib.Adodc AdoCC 
      Height          =   330
      Left            =   315
      Top             =   3885
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "CC"
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
   Begin MSAdodcLib.Adodc AdoTotales 
      Height          =   330
      Left            =   315
      Top             =   4200
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Totales"
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
   Begin VB.Label LblTipoBalance 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   105
      TabIndex        =   21
      Top             =   735
      Width           =   13560
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   210
      Top             =   6405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":045B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":0775
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":0A8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":0DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":10C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":13B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":16CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":18A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":1E1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":2135
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":244F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Balacomp.frx":4AE1
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   9555
      TabIndex        =   13
      Top             =   7140
      Width           =   1905
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
      Left            =   8820
      TabIndex        =   16
      Top             =   7140
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
      Left            =   6930
      TabIndex        =   14
      Top             =   7140
      Width           =   1905
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
      Left            =   6195
      TabIndex        =   15
      Top             =   7140
      Width           =   750
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      TabIndex        =   18
      Top             =   7140
      Width           =   855
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
      Left            =   3675
      TabIndex        =   12
      Top             =   7140
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia:"
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
      Left            =   2625
      TabIndex        =   17
      Top             =   7140
      Width           =   1065
   End
End
Attribute VB_Name = "BalanceComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OpcionBalance As Byte
Dim OpcEsMensual  As Boolean
Dim Total_MN      As Currency
Dim BalanceCC     As String
Dim sSQL_Ext      As String

Public Sub ListarTipoDeBalance(EsBalanceMes As Boolean)
  RatonReloj
  FEsperar.Show
  Imagen_Esperar "Procesar Balance de Comprobacion"
  DGBalance.Visible = False
  TextCotiza = Format(Dolar, "#,##0.00")
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If EsBalanceMes Then
     sSQL = sSQL & "AND Detalle = 'Balance Mes' "
  Else
     sSQL = sSQL & "AND Detalle = 'Balance' "
  End If
  Select_Adodc AdoCtas, sSQL
  If AdoCtas.Recordset.RecordCount > 0 Then
     MBFechaI = AdoCtas.Recordset.fields("Fecha_Inicial")
     MBFechaF = AdoCtas.Recordset.fields("Fecha_Final")
  Else
     MBFechaI = FechaSistema
     MBFechaF = FechaSistema
  End If
  Imagen_Esperar "Procesar Balance de Comprobacion"
 'MsgBox "Listar Balance: " & MBFechaI
  FechaValida MBFechaI
  FechaValida MBFechaF
  Select Case OpcionBalance
    Case 1, 2, 4: SQLMsg1 = "BALANCE DE COMPROBACION "
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Anterior,Debitos,Creditos,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Debitos<>0 OR Creditos<>0 OR Saldo_Total<>0) "
    Case 5: SQLMsg1 = "BALANCE GENERAL "
         sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Total_N6+Total_N5+Total_N4+Total_N3+Total_N2+Total_N1)<>0 " _
              & "AND TB = 'ES' "
    Case 6: SQLMsg1 = "ESTADO DE RESULTADOS "
         sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Total_N6+Total_N5+Total_N4+Total_N3+Total_N2+Total_N1)<>0 " _
              & "AND TB = 'ER' "
    Case 11: SQLMsg1 = "BALANCE DE PROMEDIOS "
         If Opcion = 1 Then SQLMsg1 = "BALANCE CONSOLIDADO "
         TextoValido TextCotiza, True
         Dolar = Round(CSng(TextCotiza.Text), 2)
         If Dolar > 0 Then
            sSQL = "UPDATE Catalogo_Cuentas " _
                 & "SET Saldo_Total_ME = Saldo_Total / " & Dolar & " " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Ejecutar_SQL_SP sSQL
         Else
            sSQL = "UPDATE Catalogo_Cuentas " _
                 & "SET Saldo_Total_ME = 0 " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Ejecutar_SQL_SP sSQL
         End If
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Total_ME,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE DG = 'G' " _
              & "AND Saldo_Total <> 0 "
    Case 12: SQLMsg1 = "BALANCE DE COMPROBACION SBS B11 "
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Anterior,Debitos,Creditos,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE LEN(Codigo) <= 9 " _
              & "AND ISNUMERIC(MidStrg(Codigo,1,1)) <> " & Val(adFalse) & " "
  End Select
  
  If OpcG.value Then sSQL = sSQL & "AND DG = 'G' "
  If OpcD.value Then sSQL = sSQL & "AND DG = 'D' "
  sSQL = sSQL & "AND Codigo <> '" & Ninguno & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
  Imagen_Esperar "Procesar Balance de Comprobacion"
  RatonReloj
  If EsBalanceMes Then SQLMsg1 = SQLMsg1 & "MENSUAL "
  
  SumaDebe = 0: SumaHaber = 0
  sSQLTotales = "SELECT DG, SUM(Debitos) As TDebitos, SUM(Creditos) As TCreditos " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND (Debitos + Creditos) <> 0 " _
              & "AND DG = 'D' " _
              & "GROUP BY DG "
  Select_Adodc AdoTotales, sSQLTotales
  With AdoTotales.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .fields("TDebitos")
          SumaHaber = SumaHaber + .fields("TCreditos")
         .MoveNext
       Loop
   End If
  End With
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  BalanceComp.Caption = "ESTADOS FINANCIEROS"
  LblTipoBalance.Caption = SQLMsg1 & vbCrLf & "DEL  " & FechaStrgCorta(MBFechaI) & "  AL  " & FechaStrgCorta(MBFechaF)
  DGBalance.Visible = True
  RatonNormal
  Unload FEsperar
End Sub

'''Public Sub ListarTipoDeBalance_Exterior(EsBalanceMes As Boolean, _
'''                                        TipoBalance As String)
'''  RatonReloj
'''  Progreso_Barra.Incremento = 0
'''  Progreso_Barra.Valor_Maximo = 100
'''  Progreso_Barra.Mensaje_Box = "Consultando el Balance"
'''  Progreso_Esperar
'''  DGBalance.Visible = False
'''  TextCotiza = Format(Dolar, "#,##0.00")
'''  sSQL = "SELECT * " _
'''       & "FROM Fechas_Balance " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  If EsBalanceMes Then
'''     sSQL = sSQL & "AND Detalle = 'Balance Mes' "
'''  Else
'''     sSQL = sSQL & "AND Detalle = 'Balance' "
'''  End If
'''  Select_Adodc AdoCtas, sSQL
'''  If AdoCtas.Recordset.RecordCount > 0 Then
'''     MBFechaI = AdoCtas.Recordset.Fields("Fecha_Inicial")
'''     MBFechaF = AdoCtas.Recordset.Fields("Fecha_Final")
'''  Else
'''     MBFechaI = FechaSistema
'''     MBFechaF = FechaSistema
'''  End If
'''  FechaValida MBFechaI
'''  FechaValida MBFechaF
'''  Progreso_Esperar
'''
'''  sSQL = "SELECT DG,Codigo_Ext, Cuenta, Saldo_Anterior, Debitos, Creditos, Saldo_Mes, Saldo_Total, TC " _
'''       & "FROM Catalogo_Cuentas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  Select Case TipoBalance
'''    Case "BS"
'''         sSQL = sSQL & "AND CC BETWEEN '1' and '3' "
'''         MensajeEncabData = "B A L A N C E   G E N E R A L"
'''    Case "BR"
'''         sSQL = sSQL & "AND CC BETWEEN '4' and '9' "
'''         MensajeEncabData = "B A L A N C E   D E   R E S U L T A D O"
'''    Case Else
'''         sSQL = sSQL & "AND (Debitos<>0 OR Creditos<>0 OR Saldo_Total<>0) "
'''         MensajeEncabData = "B A L A N C E   D E   C O M P R O B A C I O N"
'''  End Select
'''  sSQL = sSQL & "AND (Total_N6+Total_N5+Total_N4+Total_N3+Total_N2+Total_N1)<>0 "
'''  If OpcG.value Then sSQL = sSQL & "AND DG = 'G' "
'''  If OpcD.value Then sSQL = sSQL & "AND DG = 'D' "
'''  sSQL = sSQL & "AND Codigo <> '" & Ninguno & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
''' 'Recojemos la consulta que se realizo para el tipo de balance
'''  sSQL_Ext = sSQL
'''  Progreso_Esperar
'''  RatonReloj
'''  If EsBalanceMes Then SQLMsg1 = SQLMsg1 & "MENSUAL "
'''  SumaDebe = 0: SumaHaber = 0
'''  If OpcionBalance = 1 Or OpcionBalance = 4 Then
'''     With AdoBalance.Recordset
'''      If .RecordCount > 0 Then
'''          Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
'''          Progreso_Barra.Mensaje_Box = "Consultando " & ULCase(SQLMsg1)
'''          Progreso_Esperar
'''          Do While Not .EOF
'''             SumaDebe = SumaDebe + .Fields("Debitos")
'''             SumaHaber = SumaHaber + .Fields("Creditos")
'''            ' Progreso_Esperar
'''            .MoveNext
'''          Loop
'''         .MoveFirst
'''      End If
'''     End With
'''  End If
'''  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
'''  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
'''  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
'''  BalanceComp.Caption = "ESTADOS FINANCIEROS"
'''  LblTipoBalance.Caption = MensajeEncabData & vbCrLf & "DEL  " & FechaStrgCorta(MBFechaI) & "  AL  " & FechaStrgCorta(MBFechaF)
'''  DGBalance.Visible = True
'''  Progreso_Final
'''  RatonNormal
'''End Sub

Public Sub ListarTipoDeBalance_Ext(EsBalanceMes As Boolean, _
                                   TipoBalance As String, _
                                   TipoPyGCC As String)
    RatonReloj
    FEsperar.Show
    Imagen_Esperar "Consultando el Balance"
    DGBalance.Visible = False
    TextCotiza = Format(Dolar, "#,##0.00")
    sSQL = "SELECT * " _
         & "FROM Fechas_Balance " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    If EsBalanceMes Then
       sSQL = sSQL & "AND Detalle = 'Balance Mes' "
    Else
       sSQL = sSQL & "AND Detalle = 'Balance' "
    End If
    Select_Adodc AdoCtas, sSQL
    If AdoCtas.Recordset.RecordCount > 0 Then
       MBFechaI = AdoCtas.Recordset.fields("Fecha_Inicial")
       MBFechaF = AdoCtas.Recordset.fields("Fecha_Final")
    Else
       MBFechaI = FechaSistema
       MBFechaF = FechaSistema
    End If
    FechaValida MBFechaI
    FechaValida MBFechaF

    RatonReloj
    sSQL = SQL_Tipo_Balance(TipoBalance, TipoPyGCC)
   'MsgBox sSQL
    Select_Adodc_Grid DGBalance, AdoBalance, sSQL
    sSQL_Ext = sSQL

    Select Case TipoBalance
      Case "BC"
           MensajeEncabData = "BALANCE DE COMPROBACION"
      Case "BS"
           MensajeEncabData = "BALANCE GENERAL"
      Case "BR"
           MensajeEncabData = "ESTADO DE RESULTADO"
''           Codigo = SinEspaciosIzq(DCCC)
''           Codigo = MidStrg(DCCC, Len(Codigo) + 1, Len(DCCC))
''           Codigo = TrimStrg(Replace(Codigo, "-", ""))
''           MensajeEncabData = Codigo
      Case Else
           MensajeEncabData = "BALANCE NO DEFINIDO"
    End Select
    SQLMsg1 = "D E L  " & MBFechaI & "  A L  " & MBFechaI

  RatonReloj
  If EsBalanceMes Then SQLMsg1 = SQLMsg1 & "MENSUAL "
  Imagen_Esperar "Consultando el Balance"
  SumaDebe = 0: SumaHaber = 0
  Select_Adodc AdoTotales, sSQLTotales
  With AdoTotales.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .fields("TDebitos")
          SumaHaber = SumaHaber + .fields("TCreditos")
         .MoveNext
       Loop
   End If
  End With
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  BalanceComp.Caption = "ESTADOS FINANCIEROS"
  LblTipoBalance.Caption = MensajeEncabData & vbCrLf & SQLMsg1
  DGBalance.Caption = ""
  DGBalance.Visible = True
  RatonNormal
  Unload FEsperar
End Sub

'''Private Sub Imprimir_parcial()
'''  RatonReloj
'''  DGBalance.Visible = False
'''  FechaIni = MBFechaI.Text
'''  FechaFin = MBFechaF.Text
'''  SQLMsg2 = "AL " & FechaStrg(FechaFin)
'''  Select Case SSTbBalances.Tab
'''    Case 2:
'''         If OpcCoop Then
'''            Imprimir_General_Con AdoBalance, 2, True
'''         Else
'''            Imprimir_General AdoBalance, 2
'''         End If
'''    Case 3
'''         If OpcCoop Then
'''            Imprimir_General_Con AdoBalance, 2, False
'''         Else
'''            Imprimir_General AdoBalance, 2
'''         End If
'''  End Select
'''  DGBalance.Visible = True
'''  RatonNormal
'''End Sub

Private Sub Saldo_Promedio()
Dim NumDia As Integer
Dim Saldos_Prom_MN(31) As Currency
Dim Saldos_Prom_ME(31) As Currency
  RatonReloj
  DGBalance.Visible = False
  FechaValida MBFechaI
  FechaValida MBFechaF
  NumDia = Day(MBFechaF)
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "DELETE * " _
       & "FROM Saldo_Promedios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Saldo_Promedios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc AdoCtas1, sSQL
  sSQL = "SELECT Cta, Fecha, Saldo, Saldo_ME " _
       & "FROM Transacciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T <> 'A' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Cta, Fecha "
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       For I = 1 To 31
           Saldos_Prom_MN(I) = 0
           Saldos_Prom_ME(I) = 0
       Next
       Codigo = .fields("Cta")
       Do While Not .EOF
          If Codigo <> .fields("Cta") Then
             Cadena = "Codigo: " & Codigo & vbCrLf
             Saldo = Saldos_Prom_MN(1)
             Saldo_ME = Saldos_Prom_ME(1)
             For I = 1 To NumDia
                 If Saldos_Prom_MN(I) <> 0 Then
                    Saldo = Saldos_Prom_MN(I)
                 Else
                    Saldos_Prom_MN(I) = Saldo
                 End If
                 If Saldos_Prom_ME(I) <> 0 Then
                    Saldo_ME = Saldos_Prom_ME(I)
                 Else
                    Saldos_Prom_ME(I) = Saldo_ME
                 End If
             Next
             Saldo = 0
             Saldo_ME = 0
             For I = 1 To 31
                 Saldo = Saldo + Saldos_Prom_MN(I)
                 Saldo_ME = Saldo_ME + Saldos_Prom_ME(I)
             Next
             SetAddNew AdoCtas1
             SetFields AdoCtas1, "Codigo", Codigo
             SetFields AdoCtas1, "Saldo_MN", Round(Saldo / NumDia, 2)
             SetFields AdoCtas1, "Saldo_ME", Round(Saldo_ME / NumDia, 2)
             SetFields AdoCtas1, "Item", NumEmpresa
             SetFields AdoCtas1, "CodigoU", CodigoUsuario
             SetUpdate AdoCtas1
             For I = 1 To 31
                 Saldos_Prom_MN(I) = 0
                 Saldos_Prom_ME(I) = 0
             Next
             Codigo = .fields("Cta")
          End If
          Dia = Day(.fields("Fecha"))
          Saldos_Prom_MN(Dia) = .fields("Saldo")
          Saldos_Prom_ME(Dia) = .fields("Saldo_ME")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT SP.Codigo,C.Cuenta,SP.Saldo_MN,SP.Saldo_ME " _
       & "FROM Saldo_Promedios As SP,Catalogo_Cuentas As C " _
       & "WHERE C.Codigo = SP.Codigo " _
       & "AND SP.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND SP.CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY SP.Codigo "
  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
  Opcion = 2
  DGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command1_Click()
  Unload BalanceComp
End Sub

Private Sub Form_Activate()
  MBFechaI = FechaSistema
  MBFechaF = FechaSistema
  Eliminar_Nulos_SP "Catalogo_Cuentas"
  BalanceCC = Ninguno
'''  sSQL = "SELECT (Codigo & ' - ' & Detalle) CentroDeCosto " _
'''       & "FROM Catalogo_SubCtas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TC = 'CCX' " _
'''       & "ORDER BY Agrupacion DESC,Nivel,Codigo, Detalle "
'''  SelectDB_Combo DCCC, AdoCC, sSQL, "CentroDeCosto"
  
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET CC = SUBSTRING(Codigo,1,1) " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CC = '.' "
  Ejecutar_SQL_SP sSQL
  
  DGBalance.Height = MDI_Y_Max - DGBalance.Top - 300
  DGBalance.width = MDI_X_Max - DGBalance.Left
  LblTipoBalance.width = MDI_X_Max - LblTipoBalance.Left
  AdoBalance.Top = DGBalance.Top + DGBalance.Height + 30
   
  Label3.Top = DGBalance.Top + DGBalance.Height + 30
  Label6.Top = DGBalance.Top + DGBalance.Height + 30
  Label9.Top = DGBalance.Top + DGBalance.Height + 30
  Label11.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotSaldo.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotDebe.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotHaber.Top = DGBalance.Top + DGBalance.Height + 30
  OpcionBalance = 4
  OpcEsMensual = False
  
  Toolbar1.buttons("Procesar_Balance_Consolidado").Enabled = ConSucursal
  If Bloquear_Control Then
     Toolbar1.buttons("Procesar_Balance").Enabled = False
     Toolbar1.buttons("Procesar_Balance_Mensual").Enabled = False
  End If
  If CNivel(7) Then
     RatonNormal
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload BalanceComp
  Else
     ListarTipoDeBalance OpcEsMensual
     RatonNormal
  End If
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCC
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCtas1
  ConectarAdodc AdoTrans
 'ConectarAdodc AdoBalGen
 'ConectarAdodc AdoEstRes
 'ConectarAdodc AdoFechaBal
  ConectarAdodc AdoTotales
 'ConectarAdodc AdoBalGenCon
  ConectarAdodc AdoBalance
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextCotiza_GotFocus()
  MarcarTexto TextCotiza
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 Then
     DGBalance.Visible = False
     sSQL = "SELECT * " _
          & "FROM Transacciones " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Fecha "
     Select_Adodc AdoBalance, sSQL
     RatonReloj
     With AdoBalance.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             BalanceComp.Caption = .fields("Fecha")
            .fields("Debe") = Redondear(.fields("Debe"), 2)
            .fields("Haber") = Redondear(.fields("Haber"), 2)
            .fields("Parcial_ME") = Redondear(.fields("Parcial_ME"), 2)
            .Update
            .MoveNext
          Loop
      End If
     End With
     DGBalance.Visible = True
     RatonNormal
  End If
End Sub

Public Sub Procesar_Balance_Consolidado()
Dim ID_Campo As Integer
Dim ID_Max As Integer
Dim CamposEmp() As Nodo_Arbol
Dim CampoEmpresa As String
Dim CampoTotal As String
  RatonReloj
 'Borramos la Tabla temporar del Balance
  SQL1 = "DROP TABLE [Balance_Consolidado] "
  Ejecutar_SQL_SP SQL1
  SQL1 = "CREATE TABLE [Balance_Consolidado] ("
  If SQL_Server Then
     SQL1 = SQL1 _
          & "[Item] NVARCHAR (3) NULL, " _
          & "[Periodo] NVARCHAR (10) NULL, " _
          & "[TC] NVARCHAR (2) NULL, " _
          & "[DG] NVARCHAR (1) NULL, " _
          & "[Codigo] NVARCHAR (16) NULL, " _
          & "[Cuenta] NVARCHAR (80) NULL, " _
          & "[TB] NVARCHAR (3) NULL, " _
          & "[Ln] INT NULL "
  Else
     SQL1 = SQL1 _
          & "[Item] TEXT(3) NULL, " _
          & "[Periodo] TEXT(10) NULL, " _
          & "[TC] TEXT(2) NULL, " _
          & "[DG] TEXT(1) NULL, " _
          & "[Codigo] TEXT(16) NULL, " _
          & "[Cuenta] TEXT(80) NULL, " _
          & "[TB] TEXT(3) NULL, " _
          & "[Ln] LONG NULL "
  End If
  SQL1 = SQL1 & ");"
  Ejecutar_SQL_SP SQL1
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoCtas, sSQL
       
  sSQL = "SELECT * " _
       & "FROM Empresas " _
       & "WHERE Item <> '000' " _
       & "ORDER BY Sucursal,Item "
  Select_Adodc AdoCtas1, sSQL
  ID_Max = AdoCtas1.Recordset.RecordCount
  ReDim CamposEmp(ID_Max) As Nodo_Arbol
  ID_Campo = 0
  With AdoCtas1.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If CBool(.fields("Sucursal")) Then
             CampoEmpresa = "E_TOTAL"
             CampoTotal = "E_" & .fields("Item")
          Else
             CampoEmpresa = .fields("Abreviatura")
          End If
          If CampoEmpresa = Ninguno Then CampoEmpresa = "EMP_" & .fields("Item")
          CamposEmp(ID_Campo).Codigo_Aux = CampoEmpresa
          CamposEmp(ID_Campo).Item_Nodo = "E_" & .fields("Item")
          CamposEmp(ID_Campo).Valor = 0
    
         'Creamos los campos numericos para los totales del balance
          SQL1 = "ALTER TABLE Balance_Consolidado "
          If SQL_Server Then
             SQL1 = SQL1 & "ADD [" & CamposEmp(ID_Campo).Item_Nodo & "] MONEY NULL;"
          Else
             SQL1 = SQL1 & "ADD [" & CamposEmp(ID_Campo).Item_Nodo & "] CURRENCY NULL;"
          End If
          Ejecutar_SQL_SP SQL1
          ID_Campo = ID_Campo + 1
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Balance_Consolidado " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoTrans, sSQL
  RatonReloj
 'Creamos Todo el Catalogo que necesitenos
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If AdoTrans.Recordset.RecordCount > 0 Then
             AdoTrans.Recordset.MoveFirst
             AdoTrans.Recordset.Find ("Codigo = '" & .fields("Codigo") & "' ")
             If AdoTrans.Recordset.EOF Then
                SetAdoAddNew "Balance_Consolidado"
                SetAdoFields "Codigo", .fields("Codigo")
                SetAdoFields "Cuenta", .fields("Cuenta")
                SetAdoFields "TB", .fields("TB")
                SetAdoFields "DG", .fields("DG")
                SetAdoFields "TC", .fields("TC")
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "Periodo", Periodo_Contable
                SetAdoUpdate
             End If
          Else
             SetAdoAddNew "Balance_Consolidado"
             SetAdoFields "Codigo", .fields("Codigo")
             SetAdoFields "Cuenta", TrimStrg(MidStrg(.fields("Cuenta"), 1, 60))
             SetAdoFields "TB", .fields("TB")
             SetAdoFields "DG", .fields("DG")
             SetAdoFields "TC", .fields("TC")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Periodo", Periodo_Contable
             SetAdoUpdate
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "UPDATE Balance_Consolidado SET "
  For ID_Campo = 0 To ID_Max - 1
      sSQL = sSQL & CamposEmp(ID_Campo).Item_Nodo & " = 0, "
      CamposEmp(ID_Campo).Valor = 0
  Next ID_Campo
  sSQL = sSQL & "Ln = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT Item,Cta,SUM(Debe) As TDebe,SUM(Haber) As THaber " _
       & "FROM Transacciones " _
       & "WHERE T <> '" & Anulado & "' " _
       & "AND TP IN ('CD','CE','CI','ND','NC') " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Item,Cta " _
       & "ORDER BY Item,Cta "
  Select_Adodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Cta = .fields("Cta")
          Codigo = "E_" & .fields("Item")
          If OpcCoop Then
             Select Case MidStrg(Cta, 1, 1)
               Case "4"
                    Total_MN = .fields("TDebe") - .fields("THaber")
               Case "5"
                    Total_MN = .fields("THaber") - .fields("TDebe")
             End Select
          Else
             Select Case MidStrg(Cta, 1, 1)
               Case "5"
                    Total_MN = .fields("TDebe") - .fields("THaber")
               Case "4"
                    Total_MN = .fields("THaber") - .fields("TDebe")
             End Select
          End If
          Select Case MidStrg(Cta, 1, 1)
            Case "1", "6", "8"
                 Total_MN = .fields("TDebe") - .fields("THaber")
            Case "2", "3", "7", "9"
                 Total_MN = .fields("THaber") - .fields("TDebe")
          End Select
          Total_MN = Redondear(Total_MN, 2)
          For ID_Campo = 0 To ID_Max - 1
              If Codigo = CamposEmp(ID_Campo).Item_Nodo Then
                 CamposEmp(ID_Campo).Valor = CamposEmp(ID_Campo).Valor + Total_MN
                 ID_Campo = ID_Max
              End If
          Next ID_Campo
          CamposEmp(ID_Max - 1).Valor = CamposEmp(ID_Max - 1).Valor + Total_MN
          If Total_MN <> 0 Then
            'Actualizamos las cuentas generales
             Cta_Sup = Cta
             Do While (Cta_Sup <> "0")
                sSQL = "UPDATE Balance_Consolidado " _
                     & "SET " & Codigo & " = " & Codigo & " + " & Total_MN & " " _
                     & "WHERE Codigo = '" & Cta_Sup & "' " _
                     & "AND Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' "
                Ejecutar_SQL_SP sSQL
                
                sSQL = "UPDATE Balance_Consolidado " _
                     & "SET " & CampoTotal & " = " & CampoTotal & " + " & Total_MN & " " _
                     & "WHERE Codigo = '" & Cta_Sup & "' " _
                     & "AND Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' "
                Ejecutar_SQL_SP sSQL
                Cta_Sup = CodigoCuentaSup(Cta_Sup)
             Loop
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT DG,Codigo,Cuenta,"
  For ID_Campo = 0 To ID_Max - 1
      If CamposEmp(ID_Campo).Valor <> 0 Then
         sSQL = sSQL & CamposEmp(ID_Campo).Item_Nodo & " As " & CamposEmp(ID_Campo).Codigo_Aux & ","
      End If
  Next ID_Campo
  sSQL = sSQL & "TC " _
       & "FROM Balance_Consolidado " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
  RatonReloj
  BalanceComp.Caption = "ESTADOS FINANCIEROS CONSOLIDADOS"
 'DGBalance.Caption = SQLMsg1 & "DEL  " & FechaStrgCorta(MBFechaI) & "  AL  " & FechaStrgCorta(MBFechaF)
  DGBalance.Visible = True
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    BalanceCC = ""
''    If CheqBalExt.Value <> 0 Then
''       BalanceCC = SinEspaciosIzq(DCCC)
''    Else
''       BalanceCC = "00"
''    End If
''    If BalanceCC = "" Then BalanceCC = "00"
    FechaValida MBFechaI
    FechaValida MBFechaF
   'MsgBox Button.key & " - " & BalanceCC
    Select Case Button.key
      Case "Procesar_Balance"
            OpcionBalance = 1
            OpcEsMensual = False
           'Procesar_Balance OpcEsMensual
           'MsgBox BalanceCC
            Procesar_Balance_SP OpcEsMensual, MBFechaI, MBFechaF, BalanceCC
''            If CheqBalExt.Value <> 0 Then
''              'ListarTipoDeBalance_Exterior OpcEsMensual, "BC"
''               Procesar_Balance_Ext_SP
''               ListarTipoDeBalance_Ext OpcEsMensual, "BC", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BC", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Procesar_Balance_Mensual"
            OpcionBalance = 2
            OpcEsMensual = True
            MBFechaI = "01/" & Format(Month(MBFechaI), "00") & "/" & Format(Year(MBFechaI), "0000")
            MBFechaF = UltimoDiaMes(MBFechaI)
           'Procesar_Balance OpcEsMensual
            Procesar_Balance_SP OpcEsMensual, MBFechaI, MBFechaF, BalanceCC
            ListarTipoDeBalance OpcEsMensual
      Case "Procesar_Balance_Consolidado"
            OpcionBalance = 3
            Procesar_Balance_Consolidado
      Case "Presentar_Balance_Comprobacion"
            OpcionBalance = 4
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BC"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BC", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BC", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Presenta_Estado_Situacion"
           'TipoCourierNew
            OpcionBalance = 5
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BS"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BS", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BS", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Presenta_Estado_Resultado"
            OpcionBalance = 6
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BR"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BR", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BR", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
           ' MsgBox "ER..."
      Case "Presenta_Balance_Mensual"
            OpcionBalance = 7
      Case "Presenta_Balance_Semanal"
            OpcionBalance = 8
      Case "Imprimir"
            RatonReloj
            DGBalance.Visible = False
            FechaIni = MBFechaI
            FechaFin = MBFechaF
            If OpcEsMensual Then
               SQLMsg2 = "DEL " & FechaStrg(FechaIni) & " AL " & FechaStrg(FechaFin)
            Else
               SQLMsg2 = "AL " & FechaStrg(FechaFin)
            End If
            Select Case OpcionBalance
              Case 1, 2, 4: Imprimir_Balance AdoBalance
              Case 1, 2, 3: Imprimir_General_Con AdoBalance, Opcion, True
              Case 1, 2, 5: Imprimir_General AdoBalance, 1
              Case 1, 2, 6: Imprimir_General AdoBalance, 1
            End Select
            RatonNormal
            DGBalance.Visible = True
      Case "SBSB11"
           OpcionBalance = 12
           OpcEsMensual = False
           ListarTipoDeBalance OpcEsMensual
           Generar_SBS_B11
      Case "Salir"
           Unload BalanceComp
      Case "Excel"
           DGBalance.Visible = False
           GenerarDataTexto BalanceComp, AdoBalance
           DGBalance.Visible = True
    End Select
''    DGBalance.Visible = True
''    DGBalance.Visible = False
    RatonNormal
End Sub

'''Public Sub Procesar_Balance(EsBalanceMes As Boolean)
'''  RatonReloj
'''  DGBalance.Visible = False
'''  FechaValida MBFechaI
'''  FechaValida MBFechaF
'''  Ln_No = 0
'''  For I = 1 To 9
'''      Select Case I
'''        Case 1, 2, 3: sSQL = "UPDATE Catalogo_Cuentas SET TB = 'ES' "
'''        Case 4, 5, 6: sSQL = "UPDATE Catalogo_Cuentas SET TB = 'ER' "
'''        Case Else:    sSQL = "UPDATE Catalogo_Cuentas SET TB = 'EO' "
'''      End Select
'''      sSQL = sSQL _
'''           & "WHERE Item = '" & NumEmpresa & "' " _
'''           & "AND Periodo = '" & Periodo_Contable & "' " _
'''           & "AND MidStrg(Codigo,1,1)= '" & CStr(I) & "' "
'''      Ejecutar_SQL_SP sSQL
'''  Next I
'''  sSQL = "DELETE * " _
'''       & "FROM Catalogo_Cuentas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND MidStrg(Codigo,1,1) = 'x' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  SetAdoAddNew "Catalogo_Cuentas"
'''  SetAdoFields "Codigo", "x"
'''  SetAdoFields "Codigo_Ext", "x"
'''  SetAdoFields "Cuenta", " - TOTAL PASIVO Y PATRIMONIO"
'''  SetAdoFields "TB", "ES"
'''  SetAdoFields "DG", "G"
'''  SetAdoFields "TC", "N"
'''  SetAdoFields "Periodo", Periodo_Contable
'''  SetAdoUpdate
'''
'''  SetAdoAddNew "Catalogo_Cuentas"
'''  SetAdoFields "Codigo", "xx"
'''  SetAdoFields "Codigo_Ext", "xx"
'''  SetAdoFields "Cuenta", " - UTILIDAD/EXCEDENTE(Pérdida) DEL PERIODO"
'''  SetAdoFields "TB", "ES"
'''  SetAdoFields "DG", "G"
'''  SetAdoFields "TC", "N"
'''  SetAdoFields "Periodo", Periodo_Contable
'''  SetAdoUpdate
'''
'''  SetAdoAddNew "Catalogo_Cuentas"
'''  SetAdoFields "Codigo", "x"
'''  SetAdoFields "Codigo_Ext", "x"
'''  SetAdoFields "Cuenta", " - UTILIDAD/EXCEDENTE(Pérdida) DEL PERIODO"
'''  SetAdoFields "TB", "ER"
'''  SetAdoFields "DG", "G"
'''  SetAdoFields "TC", "N"
'''  SetAdoFields "Periodo", Periodo_Contable
'''  SetAdoUpdate
'''  Procesar_Balance_Comprobacion BalanceComp, MBFechaI, MBFechaF, AdoCtas, AdoTrans, EsBalanceMes
'''  If EsBalanceMes Then
'''     Fecha_Procesos "Balance Mes", MBFechaI, MBFechaF
'''  Else
'''     Fecha_Procesos "Balance", MBFechaI, MBFechaF
'''  End If
'''' Listar Balance de Comprobacion
'''  ListarTipoDeBalance EsBalanceMes
'''  RatonNormal
'''End Sub

Public Sub Generar_SBS_B11()
'cas408902
'milcambios4089
Dim NumFile As Long
Dim NombreArchivo As String
Dim TextoLinea As String
  RatonReloj
  DGBalance.Visible = False
  With AdoBalance.Recordset
   If .RecordCount > 0 Then
       CodigoDelBanco = "4089"
       Total = 0
       TotalActivo = 0
       TotalPasivo = 0
       TotalCapital = 0
       TotalIngreso = 0
       TotalEgreso = 0
       Do While Not .EOF
          Select Case MidStrg(.fields("Codigo"), 1, 1)
            Case "1": TotalActivo = TotalActivo + .fields("Saldo_Total")
            Case "2": TotalPasivo = TotalPasivo + .fields("Saldo_Total")
            Case "3": TotalCapital = TotalCapital + .fields("Saldo_Total")
            Case "4": TotalIngreso = TotalIngreso + .fields("Saldo_Total")
            Case "5", "6": TotalEgreso = TotalEgreso + .fields("Saldo_Total")
            Case Else: TotalAbonos = TotalAbonos + .fields("Saldo_Total")
          End Select
         .MoveNext
       Loop
       Total = TotalActivo + TotalPasivo + TotalCapital + TotalIngreso + TotalEgreso + TotalAbonos
      .MoveFirst
       NombreArchivo = RutaSysBases & "\SBS\B11M" & CodigoDelBanco & Replace(MBFechaF, "/", "") & ".txt"
       NumFile = FreeFile
       Open NombreArchivo For Output As #NumFile
       TextoLinea = "B11" & vbTab & CodigoDelBanco & vbTab & MBFechaF & vbTab & .RecordCount + 1 & vbTab & Format(Total, "#0.00")
       Print #NumFile, TextoLinea
       Do While Not .EOF
          TextoLinea = Replace(.fields("Codigo"), ".", "") & vbTab & Format(.fields("Saldo_Total"), "#0.00")
          Print #NumFile, TextoLinea
         .MoveNext
       Loop
       Close #NumFile
   End If
  End With
  RatonNormal
  DGBalance.Visible = True
  MsgBox "El Archivo: " & NombreArchivo & "," & vbCrLf & "Fue Generado exitosamente"
End Sub

