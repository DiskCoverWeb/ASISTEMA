VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form SaldoSubCtasVence 
   Caption         =   "ESTADO DE CUENTAS"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11280
   DrawMode        =   5  'Not Copy Pen
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
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
            Key             =   "Consultar"
            Object.ToolTipText     =   "Consultar SubModulo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Costos"
            Object.ToolTipText     =   "Presenta Resumen de Costos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Temporizada"
            Object.ToolTipText     =   "Presenta en forma temporizadas las CxC o CxP"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxCxP_Mes"
            Object.ToolTipText     =   "Presenta Resumen de Saldos por Meses"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel el Resultado"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame3 
         Caption         =   "Fechas Desde - Hasta"
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
         Width           =   4530
         Begin VB.ComboBox CTC 
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
            Left            =   2835
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   210
            Width           =   1590
         End
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   1470
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
         Begin MSMask.MaskEdBox MBoxFechaI 
            Height          =   330
            Left            =   105
            TabIndex        =   2
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Facturas"
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
         Left            =   8820
         TabIndex        =   5
         Top             =   0
         Width           =   7680
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
            Height          =   300
            Left            =   2940
            TabIndex        =   8
            Top             =   210
            Width           =   960
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
            Left            =   105
            TabIndex        =   6
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton OpcC 
            Caption         =   "Canceladas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1470
            TabIndex        =   7
            Top             =   210
            Width           =   1380
         End
         Begin VB.CheckBox CheqDSubCta 
            Caption         =   "Procesar con Detalle de SubModulo"
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
            TabIndex        =   9
            Top             =   210
            Width           =   3480
         End
      End
   End
   Begin MSDataGridLib.DataGrid DGBanco 
      Bindings        =   "SaldSbC1.frx":0000
      Height          =   3270
      Left            =   105
      TabIndex        =   22
      Top             =   1680
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   5768
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
      Height          =   300
      Left            =   210
      TabIndex        =   21
      Top             =   1680
      Width           =   435
   End
   Begin MSDataListLib.DataCombo DCDet 
      Bindings        =   "SaldSbC1.frx":0017
      DataSource      =   "AdoDet"
      Height          =   345
      Left            =   1365
      TabIndex        =   15
      Top             =   1260
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
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
   Begin VB.CheckBox CheqDet 
      Caption         =   "Por Det."
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
      Top             =   1260
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCCta 
      Bindings        =   "SaldSbC1.frx":002C
      DataSource      =   "AdoCtas"
      Height          =   345
      Left            =   1365
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "SaldSbC1.frx":0042
      DataSource      =   "AdoSubCta"
      Height          =   345
      Left            =   8505
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
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
   Begin VB.PictureBox PictFactura 
      Height          =   330
      Left            =   10605
      ScaleHeight     =   270
      ScaleWidth      =   4785
      TabIndex        =   20
      Top             =   6300
      Width           =   4845
   End
   Begin VB.CheckBox CheqCta 
      Caption         =   "Por Cta."
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
      Top             =   840
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   315
      Top             =   2940
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
      Caption         =   "SubCta"
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
   Begin VB.CheckBox CheqIndiv 
      Caption         =   "Por Benef."
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
      TabIndex        =   12
      Top             =   840
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc AdoDetCheq 
      Height          =   330
      Left            =   315
      Top             =   3255
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
      Caption         =   "DetCheq"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   6300
      Width           =   4950
      _ExtentX        =   8731
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   3885
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
   Begin MSAdodcLib.Adodc AdoDet 
      Height          =   330
      Left            =   315
      Top             =   4200
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
      Caption         =   "Det"
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
      Left            =   12705
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
            Picture         =   "SaldSbC1.frx":005A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":0374
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":068E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":09A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":0CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":0FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaldSbC1.frx":12F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelSaldo 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8820
      TabIndex        =   16
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo MN"
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
      Left            =   7875
      TabIndex        =   17
      Top             =   6300
      Width           =   960
   End
   Begin VB.Label LabelTotal 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6090
      TabIndex        =   19
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total MN"
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
      Left            =   5145
      TabIndex        =   18
      Top             =   6300
      Width           =   960
   End
End
Attribute VB_Name = "SaldoSubCtasVence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqCta_Click()
    If CheqCta.value = 1 Then DCCta.Visible = True Else DCCta.Visible = False
End Sub

Private Sub CheqDet_Click()
    If CheqDet.value = 1 Then DCDet.Visible = True Else DCDet.Visible = False
End Sub

Private Sub CheqIndiv_Click()
    If CheqIndiv.value = 1 Then DCCtas.Visible = True Else DCCtas.Visible = False
End Sub

Public Sub ImprimirCxC_P_G_I_CC()
  DGBanco.Visible = False
''  If Opcion <> 10 Then
''     If SSTab1.Tab = 0 Then Opcion = 1 Else Opcion = 2
''  End If
  SQLMsg2 = "Desde:  " & MBoxFechaI.Text & "   al   " & MBoxFechaF.Text
  If OpcP.value Then SQLMsg3 = "FACTURAS PENDIENTES" Else SQLMsg3 = "FACTURAS CANCELADAS"
 'MsgBox Opcion
  Select Case Opcion
    Case 1 ' CxC y CxP
         Select Case TipoCta
           Case "C", "P"
                If CheqDSubCta.value = 0 Then
                   Imprimir_Saldos_SubCtas_Vence AdoBanco, 1, True, True
                Else
                   Imprimir_Saldos_SubCtas_Vence AdoBanco, 1, True, False
                End If
         End Select
    Case 2
         Imprimir_Saldos_SubCtas_Vence_Temporizada AdoAux
    Case 9
         Imprimir_CxCxP_Meses AdoBanco, SinEspaciosIzq(DCCta), MBoxFechaF
    Case 10
         Imprimir_Saldos_SubCtas_Costos AdoBanco, CheqCta
    Case Else
         Imprimir_Saldos_SubCtas_IE AdoBanco, 1, True, CheqDSubCta.value
  End Select
  DGBanco.Visible = True
End Sub

Public Sub Costos()
  RatonReloj
  DGBanco.Visible = False
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaInicial = BuscarFecha(MBoxFechaI)
  FechaFinal = BuscarFecha(MBoxFechaF)
  CodigoCli = Ninguno
  LabelSaldo.Caption = "0.00"
  LabelTotal.Caption = "0.00"
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCtas.Text & "' ")
       If Not .EOF Then CodigoCli = .fields("Codigo")
   End If
  End With
  Cta = SinEspaciosIzq(DCCta)
  If Cta = "" Then Cta = Ninguno
  Beneficiario = DCDet
  If Beneficiario = "" Then Beneficiario = Ninguno
  Select Case CTC
    Case "Ingresos", "Egresos", "Centro de Costos"
         If CheqCta.value = 1 Then
            sSQL = "SELECT TS.Cta,CC.Cuenta,CS.Detalle As Sub_Modulos,TS.Codigo,"
         Else
            sSQL = "SELECT CS.Detalle As Sub_Modulos,TS.Cta,CC.Cuenta,TS.Codigo,"
         End If
         If CheqDSubCta.value = 1 Then sSQL = sSQL & "TS.Detalle_SubCta As Detalle_Auxiliar,"
         Select Case TipoCta
           Case "I": sSQL = sSQL & "SUM(TS.Creditos-TS.Debitos) As Total "
           Case "G", "CC": sSQL = sSQL & "SUM(TS.Debitos-TS.Creditos) As Total "
         End Select
         sSQL = sSQL _
              & "FROM Catalogo_SubCtas As CS, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
              & "AND TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' "
         If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo = '" & Cta & "' "
         If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
         If CheqDet.value = 1 Then sSQL = sSQL & "AND TS.Detalle_SubCta = '" & DCDet & "' "
         sSQL = sSQL _
              & "AND TS.Item = CS.Item " _
              & "AND TS.Item = CC.Item " _
              & "AND TS.Periodo = CS.Periodo " _
              & "AND TS.Periodo = CC.Periodo " _
              & "AND TS.Cta = CC.Codigo " _
              & "AND TS.Codigo = CS.Codigo "
         If CheqCta.value = 1 Then
            If CheqDSubCta.value = 1 Then
                sSQL = sSQL _
                     & "GROUP BY TS.Cta,CC.Cuenta,CS.Detalle,TS.Detalle_SubCta,TS.Codigo " _
                     & "ORDER BY TS.Cta,CC.Cuenta,CS.Detalle,TS.Detalle_SubCta "
            Else
                sSQL = sSQL _
                     & "GROUP BY TS.Cta,CC.Cuenta,CS.Detalle,TS.Codigo " _
                     & "ORDER BY TS.Cta,CC.Cuenta,CS.Detalle "
            End If
         Else
            If CheqDSubCta.value = 1 Then
                sSQL = sSQL _
                     & "GROUP BY CS.Detalle,CC.Cuenta,TS.Cta,TS.Detalle_SubCta,TS.Codigo " _
                     & "ORDER BY CS.Detalle,CC.Cuenta,TS.Cta,TS.Detalle_SubCta "
            Else
                sSQL = sSQL _
                     & "GROUP BY CS.Detalle,CC.Cuenta,TS.Cta,TS.Codigo " _
                     & "ORDER BY CS.Detalle,CC.Cuenta,TS.Cta "
            End If
         End If
         Select_Adodc_Grid DGBanco, AdoBanco, sSQL
         Total = 0
         With AdoBanco.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Total = Total + .fields("Total")
                .MoveNext
              Loop
          End If
         End With
         LabelTotal.Caption = Format(Total, "#,##0.00")
    Case Else
         RatonNormal
         MsgBox "Consuta No permitida"
  End Select
  DGBanco.Visible = True
  Opcion = 10
  RatonNormal
End Sub

Public Sub Consultar()
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  
  FechaInicial = BuscarFecha(MBoxFechaI)
  FechaFinal = BuscarFecha(MBoxFechaF)
  CodigoCli = Ninguno
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCtas.Text & "' ")
       If Not .EOF Then CodigoCli = .fields("Codigo")
   End If
  End With
  Cta = SinEspaciosIzq(DCCta)
  If Cta = "" Then Cta = Ninguno
  Beneficiario = DCDet
  If Beneficiario = "" Then Beneficiario = Ninguno
' Consultamos Facturas SubCtas
 'MsgBox CTC & vbCrLf & TipoCta
  Select Case CTC
    Case "CxC", "CxP"
         sSQL = "SELECT CC.Cuenta, C.Cliente, C.Telefono, TS.Serie, TS.Factura, MIN(TS.TP) As TP_, MIN(TS.Numero) As Numero_, MIN(TS.Fecha_E) As Fecha_Emi,MIN(TS.Fecha_V) As Fecha_Ven,"
         If CheqDSubCta.value = 1 Then sSQL = sSQL & "TS.Detalle_SubCta As Beneficiario,"
         Select Case TipoCta
           Case "C"
                sSQL = sSQL & "SUM(TS.Debitos) As Total,SUM(TS.Creditos) As Abonos,SUM(TS.Debitos-TS.Creditos) As Saldo,"
                SQL1 = "HAVING SUM(TS.Debitos-TS.Creditos) "
           Case "P"
                sSQL = sSQL & "SUM(TS.Creditos) As Total,SUM(TS.Debitos) As Abonos,SUM(TS.Creditos-TS.Debitos) As Saldo,"
                SQL1 = "HAVING SUM(TS.Creditos-TS.Debitos) "
         End Select
         If OpcP.value Then SQL1 = SQL1 & " <> 0 " Else SQL1 = SQL1 & " = 0 "
         sSQL = sSQL & "TS.TC,TS.Codigo,TS.Cta " _
              & "FROM Clientes As C, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
              & "AND TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' "
         If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo = '" & Cta & "' "
         If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
         If CheqDet.value = 1 Then sSQL = sSQL & "AND TS.Detalle_SubCta = '" & DCDet & "' "
         sSQL = sSQL & "AND TS.Codigo = C.Codigo " _
              & "AND TS.Cta = CC.Codigo " _
              & "AND TS.Item = CC.Item " _
              & "AND TS.Periodo = CC.Periodo "
         If CheqDSubCta.value = 1 Then
            sSQL = sSQL & "GROUP BY C.Cliente,TS.Codigo,CC.Cuenta,TS.Serie, TS.Factura,C.Telefono,TS.Detalle_SubCta,TS.TC,TS.Cta "
            If OpcT.value = 0 Then sSQL = sSQL & SQL1
            sSQL = sSQL & "ORDER BY CC.Cuenta,C.Cliente,TS.Detalle_SubCta,TS.Factura "
         Else
            sSQL = sSQL & "GROUP BY C.Cliente,TS.Codigo,CC.Cuenta,TS.Serie,TS.Factura,C.Telefono,TS.TC,TS.Cta "
            If OpcT.value = 0 Then sSQL = sSQL & SQL1
            sSQL = sSQL & "ORDER BY CC.Cuenta,C.Cliente,TS.Factura "
         End If
    Case "Ingresos", "Egresos"
         sSQL = "SELECT CC.Cuenta,C.Detalle As Sub_Modulos,MIN(TS.Fecha) As Fecha_Emi,"
         If CheqDSubCta.value = 1 Then sSQL = sSQL & "TS.Detalle_SubCta As Beneficiario,"
         Select Case TipoCta
           Case "I"
                sSQL = sSQL & "SUM(TS.Creditos) As Total,"
           Case "G"
                sSQL = sSQL & "SUM(TS.Debitos) As Total,"
         End Select
         sSQL = sSQL & "TS.TC,TS.Codigo,TS.Cta " _
              & "FROM Catalogo_SubCtas As C, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
              & "AND TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' "
         If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo = '" & Cta & "' "
         If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
         If CheqDet.value = 1 Then sSQL = sSQL & "AND TS.Detalle_SubCta = '" & DCDet & "' "
         sSQL = sSQL _
              & "AND TS.Codigo = C.Codigo " _
              & "AND TS.Cta = CC.Codigo " _
              & "AND TS.Item = CC.Item " _
              & "AND TS.Periodo = CC.Periodo " _
              & "GROUP BY CC.Cuenta,C.Detalle,TS.Detalle_SubCta,TS.TC,TS.Codigo,TS.Cta " _
              & "ORDER BY CC.Cuenta,C.Detalle,TS.Detalle_SubCta "
  End Select
 'MsgBox sSQL
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  Saldo = 0: Total = 0: Valor = 0
  RatonReloj
  DGBanco.Visible = False
  Anio = Year(FechaSistema)
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Select Case CTC
            Case "CxC", "CxP"
                 Total = Total + .fields("Total")
                 Saldo = Saldo + .fields("Saldo")
            Case Else
                 Total = Total + .fields("Total")
          End Select
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelSaldo.Caption = Format(Saldo, "#,##0.00")
  LabelTotal.Caption = Format(Total, "#,##0.00")
  DGBanco.Visible = True
  SaldoSubCtasVence.Caption = SQLMsg1
  Opcion = 1
  RatonNormal
End Sub

Public Sub Temporizada()
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'CCXP' "
  Ejecutar_SQL_SP sSQL
  
  FechaInicial = BuscarFecha(MBoxFechaI)
  FechaFinal = BuscarFecha(MBoxFechaF)
  CodigoCli = Ninguno
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCtas.Text & "' ")
       If Not .EOF Then CodigoCli = .fields("Codigo")
   End If
  End With
  Cta = SinEspaciosIzq(DCCta)
  If Cta = "" Then Cta = Ninguno
  Beneficiario = DCDet
  If Beneficiario = "" Then Beneficiario = Ninguno
' Consultamos Facturas SubCtas
 'MsgBox CTC & vbCrLf & TipoCta
  Select Case CTC
    Case "CxC", "CxP"
         sSQL = "SELECT CC.Cuenta,C.Cliente,C.Telefono,TS.Factura,MIN(TS.Fecha) As Fecha_Emi,MIN(TS.Fecha_V) As Fecha_Ven,"
         If CheqDSubCta.value = 1 Then sSQL = sSQL & "TS.Detalle_SubCta As Beneficiario,"
         Select Case TipoCta
           Case "C"
                sSQL = sSQL & "SUM(TS.Debitos) As Total,SUM(TS.Creditos) As Abonos,SUM(TS.Debitos-TS.Creditos) As Saldo,"
                SQL1 = "HAVING SUM(TS.Debitos-TS.Creditos) "
           Case "P"
                sSQL = sSQL & "SUM(TS.Creditos) As Total,SUM(TS.Debitos) As Abonos,SUM(TS.Creditos-TS.Debitos) As Saldo,"
                SQL1 = "HAVING SUM(TS.Creditos-TS.Debitos) "
         End Select
         If OpcP.value Then SQL1 = SQL1 & " <> 0 " Else SQL1 = SQL1 & " = 0 "
         sSQL = sSQL & "TS.TC,TS.Codigo,TS.Cta " _
              & "FROM Clientes As C, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
              & "AND TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' "
         If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo = '" & Cta & "' "
         If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
         If CheqDet.value = 1 Then sSQL = sSQL & "AND TS.Detalle_SubCta = '" & DCDet & "' "
         sSQL = sSQL & "AND TS.Codigo = C.Codigo " _
              & "AND TS.Cta = CC.Codigo " _
              & "AND TS.Item = CC.Item " _
              & "AND TS.Periodo = CC.Periodo "
         If CheqDSubCta.value = 1 Then
            sSQL = sSQL & "GROUP BY C.Cliente,TS.Codigo,CC.Cuenta,TS.Factura,C.Telefono,TS.Detalle_SubCta,TS.TC,TS.Cta " _
                 & SQL1 _
                 & "ORDER BY CC.Cuenta,C.Cliente,TS.Detalle_SubCta,TS.Factura "
         Else
            sSQL = sSQL & "GROUP BY C.Cliente,TS.Codigo,CC.Cuenta,TS.Factura,C.Telefono,TS.TC,TS.Cta " _
                 & SQL1 _
                 & "ORDER BY CC.Cuenta,C.Cliente,TS.Factura "
         End If
    Case "Ingresos", "Egresos"
         sSQL = "SELECT CC.Cuenta,C.Detalle As Sub_Modulos,MIN(TS.Fecha) As Fecha_Emi,"
         If CheqDSubCta.value = 1 Then sSQL = sSQL & "TS.Detalle_SubCta As Beneficiario,"
         Select Case TipoCta
           Case "I"
                sSQL = sSQL & "SUM(TS.Creditos) As Total,"
           Case "G"
                sSQL = sSQL & "SUM(TS.Debitos) As Total,"
         End Select
         sSQL = sSQL & "TS.TC,TS.Codigo,TS.Cta " _
              & "FROM Catalogo_SubCtas As C, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
              & "AND TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' "
         If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo = '" & Cta & "' "
         If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
         If CheqDet.value = 1 Then sSQL = sSQL & "AND TS.Detalle_SubCta = '" & DCDet & "' "
         sSQL = sSQL _
              & "AND TS.Codigo = C.Codigo " _
              & "AND TS.Cta = CC.Codigo " _
              & "AND TS.Item = CC.Item " _
              & "AND TS.Periodo = CC.Periodo " _
              & "GROUP BY CC.Cuenta,C.Detalle,TS.Detalle_SubCta,TS.TC,TS.Codigo,TS.Cta " _
              & "ORDER BY CC.Cuenta,C.Detalle,TS.Detalle_SubCta "
  End Select
 'MsgBox sSQL
  DGBanco.Visible = False
  Select_Adodc AdoBanco, sSQL
  Saldo = 0: Total = 0: Valor = 0
  RatonReloj
  Anio = Year(FechaSistema)
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SetAdoAddNew "Saldo_Diarios"
          Select Case CTC
            Case "CxC", "CxP"
                 Saldo = Saldo + .fields("Saldo")
                 FechaN = CFechaLong(.fields("Fecha_Emi"))
                 FechaFinN = CFechaLong(.fields("Fecha_Ven"))
                 SetAdoFields "Fecha_Venc", .fields("Fecha_Ven")
                 SetAdoFields "Numero", .fields("Factura")
                 SetAdoFields "Comprobante", .fields("Cliente")
                 Valor = .fields("Saldo")
            Case Else
                 FechaN = CFechaLong("01/01/" & Year(FechaSistema))
                 FechaFinN = CFechaLong(.fields("Fecha_Emi"))
                 SetAdoFields "Fecha_Venc", .fields("Fecha_Emi")
                 SetAdoFields "Comprobante", .fields("Sub_Modulos")
                 Valor = .fields("Total")
                 Saldo = Saldo + .fields("Total")
          End Select
          NumDias = FechaFinN - FechaN
          Select Case NumDias
            Case 1 To 7: SetAdoFields "Ven_1_a_7", Valor
            Case 8 To 30: SetAdoFields "Ven_8_a_30", Valor
            Case 31 To 60: SetAdoFields "Ven_31_a_60", Valor
            Case 61 To 90: SetAdoFields "Ven_61_a_90", Valor
            Case 91 To 180: SetAdoFields "Ven_91_a_180", Valor
            Case 181 To 360: SetAdoFields "Ven_181_a_360", Valor
            Case Is > 360: SetAdoFields "Ven_mas_de_360", Valor
          End Select
          SetAdoFields "T", Normal
          SetAdoFields "Fecha", .fields("Fecha_Emi")
          SetAdoFields "Dato_Aux1", .fields("Cuenta")
          SetAdoFields "Total", Valor
          SetAdoFields "Saldo_Actual", Valor
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "TP", "CCXP"
          SetAdoUpdate
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  sSQL = "SELECT Dato_Aux1 As Cuenta,Comprobante As Cliente,Fecha_Venc,Numero As Factura," _
       & "Ven_1_a_7,Ven_8_a_30,Ven_31_a_60,Ven_61_a_90,Ven_91_a_180,Ven_181_a_360,Ven_mas_de_360 " _
       & "FROM Saldo_Diarios " _
       & "WHERE Fecha_Venc <= #" & FechaFinal & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'CCXP' " _
       & "ORDER BY TC,Dato_Aux1,Comprobante,Cta,Numero "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  
  LabelSaldo.Caption = Format(Saldo, "#,##0.00")
  LabelTotal.Caption = Format(Total, "#,##0.00")
  DGBanco.Visible = True
  SaldoSubCtasVence.Caption = SQLMsg1
  Opcion = 2
  RatonNormal
End Sub

Private Sub Command1_Click()
  Unload SaldoSubCtasVence
End Sub

Private Sub CTC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CTC_LostFocus()
  Listar_Beneficiarios
End Sub

Private Sub DGBanco_DblClick()
    With AdoBanco.Recordset
     If .RecordCount > 0 Then
         If ClaveContador Then
            RatonReloj
            SerieFactura = DGBanco.Columns(3)
            FacturaNo = DGBanco.Columns(4)
            TipoDoc = DGBanco.Columns(5)
            Asiento = DGBanco.Columns(6)
            Codigo1 = DGBanco.Columns(13)
            Cta = DGBanco.Columns(14)
            Mifecha = InputBox("INGRSE FECHA DE EMISION: ", "CAMBIO DE FECHA DE EMISION", FechaSistema)
            If IsDate(Mifecha) Then
               sSQL = "UPDATE Trans_SubCtas " _
                    & "SET Fecha_E = '" & BuscarFecha(Mifecha) & "' " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND TP = '" & TipoDoc & "' " _
                    & "AND Numero = " & Asiento & " " _
                    & "AND Codigo = '" & Codigo1 & "' " _
                    & "AND Cta = '" & Cta & "' " _
                    & "AND Factura = " & FacturaNo & " "
               Ejecutar_SQL_SP sSQL
               RatonNormal
               MsgBox "Proceso realizado con exito, vuelva a consultar"
            Else
               RatonNormal
               MsgBox "Fecha ingresada incorrecta"
            End If
         End If
     Else
         RatonNormal
         MsgBox "No Existe Datos que Modificar"
     End If
    End With
End Sub

Private Sub Form_Activate()
  Toolbar1.buttons("Consultar").Enabled = False
  Toolbar1.buttons("Costos").Enabled = False
  Toolbar1.buttons("Temporizada").Enabled = False
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'CCXP' "
  Ejecutar_SQL_SP sSQL
  
  CTC.Clear
  CTC.AddItem "CxC"
  CTC.AddItem "CxP"
  CTC.AddItem "Ingresos"
  CTC.AddItem "Egresos"
  CTC.AddItem "Centro de Costos"
  CTC.Text = "CxC"
  
  Listar_Beneficiarios
  
  DGBanco.width = MDI_X_Max - 100
  DGBanco.Height = MDI_Y_Max - 2000
  
  PictFactura.width = MDI_X_Max - PictFactura.Left - 50
  PictFactura.Top = DGBanco.Top + DGBanco.Height + 30
  AdoBanco.Top = DGBanco.Top + DGBanco.Height + 30
  Label3.Top = DGBanco.Top + DGBanco.Height + 30
  LabelTotal.Top = DGBanco.Top + DGBanco.Height + 30
  Label19.Top = DGBanco.Top + DGBanco.Height + 30
  LabelSaldo.Top = DGBanco.Top + DGBanco.Height + 30
  If Bloquear_Control Then
     Toolbar1.buttons("Consultar").Enabled = False
     Toolbar1.buttons("Costos").Enabled = False
     Toolbar1.buttons("Imprimir").Enabled = False
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoDet
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoDetCheq
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

Public Sub Listar_Beneficiarios()
  RatonReloj
  Toolbar1.buttons("Consultar").Enabled = False
  Toolbar1.buttons("Costos").Enabled = False
  Toolbar1.buttons("Temporizada").Enabled = False
  Select Case CTC
    Case "CxC"
         Toolbar1.buttons("Consultar").Enabled = True
         Toolbar1.buttons("Temporizada").Enabled = True
         TipoCta = "C"
         SQLMsg1 = "SALDO DE CUENTAS POR COBRAR"
         CheqIndiv.Caption = "Beneficiario:"
    Case "CxP"
         Toolbar1.buttons("Consultar").Enabled = True
         Toolbar1.buttons("Temporizada").Enabled = True
         TipoCta = "P"
         SQLMsg1 = "SALDO DE CUENTAS POR PAGAR"
         CheqIndiv.Caption = "Beneficiario:"
    Case "Ingresos"
         Toolbar1.buttons("Costos").Enabled = True
         TipoCta = "I"
         SQLMsg1 = "SALDO DE INGRESOS"
         CheqIndiv.Caption = "Submodulo:"
    Case "Egresos"
         Toolbar1.buttons("Costos").Enabled = True
         TipoCta = "G"
         SQLMsg1 = "SALDO DE EGRESOS"
         CheqIndiv.Caption = "Submodulo:"
    Case "Centro de Costos"
         Toolbar1.buttons("Costos").Enabled = True
         TipoCta = "CC"
         SQLMsg1 = "SALDO DE COSTOS"
         CheqIndiv.Caption = "Submodulo:"
    Case Else
         TipoCta = Ninguno
  End Select
    
  DGBanco.Caption = SQLMsg1
  
  Select Case CTC
    Case "CxC", "CxP"
         sSQL = "SELECT C.Cliente,TS.Codigo " _
              & "FROM Trans_SubCtas As TS,Clientes As C " _
              & "WHERE TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' " _
              & "AND TS.Codigo = C.Codigo " _
              & "GROUP BY C.Cliente,TS.Codigo " _
              & "ORDER BY C.Cliente,TS.Codigo "
    Case "Ingresos", "Egresos", "Centro de Costos"
         sSQL = "SELECT C.Detalle As Cliente,TS.Codigo " _
              & "FROM Trans_SubCtas As TS,Catalogo_SubCtas As C " _
              & "WHERE TS.Item = '" & NumEmpresa & "' " _
              & "AND TS.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.TC = '" & TipoCta & "' " _
              & "AND TS.Codigo = C.Codigo " _
              & "AND TS.TC = C.TC " _
              & "AND TS.Item = C.Item " _
              & "AND TS.Periodo = C.Periodo " _
              & "GROUP BY C.Detalle,TS.Codigo " _
              & "ORDER BY C.Detalle,TS.Codigo "
  End Select
  SelectDB_Combo DCCtas, AdoSubCta, sSQL, "Cliente"
  
  sSQL = "SELECT (TS.Cta & '  ' & CC.Cuenta) As Nombre_Cta " _
       & "FROM Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
       & "WHERE CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.TC = '" & TipoCta & "' " _
       & "AND CC.Codigo = TS.Cta " _
       & "AND CC.Item = TS.Item " _
       & "AND CC.Periodo = TS.Periodo " _
       & "AND CC.TC = TS.TC " _
       & "GROUP BY TS.Cta,CC.Cuenta " _
       & "ORDER BY TS.Cta "
  SelectDB_Combo DCCta, AdoCtas, sSQL, "Nombre_Cta"
  
  sSQL = "SELECT Detalle_SubCta " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoCta & "' " _
       & "GROUP BY Detalle_SubCta " _
       & "ORDER BY Detalle_SubCta "
  SelectDB_Combo DCDet, AdoDet, sSQL, "Detalle_SubCta"
  RatonNormal
End Sub

Public Sub Reporte_CxCxP_x_Meses()
Dim TotalSubModulo As Currency

   If CheqCta.value <> 0 Then
      Opcion = 9
      TotalSubModulo = 0
      Cta = SinEspaciosIzq(DCCta)
      Reporte_CxCxP_x_Meses_SP Cta, MBoxFechaF
      
      sSQL = "SELECT Cta, Beneficiario, Anio, Mes, Valor_x_Mes, Categoria " _
           & "FROM Reporte_CxCxP_x_Meses " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' " _
           & "AND Cta = '" & Cta & "' " _
           & "ORDER BY Beneficiario, Anio, Mes_No "
      Select_Adodc_Grid DGBanco, AdoBanco, sSQL
      With AdoBanco.Recordset
       If .RecordCount > 0 Then
           Do While Not .EOF
              If InStr(.fields("Anio"), "TOTAL") Then TotalSubModulo = TotalSubModulo + .fields("Valor_x_Mes")
             .MoveNext
           Loop
          .MoveFirst
       End If
      End With
      LabelSaldo.Caption = Format(TotalSubModulo, "#,##0.00")
   Else
      MsgBox "Seleciones una cuenta de submodulo"
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Consultar": Consultar
      Case "Costos": Costos
      Case "Temporizada": Temporizada
      Case "Imprimir": ImprimirCxC_P_G_I_CC
      Case "CxCxP_Mes"
           DGBanco.Visible = False
           Reporte_CxCxP_x_Meses
           DGBanco.Visible = True
      Case "Excel"
           DGBanco.Visible = False
           GenerarDataTexto SaldoSubCtasVence, AdoBanco
           DGBanco.Visible = True
      Case "Salir":  Unload SaldoSubCtasVence
    End Select
End Sub
