VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Existen 
   Caption         =   "Existencias de Inventario"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12405
   Icon            =   "Existen.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   12405
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
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
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Consultar"
            Object.ToolTipText     =   "Consulta el Kardex de un Producto"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kardex_Total"
            Object.ToolTipText     =   "Presenta el Kardes de todos los productos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kardex"
            Object.ToolTipText     =   "Presenta el Resumen de Codigos de Barra"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kardex_Total_QR"
            Object.ToolTipText     =   "Kardex por Codigo QR"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir_Kardex"
            Object.ToolTipText     =   "Imprimir el Kardex de un Producto"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   4515
         TabIndex        =   1
         Top             =   0
         Width           =   9465
         Begin VB.OptionButton OpcBodega 
            Caption         =   "Bodega"
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
            TabIndex        =   3
            Top             =   210
            Width           =   1065
         End
         Begin VB.OptionButton OpcCodigoBarra 
            Caption         =   "Codigo Barra"
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
            TabIndex        =   4
            Top             =   210
            Width           =   1485
         End
         Begin VB.OptionButton OpcTodos 
            Caption         =   "Ninguno"
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
            Value           =   -1  'True
            Width           =   1170
         End
         Begin MSDataListLib.DataCombo DCBodega 
            Bindings        =   "Existen.frx":0442
            DataSource      =   "AdoBodega"
            Height          =   315
            Left            =   4305
            TabIndex        =   5
            Top             =   210
            Visible         =   0   'False
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   "DC"
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
      End
   End
   Begin VB.Frame FrmProductos 
      BackColor       =   &H00404080&
      Caption         =   "| CAMBIO DE PRODUCTOS |"
      ForeColor       =   &H0000FFFF&
      Height          =   5370
      Left            =   13125
      TabIndex        =   27
      Top             =   2940
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton Command3 
         BackColor       =   &H000080FF&
         Caption         =   "&Salir"
         Height          =   855
         Left            =   6510
         Picture         =   "Existen.frx":045A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4305
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "&Aceptar"
         Height          =   855
         Left            =   4620
         Picture         =   "Existen.frx":0D24
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4305
         Width           =   1800
      End
      Begin MSDataListLib.DataCombo DCArt 
         Bindings        =   "Existen.frx":1166
         DataSource      =   "AdoArt"
         Height          =   3495
         Left            =   210
         TabIndex        =   30
         Top             =   630
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   6165
         _Version        =   393216
         Style           =   1
         BackColor       =   12640511
         Text            =   ""
      End
      Begin VB.Label LblProducto 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Producto anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   31
         Top             =   315
         Width           =   8100
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   12915
      ScaleHeight     =   855
      ScaleWidth      =   1695
      TabIndex        =   26
      Top             =   1260
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   12600
      Picture         =   "Existen.frx":117B
      TabIndex        =   25
      Top             =   735
      Width           =   330
   End
   Begin MSDataGridLib.DataGrid DGKardex 
      Bindings        =   "Existen.frx":1A45
      Height          =   4530
      Left            =   105
      TabIndex        =   12
      ToolTipText     =   "<Ctrl+F12> Cambiar Codigo, <Ctrl+B> Cambia Codigo Barra, <Ctrl+S> Cambia Serie, <Ctrl+0> Elimina Comprobantes Cero"
      Top             =   2520
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
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
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   11130
      TabIndex        =   11
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
      Left            =   8820
      TabIndex        =   9
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   105
      Top             =   7350
      Width           =   12195
      _ExtentX        =   21511
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
      Caption         =   "Kardex"
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
      Left            =   210
      Top             =   3255
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   210
      Top             =   2940
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Art"
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
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "Existen.frx":1A5D
      DataSource      =   "AdoInv"
      Height          =   1350
      Left            =   105
      TabIndex        =   7
      Top             =   1050
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   2381
      _Version        =   393216
      Style           =   1
      BackColor       =   16777215
      ForeColor       =   4210752
      Text            =   "Productos"
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
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "Existen.frx":1A72
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      Text            =   "DC"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   210
      Top             =   3885
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Inv"
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   210
      Top             =   3570
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   210
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
      Caption         =   "Bodega"
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
      Left            =   14700
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":1A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":1DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":20BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":23D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":26F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":2A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":2D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Existen.frx":3976
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelMaximo 
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
      Left            =   11445
      TabIndex        =   13
      Top             =   1995
      Width           =   1380
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Maximo:"
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
      Left            =   10605
      TabIndex        =   19
      Top             =   1995
      Width           =   855
   End
   Begin VB.Label LabelBodega 
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
      Left            =   8820
      TabIndex        =   16
      Top             =   1995
      Width           =   1695
   End
   Begin VB.Label LabelExitencia 
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
      Left            =   11445
      TabIndex        =   14
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Existe.:"
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
      Left            =   10605
      TabIndex        =   20
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label LabelUnidad 
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
      Left            =   8820
      TabIndex        =   17
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label LabelMinimo 
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
      Left            =   11445
      TabIndex        =   15
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Minimo:"
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
      Left            =   10605
      TabIndex        =   21
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   8820
      TabIndex        =   18
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Codigo:"
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
      Left            =   7980
      TabIndex        =   24
      Top             =   1155
      Width           =   855
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
      Left            =   10290
      TabIndex        =   10
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label8 
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
      Left            =   7980
      TabIndex        =   8
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Unidad:"
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
      Left            =   7980
      TabIndex        =   23
      Top             =   1575
      Width           =   855
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bodega:"
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
      Left            =   7980
      TabIndex        =   22
      Top             =   1995
      Width           =   855
   End
End
Attribute VB_Name = "Existen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PorBodega       As Boolean
Dim PorCodigoBarra  As Boolean
Dim PorTodos        As Boolean

Public Sub ListarProductos(OpcVista As String)
 Select Case OpcVista
  Case "I"
       sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
            & "FROM Catalogo_Productos " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND TC = 'I' " _
            & "ORDER BY Codigo_Inv "
       SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
  Case "P"
      '& "AND X = 'M' "
       Codigo = SinEspaciosIzq(DCTInv.Text)
       sSQL = "SELECT Producto,Codigo_Inv,Unidad,Minimo,Maximo " _
            & "FROM Catalogo_Productos " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND MidStrg(Codigo_Inv,1," & CStr(Len(Codigo)) & ") = '" & Codigo & "' " _
            & "AND TC = 'P' " _
            & "AND Cta_Inventario <> '.' " _
            & "AND Cta_Inventario <> '0' " _
            & "ORDER BY Producto,Codigo_Inv "
       SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
 End Select
End Sub

Public Sub Listar_Bodega_CodigoBarra()
  If PorBodega Then
     sSQL = "SELECT CodBod & ' - ' & Bodega As TipoDeConsulta " _
          & "FROM Catalogo_Bodegas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY CodBod "
  Else
     sSQL = "SELECT Codigo_Barra As TipoDeConsulta " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Entrada > 0 " _
          & "GROUP BY Codigo_Barra " _
          & "ORDER BY Codigo_Barra "
  End If
  SelectDB_Combo DCBodega, AdoBodega, sSQL, "TipoDeConsulta"
End Sub

Public Sub Consultar_Kardex()
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Codigo = LabelCodigo.Caption
  Codigo1 = SinEspaciosDer(DCBodega.Text)
  Debe = 0
  Haber = 0
  If Codigo = "" Then Codigo = "."
  sSQL = "SELECT K.Codigo_Inv, K.Codigo_Barra, SUM(Entrada) As Entradas, SUM(Salida) As Salidas, SUM(Entrada-Salida) As Stock_Kardex " _
       & "FROM Trans_Kardex As K, Comprobantes As C " _
       & "WHERE K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND K.Codigo_Inv = '" & Codigo & "' " _
       & "AND K.T = '" & Normal & "' " _
       & "AND K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  If PorBodega Then sSQL = sSQL & "AND K.CodBodega = '" & Codigo1 & "' "
  sSQL = sSQL & "AND K.TP = C.TP " _
       & "AND K.Fecha = C.Fecha " _
       & "AND K.Numero = C.Numero " _
       & "AND K.Item = C.Item " _
       & "AND K.Periodo = C.Periodo " _
       & "GROUP BY K.Codigo_Inv, K.Codigo_Barra " _
       & "HAVING SUM(Entrada-Salida) >= 1 " _
       & "ORDER BY K.Codigo_Inv, K.Codigo_Barra "
  SQLDec = ""
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec, , , "Kardex_Producto"
  RatonReloj
  With AdoKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .fields("Stock_Kardex")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelExitencia.Caption = Format(Debe - Haber, "#,##0.00")
  Existen.Caption = "EXISTENCIA DE INVENTARIO"
  DGKardex.Visible = True
  RatonNormal
End Sub

Public Sub Consultar_Tipo_De_Kardex(EsKardexIndividual As Boolean)
Dim GrupoInv As String
  
  DGKardex.Visible = False
  Debe = 0
  Haber = 0
  GrupoInv = SinEspaciosIzq(DCTInv)
  Codigo = LabelCodigo.Caption
  Codigo1 = SinEspaciosIzq(DCBodega.Text)
  If Codigo = "" Then Codigo = "."
  If GrupoInv = "" Then GrupoInv = "*"
  
  sSQL = "SELECT TK.Codigo_Inv, CP.Producto, CP.Unidad, TK.CodBodega As Bodega, TK.Codigo_Barra, TK.Fecha, TK.TP, TK.Numero, TK.Entrada, TK.Salida, "
  If PorCodigoBarra Then
     sSQL = sSQL & "TK.Stock_Barra As Stock, "
  ElseIf PorBodega Then
     sSQL = sSQL & "TK.Stock_Bod As Stock, "
  Else
     sSQL = sSQL & "TK.Existencia As Stock, "
  End If
  sSQL = sSQL _
       & "TK.Costo, TK.Total As Saldo, TK.Valor_Unitario, TK.Valor_Total, TK.Fecha_Exp, TK.TC, TK.Serie, TK.Factura, TK.Cta_Inv, TK.Contra_Cta, TK.Serie_No, TK.Lote_No, " _
       & "TK.Codigo_Tra AS CI_RUC_CC, CM.Marca As 'Marca_Tipo_Proceso', TK.Detalle, TK.Centro_Costo As Beneficiario_Centro_Costo, TK.Orden_No, TK.ID " _
       & "FROM Trans_Kardex As TK, Catalogo_Productos As CP, Catalogo_Marcas As CM " _
       & "WHERE TK.Item = '" & NumEmpresa & "' " _
       & "AND TK.Periodo = '" & Periodo_Contable & "' " _
       & "AND TK.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TK.T = '" & Normal & "' "
  If EsKardexIndividual Then
     sSQL = sSQL & "AND TK.Codigo_Inv = '" & Codigo & "' "
  Else
     If GrupoInv <> "*" Then sSQL = sSQL & "AND TK.Codigo_Inv LIKE '" & GrupoInv & "%' "
  End If
  If PorBodega Then sSQL = sSQL & "AND TK.CodBodega LIKE '" & Codigo1 & "%' "
  If PorCodigoBarra Then sSQL = sSQL & "AND TK.Codigo_Barra = '" & Codigo1 & "' "
  sSQL = sSQL _
       & "AND TK.Item = CP.Item " _
       & "AND TK.Item = CM.Item " _
       & "AND TK.Periodo = CP.Periodo " _
       & "AND TK.Periodo = CM.Periodo " _
       & "AND TK.Codigo_Inv = CP.Codigo_Inv " _
       & "AND TK.CodMarca = CM.CodMar " _
       & "ORDER BY TK.Codigo_Inv, TK.Fecha,TK.Entrada DESC,TK.Salida,TK.TP,TK.Numero,TK.ID "
  SQLDec = "TK.Costo " & CStr(Dec_Costo) & "| TK.Valor_Unitario " & CStr(Dec_Costo) & "|,TK.Valor_Total 2|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  DGKardex.Visible = False
  RatonReloj
  If EsKardexIndividual Then
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Debe = Debe + .fields("Entrada")
             Haber = Haber + .fields("Salida")
            .MoveNext
          Loop
         .MoveFirst
      End If
     End With
     Existen.Caption = "EXISTENCIA DE INVENTARIO"
  Else
     Existen.Caption = "EXISTENCIA DE TODOS LOS INVENTARIOS"
  End If
  LabelExitencia.Caption = Format(Debe - Haber, "#,##0.00")
  DGKardex.Visible = True
  RatonNormal
End Sub

Public Sub Kardex_Total_QR()
Dim GrupoInv As String
  
  DGKardex.Visible = False
  Debe = 0
  Haber = 0
  GrupoInv = SinEspaciosIzq(DCTInv)
  Codigo = LabelCodigo.Caption
  Codigo1 = SinEspaciosIzq(DCBodega.Text)
  If Codigo = "" Then Codigo = "."
  If GrupoInv = "" Then GrupoInv = "*"
  
  sSQL = "SELECT TK.Codigo_Inv, CP.Producto, CP.Unidad, TK.CodBodega As Bodega, TK.Codigo_Barra, TK.Fecha, TK.TP, TK.Numero, TK.Entrada, TK.Salida, "
  If PorCodigoBarra Then
     sSQL = sSQL & "TK.Stock_Barra As Stock, "
  ElseIf PorBodega Then
     sSQL = sSQL & "TK.Stock_Bod As Stock, "
  Else
     sSQL = sSQL & "TK.Existencia As Stock, "
  End If
  sSQL = sSQL _
       & "TK.Costo, TK.Total As Saldo, TK.Valor_Unitario, TK.Valor_Total, TK.Fecha_Exp, TK.TC, TK.Serie, TK.Factura, TK.Cta_Inv, TK.Contra_Cta, TC.Porc_C As Temperatura, TK.Serie_No, TK.Lote_No, " _
       & "TK.Codigo_Tra AS CI_RUC_CC, CM.Marca As 'Marca_Tipo_Proceso', TK.Detalle, TK.Centro_Costo As Beneficiario_Centro_Costo, TK.Orden_No, TK.ID " _
       & "FROM Trans_Kardex As TK, Catalogo_Productos As CP, Catalogo_Marcas As CM, Trans_Correos As TC " _
       & "WHERE TK.Item = '" & NumEmpresa & "' " _
       & "AND TK.Periodo = '" & Periodo_Contable & "' " _
       & "AND TK.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TK.T = '" & Normal & "' "
'  If EsKardexIndividual Then
'     sSQL = sSQL & "AND TK.Codigo_Inv = '" & Codigo & "' "
'  Else
'     If GrupoInv <> "*" Then sSQL = sSQL & "AND TK.Codigo_Inv LIKE '" & GrupoInv & "%' "
'  End If
  If PorBodega Then sSQL = sSQL & "AND TK.CodBodega LIKE '" & Codigo1 & "%' "
  If PorCodigoBarra Then sSQL = sSQL & "AND TK.Codigo_Barra = '" & Codigo1 & "' "
  sSQL = sSQL _
       & "AND TK.Item = CP.Item " _
       & "AND TK.Item = CM.Item " _
       & "AND TK.Item = TC.Item " _
       & "AND TK.Periodo = CP.Periodo " _
       & "AND TK.Periodo = CM.Periodo " _
       & "AND TK.Periodo = TC.Periodo " _
       & "AND TK.Codigo_Inv = CP.Codigo_Inv " _
       & "AND TC.Envio_No = SUBSTRING(TK.Codigo_Barra,1,LEN(TC.Envio_No)) " _
       & "AND TK.CodMarca = CM.CodMar " _
       & "ORDER BY TK.Codigo_Inv, TK.Fecha,TK.Entrada DESC,TK.Salida,TK.TP,TK.Numero,TK.ID "
  SQLDec = "TK.Costo " & CStr(Dec_Costo) & "| TK.Valor_Unitario " & CStr(Dec_Costo) & "|,TK.Valor_Total 2|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  DGKardex.Visible = False
  RatonReloj
'  If EsKardexIndividual Then
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Debe = Debe + .fields("Entrada")
             Haber = Haber + .fields("Salida")
            .MoveNext
          Loop
         .MoveFirst
      End If
     End With
     Existen.Caption = "EXISTENCIA DE INVENTARIO"
'  Else
'     Existen.Caption = "EXISTENCIA DE TODOS LOS INVENTARIOS"
'  End If
  LabelExitencia.Caption = Format(Debe - Haber, "#,##0.00")
  DGKardex.Visible = True
  RatonNormal
End Sub

Private Sub Command1_Click()
    Titulo = "PREGUNTA DE ACTUALIZACION"
    Mensajes = "Esta seguro de cambiar" & vbCrLf & vbCrLf _
             & DGKardex.Caption & vbCrLf & vbCrLf _
             & "por el Producto:" & vbCrLf & vbCrLf _
             & DCArt.Text & vbCrLf
    If BoxMensaje = vbYes Then
       With AdoArt.Recordset
        If .RecordCount > 1 Then
           .MoveFirst
           .Find ("Producto = '" & DCArt.Text & "' ")
            If Not .EOF Then
               sSQL = "UPDATE Trans_Kardex " _
                    & "SET Codigo_Inv = '" & .fields("Codigo_Inv") & "' " _
                    & "WHERE ID = " & ID_Reg & " " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' "
               Ejecutar_SQL_SP sSQL
               If Len(FA.TC) = 2 And Len(FA.Serie) = 6 And FA.Factura > 0 Then
                  sSQL = "UPDATE Detalle_Factura " _
                       & "SET Codigo = '" & .fields("Codigo_Inv") & "' " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND Periodo = '" & Periodo_Contable & "' " _
                       & "AND TC = '" & FA.TC & "' " _
                       & "AND Serie = '" & FA.Serie & "' " _
                       & "AND Factura = " & FA.Factura & " " _
                       & "AND Codigo = '" & CodigoInv & "' "
                  Ejecutar_SQL_SP sSQL
                 'MsgBox sSQL & vbCrLf & Len(FA.TC) & vbCrLf & Len(FA.Serie) & vbCrLf & FA.Factura
               End If
               Listar_Articulos
               RatonNormal
               MsgBox "Proceso Terminado, vuelva a listar el Kardex"
               FrmProductos.Visible = False
            End If
        End If
       End With
    End If
End Sub

Private Sub Command2_Click()
  Unload Existen
End Sub

Private Sub Command3_Click()
  FrmProductos.Visible = False
  Listar_Articulos
  RatonNormal
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
  ListarProductos "P"
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCInv_LostFocus()
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto = '" & DCInv & "' ")
       If Not .EOF Then
          Codigo = .fields("Codigo_Inv")
          DGKardex.Caption = UCase(.fields("Producto"))
          LabelCodigo.Caption = .fields("Codigo_Inv")
          LabelUnidad.Caption = .fields("Unidad")
          'LabelBodega.Caption = .Fields("Bodega")
          LabelMinimo.Caption = Format(.fields("Minimo"), "#,##0.00")
          LabelMaximo.Caption = Format(.fields("Maximo"), "#,##0.00")
          LabelExitencia.Caption = "0.00"
       Else
         .MoveFirst
         .Find ("Codigo_Inv = '" & DCInv & "' ")
          If Not .EOF Then
             Codigo = .fields("Codigo_Inv")
             DCInv.Text = UCase(.fields("Producto"))
             DGKardex.Caption = DCInv.Text
             LabelCodigo.Caption = .fields("Codigo_Inv")
             LabelUnidad.Caption = .fields("Unidad")
             'LabelBodega.Caption = .Fields("Bodega")
             LabelMinimo.Caption = Format(.fields("Minimo"), "#,##0.00")
             LabelMaximo.Caption = Format(.fields("Maximo"), "#,##0.00")
             LabelExitencia.Caption = "0.00"
          Else
             MsgBox "No existen este producto"
          End If
       End If
   End If
  End With
End Sub

Private Sub DGKardex_KeyDown(KeyCode As Integer, Shift As Integer)
 Keys_Especiales Shift
 If AdoKardex.Recordset.RecordCount > 0 Then
    CodigoInv = DGKardex.Columns(0)
    FA.TC = DGKardex.Columns(15)
    FA.Serie = DGKardex.Columns(16)
    FA.Factura = DGKardex.Columns(18)
    CodigoN = DGKardex.Columns(20)
    Codigos = DGKardex.Columns(21)
    Existencia = Val(DGKardex.Columns(9).Text)
    ID_Reg = Val(DGKardex.Columns(CantCampos - 1).Text)
    FechaIni = BuscarFecha(MBoxFechaI.Text)
    FechaFin = BuscarFecha(MBoxFechaF.Text)

    If CtrlDown And KeyCode = vbKeyF11 Then
       If Leer_Codigo_Inv(CodigoInv, FechaSistema) Then
          DatInv.Codigo_Barra = Codigos
          Imprimir_Codigos_Barras_Kardex Picture1, CInt(Existencia)
       End If
    End If
    
    If CtrlDown And KeyCode = vbKeyF12 Then
       LblProducto = DGKardex.Caption
       Listar_Articulos True
       RatonNormal
       FrmProductos.Visible = True
       DCArt.SetFocus
    End If
    
    If CtrlDown And KeyCode = vbKeyB Then
       CodigoB = UCase(TrimStrg(MidStrg(InputBox("INGRESE EL CODIGO DE BARRAS DE ESTE PRODUCTO", "INGRESO DE CODIGO DE BARRAS", Codigos), 1, 25)))
       If Len(CodigoB) > 1 Then
          sSQL = "UPDATE Trans_Kardex " _
               & "SET Codigo_Barra = '" & CodigoB & "' " _
               & "WHERE ID = " & ID_Reg & " " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Ejecutar_SQL_SP sSQL
          If Len(FA.TC) = 2 And Len(FA.Serie) = 6 And FA.Factura > 0 Then
             sSQL = "UPDATE Detalle_Factura " _
                  & "SET Codigo_Barra = '" & CodigoB & "' " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TC = '" & FA.TC & "' " _
                  & "AND Serie = '" & FA.Serie & "' " _
                  & "AND Factura = " & FA.Factura & " " _
                  & "AND Codigo = '" & CodigoInv & "' "
             Ejecutar_SQL_SP sSQL
          End If
          RatonNormal
          MsgBox "Proceso Terminado, vuelva a listar el documento"
       End If
    End If
    
    If CtrlDown And KeyCode = vbKeyS Then
       CodigoP = UCase(TrimStrg(InputBox("INGRESE LA SERIE DE ESTE PRODUCTO", "INGRESO DE SERIE", CodigoN)))
       If Len(CodigoP) > 1 Then
          sSQL = "UPDATE Trans_Kardex " _
               & "SET Serie_No = '" & CodigoP & "' " _
               & "WHERE ID = " & ID_Reg & " " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Ejecutar_SQL_SP sSQL
          If Len(FA.TC) = 2 And Len(FA.Serie) = 6 And FA.Factura > 0 Then
             sSQL = "UPDATE Detalle_Factura " _
                  & "SET Serie_No = '" & CodigoP & "' " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TC = '" & FA.TC & "' " _
                  & "AND Serie = '" & FA.Serie & "' " _
                  & "AND Factura = " & FA.Factura & " " _
                  & "AND Codigo = '" & CodigoInv & "' "
             Ejecutar_SQL_SP sSQL
          End If
          RatonNormal
          MsgBox "Proceso Terminado, vuelva a listar el documento"
       End If
    End If
    
    If CtrlDown And KeyCode = vbKey0 Then
       Titulo = "PREGUNTA DE ELIMINACION"
       Mensajes = "Esta seguro de Eliminar el Kardex de comprobantes en cero de " & MBoxFechaI.Text & " - " & MBoxFechaF.Text & "?"
       If BoxMensaje = vbYes Then
          RatonReloj
          sSQL = "DELETE * " _
               & "FROM Trans_Kardex " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
               & "AND Numero = 0 "
          Ejecutar_SQL_SP sSQL
          RatonNormal
          MsgBox "Proceso Terminado, vuelva a listar el Kardex"
       End If
    End If
 Else
    MsgBox "No hay regstros que procesar"
 End If
End Sub

Private Sub Form_Activate()
  MBoxFechaI = FechaSistema
  MBoxFechaF = FechaSistema
  PorBodega = False
  PorCodigoBarra = False
  PorTodos = False
  
''  sSQL = "UPDATE Catalogo_Productos " _
''       & "SET X = '.' " _
''       & "WHERE Item = '" & NumEmpresa & "' " _
''       & "AND Periodo = '" & Periodo_Contable & "' "
''  Ejecutar_SQL_SP sSQL
''
''  sSQL = "UPDATE Catalogo_Productos " _
''       & "SET X = 'M' " _
''       & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
''       & "WHERE CP.Item = '" & NumEmpresa & "' " _
''       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
''       & "AND CP.Item = TK.Item " _
''       & "AND CP.Periodo = TK.Periodo " _
''       & "AND CP.Codigo_Inv = TK.Codigo_Inv "
''  Ejecutar_SQL_SP sSQL
''
''  sSQL = "SELECT Codigo_Inv " _
''       & "FROM Catalogo_Productos " _
''       & "WHERE Item = '" & NumEmpresa & "' " _
''       & "AND Periodo = '" & Periodo_Contable & "' " _
''       & "AND TC = 'P' " _
''       & "AND X = 'M' " _
''       & "GROUP BY Codigo_Inv " _
''       & "ORDER BY Codigo_Inv "
''  Select_Adodc AdoKardex, sSQL
''  With AdoKardex.Recordset
''   If .RecordCount > 0 Then
''       Codigo = CodigoCuentaSup(.Fields("Codigo_Inv"))
''       Do While Not .EOF
''          If Codigo <> CodigoCuentaSup(.Fields("Codigo_Inv")) Then
''             Do While Len(Codigo) > 1
''                sSQL = "UPDATE Catalogo_Productos " _
''                     & "SET X = 'M' " _
''                     & "WHERE Item = '" & NumEmpresa & "' " _
''                     & "AND Periodo = '" & Periodo_Contable & "' " _
''                     & "AND X <> 'M' " _
''                     & "AND Codigo_Inv = '" & Codigo & "' "
''                Ejecutar_SQL_SP sSQL
''               ' MsgBox sSQL
''                Codigo = CodigoCuentaSup(Codigo)
''             Loop
''             Codigo = CodigoCuentaSup(.Fields("Codigo_Inv"))
''          End If
''         .MoveNext
''       Loop
''       Do While Len(Codigo) > 1
''          sSQL = "UPDATE Catalogo_Productos " _
''               & "SET X = 'M' " _
''               & "WHERE Item = '" & NumEmpresa & "' " _
''               & "AND Periodo = '" & Periodo_Contable & "' " _
''               & "AND X <> 'M' " _
''               & "AND Codigo_Inv = '" & Codigo & "' "
''          Ejecutar_SQL_SP sSQL
''         ' MsgBox sSQL
''          Codigo = CodigoCuentaSup(Codigo)
''       Loop
''   End If
''  End With

  ListarProductos "I"
  ListarProductos "P"
  Listar_Articulos
    
  DGKardex.width = MDI_X_Max - 100
  DGKardex.Height = (MDI_Y_Max - DGKardex.Top - 500)
  AdoKardex.Top = DGKardex.Top + DGKardex.Height + 10
  AdoKardex.width = MDI_X_Max - 100
  
  FrmProductos.Left = ((Screen.width - FrmProductos.width) / 2)
  FrmProductos.Top = ((Screen.Height - FrmProductos.Height) / 2)
  
'''   MSFlexGrid1.width = MDI_X_Max - 100
'''   DGAsiento.width = MDI_X_Max - 100
'''   LblConcepto.width = MDI_X_Max - (Command1.width + Command4.width) - 130
'''
'''   MSFlexGrid1.Height = (MDI_Y_Max - MSFlexGrid1.Top - 900) / 2
'''   DGAsiento.Height = MSFlexGrid1.Height
  
  RatonNormal
  Existen.Caption = "EXISTENCIA DE INVENTARIO"
End Sub

Private Sub Form_Load()
 'CentrarForm Existen
  ConectarAdodc AdoArt
  ConectarAdodc AdoInv
  ConectarAdodc AdoTInv
  ConectarAdodc AdoSaldos
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  DGKardex.ToolTipText = "<Ctrl><F11>: Imprime Codigos de Barra, <Ctrl><F12>: Cambiar Articulo, <Ctrl><B>: Cambia Codigo de Barra, <Ctrl><S>: Cambia la Serie"
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

Private Sub OpcBodega_Click()
    PorBodega = True
    PorCodigoBarra = False
    PorTodos = False
    Listar_Bodega_CodigoBarra
    DCBodega.Visible = True
End Sub

Private Sub OpcCodigoBarra_Click()
    PorBodega = False
    PorCodigoBarra = True
    PorTodos = False
    Listar_Bodega_CodigoBarra
    DCBodega.Visible = True
End Sub

Private Sub OpcTodos_Click()
    PorBodega = False
    PorCodigoBarra = False
    PorTodos = True
    DCBodega.Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  
   FechaValida MBoxFechaI
   FechaValida MBoxFechaF
     
   FechaIni = BuscarFecha(MBoxFechaI)
   FechaFin = BuscarFecha(MBoxFechaF)
   Codigo1 = SinEspaciosDer(DCBodega)
   If Codigo1 = "" Then Codigo1 = "."
     
  'MsgBox Button.key
   Select Case Button.key
     Case "Consultar"
          Consultar_Tipo_De_Kardex True
     Case "Kardex"
          Consultar_Kardex
     Case "Kardex_Total"
          Consultar_Tipo_De_Kardex False
     Case "Kardex_Total_QR"
          Kardex_Total_QR
     Case "Imprimir_Kardex"
          DGKardex.Visible = False
          ImprimirKardex AdoKardex, AdoArt
          Existen.Caption = "EXISTENCIA DE INVENTARIO"
          DGKardex.Visible = True
     Case "Excel"
          DGKardex.Visible = False
          Exportar_AdoDB_Excel AdoKardex.Recordset, "Kardex del " & BuscarFecha(MBoxFechaI) & " al " & BuscarFecha(MBoxFechaF)
          DGKardex.Visible = True
     Case "Salir"
          Unload Existen
   End Select
End Sub

Public Sub Listar_Articulos(Optional SoActivos As Boolean)
  sSQL = "SELECT Codigo_Inv,Producto,Unidad,Bodega,Minimo,Maximo,Cta_Inventario,Cta_Costo_Venta " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If SoActivos Then sSQL = sSQL & "AND T = 'N' "
  sSQL = sSQL _
       & "AND TC = 'P' " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "AND LEN(Cta_Inventario) > 1 " _
       & "AND LEN(Cta_Costo_Venta) > 1 " _
       & "ORDER BY Codigo_Inv "
  SelectDB_Combo DCArt, AdoArt, sSQL, "Producto"
End Sub
