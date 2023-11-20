VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FComprobantes 
   Caption         =   "COMPROBANTE DE EGRESO"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   11595
   WindowState     =   1  'Minimized
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Comproba.frx":0000
      DataSource      =   "AdoBenef"
      Height          =   345
      Left            =   1470
      TabIndex        =   10
      Top             =   840
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Beneficiario"
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
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
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
      Picture         =   "Comproba.frx":0017
      TabIndex        =   56
      Top             =   8400
      Width           =   1380
   End
   Begin VB.TextBox TxtEmail 
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
      MaxLength       =   60
      TabIndex        =   14
      ToolTipText     =   "Escriba el Nombre o C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   1575
      Width           =   8625
   End
   Begin VB.CheckBox CheqCopia 
      Caption         =   "Imprimir con copia"
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
      Left            =   11970
      TabIndex        =   71
      Top             =   105
      Width           =   2010
   End
   Begin VB.OptionButton OpcTP 
      Caption         =   "&5.- Nota de Crédito"
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
      Index           =   4
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.OptionButton OpcTP 
      Caption         =   "&4.- Nota de Débito"
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
      Index           =   3
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1905
   End
   Begin VB.OptionButton OpcTP 
      Caption         =   "&3.- Egreso"
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
      Index           =   2
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.OptionButton OpcTP 
      Caption         =   "&2.- Ingreso"
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
      Index           =   1
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   1170
   End
   Begin VB.OptionButton OpcTP 
      Caption         =   "&1.- Diario"
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
      Index           =   0
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Value           =   -1  'True
      Width           =   1170
   End
   Begin MSDataListLib.DataList DLCuentas 
      Bindings        =   "Comproba.frx":0119
      DataSource      =   "AdoCuentas"
      Height          =   2700
      Left            =   1890
      TabIndex        =   42
      Top             =   5250
      Visible         =   0   'False
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   4763
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameAsigna 
      Height          =   1905
      Left            =   11550
      TabIndex        =   45
      Top             =   5250
      Visible         =   0   'False
      Width           =   2325
      Begin VB.TextBox TxtCheqDep 
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
         Left            =   945
         MaxLength       =   16
         TabIndex        =   49
         Text            =   "0"
         Top             =   525
         Width           =   1275
      End
      Begin VB.TextBox TextOpcTM 
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
         Height          =   405
         Left            =   1785
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "Comproba.frx":0132
         Top             =   945
         Width           =   435
      End
      Begin VB.TextBox TextOpcDH 
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
         Height          =   405
         Left            =   1785
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   55
         Text            =   "Comproba.frx":0134
         Top             =   1365
         Width           =   435
      End
      Begin MSMask.MaskEdBox MBEfectivizar 
         Height          =   330
         Left            =   945
         TabIndex        =   47
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
      Begin VB.Label Label21 
         Caption         =   " Chq/Dep"
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
         TabIndex        =   48
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   " Efectiv."
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
         TabIndex        =   46
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   " Haber    2"
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
         TabIndex        =   54
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   " 2.- M/E"
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
         Left            =   735
         TabIndex        =   51
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   " Debe     1"
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
         TabIndex        =   53
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Valores: 1.- M/N"
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
         TabIndex        =   50
         Top             =   945
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Co&nversión"
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
      Left            =   10290
      TabIndex        =   17
      Top             =   1155
      Width           =   1905
      Begin VB.OptionButton OpcDiv 
         Caption         =   "(/)"
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
         Left            =   210
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton OpcMult 
         Caption         =   "(x)"
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
         Left            =   1050
         TabIndex        =   19
         Top             =   315
         Width           =   645
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
      Left            =   8820
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1575
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   840
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2745
      Left            =   105
      TabIndex        =   61
      Top             =   5355
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   4842
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&4.- CONTABILIZACION"
      TabPicture(0)   =   "Comproba.frx":0136
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGAsientos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&5.- SUBCUENTAS"
      TabPicture(1)   =   "Comproba.frx":0152
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGAsientosSC"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&6.- RETENCIONES"
      TabPicture(2)   =   "Comproba.frx":016E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGAsientosR"
      Tab(2).Control(1)=   "DGAC"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&7.- AC - AV - AI - AE"
      TabPicture(3)   =   "Comproba.frx":018A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGAE"
      Tab(3).Control(1)=   "DGAI"
      Tab(3).Control(2)=   "DGAV"
      Tab(3).ControlCount=   3
      Begin MSDataGridLib.DataGrid DGAC 
         Bindings        =   "Comproba.frx":01A6
         Height          =   750
         Left            =   -74895
         TabIndex        =   69
         Top             =   420
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   1323
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAE 
         Bindings        =   "Comproba.frx":01BA
         Height          =   750
         Left            =   -74895
         TabIndex        =   68
         Top             =   1890
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   1323
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAI 
         Bindings        =   "Comproba.frx":01CE
         Height          =   750
         Left            =   -74895
         TabIndex        =   67
         Top             =   1155
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   1323
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAV 
         Bindings        =   "Comproba.frx":01E2
         Height          =   750
         Left            =   -74895
         TabIndex        =   70
         Top             =   420
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   1323
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAsientosR 
         Bindings        =   "Comproba.frx":01F6
         Height          =   1380
         Left            =   -74895
         TabIndex        =   66
         Top             =   1155
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   2434
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAsientosSC 
         Bindings        =   "Comproba.frx":0211
         Height          =   2220
         Left            =   -74895
         TabIndex        =   65
         Top             =   420
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   3916
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
      Begin MSDataGridLib.DataGrid DGAsientos 
         Bindings        =   "Comproba.frx":022D
         Height          =   2220
         Left            =   105
         TabIndex        =   64
         Top             =   315
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   3916
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
   End
   Begin MSAdodcLib.Adodc AdoSQL 
      Height          =   330
      Left            =   210
      Top             =   5670
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
      Caption         =   "SQL"
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
      Height          =   540
      Left            =   1785
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4095
      Width           =   12195
   End
   Begin VB.CommandButton CmdCancelar 
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
      Height          =   330
      Left            =   1575
      Picture         =   "Comproba.frx":0247
      TabIndex        =   57
      Top             =   8400
      Width           =   1275
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
      Left            =   11550
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   4935
      Width           =   2325
   End
   Begin VB.TextBox TextCuenta 
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
      Left            =   1890
      TabIndex        =   41
      Top             =   4935
      Width           =   9570
   End
   Begin VB.TextBox TextCodigo 
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
      TabIndex        =   39
      Text            =   "0"
      ToolTipText     =   "<Ctrl+B> Por Patrón de Búsqueda"
      Top             =   4935
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   22
      Top             =   1995
      Visible         =   0   'False
      Width           =   13875
      Begin VB.TextBox TextDeposito 
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
         Left            =   12285
         MaxLength       =   16
         TabIndex        =   34
         Text            =   "0"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid DGAsientosB 
         Bindings        =   "Comproba.frx":0689
         Height          =   960
         Left            =   105
         TabIndex        =   35
         Top             =   945
         Visible         =   0   'False
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   1693
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
         AllowDelete     =   -1  'True
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
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "Comproba.frx":06A4
         DataSource      =   "AdoBanco"
         Height          =   345
         Left            =   1260
         TabIndex        =   28
         Top             =   525
         Visible         =   0   'False
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "Banco"
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
      Begin MSDataListLib.DataCombo DCCaja 
         Bindings        =   "Comproba.frx":06BB
         DataSource      =   "AdoCaja"
         Height          =   345
         Left            =   1260
         TabIndex        =   24
         Top             =   210
         Visible         =   0   'False
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "Caja"
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
      Begin VB.TextBox TextCheque 
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
         Left            =   12075
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "Comproba.frx":06D1
         Top             =   525
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TextCantidad 
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
         Left            =   12075
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "Comproba.frx":06D3
         Top             =   210
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox CheckEfect 
         Caption         =   "&Efectivo"
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
         TabIndex        =   23
         Top             =   210
         Width           =   1065
      End
      Begin VB.CheckBox CheckBco 
         Caption         =   "&Banco "
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
         TabIndex        =   27
         Top             =   525
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBFechaEfec 
         Height          =   330
         Left            =   11025
         TabIndex        =   32
         Top             =   1260
         Visible         =   0   'False
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
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Efecti&vizar"
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
         TabIndex        =   31
         Top             =   945
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Depósito No."
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
         Left            =   12285
         TabIndex        =   33
         Top             =   945
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQ."
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
         TabIndex        =   29
         Top             =   525
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQ."
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
         TabIndex        =   25
         Top             =   210
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   210
      Top             =   5985
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
      Caption         =   "Caja"
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
      Left            =   210
      Top             =   6300
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
      Left            =   210
      Top             =   6615
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   210
      Top             =   6930
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
   Begin MSAdodcLib.Adodc AdoAsientosSC 
      Height          =   330
      Left            =   2205
      Top             =   5670
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
      Caption         =   "AsientosSC"
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
   Begin MSAdodcLib.Adodc AdoAsientosB 
      Height          =   330
      Left            =   2205
      Top             =   5985
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
      Caption         =   "AsientosB"
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
      Left            =   2205
      Top             =   6300
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin MSAdodcLib.Adodc AdoAsientosR 
      Height          =   330
      Left            =   2205
      Top             =   6615
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
      Caption         =   "AsientosR"
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   4410
      Top             =   6930
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
      Caption         =   "Benef"
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
   Begin MSAdodcLib.Adodc AdoRolPago 
      Height          =   330
      Left            =   2205
      Top             =   6930
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
      Caption         =   "RolPago"
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
   Begin MSAdodcLib.Adodc AdoAC 
      Height          =   330
      Left            =   4410
      Top             =   5670
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
      Caption         =   "AC"
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
   Begin MSAdodcLib.Adodc AdoAV 
      Height          =   330
      Left            =   4410
      Top             =   5985
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
      Caption         =   "AV"
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
   Begin MSAdodcLib.Adodc AdoAI 
      Height          =   330
      Left            =   4410
      Top             =   6300
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
      Caption         =   "AI"
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
   Begin MSAdodcLib.Adodc AdoAE 
      Height          =   330
      Left            =   4410
      Top             =   6615
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
      Caption         =   "AE"
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
   Begin MSAdodcLib.Adodc AdoCentroCostos 
      Height          =   330
      Left            =   210
      Top             =   7245
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
      Caption         =   "CentroCostos"
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
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999999999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7350
      TabIndex        =   72
      Top             =   525
      Width           =   4320
   End
   Begin VB.Label LblRUC 
      Alignment       =   2  'Center
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
      Left            =   11760
      TabIndex        =   12
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label Label22 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " E&MAIL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   1260
      Width           =   8625
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
      Height          =   330
      Left            =   12285
      TabIndex        =   21
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &COTIZACION:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8820
      TabIndex        =   15
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label20 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &R.U.C./C.I."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11760
      TabIndex        =   11
      Top             =   525
      Width           =   2220
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &PAGADO A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1470
      TabIndex        =   9
      Top             =   525
      Width           =   5895
   End
   Begin VB.Label LabelComp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   10290
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " Ingreso No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7455
      TabIndex        =   5
      Top             =   105
      Width           =   2850
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
      Left            =   12075
      TabIndex        =   58
      Top             =   8190
      Width           =   1800
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
      Left            =   10290
      TabIndex        =   59
      Top             =   8190
      Width           =   1800
   End
   Begin VB.Label LabelDiferencia 
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
      Left            =   7140
      TabIndex        =   62
      Top             =   8190
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808080&
      Caption         =   "VALOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   11550
      TabIndex        =   43
      Top             =   4725
      Width           =   2325
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      Caption         =   " DIGITE LA CLAVE O SELECCIONE LA CUENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1890
      TabIndex        =   40
      Top             =   4725
      Width           =   9570
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12285
      TabIndex        =   20
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia"
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
      Left            =   5985
      TabIndex        =   63
      Top             =   8190
      Width           =   1170
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      Caption         =   " CO&DIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   38
      Top             =   4725
      Width           =   1695
   End
   Begin VB.Label Label19 
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
      Left            =   9240
      TabIndex        =   60
      Top             =   8190
      Width           =   1065
   End
   Begin VB.Label Label7 
      Caption         =   " P&or concepto de:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   36
      Top             =   4095
      Width           =   1590
   End
End
Attribute VB_Name = "FComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim EmailOld As String
'1768014170001
'TAME
Dim TipoBusqueda As String

Public Sub Asientos_Grabados()
  sSQL = "SELECT " & Full_Fields("Asiento") & " " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY A_No "
  Select_Adodc_Grid DGAsientos, AdoAsientos, sSQL
  
  SQL2 = "SELECT " & Full_Fields("Asiento_SC") & " " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientosSC, SQL2
End Sub

Private Sub Tipo_De_Comprobante_No(C1 As Comprobantes)

  TextCotiza = Dolar
  Frame2.Visible = False
  'FrmBenef.Visible = False
  CheckEfect.Visible = False
  CheckBco.Value = False
  Label6.Caption = Moneda
 'Determinamos que tipo de comprobante realizamos
  Select Case C1.TP
    Case CompDiario
         NumComp = ReadSetDataNum("Diario", True, False)
         Label2.Caption = "Diario No. "
         Label3.Caption = " &BENEFICIARIO:"
         FComprobantes.Caption = "COMPROBANTE DE DIARIO"
         Label2.Caption = "Diario No. " & Format(Year(MBoxFecha), "0000") & "- "
         OpcTP(0).Value = True
    Case CompIngreso
         Frame2.Visible = True
         NumComp = ReadSetDataNum("Ingresos", True, False)
         Label2.Caption = "Ingreso No. "
         Label3.Caption = " RECI&BI DE:"
         CheckEfect.Visible = 1
'         CheckBco.value = 1
         FComprobantes.Caption = "COMPROBANTE DE INGRESO"
         Label2.Caption = "Ingreso No. " & Format(Year(MBoxFecha), "0000") & "- "
         OpcTP(1).Value = True
    Case CompEgreso
         Frame2.Visible = True
         NumComp = ReadSetDataNum("Egresos", True, False)
         Label2.Caption = "Egreso No. "
         Label3.Caption = " &PAGADO A:"
         CheckEfect.Visible = 1
 '        CheckBco.value = 1
         FComprobantes.Caption = "COMPROBANTE DE EGRESO"
         Label2.Caption = "Egreso No. " & Format(Year(MBoxFecha), "0000") & "- "
         OpcTP(2).Value = True
    Case CompNotaDebito
         NumComp = ReadSetDataNum("NotaDebito", True, False)
         Label2.Caption = "Nota de Débito No. "
         Label3.Caption = " &BENEFICIARIO:"
         FComprobantes.Caption = "COMPROBANTE DE NOTA DE DEBITO"
         Label2.Caption = "Nota de Debito No. " & Format(Year(MBoxFecha), "0000") & "- "
         OpcTP(3).Value = True
    Case CompNotaCredito
         NumComp = ReadSetDataNum("NotaCredito", True, False)
         Label2.Caption = "Nota de Débito No. "
         Label3.Caption = " &BENEFICIARIO:"
         FComprobantes.Caption = "COMPROBANTE DE NOTA DE CREDITO"
         Label2.Caption = "Nota de Credito No. " & Format(Year(MBoxFecha), "0000") & "- "
         OpcTP(4).Value = True
  End Select
  If ModificarComp Then NumComp = C1.Numero
    
 'Presentamos la informacion del Asiento y Anexo en este comprobante
  SQL2 = "SELECT " & Full_Fields("Asiento") & " " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY A_No "
  Select_Adodc_Grid DGAsientos, AdoAsientos, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_SC") & " " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY SC_No "
  Select_Adodc_Grid DGAsientosSC, AdoAsientosSC, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_B") & " " _
       & "FROM Asiento_B " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAsientosB, AdoAsientosB, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_Air") & " " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAsientosR, AdoAsientosR, SQL2
    
  SQL2 = "SELECT " & Full_Fields("Asiento_Compras") & " " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAC, AdoAC, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_Ventas") & " " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAV, AdoAV, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_Exportaciones") & " " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAE, AdoAE, SQL2
  
  SQL2 = "SELECT " & Full_Fields("Asiento_Importaciones") & " " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGAI, AdoAI, SQL2
  LabelComp.Caption = Format(NumComp, "00000000")
  MBoxFecha.SetFocus
End Sub

Private Sub CheckBco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CheckEfect_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  TipoBusqueda = ""
End Sub

Private Sub DCCliente_GotFocus()
  TipoBusqueda = "%"
  Label23.Caption = ""
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) >= 1 Then
       sSQL = "SELECT TOP 50 Cliente,CI_RUC,Codigo " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoBenef, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
    With AdoBenef.Recordset
      If Len(DCCliente) <= 1 Then
         Co.CodigoB = Ninguno
         Co.Beneficiario = Ninguno
         Co.RUC_CI = Ninguno
         Co.Email = Ninguno
         TipoSRI.Estado = Ninguno
      Else
         If .RecordCount >= 1 Then
             Co.CodigoB = Ninguno
             RatonReloj
            .MoveFirst
             If IsNumeric(DCCliente.Text) Then DCCliente.Text = .fields("Cliente")
            .Find ("Cliente = '" & DCCliente & "' ")
             If Not .EOF Then Co.CodigoB = .fields("Codigo")
             DCCliente.Height = 0
             Datos_del_Cliente Co
             Llenar_Encabezado_Comprobante
             MarcarTexto TxtEmail
             RatonNormal
             TxtEmail.SetFocus
         Else
            'FrmBenef.Visible = False
             NombreCliente = DCCliente
             Nuevo = True
             FClientesFlash.Show 1
             TipoBusqueda = "%"
             DCCliente.Height = 0
             DCCliente = NombreCliente
             DCCliente.SetFocus
         End If
      End If
    End With
    If UCase(TipoSRI.Estado) = "ACTIVO" Then Label23.ForeColor = &HC0FFC0 Else Label23.ForeColor = &HFFFFFF
    Label23.Caption = TipoSRI.Estado
End Sub

Private Sub DGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGAsientos.Visible = False
     GenerarDataTexto FComprobantes, AdoAsientos
     DGAsientos.Visible = True
  End If
End Sub

Private Sub DGAsientosR_BeforeDelete(Cancel As Integer)
  Codigo = AdoAsientosR.Recordset.fields("Cta")
  Cancel = DeleteSiNo(AdoAsientosR)
End Sub

Private Sub DGAsientosSC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF5 Then Asientos_Grabados
  If KeyCode = vbKeyF1 Then
     DGAsientosSC.Visible = False
     GenerarDataTexto FComprobantes, AdoAsientosSC
     DGAsientosSC.Visible = True
  End If
End Sub

Private Sub DLCuentas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEsc = False
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     PresionoEsc = True
     TextCodigo.SetFocus
  End If
End Sub

Private Sub DLCuentas_LostFocus()
  DLCuentas.Visible = False
  If PresionoEsc Then
     PresionoEsc = False
     TextCodigo.Text = ""
     TextCodigo.SetFocus
  Else
     Cadena = SinEspaciosIzq(DLCuentas.Text)
     Codigo = Leer_Cta_Catalogo(Cadena)
     TextCodigo.Text = Codigo
     TextCuenta.Text = Cuenta
     FrameAsigna.Visible = True
     If OpcCoop Then
        Label11.Visible = False
        Label16.Visible = False
        TextOpcTM.Visible = False
        TextOpcDH.SetFocus
     Else
        If SubCta = "BA" Then
           Label12.Visible = True
           Label21.Visible = True
           TxtCheqDep.Visible = True
           MBEfectivizar.Visible = True
           MBEfectivizar.SetFocus
        Else
           TextOpcTM.SetFocus
        End If
     End If
  End If
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         If OpcTP(1).Value Or OpcTP(4).Value Then
            CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
            If CheckBco.Value Then
               With AdoAsientosB.Recordset
                If .RecordCount > 0 Then
                   .MoveFirst
                    Do While Not .EOF
                       Cadena = .fields("CTA_BANCO")
                       Valor = .fields("VALOR")
                       Codigo = Leer_Cta_Catalogo(Cadena)
                       OpcDH = 1: ValorDH = Valor
                       FechaValida MBFechaEfec
                       TextoValido TextDeposito, , True
                       Fecha_Vence = .fields("EFECTIVIZAR")
                       NoCheque = .fields("CHEQ_DEP")
                       InsertarAsientosC AdoAsientos
                      .MoveNext
                    Loop
                End If
               End With
            End If
            SumaBancos = CalculosTotalBancos(AdoAsientosB)
            Monto_Total = SumaBancos + Abono
            LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
            TextConcepto.SetFocus
         End If
         If OpcTP(4).Value Then TextConcepto.SetFocus
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub DCBanco_LostFocus()
Dim AdoCheq As ADODB.Recordset
  LeerBanco AdoCtas, DCBanco.Text
  DCBanco.Text = Codigo & Space(5) & NombreBanco
  If OpcTP(2).Value Then
     Label24.Caption = "Cheq. No."
     Set AdoCheq = New ADODB.Recordset
     AdoCheq.CursorType = adOpenStatic
     AdoCheq.CursorLocation = adUseClient
     Cta_Aux = SinEspaciosIzq(DCBanco)
     If Cta_Aux = "" Then Cta_Aux = "."
     TextCheque = "00000001"
     sSQL = "SELECT MAX(Cheq_Dep) As Ultimo_Chep " _
          & "FROM Transacciones " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cta = '" & Cta_Aux & "' " _
          & "AND TP = 'CE' " _
          & "AND ISNUMERIC(Cheq_Dep) <> 0 " _
          & "AND Haber > 0 " _
          & "AND Fecha <= #" & BuscarFecha(MBoxFecha) & "# "
          '& "ORDER BY Cheq_Dep DESC "   'Fecha DESC
     sSQL = CompilarSQL(sSQL)
     
     AdoCheq.open sSQL, AdoStrCnn, , , adCmdText
    'Seteamos los encabezados para las facturas
     If AdoCheq.RecordCount > 0 Then
        If Not IsNull(AdoCheq.fields("Ultimo_Chep")) Then
           TextCheque = Format(Val(AdoCheq.fields("Ultimo_Chep")) + 1, "00000000")
        End If
     End If
     AdoCheq.Close
  Else
     If Moneda_US Then Label24.Caption = "M/E" Else Label24.Caption = Moneda
  End If
End Sub

Private Sub DCCaja_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCaja_LostFocus()
  LeerBanco AdoCtas, DCCaja.Text
  If Moneda_US Then Label6.Caption = "M/E" Else Label6.Caption = Moneda
End Sub

Private Sub DGAsientosB_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBEfectivizar_GotFocus()
  MBEfectivizar.Text = MBoxFecha.Text
  MarcarTexto MBEfectivizar
End Sub

Private Sub MBEfectivizar_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBEfectivizar_LostFocus()
  FechaValida MBEfectivizar
End Sub

Private Sub MBFechaEfec_GotFocus()
  MBFechaEfec.Text = MBoxFecha.Text
  MarcarTexto MBFechaEfec
End Sub

Private Sub MBFechaEfec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEfec_LostFocus()
  FechaValida MBFechaEfec
  If OpcTP(2).Value Then
     NoCheque = TextCheque.Text
     Fecha_Vence = MBFechaEfec.Text
     SQL1 = "DELETE * " _
          & "FROM Asiento_B " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Ejecutar_SQL_SP SQL1
     InsertarAsientoBanco AdoAsientosB, 0
     TextConcepto.SetFocus
  Else
     TextCheque.Text = Format(TextCheque.Text, "#,##0.00")
  End If
End Sub

Private Sub OpcTP_Click(Index As Integer)
  Select Case Index
    Case 0: Co.TP = CompDiario
    Case 1: Co.TP = CompIngreso
    Case 2: Co.TP = CompEgreso
    Case 3: Co.TP = CompNotaDebito
    Case 4: Co.TP = CompNotaCredito
  End Select
  Tipo_De_Comprobante_No Co
End Sub

Private Sub OpcTP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then MBoxFecha.SetFocus
End Sub

Private Sub TextCotiza_GotFocus()
  MarcarTexto TextCotiza
End Sub

Private Sub TextOpcDH_Change()
  If 1 > Val(TextOpcDH) Or Val(TextOpcDH) > 2 Then
     TextOpcDH.Text = ""
  Else
     OpcDH = Val(TextOpcDH)
     SiguienteControl
  End If
End Sub

Private Sub TextOpcDH_GotFocus()
  TextOpcDH.Text = ""
End Sub

Private Sub TextOpcDH_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then TextOpcDH.Text = "1"
End Sub

Private Sub TextOpcDH_LostFocus()
Dim EsRetencion As Boolean
  EsRetencion = True
  TextOpcDH = UCaseStrg(TextOpcDH)
  TextoValido TxtCheqDep, , True
  FechaValida MBEfectivizar
  Fecha_Vence = MBEfectivizar.Text
  If SubCta = "BA" Then NoCheque = TxtCheqDep.Text Else NoCheque = Ninguno
  If OpcCoop Then
     If Moneda_US Then OpcTM = 2 Else OpcTM = 1
  End If
  If IsNumeric(TextOpcDH) Then OpcDH = Val(TextOpcDH)
  If OpcTM >= 1 And OpcDH >= 1 Then
     FrameAsigna.Visible = False
     Select Case SubCta
       Case "C", "P", "G", "I", "PM"
            Label17.Caption = "VALOR M/N"
            If Moneda_US Or OpcTM = 2 Then Label17.Caption = "VALOR M/E"
            FechaTexto = MBoxFecha
            SubCtaGen = Codigo
            FSubCtas.Show 1
            TextCuenta.Text = ""
            Asientos_Grabados
       Case "CP"
            FechaTexto = MBoxFecha
            Nombre_Cta_Ret = Cuenta
            SubCtaGen = Codigo
            If OpcDH > 1 Then FSubCtas.Show 1
'''            Else
'''               Aprobacion.Show 1
'''            End If
       Case "CC"
            If CentroDeCosto Then
               CodigoCC = Ninguno
               FechaTexto = MBoxFecha
               SubCtaGen = Codigo
               FCentroCostos.Show 1
               Asientos_Grabados
            End If
     End Select
     SQL2 = "SELECT * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "ORDER BY Cta, Codigo, SC_No "
     Select_Adodc_Grid DGAsientosSC, AdoAsientosSC, SQL2
     TextValor.Visible = True
     TextValor.SetFocus
  Else
     If OpcCoop Then
        FrameAsigna.Visible = False
     Else
        TextOpcTM.SetFocus
     End If
  End If
End Sub

Private Sub TextOpcTM_Change()
  If 1 > Val(TextOpcTM.Text) Or Val(TextOpcTM.Text) > 2 Then
     TextOpcTM.Text = ""
  Else
     TextOpcDH.SetFocus
  End If
End Sub

Private Sub TextOpcTM_GotFocus()
  TextOpcTM.Text = ""
  TextOpcDH.Text = ""
End Sub

Private Sub TextOpcTM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then TextOpcTM.Text = "1"
End Sub

Private Sub TextOpcTM_LostFocus()
  OpcTM = Val(TextOpcTM.Text)
End Sub

Private Sub Form_Deactivate()
  FComprobantes.WindowState = vbMinimized
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
 ' MsgBox FechaComp
''  If FechaComp <> Ninguno Then
''
''  Else
''     FechaValida MBoxFecha, False
''  End If
   
  FechaValida MBoxFecha, True
  Validar_Porc_IVA MBoxFecha
  MBFechaEfec = MBoxFecha
  MBEfectivizar = MBoxFecha
  NombreCliente = DCCliente
  Co.Beneficiario = NombreCliente
  Co.Fecha = MBoxFecha
  FechaComp = MBoxFecha
End Sub

Private Sub TextCantidad_GotFocus()
  TextCantidad.Text = ""
  MarcarTexto TextCantidad
End Sub

Private Sub TextCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheque_GotFocus()
  MarcarTexto TextCheque
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheque_LostFocus()
Dim AdoCheq As ADODB.Recordset
  Set AdoCheq = New ADODB.Recordset
  AdoCheq.CursorType = adOpenStatic
  AdoCheq.CursorLocation = adUseClient
  If OpcTP(1).Value Then
     TextoValido TextCheque, True
  Else
     TextoValido TextCheque, False
  End If
  Cta_Aux = SinEspaciosIzq(DCBanco)
  If Cta_Aux = "" Then Cta_Aux = "."
  sSQL = "SELECT * " _
       & "FROM Transacciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Cta = '" & Cta_Aux & "' " _
       & "AND Cheq_Dep = '" & TextCheque & "' " _
       & "AND TP = 'CE' " _
       & "AND Haber > 0 "
  'MsgBox sSQL
  AdoCheq.open sSQL, AdoStrCnn, , , adCmdText
 'Seteamos los encabezados para las facturas
  If AdoCheq.RecordCount > 0 Then
     MsgBox "Warning: El Cheque No. " & TextCheque & ", ya esta ingresado."
  End If
  AdoCheq.Close
  If OpcTP(2).Value And IsNumeric(TextCheque) Then TextCheque = Format(Val(TextCheque), "00000000")
End Sub

Private Sub TextDeposito_GotFocus()
  TextDeposito.Text = ""
End Sub

Private Sub TextDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDeposito_LostFocus()
  TextoValido TextDeposito, False
  Valor = CDbl(TextCheque.Text)
  If Valor <> 0 Then
     NoCheque = TextDeposito.Text
     Fecha_Vence = MBFechaEfec.Text
     InsertarAsientoBanco AdoAsientosB, Valor
  End If
  If IsNumeric(TextDeposito) Then TextDeposito = Format(Val(TextDeposito), "00000000")
  DCBanco.SetFocus
End Sub

Private Sub TxtCheqDep_GotFocus()
  TxtCheqDep.Text = ""
End Sub

Private Sub TxtCheqDep_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCheqDep_LostFocus()
  TextoValido TxtCheqDep
End Sub

Private Sub TextConcepto_Change()
  If TextoMaximo(TextConcepto) Then TextCodigo.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
  Monto_Total = SumaBancos + Abono
  LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_LostFocus()
  TextoValido TextCotiza, True
  If Val(TextCotiza.Text) <= 0 Then TextCotiza.Text = Dolar
  Dolar = Val(TextCotiza.Text)
End Sub

Private Sub TextValor_GotFocus()
  RatonNormal
  If Datos_De_Empresa("Det_Comp") Then
     DetalleComp = TrimStrg(MidStrg(InputBox("Detalle del Concepto:", "CONCEPTO AUXILIAR", ""), 1, 45))
  Else
     DetalleComp = Ninguno
  End If
  If DetalleComp = "" Then DetalleComp = Ninguno
  Select Case SubCta
    Case "C", "P", "G", "I", "CP", "PM", "CC"
         Codigo = Leer_Cta_Catalogo(TextCodigo)
         ValorDH = SumatoriaSC
         InsertarAsiento AdoAsientos
         SiguienteControl
    Case Else
         TextValor.Text = ""
  End Select
End Sub

Private Sub TextValor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValor_LostFocus()
  Select Case SubCta
    Case "C", "P", "G", "I", "CP", "PM", "CC"
        'No inserta nada porque ya se inserto
    Case Else
         TextoValido TextValor, True
         TotalSubCta = 0: SubCtaGen = Ninguno
         ValorDH = CCur(Val(TextValor))
         InsertarAsiento AdoAsientos
  End Select
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
  TextCodigo.SetFocus
End Sub

Private Sub DGAsientos_BeforeDelete(Cancel As Integer)
Dim OpcDH_A As Byte
  OpcDH_A = 0
  Codigo = AdoAsientos.Recordset.fields("CODIGO")
 ' Ln_No_A = AdoAsientos.Recordset.Fields("SC_No")
  If AdoAsientos.Recordset.fields("DEBE") > 0 Then OpcDH_A = 1
  If AdoAsientos.Recordset.fields("HABER") > 0 Then OpcDH_A = 2
  Cancel = DeleteSiNo(AdoAsientos)
  If Cancel = False Then
     EliminarSubCta Codigo, OpcDH_A
     SQL2 = "SELECT * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc_Grid DGAsientosSC, AdoAsientosSC, SQL2
     SQL2 = "SELECT * " _
          & "FROM Asiento_Air " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc_Grid DGAsientosR, AdoAsientosR, SQL2
  End If
  RatonNormal
End Sub

Private Sub DGAsientos_GotFocus()
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub DGAsientos_LostFocus()
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub CmdCancelar_Click()
  ModificarComp = False
  CopiarComp = False
  NuevoComp = True
  Unload FComprobantes
End Sub

Private Sub CmdGrabar_Click()
  Asientos_Grabados
  OpcSubCtaDH = 1
  Monto_Total = CCur(LabelTotal.Caption)
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
  If Round(SumaDebe - SumaHaber, 2) <> 0 Then
     Mensajes = "Las transacciones no cuadran correctamente" & vbCrLf _
              & "corrija los resultados de las cuentas"
     MsgBox Mensajes
     TextCodigo.SetFocus
  Else
     Mensajes = "Esta seguro de Grabar el Comprobante" & vbCrLf _
              & "de " & Label2.Caption & " No. " & LabelComp.Caption
     Titulo = "Pregunta de grabación"
     If BoxMensaje = vbYes Then
        TextoImprimio = ""
'''     MsgBox "Modificar Comprobante" & vbCrLf _
'''          & "TP: " & Co.TP & vbCrLf _
'''          & "Numero: " & Co.Numero & vbCrLf _
'''          & "Fcha: " & Co.Fecha & vbCrLf _
'''          & "CodigoB: " & Co.CodigoB & vbCrLf _
'''          & "Beneficiario: " & Co.Beneficiario & vbCrLf _
'''          & "Item: " & Co.Item & vbCrLf _
'''          & "New: " & NuevoComp & ", Modify: " & ModificarComp & ", Copy: " & CopiarComp
     
        'If Len(TxtEmail) > 1 And EmailOld <> TxtEmail Then Co.Email = TrimStrg(TxtEmail)
        DGAsientos.Visible = False
        If AdoAsientos.Recordset.RecordCount > 0 Then
          RatonReloj
          If NuevoComp Then
             If OpcTP(0).Value Then NumComp = ReadSetDataNum("Diario", True, True)
             If OpcTP(1).Value Then NumComp = ReadSetDataNum("Ingresos", True, True)
             If OpcTP(2).Value Then NumComp = ReadSetDataNum("Egresos", True, True)
             If OpcTP(3).Value Then NumComp = ReadSetDataNum("NotaDebito", True, True)
             If OpcTP(4).Value Then NumComp = ReadSetDataNum("NotaCredito", True, True)
          End If
          FechaTexto = MBoxFecha
          Co.T = Normal
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Monto_Total = Monto_Total
          Co.Concepto = TextConcepto
         'Co.CodigoB = CodigoBenef
          Co.Efectivo = Abono
          Co.Cotizacion = TextCotiza
          Co.Item = NumEmpresa
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
         'Grabamos el Comprobante
          GrabarComprobante Co
        ' Seteamos para el siguiente comprobante
          DGAsientosB.Visible = False
          RatonNormal
          ImprimirComprobantesDe False, Co
          If CheqCopia.Value Then ImprimirComprobantesDe False, Co
          BorrarAsientos True
          NumComp = NumComp + 1
          Co.Numero = NumComp
          LabelComp.Caption = Format(NumComp, "00000000")
          LabelTotal.Caption = "0.00"
          Label6.Visible = False
          DGAsientos.Visible = True
          If Len(CtaConciliada) Then MsgBox "ADVERTENCIA:" & vbCrLf & vbCrLf & "Revise la Conciliacion Bancaria de la(s) siguiente(s) Cuenta(s):" & vbCrLf & vbCrLf _
                                          & CtaConciliada
          If ModificarComp Then
             ModificarComp = False
             CopiarComp = False
             NuevoComp = True
             Unload FComprobantes
             Exit Sub
          Else
             ModificarComp = False
             CopiarComp = False
             NuevoComp = True
             Tipo_De_Comprobante_No Co
             MBoxFecha.SetFocus
          End If
       Else
          MsgBox "Warning: Falta de Ingresar datos."
          DGAsientos.Visible = True
          TextCodigo.SetFocus
       End If
     Else
       TextCodigo.SetFocus
     End If
  End If
End Sub

Private Sub TextCodigo_GotFocus()
  RatonNormal
  FechaComp = MBoxFecha
  Label12.Visible = False
  Label21.Visible = False
  TxtCheqDep.Visible = False
  MBEfectivizar.Visible = False
  TextCodigo.Text = ""
  TextValor.Visible = False
  Opcion_Mulp = OpcMult.Value
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DetCtaBusqueda As String
  Keys_Especiales Shift
  Select Case KeyCode
    Case vbKeyEscape
         If UCaseStrg(Modulo) = "GASTOS" Then
            CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
            Diferencia = SumaDebe - SumaHaber
            If Diferencia > 0 Then
               OpcDH = 2: ValorDH = Diferencia
               Codigo = Leer_Cta_Catalogo(Cta_CajaG)
               If Moneda_US Then ValorDH = Round(ValorDH * Dolar, 2)
               Fecha_Vence = MBoxFecha
               NoCheque = Ninguno
               InsertarAsientosC AdoAsientos
            End If
         Else
            If OpcTP(2).Value Or OpcTP(3).Value Then
               TextValor.Visible = True
               If CheckEfect.Value Then
                  OpcDH = 2: ValorDH = Abono
                  Cadena = SinEspaciosIzq(DCCaja.Text)
                  Codigo = Leer_Cta_Catalogo(Cadena)
                  If Moneda_US Then ValorDH = Round(ValorDH * Dolar, 2)
                  Fecha_Vence = MBoxFecha.Text
                  NoCheque = Ninguno
                  InsertarAsientosC AdoAsientos
               End If
               SumaBancos = 0
               CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
               If CheckBco.Value Then
                  SumaBancos = SumaDebe - SumaHaber
                  OpcDH = 2: ValorDH = SumaDebe - SumaHaber
                  Cadena = SinEspaciosIzq(DCBanco.Text)
                  Codigo = Leer_Cta_Catalogo(Cadena)
                  SQL2 = "SELECT * " _
                       & "FROM Asiento_B " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND CodigoU = '" & CodigoUsuario & "' " _
                       & "AND T_No = " & Trans_No & " "
                  Select_Adodc_Grid DGAsientosB, AdoAsientosB, SQL2
                  With AdoAsientosB.Recordset
                   If .RecordCount > 0 Then
                      .MoveFirst
                       Fecha_Vence = .fields("EFECTIVIZAR")
                       NoCheque = .fields("CHEQ_DEP")
                      .fields("VALOR") = SumaBancos
                      .Update
                   End If
                  End With
                  If OpcCoop And Moneda_US Then ValorDH = ValorDH * Dolar
                  InsertarAsientosC AdoAsientos
               End If
               Monto_Total = SumaBancos + Abono
               LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
               CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
            End If
         End If
         TextCodigo.Text = "-1"
         CmdGrabar.SetFocus
    Case vbKeyF2
        ' MsgBox TextCodigo
         Select_Cuentas DLCuentas, AdoCuentas, TextCodigo
         TextCodigo.Text = "0"
         SiguienteControl
    Case vbKeyReturn
         SiguienteControl
    Case Else
         Select_Cuentas DLCuentas, AdoCuentas
  End Select
  If CtrlDown And KeyCode = vbKeyB Then
     DetCtaBusqueda = InputBox("", "INGRESE EL PATRON DE BUSQUEDA")
     Select_Cuentas DLCuentas, AdoCuentas, DetCtaBusqueda
     TextCodigo.Text = "0"
     SiguienteControl
  End If
End Sub

Private Sub TextCodigo_LostFocus()
  RatonNormal
  TextoValido TextCodigo
  Select Case UCaseStrg(TextCodigo)
    Case ".": Select_Cuentas DLCuentas, AdoCuentas
              DLCuentas.Visible = True
              DLCuentas.SetFocus
    Case "AC": FComprasAT.Show 1
               Asientos_Grabados
               TextCodigo.SetFocus
    Case "AV": FVentasAT.Show 1
               Asientos_Grabados
               TextCodigo.SetFocus
    Case "AE": FExportacionesAT.Show 1
               Asientos_Grabados
               TextCodigo.SetFocus
    Case "AI": FImportacionesAT.Show 1
               Asientos_Grabados
               TextCodigo.SetFocus
    Case Else: LeerCodigoCta TextCodigo, TextCuenta, TextValor, DLCuentas, FrameAsigna, TextOpcTM, TextOpcDH
               If SubCta = "BA" Then
                  Label12.Visible = True
                  Label21.Visible = True
                  TxtCheqDep.Visible = True
                  MBEfectivizar.Visible = True
                  MBEfectivizar.SetFocus
               End If
  End Select
End Sub

Private Sub Form_Activate()
    CodigoCC = Ninguno
    Trans_No = 1: Ln_No = 1: Ret_No = 1: LnSC_No = 1
    MBoxFecha = FechaSistema
    
    Co.Item = NumEmpresa
    Co.RetNueva = True
    Co.Ctas_Modificar = ""
    Co.CodigoInvModificar = ""
    Co.TipoContribuyente = ""
    Co.RUC_CI = Ninguno
    Co.CodigoB = Ninguno
    Co.Beneficiario = Ninguno
    Co.Email = Ninguno
    Co.TD = Ninguno
    Co.Direccion = Ninguno
    Co.Telefono = Ninguno
    Co.Grupo = Ninguno
    Co.AgenteRetencion = Ninguno
    Co.MicroEmpresa = Ninguno
    Co.Estado = Ninguno
    Co.Cotizacion = 0
    Co.Concepto = ""
    Co.Efectivo = 0
    Co.Total_Banco = 0
    Co.Monto_Total = 0
          
  If Len(Co.TP) < 2 Then Co.TP = CompDiario
  If Len(Co.Fecha) = 10 And IsDate(Co.Fecha) Then MBoxFecha = Co.Fecha Else MBoxFecha = FechaSistema
 'Leer los datos del comprobante a modificar o copiar

  If ModificarComp Then
     Control_Procesos Normal, "Modificar Comprobante de: " & Co.TP & " No. " & Co.Numero
     Listar_Comprobante_SP Co
  End If
  
  If CopiarComp Then
     Control_Procesos Normal, "Copiando Comprobante de: " & Co.TP & " No. " & Co.Numero
     Listar_Comprobante_SP Co
     NuevoComp = True
  End If

  If NuevoComp Then
     Co.Numero = 0
     If ExistenMovimientos And Not CopiarComp Then
        Mensajes = "El Sistema se cerro de forma inesperada, existen movimientos " _
                 & "en transito con su codigo de usuario." & vbCrLf & vbCrLf _
                 & "Desea recuperarlos?"
        Titulo = "PREGUNTA DE CONFIRMACION"
        If BoxMensaje <> vbYes Then BorrarAsientos True
     End If
  End If
  
  sSQL = "SELECT TOP 50 Cliente,CI_RUC,Codigo " _
       & "FROM Clientes " _
       & "WHERE Cliente LIKE '%" & Co.Beneficiario & "%' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
  
  TipoBusqueda = "%"
  
  Tipo_De_Comprobante_No Co

'  Datos_del_Cliente Co    'Solo con el CodigoB = Codigo del Cliente se busca
  'Select_Cuentas DLCuentas, AdoCuentas
  Llenar_Encabezado_Comprobante
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
  Una_Vez = True
 'Listamos lista de clientes para procesar comprobantes
  
  If UCaseStrg(Modulo) = "GASTOS" Then
     OpcTP(0).Visible = True
     OpcTP(1).Visible = False
     OpcTP(2).Visible = False
     OpcTP(3).Visible = False
     OpcTP(4).Visible = False
  End If
  If Bloquear_Control Then CmdGrabar.Enabled = False
  RatonNormal
  FComprobantes.WindowState = vbMaximized

  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
 'Abriendo Bases Relacionadas
  ConectarAdodc AdoAC
  ConectarAdodc AdoAV
  ConectarAdodc AdoAE
  ConectarAdodc AdoAI
  ConectarAdodc AdoSQL
  ConectarAdodc AdoCaja
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoBenef
  ConectarAdodc AdoRolPago
  ConectarAdodc AdoCuentas
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoAsientosB
  ConectarAdodc AdoAsientosR
  ConectarAdodc AdoAsientosSC
  ConectarAdodc AdoCentroCostos
  
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 300
  SSTab1.width = MDI_X_Max - SSTab1.Left
  DGAsientosSC.width = SSTab1.width - 200
  DGAsientosSC.Height = SSTab1.Height - DGAsientosSC.Top - 100
  DLCuentas.Height = SSTab1.Height - DGAsientosSC.Top
  DGAC.width = SSTab1.width - 200
  DGAV.width = SSTab1.width - 200
  DGAI.width = SSTab1.width - 200
  DGAE.width = SSTab1.width - 200
  
  DGAsientosR.width = SSTab1.width - 200
  DGAsientosR.Height = SSTab1.Height - DGAsientosR.Top - 100
  
  DGAsientos.width = SSTab1.width - 200
  DGAsientos.Height = SSTab1.Height - DGAsientos.Top - 100
  
  CmdGrabar.Top = SSTab1.Top + SSTab1.Height + 50
  CmdCancelar.Top = SSTab1.Top + SSTab1.Height + 50
  Label8.Top = SSTab1.Top + SSTab1.Height + 50
  Label19.Top = SSTab1.Top + SSTab1.Height + 50
  LabelDiferencia.Top = SSTab1.Top + SSTab1.Height + 50
  LabelDebe.Top = SSTab1.Top + SSTab1.Height + 50
  LabelHaber.Top = SSTab1.Top + SSTab1.Height + 50
End Sub

Private Sub TextCantidad_LostFocus()
  Fecha_Vence = MBoxFecha.Text
  NoCheque = Ninguno
  TextoValido TextCantidad, True
  Abono = Round(CDbl(TextCantidad.Text), 2)
  TextCantidad.Text = Format(Abono, "#,##0.00")
  Monto_Total = Abono
  LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
  If CheckEfect.Value And OpcTP(1).Value Then
     OpcDH = 1: ValorDH = Abono
     Cadena = SinEspaciosIzq(DCCaja.Text)
     Codigo = Leer_Cta_Catalogo(Cadena)
     If Moneda_US Then ValorDH = Round(ValorDH * Dolar, 2)
     InsertarAsientosC AdoAsientos
  End If
  CheckBco.SetFocus
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
End Sub

Private Sub CheckBco_Click()
  If CheckBco.Value Then
     sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE TC = 'BA' " _
          & "AND DG = 'D' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo "
     SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
     If OpcTP(2).Value Then
        Label24.Caption = "Cheq. No."
        Label5.Visible = False
        TextDeposito.Visible = False
     Else
        Label24.Caption = UCaseStrg(Moneda)
        Label5.Visible = True
        TextDeposito.Visible = True
     End If
     
     MBFechaEfec.Visible = True
     Label10.Visible = True
     DCBanco.Visible = True
     DGAsientosB.Visible = True
     TextCheque.Visible = True
     Label24.Visible = True
     DCBanco.SetFocus
  Else
     MBFechaEfec.Visible = False
     Label10.Visible = False
     DCBanco.Visible = False
     DGAsientosB.Visible = False
     Label24.Visible = False
     Label5.Visible = False
     TextDeposito.Visible = False
     TextCheque.Visible = False
  End If
End Sub

Private Sub CheckEfect_Click()
  Label6.Caption = Moneda
  If CheckEfect.Value Then
     sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCaja " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE TC = 'CJ' " _
          & "AND DG = 'D' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo "
     SelectDB_Combo DCCaja, AdoCaja, sSQL, "NomCaja"
     Label6.Visible = True
     TextCantidad.Visible = True
     DCCaja.Visible = True
     DCCaja.SetFocus
  Else
     TextCantidad.Text = "0.00"
     Abono = 0
     TextCantidad.Visible = False
     DCCaja.Visible = False
     Label6.Visible = False
  End If
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  TextoValido TxtEmail
  TxtEmail = TrimStrg(LCase(TxtEmail))
  Co.Email = TxtEmail
End Sub

'''Public Sub Lista_de_Clientes()
''''''  IniTiempo = Time
'''  sSQL = "SELECT Cliente,CI_RUC,Codigo " _
'''       & "FROM Clientes " _
'''       & "WHERE Cliente = '.' " _
'''       & "ORDER BY Cliente "
'''  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
''''''  sSQL = "SELECT * " _
''''''       & "FROM fn_Lista_de_Clientes() " _
''''''       & "ORDER BY Cliente "
''''''  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
'''End Sub

Public Sub Llenar_Encabezado_Comprobante()
  DCCliente = Co.Beneficiario
  LblRUC.Caption = Co.RUC_CI
  TxtEmail = Co.Email
  TextConcepto = Co.Concepto
  TextCotiza = Co.Cotizacion
  LabelTotal.Caption = Monto_Total
  Monto_Total = Co.Monto_Total
  Abono = Co.Efectivo
  TextCantidad = Abono
'  EmailOld = TxtEmail
  TipoDoc = Co.TD
  TipoBenef = Co.TD
  CICliente = Co.RUC_CI
  CodigoBenef = Co.CodigoB
  CodigoCliente = Co.CodigoB
  NombreCliente = Co.Beneficiario
End Sub
