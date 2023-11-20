VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{64AED23E-31A2-4023-8C7D-E628B15843D8}#1.0#0"; "Code39X.ocx"
Begin VB.Form FComprobantes 
   Caption         =   "COMPROBANTE DE EGRESO"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Comproba.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
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
      Picture         =   "Comproba.frx":0342
      TabIndex        =   58
      Top             =   8085
      Width           =   1380
   End
   Begin VB.Frame FrmBenef 
      BackColor       =   &H00800000&
      Caption         =   "BUSCAR BENEFICIARIO"
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
      Height          =   3060
      Left            =   1995
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   6315
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "Comproba.frx":0444
         DataSource      =   "AdoBenef"
         Height          =   2715
         Left            =   105
         TabIndex        =   12
         Top             =   210
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   4789
         _Version        =   393216
         Style           =   1
         Text            =   "Beneficiario"
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
      Left            =   1575
      MaxLength       =   60
      TabIndex        =   18
      ToolTipText     =   "Escriba el Nombre o C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   1470
      Width           =   6210
   End
   Begin Code39X.Code39Clt Code39Clt1 
      Left            =   105
      Top             =   8400
      _ExtentX        =   1905
      _ExtentY        =   1085
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
      Left            =   2940
      TabIndex        =   73
      Top             =   525
      Width           =   2430
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
      Bindings        =   "Comproba.frx":045B
      DataSource      =   "AdoCuentas"
      Height          =   2220
      Left            =   210
      TabIndex        =   44
      Top             =   5670
      Visible         =   0   'False
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   3916
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
      Left            =   9135
      TabIndex        =   47
      Top             =   5145
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
         TabIndex        =   51
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
         TabIndex        =   54
         Text            =   "Comproba.frx":0474
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
         TabIndex        =   57
         Text            =   "Comproba.frx":0476
         Top             =   1365
         Width           =   435
      End
      Begin MSMask.MaskEdBox MBEfectivizar 
         Height          =   330
         Left            =   945
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   48
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
         TabIndex        =   56
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
         TabIndex        =   53
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
         TabIndex        =   55
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
         TabIndex        =   52
         Top             =   945
         Width           =   1485
      End
   End
   Begin VB.TextBox TextBenef 
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
      MaxLength       =   50
      TabIndex        =   10
      ToolTipText     =   "Escriba el Nombre o C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   840
      Width           =   6210
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
      Height          =   645
      Left            =   7875
      TabIndex        =   19
      Top             =   840
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
         TabIndex        =   20
         Top             =   210
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
         TabIndex        =   21
         Top             =   210
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
      Left            =   6510
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1155
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1575
      TabIndex        =   8
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2745
      Left            =   105
      TabIndex        =   63
      Top             =   5250
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4842
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&4.- CONTABILIZACION"
      TabPicture(0)   =   "Comproba.frx":0478
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGAsientos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&5.- SUBCUENTAS"
      TabPicture(1)   =   "Comproba.frx":0494
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGAsientosSC"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&6.- RETENCIONES"
      TabPicture(2)   =   "Comproba.frx":04B0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGAC"
      Tab(2).Control(1)=   "DGAsientosR"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&7.- AC - AV - AI - AE"
      TabPicture(3)   =   "Comproba.frx":04CC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGAE"
      Tab(3).Control(1)=   "DGAI"
      Tab(3).Control(2)=   "DGAV"
      Tab(3).ControlCount=   3
      Begin MSDataGridLib.DataGrid DGAC 
         Bindings        =   "Comproba.frx":04E8
         Height          =   750
         Left            =   -74895
         TabIndex        =   71
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
         Bindings        =   "Comproba.frx":04FC
         Height          =   750
         Left            =   -74895
         TabIndex        =   70
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
         Bindings        =   "Comproba.frx":0510
         Height          =   750
         Left            =   -74895
         TabIndex        =   69
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
         Bindings        =   "Comproba.frx":0524
         Height          =   750
         Left            =   -74895
         TabIndex        =   72
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
         Bindings        =   "Comproba.frx":0538
         Height          =   1380
         Left            =   -74895
         TabIndex        =   68
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
         Bindings        =   "Comproba.frx":0553
         Height          =   2220
         Left            =   -74895
         TabIndex        =   67
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
         Bindings        =   "Comproba.frx":056F
         Height          =   2220
         Left            =   105
         TabIndex        =   66
         Top             =   420
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
      Top             =   5355
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
      TabIndex        =   39
      Top             =   3990
      Width           =   9675
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
      Picture         =   "Comproba.frx":0589
      TabIndex        =   59
      Top             =   8085
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
      Left            =   9135
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   4830
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
      Left            =   1785
      TabIndex        =   43
      Top             =   4830
      Width           =   7365
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
      TabIndex        =   41
      Text            =   "0"
      ToolTipText     =   "<Ctrl+B> Por Patrón de Búsqueda"
      Top             =   4830
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
      TabIndex        =   24
      Top             =   1890
      Visible         =   0   'False
      Width           =   11355
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
         Left            =   9765
         MaxLength       =   16
         TabIndex        =   36
         Text            =   "0"
         Top             =   1260
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid DGAsientosB 
         Bindings        =   "Comproba.frx":09CB
         Height          =   960
         Left            =   105
         TabIndex        =   37
         Top             =   945
         Width           =   8310
         _ExtentX        =   14658
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
         Bindings        =   "Comproba.frx":09E6
         DataSource      =   "AdoBanco"
         Height          =   315
         Left            =   1260
         TabIndex        =   30
         Top             =   525
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Banco"
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
      Begin MSDataListLib.DataCombo DCCaja 
         Bindings        =   "Comproba.frx":09FD
         DataSource      =   "AdoCaja"
         Height          =   315
         Left            =   1260
         TabIndex        =   26
         Top             =   210
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Caja"
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
         Left            =   9555
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "Comproba.frx":0A13
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
         Left            =   9555
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "Comproba.frx":0A15
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
         TabIndex        =   25
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
         TabIndex        =   29
         Top             =   525
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBFechaEfec 
         Height          =   330
         Left            =   8505
         TabIndex        =   34
         Top             =   1260
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
         Left            =   8505
         TabIndex        =   33
         Top             =   945
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
         Left            =   9765
         TabIndex        =   35
         Top             =   945
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
         Left            =   8505
         TabIndex        =   31
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
         Left            =   8505
         TabIndex        =   27
         Top             =   210
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc AdoCaja 
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
      Top             =   5355
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
      Top             =   5355
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
      TabIndex        =   17
      Top             =   1470
      Width           =   1485
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
      Left            =   9870
      TabIndex        =   23
      Top             =   1155
      Width           =   1590
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
      Left            =   1575
      TabIndex        =   14
      Top             =   1155
      Width           =   3585
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
      Left            =   5145
      TabIndex        =   15
      Top             =   1155
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
      Left            =   105
      TabIndex        =   13
      Top             =   1155
      Width           =   1485
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
      Left            =   105
      TabIndex        =   9
      Top             =   840
      Width           =   1485
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
      Left            =   10080
      TabIndex        =   6
      Top             =   105
      Width           =   1380
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
      Width           =   2640
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
      Left            =   9345
      TabIndex        =   60
      Top             =   8085
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
      Left            =   7560
      TabIndex        =   61
      Top             =   8085
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
      Left            =   4410
      TabIndex        =   64
      Top             =   8085
      Width           =   1695
   End
   Begin VB.Label Label17 
      BackColor       =   &H00808080&
      Caption         =   " VALOR"
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
      Left            =   9135
      TabIndex        =   45
      Top             =   4620
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
      Left            =   1785
      TabIndex        =   42
      Top             =   4620
      Width           =   7365
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
      Left            =   9870
      TabIndex        =   22
      Top             =   840
      Width           =   1590
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
      Width           =   1485
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
      Left            =   3255
      TabIndex        =   65
      Top             =   8085
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
      TabIndex        =   40
      Top             =   4620
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
      Left            =   6510
      TabIndex        =   62
      Top             =   8085
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
      TabIndex        =   38
      Top             =   3990
      Width           =   1590
   End
End
Attribute VB_Name = "FComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Asientos_Grabados()
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY A_No "
  SelectDataGrid DGAsientos, AdoAsientos, sSQL
End Sub

Public Sub CalculosTotalAsientosX()
  SumaDebe = 0: SumaHaber = 0
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelDiferencia.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
End Sub

Public Sub Tipo_De_Comprobante_No(XIndex As Integer, Optional XTipoComp As String)
  IniciarAsientos DGAsientos, DGAsientosSC, DGAsientosB, DGAsientosR, AdoAsientos, AdoAsientosB, AdoAsientosSC, AdoAsientosR
  Select_Cuentas DLCuentas, AdoCuentas
  FechaComp = MBoxFecha
  TextCotiza = Dolar
  FrmBenef.Visible = False
  CheckEfect.Visible = False
  Label6.Caption = Moneda
 'Verificamos si esta activado Caja
  If CheckEfect.value Then
     TextCantidad.Visible = True
     DCCaja.Visible = True
     Label6.Visible = True
  Else
     TextCantidad.Text = "0.00"
     Abono = 0
     TextCantidad.Visible = False
     DCCaja.Visible = False
     Label6.Visible = False
  End If
 'Verificamos si esta activado Bancos
  If CheckBco.value Then
     If OpcTP(2).value Then
        Label24.Caption = "Cheq. No."
        Label5.Visible = False
        TextDeposito.Visible = False
     Else
        Label24.Caption = UCase$(Moneda)
        Label5.Visible = True
        TextDeposito.Visible = True
     End If
     MBFechaEfec.Visible = True
     Label10.Visible = True
     DCBanco.Visible = True
     DGAsientosB.Visible = True
     TextCheque.Visible = True
     Label24.Visible = True
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
  Frame2.Visible = False
 'Solo si estamos modificando un comprobante
  Select Case XTipoComp
    Case CompDiario:      XIndex = 0
    Case CompIngreso:     XIndex = 1
    Case CompEgreso:      XIndex = 2
    Case CompNotaDebito:  XIndex = 3
    Case CompNotaCredito: XIndex = 4
  End Select
 'Determinamos que tipo de comprobante realizamos
  Select Case XIndex
    Case 0
         NumComp = ReadSetDataNum("Diario", True, False)
         Label2.Caption = "Diario No. "
         Label3.Caption = " &BENEFICIARIO:"
         Frame2.Visible = False
    Case 1
         NumComp = ReadSetDataNum("Ingresos", True, False)
         Label2.Caption = "Ingreso No. "
         Label3.Caption = " RECI&BI DE:"
         CheckEfect.Visible = True
         Frame2.Visible = True
    Case 2
         NumComp = ReadSetDataNum("Egresos", True, False)
         Label2.Caption = "Egreso No. "
         Label3.Caption = " &PAGADO A:"
         Frame2.Visible = True
         CheckEfect.Visible = True
         Frame2.Visible = True
    Case 3
         NumComp = ReadSetDataNum("NotaDebito", True, False)
         Label2.Caption = "Nota de Débito No. "
         Label3.Caption = " &BENEFICIARIO:"
         'Frame2.Visible = True
    Case 4
         NumComp = ReadSetDataNum("NotaCredito", True, False)
         Label2.Caption = "Nota de Débito No. "
         Label3.Caption = " &BENEFICIARIO:"
         'Frame2.Visible = True
  End Select
  If NumeroComp > 0 Then NumComp = NumeroComp
 'Presentamos la informacion del Anexo en este comprobante
  SQL2 = "SELECT * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGAsientosR, AdoAsientosR, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGAC, AdoAC, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGAV, AdoAV, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGAE, AdoAE, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGAI, AdoAI, SQL2
  FComprobantes.Caption = "COMPROBANTE DE " & UCase$(Mid$(OpcTP(XIndex).Caption, 6, Len(OpcTP(XIndex).Caption)))
  Label2.Caption = Mid$(OpcTP(XIndex).Caption, 6, Len(OpcTP(XIndex).Caption)) & " No. " _
                 & Format(Year(MBoxFecha), "0000") & "- "
  If XTipoComp <> "" Then NumComp = NumeroComp
  LabelComp.Caption = Format(NumComp, "00000000")
  MBoxFecha.SetFocus
End Sub

Private Sub CheckBco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CheckEfect_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCliente.Text & "'")
       If Not .EOF Then
          TipoContribuyente = ""
          LblRUC.Caption = .Fields("CI_RUC")
          CICliente = .Fields("CI_RUC")
          CodigoBenef = .Fields("Codigo")
          TextBenef.Text = .Fields("Cliente")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          TxtEmail = .Fields("Email")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          TipoBenef = .Fields("TD")
          If .Fields("RISE") Then TipoContribuyente = TipoContribuyente & " RISE"
          If .Fields("Especial") Then TipoContribuyente = TipoContribuyente & " Contribuyente especial"
          FrmBenef.Visible = False
          TextCotiza.SetFocus
       Else
          FrmBenef.Visible = False
          FComprobantes.Visible = False
          NombreCliente = DCCliente.Text
          Nuevo = True
          'Unload FComprobantes
          FClientesFlash.Show 1
          FComprobantes.Visible = True
          sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,RISE,Especial,Telefono,Direccion,Email " _
               & "FROM Clientes " _
               & "WHERE T <> '.' " _
               & "ORDER BY Cliente "
          SelectDBCombo DCCliente, AdoBenef, sSQL, "Cliente"
       End If
   Else
       FrmBenef.Visible = False
       FComprobantes.Visible = False
       NombreCliente = DCCliente.Text
       Nuevo = True
       'Unload FComprobantes
       FClientesFlash.Show 1
       FComprobantes.Visible = True
       sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,RISE,Especial,Telefono,Direccion,Email " _
            & "FROM Clientes " _
            & "WHERE T <> '.' " _
            & "ORDER BY Cliente "
       SelectDBCombo DCCliente, AdoBenef, sSQL, "Cliente"
   End If
  End With
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
  Codigo = AdoAsientosR.Recordset.Fields("Cta")
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
     LeerCta Cadena
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
         If OpcTP(1).value Or OpcTP(4).value Then
            CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
            If CheckBco.value Then
               With AdoAsientosB.Recordset
                If .RecordCount > 0 Then
                   .MoveFirst
                    Do While Not .EOF
                       Cadena = .Fields("CTA_BANCO")
                       Valor = .Fields("VALOR")
                       LeerCta Cadena
                       OpcDH = 1: ValorDH = Valor
                       FechaValida MBFechaEfec
                       TextoValido TextDeposito, , True
                       Fecha_Vence = .Fields("EFECTIVIZAR")
                       NoCheque = .Fields("CHEQ_DEP")
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
         If OpcTP(4).value Then TextConcepto.SetFocus
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub DCBanco_LostFocus()
Dim AdoCheq As ADODB.Recordset
  LeerBanco AdoCtas, DCBanco.Text
  DCBanco.Text = Codigo & Space(5) & NombreBanco
  If OpcTP(2).value Then
     Label24.Caption = "Cheq. No."
     Set AdoCheq = New ADODB.Recordset
     AdoCheq.CursorType = adOpenStatic
     AdoCheq.CursorLocation = adUseClient
     Cta_Aux = SinEspaciosIzq(DCBanco)
     If Cta_Aux = "" Then Cta_Aux = "."
     TextCheque = "00000001"
     sSQL = "SELECT * " _
          & "FROM Transacciones " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cta = '" & Cta_Aux & "' " _
          & "AND TP = 'CE' " _
          & "AND ISNUMERIC(Cheq_Dep) <> 0 " _
          & "AND Haber > 0 " _
          & "AND Fecha <= #" & BuscarFecha(MBoxFecha) & "# " _
          & "ORDER BY Cheq_Dep DESC "   'Fecha DESC
     sSQL = CompilarSQL(sSQL)
     AdoCheq.Open sSQL, AdoStrCnn, , , adCmdText
    'Seteamos los encabezados para las facturas
     If AdoCheq.RecordCount > 0 Then
        TextCheque = Format(Val(AdoCheq.Fields("Cheq_Dep")) + 1, "00000000")
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
  If OpcTP(2).value Then
     NoCheque = TextCheque.Text
     Fecha_Vence = MBFechaEfec.Text
     SQL1 = "DELETE * " _
          & "FROM Asiento_B " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute SQL1
     InsertarAsientoBanco AdoAsientosB, 0
     TextConcepto.SetFocus
  Else
     TextCheque.Text = Format(TextCheque.Text, "#,##0.00")
  End If
End Sub

Private Sub OpcTP_Click(Index As Integer)
  Tipo_De_Comprobante_No Index
End Sub

Private Sub OpcTP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then MBoxFecha.SetFocus
End Sub

Private Sub TextBenef_GotFocus()
  MarcarTexto TextBenef
End Sub

Private Sub TextBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyB Then TextBenef.Text = "Ninguno"
End Sub

Private Sub TextBenef_LostFocus()
  TextoValido TextBenef, , True
  
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextBenef.Text & "'")
       If Not .EOF Then
          LblRUC.Caption = .Fields("CI_RUC")
          CICliente = .Fields("CI_RUC")
          CodigoBenef = .Fields("Codigo")
          TextBenef.Text = .Fields("Cliente")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          TxtEmail = .Fields("Email")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          TipoBenef = .Fields("TD")
         'MsgBox TextBenef.Text
          Co.CodigoB = CodigoBenef
          TextCotiza.SetFocus
       Else
         .MoveFirst
         .Find ("CI_RUC = '" & TextBenef.Text & "'")
          If Not .EOF Then
             LblRUC.Caption = .Fields("CI_RUC")
             CICliente = .Fields("CI_RUC")
             CodigoBenef = .Fields("Codigo")
             TextBenef.Text = .Fields("Cliente")
             CodigoCliente = .Fields("Codigo")
             NombreCliente = .Fields("Cliente")
             TxtEmail = .Fields("Email")
             Grupo_No = .Fields("Grupo")
             TipoDoc = .Fields("TD")
             TipoBenef = .Fields("TD")
             Co.CodigoB = CodigoBenef
             TextCotiza.SetFocus
          Else
             DCCliente.Text = TextBenef.Text
             FrmBenef.Visible = True
             DCCliente.SetFocus
          End If
       End If
   Else
       DCCliente.Text = TextBenef.Text
       FrmBenef.Visible = True
       DCCliente.SetFocus
   End If
  End With
End Sub

Private Sub TextCotiza_GotFocus()
  MarcarTexto TextCotiza
End Sub

Private Sub TextOpcDH_Change()
  If 1 > Val(TextOpcDH.Text) Or Val(TextOpcDH.Text) > 2 Then
     TextOpcDH.Text = ""
  Else
     OpcDH = Val(TextOpcDH.Text)
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
  TextOpcDH = UCase$(TextOpcDH)
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
            If OpcDH > 1 Then
               FSubCtas.Show 1
            Else
               Aprobacion.Show 1
            End If
     End Select
     SQL2 = "SELECT * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     SelectDataGrid DGAsientosSC, AdoAsientosSC, SQL2
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
  FComprobantes.WindowState = 1
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
  MBFechaEfec.Text = MBoxFecha.Text
  MBEfectivizar.Text = MBoxFecha.Text
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
  If OpcTP(1).value Then
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
  AdoCheq.Open sSQL, AdoStrCnn, , , adCmdText
 'Seteamos los encabezados para las facturas
  If AdoCheq.RecordCount > 0 Then
     MsgBox "Warning: El Cheque No. " & TextCheque & ", ya esta ingresado."
  End If
  AdoCheq.Close
  If OpcTP(2).value And IsNumeric(TextCheque) Then TextCheque = Format(Val(TextCheque), "00000000")
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
     DetalleComp = Trim$(Mid$(InputBox("Detalle del Concepto:", "CONCEPTO AUXILIAR", ""), 1, 45))
  Else
     DetalleComp = Ninguno
  End If
  If DetalleComp = "" Then DetalleComp = Ninguno
  Select Case SubCta
    Case "C", "P", "G", "I", "CP", "PM"
         LeerCta TextCodigo.Text
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
    Case "N", "CJ", "BA", "CS", "PS", "TJ", "RF", "RI", "CI", "CF"
         TextoValido TextValor, True
         TotalSubCta = 0: SubCtaGen = Ninguno
         ValorDH = Val(CCur(TextValor))
         InsertarAsiento AdoAsientos
  End Select
  TextCodigo.SetFocus
End Sub

Private Sub DGAsientos_BeforeDelete(Cancel As Integer)
  Codigo = AdoAsientos.Recordset.Fields("CODIGO")
  Cancel = DeleteSiNo(AdoAsientos)
  If Cancel = False Then
     EliminarSubCta Codigo
     SQL2 = "SELECT * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     SelectDataGrid DGAsientosSC, AdoAsientosSC, SQL2
     SQL2 = "SELECT * " _
          & "FROM Asiento_Air " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     SelectDataGrid DGAsientosR, AdoAsientosR, SQL2
  End If
  RatonNormal
End Sub

Private Sub DGAsientos_GotFocus()
  CalculosTotalAsientosX
End Sub

Private Sub DGAsientos_LostFocus()
  CalculosTotalAsientosX
End Sub

Private Sub CmdCancelar_Click()
  NumEmpresa = NumItemTemp
  Unload FComprobantes
End Sub

Private Sub CmdGrabar_Click()
  Asientos_Grabados
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo = '" & CodigoBenef & "' ")
       If Not .EOF Then
         .Fields("Email") = TxtEmail
         .Update
       End If
   End If
  End With
  
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectAdodc AdoAsientosSC, SQL2
  OpcSubCtaDH = 1
  Monto_Total = CCur(LabelTotal.Caption)
  CalculosTotalAsientosX
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
        DGAsientos.Visible = False
        For I = 0 To 4
            OpcTP(I).Enabled = True
        Next I
        If AdoAsientos.Recordset.RecordCount > 0 Then
          RatonReloj
          If NumeroComp <> 0 Then
             NumComp = NumeroComp
          Else
             If OpcTP(0).value Then NumComp = ReadSetDataNum("Diario", True, True)
             If OpcTP(1).value Then NumComp = ReadSetDataNum("Ingresos", True, True)
             If OpcTP(2).value Then NumComp = ReadSetDataNum("Egresos", True, True)
             If OpcTP(3).value Then NumComp = ReadSetDataNum("NotaDebito", True, True)
             If OpcTP(4).value Then NumComp = ReadSetDataNum("NotaCredito", True, True)
          End If
          FechaTexto = MBoxFecha
         'Grabacion del Comp
          If OpcTP(0).value Then Co.TP = CompDiario
          If OpcTP(1).value Then Co.TP = CompIngreso
          If OpcTP(2).value Then Co.TP = CompEgreso
          If OpcTP(3).value Then Co.TP = CompNotaDebito
          If OpcTP(4).value Then Co.TP = CompNotaCredito
          Co.T = Normal
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Monto_Total = Monto_Total
          Co.Concepto = TextConcepto.Text
          Co.CodigoB = CodigoBenef
          Co.Efectivo = Abono
          Co.Cotizacion = TextCotiza.Text
          Co.Item = NumEmpresa
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
        
          SQL2 = "SELECT * " _
               & "FROM Clientes " _
               & "WHERE Codigo = '" & Co.CodigoB & "' "
          SelectAdodc AdoSQL, SQL2
          With AdoSQL.Recordset
           If .RecordCount > 0 Then
               Co.Beneficiario = .Fields("Cliente")
               Co.Direccion = .Fields("Direccion")
               Co.RUC_CI = .Fields("CI_RUC")
               Co.TD = .Fields("TD")
               Co.Telefono = .Fields("Telefono")
               Co.Email = .Fields("Email")
           End If
          End With
          GrabarComprobante Co
          Control_Procesos Normal, "Grabar Comprobante de: " & Co.TP & "No. " & NumComp
        ' Seteamos para el siguiente comprobante
          DGAsientosB.Visible = False
          RatonNormal
          ImprimirComprobantesDe False, Co
          If CheqCopia.value Then ImprimirComprobantesDe False, Co
          NuevoComp = True
          NumEmpresa = NumItemTemp
          BorrarAsientos True
          IniciarAsientos DGAsientos, DGAsientosSC, DGAsientosB, DGAsientosR, AdoAsientos, AdoAsientosB, AdoAsientosSC, AdoAsientosR
          NumComp = NumComp + 1
          LabelComp.Caption = Format(NumComp, "00000000")
          LabelTotal.Caption = "0.00"
          Label6.Visible = False
          CheckEfect.value = False
          CheckBco.value = False
          DGAsientosB.Visible = False
          TextCantidad.Text = ""
          TextCheque.Text = ""
          TextDeposito.Text = ""
          DGAsientos.Visible = True
          If NumeroComp <> 0 Then
             Unload FComprobantes
             Exit Sub
          Else
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
  Opcion_Mulp = OpcMult.value
  CalculosTotalAsientos AdoAsientos, LabelDebe, LabelHaber, LabelDiferencia
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DetCtaBusqueda As String
  Keys_Especiales Shift
  Select Case KeyCode
    Case vbKeyEscape
         If OpcTP(2).value Or OpcTP(3).value Then
            TextValor.Visible = True
            If CheckEfect.value Then
               OpcDH = 2: ValorDH = Abono
               Cadena = SinEspaciosIzq(DCCaja.Text)
               LeerCta Cadena
               If Moneda_US Then ValorDH = Round(ValorDH * Dolar, 2)
               Fecha_Vence = MBoxFecha.Text
               NoCheque = Ninguno
               InsertarAsientosC AdoAsientos
            End If
            SumaBancos = 0
            CalculosTotalAsientosX
            If CheckBco.value Then
               SumaBancos = SumaDebe - SumaHaber
               OpcDH = 2: ValorDH = SumaDebe - SumaHaber
               Cadena = SinEspaciosIzq(DCBanco.Text)
               LeerCta Cadena
               SQL2 = "SELECT * " _
                    & "FROM Asiento_B " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND CodigoU = '" & CodigoUsuario & "' " _
                    & "AND T_No = " & Trans_No & " "
               SelectDataGrid DGAsientosB, AdoAsientosB, SQL2
               With AdoAsientosB.Recordset
                If .RecordCount > 0 Then
                   .MoveFirst
                    Fecha_Vence = .Fields("EFECTIVIZAR")
                    NoCheque = .Fields("CHEQ_DEP")
                   .Fields("VALOR") = SumaBancos
                   .Update
                End If
               End With
               If OpcCoop And Moneda_US Then ValorDH = ValorDH * Dolar
               InsertarAsientosC AdoAsientos
            End If
            Monto_Total = SumaBancos + Abono
            LabelTotal.Caption = Format(Monto_Total, "#,##0.00")
            CalculosTotalAsientosX
         End If
         TextCodigo.Text = "-1"
         CmdGrabar.SetFocus
    Case vbKeyF2
         Select_Cuentas DLCuentas, AdoCuentas, TextCodigo
         TextCodigo.Text = "0"
         SiguienteControl
    Case vbKeyReturn
         SiguienteControl
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
  Select Case UCase$(TextCodigo)
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
  MBoxFecha = FechaSistema
  If FechaComp <> Ninguno Then MBoxFecha = FechaComp
  Trans_No = 1: Ln_No = 1: Ret_No = 1: LnSC_No = 1
  If CodigoCliente = "" Then
     CodigoCliente = Ninguno
     NombreCliente = Ninguno
  End If
  IniciarAsientos DGAsientos, DGAsientosSC, DGAsientosB, DGAsientosR, AdoAsientos, AdoAsientosB, AdoAsientosSC, AdoAsientosR
  If NumeroComp <> 0 And FechaComp <> Ninguno Then
     NumEmpresa = NumItem
     NumComp = NumeroComp
     Tipo_De_Comprobante_No 0, TipoComp
     For I = 0 To 4
         OpcTP(I).Enabled = False
     Next I
    'Ver cual esta activado
     Select Case TipoComp
       Case CompDiario: OpcTP(0).value = True
       Case CompIngreso: OpcTP(1).value = True
       Case CompEgreso: OpcTP(2).value = True
       Case CompNotaDebito: OpcTP(3).value = True
       Case CompNotaCredito: OpcTP(4).value = True
     End Select
     ListarComprobante TipoComp, NumItem, NumeroComp, AdoAsientos, AdoAsientosB, AdoAsientosSC, AdoAsientosR
     TextBenef = NombreCliente
     TextConcepto = Co.Concepto
     LblRUC.Caption = Co.RUC_CI
     TextCotiza = Co.Cotizacion
     Monto_Total = Co.Monto_Total
     Abono = Co.Efectivo
     LabelTotal.Caption = Monto_Total
     TextCantidad = Abono
     NumComp = NumeroComp
     
     If CopiarComp Then
        If OpcTP(0).value Then NumComp = ReadSetDataNum("Diario", True, False)
        If OpcTP(1).value Then NumComp = ReadSetDataNum("Ingresos", True, False)
        If OpcTP(2).value Then NumComp = ReadSetDataNum("Egresos", True, False)
        If OpcTP(3).value Then NumComp = ReadSetDataNum("NotaDebito", True, False)
        If OpcTP(4).value Then NumComp = ReadSetDataNum("NotaCredito", True, False)
        NumeroComp = 0
        NumEmpresa = NumItemTemp
        FechaComp = Ninguno
     End If
     CalculosTotalAsientosX
  Else
     NumEmpresa = NumItemTemp
     Tipo_De_Comprobante_No 0   'Ininciamos con Diarios
  End If
  TipoDoc = TipoComp
  Una_Vez = True
'''  LabelComp.Caption = Format(NumComp, "00000000")
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCaja " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'CJ' AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCaja, AdoCaja, sSQL, "NomCaja"
  
  sSQL = "SELECT Codigo,IEESS_Per,IEESS_Pat,Salario,Item " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoRolPago, sSQL
  
  sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,RISE,Especial,Telefono,Direccion,Email " _
       & "FROM Clientes " _
       & "WHERE T <> '.' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCCliente, AdoBenef, sSQL, "Cliente"
  CodigoBenef = CodigoCliente
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & NombreCliente & "' ")
       If Not .EOF Then
          CICliente = .Fields("CI_RUC")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          TipoBenef = .Fields("TD")
       End If
   End If
  End With
  FComprobantes.WindowState = 2
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
 'Abriendo Bases Relacionadas
  FComprobantes.WindowState = 1
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
  If CheckEfect.value And OpcTP(1).value Then
     OpcDH = 1: ValorDH = Abono
     Cadena = SinEspaciosIzq(DCCaja.Text)
     LeerCta Cadena
     If Moneda_US Then ValorDH = Round(ValorDH * Dolar, 2)
     InsertarAsientosC AdoAsientos
  End If
  CheckBco.SetFocus
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto, False
End Sub

Private Sub CheckBco_Click()
  If CheckBco.value Then
     If OpcTP(2).value Then
        Label24.Caption = "Cheq. No."
        Label5.Visible = False
        TextDeposito.Visible = False
     Else
        Label24.Caption = UCase$(Moneda)
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
  If CheckEfect.value Then
     TextCantidad.Visible = True
     DCCaja.Visible = True
     DCCaja.SetFocus
     Label6.Visible = True
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
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo = '" & FA.CodigoC & "' ")
       If Not .EOF Then
         .Fields("Email") = TxtEmail
         .Update
       End If
   End If
  End With
End Sub
