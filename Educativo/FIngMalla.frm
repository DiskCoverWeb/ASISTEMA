VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FMalla_Curricular 
   Caption         =   "Ingreso de Cuentas Contables"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   13035
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView TVCatalogo 
      Height          =   2745
      Left            =   105
      TabIndex        =   22
      Top             =   315
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   4842
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11655
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":0000
            Key             =   "K1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":031A
            Key             =   "K2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":0634
            Key             =   "K3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":0A86
            Key             =   "K4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":1360
            Key             =   "K5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngMalla.frx":1C3A
            Key             =   "K6"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   105
      TabIndex        =   3
      Top             =   3255
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DATOS PRINCIPALES"
      TabPicture(0)   =   "FIngMalla.frx":22E0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LabelCtaSup"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LabelTipoCta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DCMateriasP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "MBoxCta"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CheqCualitativa"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CheqPromedia"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CheqImprimir"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DCMaterias"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtHoras"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CheqCualitativa2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtProfesor"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "CURSOS"
      TabPicture(1)   =   "FIngMalla.frx":22FC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGCursos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "MATERIAS"
      TabPicture(2)   =   "FIngMalla.frx":2318
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGMaterias"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "TABLA DE EQUIVALENCIAS"
      TabPicture(3)   =   "FIngMalla.frx":2334
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGEquivalencia"
      Tab(3).ControlCount=   1
      Begin VB.TextBox TxtProfesor 
         Height          =   330
         Left            =   5355
         TabIndex        =   26
         Top             =   735
         Width           =   4005
      End
      Begin VB.CheckBox CheqCualitativa2 
         Caption         =   "Cualitativa 2"
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
         Left            =   9450
         TabIndex        =   25
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox TxtHoras 
         Height          =   330
         Left            =   4620
         TabIndex        =   24
         Text            =   "0"
         Top             =   735
         Width           =   645
      End
      Begin MSDataListLib.DataCombo DCMaterias 
         Bindings        =   "FIngMalla.frx":2350
         DataSource      =   "AdoMaterias"
         Height          =   315
         Left            =   105
         TabIndex        =   18
         Top             =   1575
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Materia"
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
      Begin VB.Frame Frame1 
         Caption         =   "DATOS DEL CURSO"
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
         TabIndex        =   16
         Top             =   2520
         Width           =   12090
         Begin VB.Label Label4 
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   1170
            Left            =   105
            TabIndex        =   17
            Top             =   315
            Width           =   11880
         End
      End
      Begin MSDataGridLib.DataGrid DGCursos 
         Bindings        =   "FIngMalla.frx":236A
         Height          =   2850
         Left            =   -74895
         TabIndex        =   14
         ToolTipText     =   "<Ctrl+F5> Modificar Datos"
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   5027
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
      Begin VB.CheckBox CheqImprimir 
         Caption         =   "Imprimir"
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
         TabIndex        =   13
         Top             =   840
         Width           =   1065
      End
      Begin VB.CheckBox CheqPromedia 
         Caption         =   "Promedia"
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
         TabIndex        =   12
         Top             =   420
         Width           =   1170
      End
      Begin VB.CheckBox CheqCualitativa 
         Caption         =   "Cualitativa"
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
         Left            =   9450
         TabIndex        =   11
         Top             =   420
         Width           =   1275
      End
      Begin MSMask.MaskEdBox MBoxCta 
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   735
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGMaterias 
         Bindings        =   "FIngMalla.frx":2382
         Height          =   2850
         Left            =   -74895
         TabIndex        =   15
         ToolTipText     =   "<Ctrl+F5> Modificar Datos"
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   5027
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
      Begin MSDataListLib.DataCombo DCMateriasP 
         Bindings        =   "FIngMalla.frx":239C
         DataSource      =   "AdoMaterias"
         Height          =   315
         Left            =   105
         TabIndex        =   19
         Top             =   2205
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Materia"
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
      Begin MSDataGridLib.DataGrid DGEquivalencia 
         Bindings        =   "FIngMalla.frx":23B6
         Height          =   2850
         Left            =   -74895
         TabIndex        =   21
         ToolTipText     =   "<Ctrl+F5> Modificar Datos"
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   5027
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Docente de la Materia"
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
         TabIndex        =   27
         Top             =   420
         Width           =   4005
      End
      Begin VB.Label LabelTipoCta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
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
         Left            =   2835
         TabIndex        =   10
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label LabelCtaSup 
         BackColor       =   &H80000005&
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
         Left            =   1470
         TabIndex        =   8
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas"
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
         Left            =   4620
         TabIndex        =   23
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M A T E R I A    S U P E R I O R"
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
         Top             =   1890
         Width           =   12090
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nivel/Curso"
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
         TabIndex        =   4
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C U R S O / S E C C I O N / M A T E R I A"
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
         Top             =   1260
         Width           =   12090
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Superior"
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
         TabIndex        =   7
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Cuenta"
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
         Left            =   2835
         TabIndex        =   9
         Top             =   420
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoMalla 
      Height          =   330
      Left            =   210
      Top             =   525
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Malla"
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
   Begin VB.CommandButton Command2 
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
      Left            =   11445
      Picture         =   "FIngMalla.frx":23D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   11445
      Picture         =   "FIngMalla.frx":2C9E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   210
      Top             =   1155
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Materias"
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
   Begin MSAdodcLib.Adodc AdoCursos 
      Height          =   330
      Left            =   210
      Top             =   840
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Cursos"
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
   Begin MSAdodcLib.Adodc AdoEmp 
      Height          =   330
      Left            =   210
      Top             =   1470
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Emp"
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
   Begin MSAdodcLib.Adodc AdoEquivalencia 
      Height          =   330
      Left            =   210
      Top             =   1785
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Equivalencia"
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
   Begin MSAdodcLib.Adodc AdoPresupuesto 
      Height          =   330
      Left            =   210
      Top             =   2100
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Presupuesto"
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
      Left            =   210
      Top             =   2415
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M A L L A     C U R R I C U L A R"
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
      TabIndex        =   0
      Top             =   105
      Width           =   11250
   End
End
Attribute VB_Name = "FMalla_Curricular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Cta_Ini As String
Dim Cta_Fin As String
Dim Codigo_Ini As String
Dim Codigo_Fin As String
Dim CodMat As String
Dim CodMatP As String
Dim nodX As Node

Private Sub Command1_Click()
  If Nuevo Then GrabarCta (True) Else GrabarCta (False)
End Sub

Private Sub Command2_Click()
  Unload FMalla_Curricular
End Sub

Private Sub DCMaterias_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGCursos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGCursos.AllowUpdate = True
     DGCursos.AllowAddNew = True
     DGCursos.AllowDelete = True
     MsgBox "Proceso Aceptado" & vbCrLf & vbCrLf & "puede Modificar Los Cursos"
     DGCursos.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyC Then Ctrl_C_Grid DGCursos
  If CtrlDown And KeyCode = vbKeyV Then Ctrl_V_Grid "Catalogo_Cursos"
End Sub

Private Sub DGEquivalencia_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGEquivalencia.AllowUpdate = True
     DGEquivalencia.AllowAddNew = True
     DGEquivalencia.AllowDelete = True
     MsgBox "Proceso Aceptado" & vbCrLf & vbCrLf & "puede Modificar Las Equivalencias"
     DGEquivalencia.SetFocus
  End If
End Sub

Private Sub DGMaterias_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGMaterias.AllowUpdate = True
     DGMaterias.AllowAddNew = True
     DGMaterias.AllowDelete = True
     MsgBox "Proceso Aceptado" & vbCrLf & vbCrLf & "puede Modificar Las Materias"
     DGMaterias.SetFocus
  End If
End Sub

Private Sub Form_Activate()
Dim CodigoCtas() As String
  RatonReloj
 'DGDetalle.width = MDI_X_Max - DGDetalle.Left
  TVCatalogo.Height = ((MDI_Y_Max - TVCatalogo.Top) / 2) - 200
  SSTab1.Top = TVCatalogo.Top + TVCatalogo.Height + 50
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 50
  'MDI_Y_Max -SSTab1.Top
  Frame1.Height = SSTab1.Height - Frame1.Top - 100
  Label4.Height = Frame1.Height - Label4.Top - 100
  
  DGEquivalencia.Height = SSTab1.Height - 550
  DGMaterias.Height = SSTab1.Height - 550
  DGCursos.Height = SSTab1.Height - 550
  FormatoMaskCurso MBoxCta
  Si_No = False
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Catalogo_Materias " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Materia "
  SelectDataGrid DGMaterias, AdoMaterias, sSQL
  SelectDBCombo DCMaterias, AdoMaterias, sSQL, "Materia"
  SelectDBCombo DCMateriasP, AdoMaterias, sSQL, "Materia"
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Curso "
  SelectDataGrid DGCursos, AdoCursos, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Equivalencia " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Desde,Hasta "
  SelectDataGrid DGEquivalencia, AdoEquivalencia, sSQL
  
  sSQL = "SELECT CE.CodigoE,CE.CodMat,CE.TC,CE.Horas_Clase,CM.Materia,CE.CodMatP,C.Cliente " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias AS CM,Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Profesor = C.Codigo " _
       & "ORDER BY CE.CodigoE "
  SelectAdodc AdoMalla, sSQL
  With AdoMalla.Recordset
   If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Codigo = "C" & .Fields("CodigoE")
           Cta_Sup = "C" & CodigoCuentaSup(.Fields("CodigoE"))
           Cuenta = .Fields("CodigoE")
           Select Case .Fields("TC")
             Case "C", "P", "N"
                  Cuenta = .Fields("CodigoE")
                  If AdoCursos.Recordset.RecordCount > 0 Then
                     AdoCursos.Recordset.MoveFirst
                     AdoCursos.Recordset.Find ("Curso = '" & .Fields("CodigoE") & "' ")
                     If Not AdoCursos.Recordset.EOF Then
                        Cuenta = .Fields("CodigoE") & " " & AdoCursos.Recordset.Fields("Descripcion")
                     End If
                  End If
             Case Else
                  Cuenta = .Fields("CodigoE") & " (" & .Fields("CodMat") & ") " & .Fields("Materia")
           End Select
           AddNewCta .Fields("TC"), .Fields("CodMatP")
          .MoveNext
        Loop
    End If
   End With
   Select Case CodigoUsuario
     Case "ACCESO01", "ACCESO02", "ACCESO03", "ACCESO04", "ACCESO05", "0702164179"
          'Command3.Enabled = True
   End Select
   Command1.SetFocus
   RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoMalla
  ConectarAdodc AdoCursos
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoEquivalencia
End Sub

Private Sub MBoxCta_GotFocus()
  MarcarTexto MBoxCta
End Sub

Private Sub MBoxCta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_LostFocus()
  Codigo = CodigoCuentaSup(CambioCodigoCta(MBoxCta.Text))
  If Codigo = "0" Then Codigo = CambioCodigoCta(MBoxCta.Text)
  sSQL = "SELECT CodigoE " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE CodigoE = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAux, sSQL, False
  If (AdoAux.Recordset.RecordCount <= 0) And (Len(Codigo) > 1) Then
     Cadena = "Warnign: No puede crear este Código," & vbCrLf _
            & "no existe Nivel Superior "
     MsgBox Cadena
     MBoxCta.SetFocus
  Else
     LabelCtaSup.Caption = CambioCodigoCtaSup(CambioCodigoCta(MBoxCta.Text))
     Codigos = CambioCodigoCta(MBoxCta.Text)
     sSQL = "SELECT CodigoE " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE CodigoE = '" & Codigos & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     SelectAdodc AdoAux, sSQL
     If (AdoAux.Recordset.RecordCount > 0) And (Nuevo) Then
        MsgBox "Esta Cuenta ya existe, vuelva a ingresar otra cuenta."
        MBoxCta.SetFocus
     End If
  End If
End Sub

Public Sub LlenarCta()
Dim DatosCurso As String
Dim CodCurso As String
  With AdoMalla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Codigo1 & "' ")
       If Not .EOF Then
          TipoSubCta = .Fields("TC")
          MBoxCta.Text = FormatoCodigoCurso(.Fields("CodigoE"))
          CodMat = .Fields("CodMat")
          CodMatP = .Fields("CodMatP")
          TxtHoras = .Fields("Horas_Clase")
          TxtProfesor = .Fields("Cliente")
          If .Fields("TC") = "M" Then
              CodCurso = CodigoCuentaSup(.Fields("CodigoE"))
          Else
              CodCurso = .Fields("CodigoE")
          End If
          If CodCurso = "" Then CodCurso = Ninguno
          Select Case TipoSubCta
            Case "C": LabelTipoCta.Caption = "SECCION"
            Case "P": LabelTipoCta.Caption = "CURSO"
            Case "M": LabelTipoCta.Caption = "MATERIA"
            Case "N": LabelTipoCta.Caption = "NIVEL"
            Case Else: LabelTipoCta.Caption = "NINGUNO"
          End Select
          DatosCurso = ""
          If AdoCursos.Recordset.RecordCount > 0 Then
             AdoCursos.Recordset.MoveFirst
             AdoCursos.Recordset.Find ("Curso = '" & CodCurso & "' ")
             If Not AdoCursos.Recordset.EOF Then
                DCMaterias.Text = AdoCursos.Recordset.Fields("Descripcion")
                If Len(AdoCursos.Recordset.Fields("Descripcion")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Descripcion") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Paralelo")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Paralelo") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Bachiller")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Bachiller") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Especialidad")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Especialidad") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Ciclo")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Ciclo") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Seccion")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Seccion") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Titulo")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Titulo") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Tipo_Titulo")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Tipo_Titulo") & vbCrLf
                If Len(AdoCursos.Recordset.Fields("Codigo_Titulo")) > 1 Then DatosCurso = DatosCurso & AdoCursos.Recordset.Fields("Codigo_Titulo") & vbCrLf
             End If
          End If
          DCMaterias.Text = Ninguno
          If Len(CodMat) > 1 Then
             If AdoMaterias.Recordset.RecordCount > 0 Then
                AdoMaterias.Recordset.MoveFirst
                AdoMaterias.Recordset.Find ("CodMat = '" & CodMat & "' ")
                If Not AdoMaterias.Recordset.EOF Then
                   DCMaterias.Text = AdoMaterias.Recordset.Fields("Materia")
                   If AdoMaterias.Recordset.Fields("C") Then CheqCualitativa.value = 1 Else CheqCualitativa.value = 0
                   If AdoMaterias.Recordset.Fields("C2") Then CheqCualitativa2.value = 1 Else CheqCualitativa2.value = 0
                   If AdoMaterias.Recordset.Fields("P") Then CheqPromedia.value = 1 Else CheqPromedia.value = 0
                   If AdoMaterias.Recordset.Fields("I") Then CheqImprimir.value = 1 Else CheqImprimir.value = 0
                End If
             End If
          End If
          DCMateriasP.Text = Ninguno
          If Len(CodMatP) > 1 Then
             If AdoMaterias.Recordset.RecordCount > 0 Then
                AdoMaterias.Recordset.MoveFirst
                AdoMaterias.Recordset.Find ("CodMat = '" & CodMatP & "' ")
                If Not AdoMaterias.Recordset.EOF Then
                   DCMateriasP.Text = AdoMaterias.Recordset.Fields("Materia")
                End If
             End If
          End If
          Label4.Caption = DatosCurso
          Nuevo = False
       End If
   End If
  End With
End Sub

Public Sub GrabarCta(NuevaCta As Boolean)
  NuevaCta = False
  CodMat = Ninguno
  CodMatP = Ninguno
  Select Case LabelTipoCta.Caption
    Case "SECCION": TipoSubCta = "C"
    Case "CURSO": TipoSubCta = "P"
    Case "MATERIA": TipoSubCta = "M"
    Case "NIVEL": TipoSubCta = "N"
    Case Else: TipoSubCta = Ninguno
  End Select
  If TipoSubCta = "M" Then
     If AdoMaterias.Recordset.RecordCount > 0 Then
        AdoMaterias.Recordset.MoveFirst
        AdoMaterias.Recordset.Find ("Materia = '" & DCMaterias & "' ")
        If Not AdoMaterias.Recordset.EOF Then
           CodMat = AdoMaterias.Recordset.Fields("CodMat")
           If CheqCualitativa.value = 1 Then AdoMaterias.Recordset.Fields("C") = True Else AdoMaterias.Recordset.Fields("C") = False
           If CheqCualitativa2.value = 1 Then AdoMaterias.Recordset.Fields("C2") = True Else AdoMaterias.Recordset.Fields("C2") = False
           If CheqPromedia.value = 1 Then AdoMaterias.Recordset.Fields("P") = True Else AdoMaterias.Recordset.Fields("P") = False
           If CheqImprimir.value = 1 Then AdoMaterias.Recordset.Fields("I") = True Else AdoMaterias.Recordset.Fields("I") = False
           AdoMaterias.Recordset.Update
        End If
        AdoMaterias.Recordset.MoveFirst
        AdoMaterias.Recordset.Find ("Materia = '" & DCMateriasP & "' ")
        If Not AdoMaterias.Recordset.EOF Then CodMatP = AdoMaterias.Recordset.Fields("CodMat")
     End If
  End If
  Codigo1 = CambioCodigoCta(MBoxCta.Text)
  Codigo = "C" & Codigo1
  Cta_Sup = "C" & CodigoCuentaSup(Codigo1)
  Cuenta = Codigo1 & " - " & DCMaterias.Text
  Mensajes = "Esta seguro de Grabar la cuenta" & vbCrLf _
           & "No. [" & Codigo1 & "] - " & DCMaterias
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY CodigoE "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("CodigoE = '" & Codigo1 & "' ")
          If .EOF Then
             .AddNew
             .Fields("CodigoE") = Codigo1
             .Fields("Profesor") = Ninguno
             .Fields("Lleno") = vbFalse
             .Fields("Cupo") = 50
             .Fields("NG") = vbFalse
             .Fields("Horas_Clase") = 0
              NuevaCta = True
          End If
      Else
         .AddNew
         .Fields("CodigoE") = Codigo1
         .Fields("Profesor") = Ninguno
         .Fields("Lleno") = vbFalse
         .Fields("Cupo") = 50
         .Fields("NG") = vbFalse
         .Fields("Horas_Clase") = 0
      End If
     .Fields("TC") = TipoSubCta
      If Mid$(Codigo1, 1, 1) <= 1 Then .Fields("Seccion") = 1 Else .Fields("Seccion") = 2
     .Fields("CodMat") = CodMat
     .Fields("CodMatP") = CodMatP
      If TipoSubCta = "M" Then
        .Fields("Id_No") = Val(Mid$(Codigo1, Len(Codigo1) - 1, 2))
      Else
        .Fields("Id_No") = 0
      End If
     .Fields("Horas_Clase") = Val(TxtHoras)
     .Fields("Periodo") = Periodo_Contable
     .Fields("Item") = NumEmpresa
     .Update
     End With
  End If
  If NuevaCta Then
     Control_Procesos Normal, "Nuva Cuenta: " & Codigo1 & " - " & DCMaterias
  Else
     Control_Procesos Normal, "Modificacion de Cuenta: " & Codigo1 & " - " & DCMaterias
  End If
  sSQL = "SELECT CE.CodigoE,CE.CodMat,CE.TC,CE.Horas_Clase,CM.Materia,CE.CodMatP,C.Cliente " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias AS CM,Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Profesor = C.Codigo " _
       & "ORDER BY CE.CodigoE "
  SelectAdodc AdoMalla, sSQL
  IE = TVCatalogo.SelectedItem.Index
  If NuevaCta Then TVCatalogo.Nodes(IE).Text = Codigo1 & " (" & CodMat & ") " & DCMaterias
  TVCatalogo.Refresh
  Label6.Visible = True
  Nuevo = False
End Sub

Public Sub NuevaCta()
  OpcNor.value = True
  LabelNumero.Caption = "0"
  LabelNumero.Caption = ""
  TextConcepto.Text = ""
  TextPresupuesto.Text = ""
  LabelCtaSup.Caption = ""
  MBoxCta.Text = LimpiarCtas
  Nuevo = True
  MBoxCta.SetFocus
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IdCurso As Integer
  Keys_Especiales Shift
  PresionoEnter KeyCode
  Codigo1 = SinEspaciosIzq(TVCatalogo.SelectedItem)
  If CtrlDown And KeyCode = vbKeyI Then Codigo_Ini = Codigo1
  If CtrlDown And KeyCode = vbKeyU Then Codigo_Fin = Codigo1
  If CtrlDown And KeyCode = vbKeyC Then
     Codigo_Ini = Codigo1
     Cta_Ini = Codigo2
     Cadena = "COPIAR EL CURSO: " & Codigo_Ini & vbCrLf & vbCrLf _
            & "INGRESE EL NUEVO CURSO:"
     Codigo2 = UCase(InputBox(Cadena, "COPIAR CURSOS", Codigo_Ini))
     If Len(Codigo2) = 7 And Codigo2 <> Codigo_Ini Then
        sSQL = "SELECT * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Mid$(CodigoE,1," & Len(Codigo_Ini) & ") = '" & Codigo_Ini & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY CodigoE "
        SelectAdodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                SetAdoAddNew "Catalogo_Estudiantil"
                For IdCurso = 0 To .Fields.Count - 1
                    If .Fields(IdCurso).Name = "CodigoE" Then
                        SetAdoFields .Fields(IdCurso).Name, Codigo2 & Mid$(.Fields(IdCurso), 8, 3)
                    Else
                        SetAdoFields .Fields(IdCurso).Name, .Fields(IdCurso)
                    End If
                Next IdCurso
                SetAdoUpdate
               .MoveNext
             Loop
             MsgBox "Copia Realizada con exito"
         End If
        End With
     End If
  End If
 'If CtrlDown And KeyCode = vbKeyDelete Then EliminarCtaGrupo
  If KeyCode = vbKeyDelete Then EliminarCta
End Sub

Private Sub TVCatalogo_LostFocus()
    Codigo1 = SinEspaciosIzq(TVCatalogo.SelectedItem)
    Cadena = TVCatalogo.SelectedItem
    LlenarCta
End Sub

Public Sub AddNewCta(TipoTC As String, CodMatP As String)
    Select Case TipoTC
      Case "C": IE = 1
      Case "P": IE = 2
      Case "N": IE = 3
      Case "M": IE = 4
      Case Else: IE = 5
    End Select
    If CodMatP <> Ninguno Then IE = 6
    If Len(Codigo) = 2 Then
       Set nodX = TVCatalogo.Nodes.Add(, , Codigo, Cuenta, ImageList1.ListImages(IE).key, ImageList1.ListImages(IE).key)
    Else
       Set nodX = TVCatalogo.Nodes.Add(Cta_Sup, tvwChild, Codigo, Cuenta, ImageList1.ListImages(IE).key, ImageList1.ListImages(IE).key)
    End If
    nodX.Tag = Mid$(Codigo, 2, Len(Codigo))
End Sub

Public Sub UpdateCta(TipoTC As String)
 ' TVCatalogo.SelectedItem = Cuenta
  Select Case TipoTC
    Case "DG": IE = 9
    Case "RF": IE = 11
    Case "CF": IE = 11
    Case "RI": IE = 21
    Case "RP": IE = 11
    Case "RI": IE = 11
    Case "CI": IE = 11
    Case "C": IE = 12
    Case "P": IE = 13
    Case "I": IE = 14
    Case "G": IE = 15
    Case "CJ": IE = 16
    Case "BA": IE = 17
    Case "CS": IE = 18
    Case "PS": IE = 19
    Case "CP": IE = 20
    Case "PM": IE = 22
    Case "TJ": IE = 23
    Case Else: IE = 10
  End Select
'  nodX.Image = ImageList1.ListImages(IE).key
'  nodX.SelectedImage = ImageList1.ListImages(IE).key
End Sub

Public Sub EliminarCta()
  Codigo1 = CambioCodigoCta(MBoxCta.Text)
  Cadena = SinEspaciosIzq(TVCatalogo.SelectedItem)
  Codigo2 = Trim(Mid$(TVCatalogo.SelectedItem, Len(Cadena) + 1, Len(TVCatalogo.SelectedItem)))
  With AdoMalla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE = '" & Cadena & "' ")
       If Not .EOF Then
          sSQL = "SELECT CodMat " _
               & "FROM Trans_Notas " _
               & "WHERE CodE = '" & Cadena & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "GROUP BY CodMat " _
               & "ORDER BY CodMat "
          SelectAdodc AdoAux, sSQL, False
          If AdoAux.Recordset.RecordCount > 0 Then
             Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
                      & "No se puede eliminar esta(s) Materia(s): " & vbCrLf _
                      & "porque tiene(n) novimiento(s)."
             MsgBox Mensajes
          Else
             Mensajes = "Esta seguro que desea eliminar la Materia:" & vbCrLf & vbCrLf _
                      & Cadena & ": " & Codigo2 & vbCrLf & vbCrLf _
                      & "y sus grupos "
             Titulo = "Pregunta de Eliminacion"
             If BoxMensaje = vbYes Then
                Cadena1 = TVCatalogo.Nodes(TVCatalogo.SelectedItem.Index).Tag
                For I = TVCatalogo.Nodes.Count To 1 Step -1
                    If Cadena1 = TVCatalogo.Nodes(I).Tag Then
                       SQL1 = "DELETE * " _
                            & "FROM Catalogo_Estudiantil " _
                            & "WHERE CodigoE = '" & TVCatalogo.Nodes(I).Tag & "' " _
                            & "AND Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' "
                       ConectarAdoExecute SQL1
                       TVCatalogo.Nodes.Remove I
                    End If
                Next I
             End If
          End If
       End If
   End If
  End With
End Sub

'''Public Sub EliminarCtaGrupo()
'''  With AdoCta.Recordset
'''   If .RecordCount > 0 Then
'''        sSQL = "SELECT Cta,Count(Cta) As Cant_Cta " _
'''             & "FROM Transacciones " _
'''             & "WHERE Cta BETWEEN '" & Codigo_Ini & "' and '" & codigofin & "' " _
'''             & "AND Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "GROUP BY Cta " _
'''             & "ORDER BY Cta "
'''        SelectAdodc AdoCtas, sSQL, False
'''        If AdoCtas.Recordset.RecordCount > 0 Then
'''           Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
'''                    & "No se puede eliminar esta(s) Cuenta(s): " & vbCrLf
'''           Do While Not AdoCtas.Recordset.EOF
'''              Mensajes = Mensajes & AdoCtas.Recordset.Fields("Cta") & " Cantidad de Movimientos: " & AdoCtas.Recordset.Fields("Cant_Cta") & vbCrLf
'''              AdoCtas.Recordset.MoveNext
'''           Loop
'''           Mensajes = Mensajes & vbCrLf & "porque tiene(n) novimiento(s)."
'''           MsgBox Mensajes
'''        Else
'''           Mensajes = "Esta seguro que desea eliminar" & vbCrLf & vbCrLf _
'''                    & "Desde: " & Codigo_Ini & " hasta " & Codigo_Fin & vbCrLf & vbCrLf _
'''                    & "y sus grupos "
'''           Titulo = "Pregunta de Eliminacion"
'''           If BoxMensaje = vbYes Then
'''              For I = TVCatalogo.Nodes.Count To 1 Step -1
'''                  If Codigo_Ini <= TVCatalogo.Nodes(I).Tag And TVCatalogo.Nodes(I).Tag <= Codigo_Fin Then
'''                     SQL1 = "DELETE * " _
'''                          & "FROM Trans_Presupuestos " _
'''                          & "WHERE Cta = '" & TVCatalogo.Nodes(I).Tag & "' " _
'''                          & "AND Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' "
'''                     ConectarAdoExecute SQL1
'''                    .MoveFirst
'''                    .Find ("Codigo like '" & TVCatalogo.Nodes(I).Tag & "' ")
'''                     If Not .EOF Then
'''                       .Delete
'''                        TVCatalogo.Nodes.Remove I
'''                     End If
'''                  End If
'''              Next I
'''           End If
'''        End If
'''   End If
'''  End With
'''End Sub
'''
Private Sub TxtHoras_GotFocus()
  MarcarTexto TxtHoras
End Sub

Private Sub TxtHoras_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtHoras_LostFocus()
  TextoValido TxtHoras, True, , 0
End Sub
