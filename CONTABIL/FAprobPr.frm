VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Aprobacion 
   Caption         =   "PRESTAMO"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtConyugue 
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
      MaxLength       =   49
      TabIndex        =   41
      Top             =   1890
      Width           =   7680
   End
   Begin VB.TextBox TxtRUCS 
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
      MaxLength       =   15
      TabIndex        =   40
      Text            =   "0"
      Top             =   2625
      Width           =   1800
   End
   Begin MSDataListLib.DataCombo DCTipoPrestamo 
      Bindings        =   "FAprobPr.frx":0000
      DataSource      =   "AdoTipoPrest"
      Height          =   315
      Left            =   1365
      TabIndex        =   3
      Top             =   420
      Width           =   5790
      _ExtentX        =   10213
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3585
      Left            =   105
      TabIndex        =   25
      Top             =   3045
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   6324
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Tabla de Pagos"
      TabPicture(0)   =   "FAprobPr.frx":001B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGTabla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Impresion de Comprobantes"
      TabPicture(1)   =   "FAprobPr.frx":0037
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGGarantes"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "Command6"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Asiento Contable"
      TabPicture(2)   =   "FAprobPr.frx":0053
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGAsiento"
      Tab(2).Control(1)=   "TextConcepto"
      Tab(2).Control(2)=   "Command2"
      Tab(2).Control(3)=   "LabelEgresos"
      Tab(2).Control(4)=   "LabelIngresos"
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(6)=   "Label13"
      Tab(2).ControlCount=   7
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "FAprobPr.frx":006F
         Height          =   1905
         Left            =   -74895
         TabIndex        =   39
         Top             =   1155
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   3360
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
      Begin MSDataGridLib.DataGrid DGGarantes 
         Bindings        =   "FAprobPr.frx":0088
         Height          =   2115
         Left            =   -74895
         TabIndex        =   38
         Top             =   1365
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   3731
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
      Begin MSDataGridLib.DataGrid DGTabla 
         Bindings        =   "FAprobPr.frx":00A2
         Height          =   3060
         Left            =   105
         TabIndex        =   37
         Top             =   420
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   5398
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
         Height          =   645
         Left            =   -73845
         MaxLength       =   119
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   420
         Width           =   8205
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Acreditar &Prestamo"
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
         Left            =   -65550
         Picture         =   "FAprobPr.frx":00B9
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   420
         Width           =   1905
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&1.- Imprimir Pagaré"
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
         Left            =   -74895
         Picture         =   "FAprobPr.frx":04FB
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   420
         Width           =   2010
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&2.- Imprimir Liquidación"
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
         Left            =   -72795
         Picture         =   "FAprobPr.frx":093D
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   420
         Width           =   2220
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&3.- Imprimir Tabla Cliente"
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
         Left            =   -70485
         Picture         =   "FAprobPr.frx":0D7F
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   420
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&4.- Imprimir Tabla de Amortizacion"
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
         Left            =   -67860
         Picture         =   "FAprobPr.frx":11C1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   420
         Width           =   3165
      End
      Begin VB.Label LabelEgresos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -65865
         TabIndex        =   35
         Top             =   3150
         Width           =   1905
      End
      Begin VB.Label LabelIngresos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -67755
         TabIndex        =   36
         Top             =   3150
         Width           =   1905
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Concepto"
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
         Left            =   -74895
         TabIndex        =   26
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTALES"
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
         Left            =   -68805
         TabIndex        =   33
         Top             =   3150
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc AdoTabla 
      Height          =   330
      Left            =   315
      Top             =   3255
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
      Caption         =   "Tabla"
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
   Begin VB.TextBox TxtRazonSocial 
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
      MaxLength       =   49
      TabIndex        =   15
      Top             =   1575
      Width           =   7680
   End
   Begin VB.TextBox TxtNombresS 
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
      TabIndex        =   13
      Top             =   1260
      Width           =   7680
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   330
      Left            =   4305
      TabIndex        =   11
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   840
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   192
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCC-C"
      Mask            =   "CCCCCCCC-C"
      PromptChar      =   "0"
   End
   Begin VB.TextBox TextImpuesto 
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
      Left            =   7245
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   1905
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1575
      TabIndex        =   9
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
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
   Begin VB.TextBox TextTP 
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
      Left            =   1995
      MaxLength       =   30
      TabIndex        =   21
      Top             =   2625
      Width           =   3585
   End
   Begin VB.CommandButton Command3 
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
      Left            =   9345
      Picture         =   "FAprobPr.frx":1603
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   840
      Width           =   1170
   End
   Begin VB.TextBox TextInt 
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
      Left            =   5670
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   2625
      Width           =   750
   End
   Begin VB.TextBox TextMonto 
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
      Left            =   7350
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   2625
      Width           =   1905
   End
   Begin VB.TextBox TextMeses 
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
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   2625
      Width           =   750
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   315
      Top             =   3570
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   315
      Top             =   3885
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
      Caption         =   "Asiento"
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   315
      Top             =   4200
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
      Caption         =   "TipoPrest"
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
   Begin MSAdodcLib.Adodc AdoConyugue 
      Height          =   330
      Left            =   315
      Top             =   4515
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
      Caption         =   "Conyugue"
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
      Top             =   4830
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
   Begin MSAdodcLib.Adodc AdoGarantes 
      Height          =   330
      Left            =   315
      Top             =   5145
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
      Caption         =   "Garantes"
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
   Begin MSAdodcLib.Adodc AdoPagare 
      Height          =   330
      Left            =   315
      Top             =   5460
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
      Caption         =   "Pagare"
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Conyugue"
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
      TabIndex        =   42
      Top             =   1890
      Width           =   1485
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Apro."
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
      Width           =   1170
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA No."
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
      TabIndex        =   10
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Inspección/Avalúo"
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
      Left            =   7245
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Representante"
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
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombres"
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
      TabIndex        =   12
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label LabelCredNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   9240
      TabIndex        =   7
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Credito No."
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
      TabIndex        =   6
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &VALIDACION DE PRESTAMO"
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
      TabIndex        =   2
      Top             =   105
      Width           =   5790
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE PRESTAMO"
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
      Left            =   1995
      TabIndex        =   17
      Top             =   2310
      Width           =   3585
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interes"
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
      Left            =   5670
      TabIndex        =   18
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto de Prestamo"
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
      Left            =   7350
      TabIndex        =   20
      Top             =   2310
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
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
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I./R.U.C."
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
      TabIndex        =   16
      Top             =   2310
      Width           =   1800
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA APERT."
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
      TabIndex        =   8
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "Aprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InsertarMontosPrestamo(DtaCta As Adodc, _
                                  CuentaNo As String, _
                                  TDebe As Currency, _
                                  THaber As Currency)
  If CuentaNo <> "00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
  sSQL = "SELECT TOP 1 * " _
       & "FROM Trans_Libretas " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       SaldoDisp = 0
       SaldoCont = 0
       ID_Trans = 0
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           ID_Trans = .Fields("IDT")
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Cuenta_No") = CuentaNo
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
      .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
      .Fields("ID") = NumeroLineas + 1
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
       SetUpdate DtaCta
  End With
  End If
End Sub

Public Sub ListarCuenta(Cuenta_No As String)
   Codigo = SinEspaciosIzq(DCTipoPrestamo.Text)
   Contrato_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 2)
   Cuenta_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 3)
   TxtNombresS.Text = ""
   TxtRUCS.Text = "000000000-0-000"
   MBoxCuenta.Text = Cuenta_No
   TxtNombresS.Text = ""
   TxtRazonSocial.Text = ""
   sSQL = "SELECT * " _
        & "FROM Tipo_Prestamo " _
        & "WHERE TP = '" & Codigo & "' " _
        & "AND Item = '" & NumEmpresa & "' "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Si_No = .Fields("DM")
        TipoProc = .Fields("CTP")
        Cta_Prestamos = .Fields("Cta_Prestamo")
        Cta_Intereses = .Fields("Cta_Interes")
        Cta_Amortizacion = .Fields("Cta_Capital")
        Cta_Impuestos = .Fields("Cta_Impuesto")
        Cta_Provision = .Fields("Cta_Provision")
        TextTP.Text = Codigo & "  " & .Fields("Descripcion")
        If Si_No Then Label5.Caption = " Dias" Else Label5.Caption = " Meses"
        'MsgBox Codigo & vbCrLf & NumEmpresa
    End If
   End With
   CodigoCli = Ninguno
   sSQL = "SELECT * " _
        & "FROM Cuentas " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        CodigoCli = .Fields("Codigo")
        Moneda_US = False '.Fields("ME")
    End If
   End With
   sSQL = "SELECT * FROM Clientes " _
        & "WHERE Codigo = '" & CodigoCli & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        TxtNombresS.Text = .Fields("Cliente")
        TxtRUCS.Text = .Fields("CI_RUC")
        TxtRazonSocial.Text = .Fields("Representante")
        Cod_Remit = .Fields("Est_Civil")
        sSQL = "SELECT * FROM Garantes " _
             & "WHERE TP = '" & Codigo & "' " _
             & "AND Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "ORDER BY Num,GC "
        SelectDataGrid DGGarantes, AdoGarantes, sSQL
        sSQL = "SELECT * " _
             & "FROM Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND T = 'N' " _
             & "ORDER BY T,TP,Credito_No,Fecha "
        SelectDataGrid DGTabla, AdoTabla, sSQL
        With AdoTabla.Recordset
         If .RecordCount > 0 Then
             MBoxFecha.Text = .Fields("Fecha")
             TextMonto.Text = .Fields("Saldo_Pendiente")
             LabelCredNo.Caption = .Fields("Credito_No")
             TextInt.Text = Format(.Fields("Tasa"), "#,##0.00")
             TextMeses.Text = .Fields("Meses")
         End If
        End With
    End If
   End With
   TxtConyugue.Text = ""
   If Cod_Remit <> "S" Then Cod_Remit = CodigoCli Else Cod_Remit = "_"
   sSQL = "SELECT * FROM Conyugue " _
        & "WHERE Codigo = '" & Cod_Remit & "' "
   SelectAdodc AdoConyugue, sSQL
   If AdoConyugue.Recordset.RecordCount > 0 Then
      TxtConyugue.Text = AdoConyugue.Recordset.Fields("Nombres")
   End If
End Sub

Private Sub Command1_Click()
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Liquidación"
If BoxMensaje = vbYes Then
RatonReloj
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     MiFecha = FechaStrg(.Fields("Fecha"))
 End If
End With
InicioX = 0.5: InicioY = 0
sSQL = "SELECT * FROM Pagare " _
     & "ORDER BY LN "
SelectAdodc AdoPagare, sSQL
DataAnchoCampos 1, AdoPagare, 8, TipoTimes, 1
Ancho(3) = 19
CantCampos = 3
Pagina = 1
Printer.FontBold = True
'Iniciamos la impresion
Printer.FontBold = False
If AdoTabla.Recordset.RecordCount > 0 Then
   PosLinea = 1
   Printer.FontSize = 14
   Printer.FontBold = True
   Printer.CurrentX = CentrarTextoEncab(Empresa, 1, 18)
   Printer.CurrentY = 0.3
   Printer.Print Empresa
   Printer.Line (1, 1)-(18, 1), QBColor(0)
   Printer.FontSize = 12
   PrinterTexto 1, 1.1, "LIQUIDACION PRESTAMO"
   Printer.FontSize = 10
   PrinterTexto 14, 1.1, "Crédito No."
   PrinterTexto 1.5, 2, "Cliente:"
   PrinterTexto 1.5, 2.4, "Garantes:"
   Printer.FontBold = False
   PrinterVariables 16, 1.1, LabelCredNo.Caption
   PrinterVariables 5, 2, TxtNombresS.Text
   PosLinea = 2.4
   With AdoGarantes.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Cta = .Fields("Nombres")
           If Cta <> "" Then
              PrinterVariables 5, PosLinea, Cta
              PosLinea = PosLinea + 0.4
           End If
          .MoveNext
        Loop
    End If
   End With
   Printer.FontBold = True
   PrinterTexto 1.5, PosLinea, "Emisión:"
   PrinterTexto 1.5, PosLinea + 0.4, "Vencimiento:"
   PrinterTexto 1.5, PosLinea + 0.8, "Concepto:"
   Printer.FontBold = False
   PrinterTexto 5, PosLinea, FechaStrg(FechaSistema)
   PrinterTexto 5, PosLinea + 0.4, MiFecha
   PrinterLineas 5, PosLinea + 0.8, TextConcepto.Text, 13
   Printer.Line (1, PosLinea)-(18, PosLinea), QBColor(0)
   PosLinea = PosLinea + 0.1
End If
Printer.FontBold = True
PrinterTexto 1.5, PosLinea, "CODIGO"
PrinterTexto 4, PosLinea, "CUENTA"
PrinterTexto 12.5, PosLinea, "DEBITO"
PrinterTexto 15.5, PosLinea, "CREDITO"
PosLinea = PosLinea + 0.4
Printer.Line (1, PosLinea)-(18, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.FontSize = 9
Printer.FontBold = False
Debe = 0: Haber = 0
With AdoAsiento.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Do While Not .EOF
        PrinterFields 1.5, PosLinea, .Fields("CODIGO")
        PrinterFields 4, PosLinea, .Fields("CUENTA")
        PrinterFields 12, PosLinea, .Fields("DEBE")
        PrinterFields 15, PosLinea, .Fields("HABER")
        Debe = Debe + .Fields("DEBE")
        Haber = Haber + .Fields("HABER")
        PosLinea = PosLinea + 0.4
       .MoveNext
     Loop
 End If
End With
Printer.Line (1.5, PosLinea)-(18, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto 11, PosLinea, "TOTALES"
PrinterVariables 12, PosLinea, Debe
PrinterVariables 15, PosLinea, Haber
Printer.FontSize = 10
PosLinea = PosLinea + 0.5
Printer.Line (1.5, PosLinea)-(18, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Cadena = "Valor que acreditamos en la cuenta No. " & Cuenta_No
PrinterTexto 1.5, PosLinea, Cadena
PosLinea = PosLinea + 0.5
Cadena = "Autorizo debitar de mi cuenta de Ahorros las cuotas y gastos por la " _
       & "cancelación de este crédito. " & Chr(10) _
       & "Declaro conocer y aceptar los costos y gastos de este crédito, " _
       & "que me han sido informados por la " & Empresa & "."
PrinterLineas 1.5, PosLinea, Cadena, 16
PosLinea = PosLinea + 1
Printer.FontSize = 8
Printer.Line (1.5, PosLinea)-(18, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterVariables 1.5, PosLinea, "ELABORADO"
PrinterVariables 4.5, PosLinea, "REVISADO"
PrinterVariables 7, PosLinea, "CONTABILIZADO"
PrinterVariables 10.5, PosLinea, "AUTORIZADO"
PrinterVariables 13.5, PosLinea, "RECIBO CONFORME Y AUTORIZO"
PosLinea = PosLinea + 0.5
Printer.Line (1.5, PosLinea)-(18, PosLinea), QBColor(0)
RatonNormal
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Private Sub Command2_Click()
Titulo = "GRABACION"
Mensajes = "Seguro Acreditar Prestamo"
If BoxMensaje = vbYes Then
  Contrato_No = LabelCredNo.Caption
  TipoDoc = SinEspaciosIzq(TextTP.Text)
  Tasa = TextInt.Text
  sSQL = "SELECT * " _
       & "FROM Trans_Prestamos " _
       & "WHERE Credito_No = '" & Contrato_No & "' " _
       & "AND TP = '" & TipoDoc & "' "
  SelectAdodc AdoAux, sSQL
  With AdoTabla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       RatonReloj
       Do While Not .EOF
          If .Fields("Mes_No") <> 0 Then
              AdoAux.Recordset.AddNew
              AdoAux.Recordset.Fields("Fecha_V") = Ninguno
              AdoAux.Recordset.Fields("T") = Procesado
              AdoAux.Recordset.Fields("TP") = TipoDoc
              AdoAux.Recordset.Fields("ME") = False
              AdoAux.Recordset.Fields("Credito_No") = Contrato_No
              AdoAux.Recordset.Fields("Cuenta_No") = MBoxCuenta.Text
              AdoAux.Recordset.Fields("Mes_No") = .Fields("Mes_No")
              AdoAux.Recordset.Fields("Dia") = .Fields("Dia")
              AdoAux.Recordset.Fields("Fecha") = .Fields("Fecha")
              AdoAux.Recordset.Fields("Fecha_C") = FechaSistema
              AdoAux.Recordset.Fields("Capital") = .Fields("Capital")
              AdoAux.Recordset.Fields("Interes") = .Fields("Interes")
              AdoAux.Recordset.Fields("Comision") = .Fields("Comision")
              AdoAux.Recordset.Fields("Pagos") = .Fields("Pagos")
              AdoAux.Recordset.Fields("Saldo") = .Fields("Saldo")
              AdoAux.Recordset.Fields("Item") = NumEmpresa
              SetUpdate AdoAux
          End If
         .MoveNext
       Loop
       RatonNormal
   End If
  End With
  sSQL = "SELECT * FROM Prestamos " _
       & "WHERE Credito_No = '" & Contrato_No & "' " _
       & "AND TP = '" & SinEspaciosIzq(DCTipoPrestamo.Text) & "' "
  SelectAdodc AdoAux, sSQL
  With AdoTabla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       RatonReloj
       Do While Not .EOF
          If .Fields("Mes_No") = 0 Then
              'AdoAux.Recordset.AddNew
              AdoAux.Recordset.Fields("T") = Procesado
              AdoAux.Recordset.Fields("TP") = TipoDoc
              AdoAux.Recordset.Fields("ME") = False
              AdoAux.Recordset.Fields("Credito_No") = Contrato_No
              AdoAux.Recordset.Fields("Tasa") = Tasa
              AdoAux.Recordset.Fields("Plazo") = Int(Val(TextMeses.Text) * 30)
              AdoAux.Recordset.Fields("Cuenta_No") = MBoxCuenta.Text
              AdoAux.Recordset.Fields("Meses") = Val(TextMeses.Text)
              AdoAux.Recordset.Fields("Dia") = .Fields("Dia")
              AdoAux.Recordset.Fields("Fecha") = .Fields("Fecha")
              AdoAux.Recordset.Fields("Capital") = Round(CCur(TextMonto.Text), 2)
              AdoAux.Recordset.Fields("Interes") = Round(TextInt.Text, 2)
              AdoAux.Recordset.Fields("Pagos") = .Fields("Pagos")
              AdoAux.Recordset.Fields("Saldo_Pendiente") = .Fields("Saldo")
              AdoAux.Recordset.Fields("Encaje") = Round(CCur(TextMonto.Text) * 0.2, 2)
              AdoAux.Recordset.Fields("Patrimonio") = 0
              AdoAux.Recordset.Fields("Item") = NumEmpresa
              AdoAux.Recordset.Update
          End If
         .MoveNext
       Loop
       RatonNormal
   End If
  End With
  TextoValido TxtRazonSocial, , True
  If Round(SumaDebe - SumaHaber, 2) = 0 Then
     RatonReloj
     InsertarMontosPrestamo AdoAux, Cuenta_No, 0, TotalLibreta
     If Si_No = False Then
        Valor = CCur(TextMonto.Text) * 0.2
        sSQL = "SELECT * FROM Bloqueos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Item = '" & NumEmpresa & "' "
        SelectAdodc AdoCta, sSQL
        With AdoCta.Recordset
            .AddNew
            .Fields("T") = Normal
            .Fields("Fecha") = FechaSistema
            .Fields("Cuenta_No") = Cuenta_No
            .Fields("Valor") = Round(Valor, 2)
            .Fields("Item") = NumEmpresa
            .Update
        End With
     End If
     Trans_No = 51
     Co.T = Normal
     Co.TP = CompDiario
     Co.Fecha = FechaSistema
     Co.CodigoB = Ninguno
     Co.Efectivo = 0
     Co.Monto_Total = 0
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = TextConcepto.Text
     Co.T_No = Trans_No
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     GrabarComprobante Co, AdoAsiento
  End If
  sSQL = "SELECT TP & '  ' & Credito_No & '  ' & Cuenta_No As TipoP " _
       & "FROM Prestamos " _
       & "WHERE T = 'N' " _
       & "ORDER BY TP,Credito_No "
  SelectDBCombo DCTipoPrestamo, AdoTipoPrest, sSQL, "TipoP", False
  RatonNormal
  MsgBox "Prestamo Otorgado con exito"
  DCTipoPrestamo.SetFocus
End If
End Sub

Private Sub Command3_Click()
 Unload Aprobacion
End Sub

Private Sub Command4_Click()
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Pagaré"
If BoxMensaje = vbYes Then
RatonReloj
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     MiFecha = .Fields("Fecha")
 End If
End With
InicioX = 0.5: InicioY = 0
DataAnchoCampos 1, AdoConyugue, 9, TipoTimes, 1, True
LimiteAlto = Printer.ScaleHeight - 1 'Limite de impresión a lo largo
LimiteAncho = Printer.ScaleWidth     'Limite de impresión a lo largo
AnchoPapel = Printer.ScaleWidth      'Limite de impresión a lo largo

Ancho(3) = 19
CantCampos = 3
Pagina = 1
Printer.FontSize = 9
Printer.FontBold = True
'Iniciamos la impresion
Printer.FontBold = True
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
    .MoveNext
      Contrato_No = LabelCredNo.Caption
      Credito_No = LabelCredNo.Caption
     'Encabezadoado adoPagare
      PosLinea = 1
      Total = CDbl(TextMonto.Text)
      PrinterTexto 4.5, 3.6, LabelCredNo.Caption
      PrinterVariables 4.5, 4, Format(Total, "#,##0.00")
      PrinterVariables 9, 4.8, Format(Total, "#,##0.00")
      PrinterVariables 9.2, 5.5, Format(CDbl(TextInt.Text), "#,##0.00")
      PrinterVariables 16.1, 5.5, "1.20"
      If Si_No Then
         PrinterVariables 14.4, 4.5, TextMeses.Text
      Else
         PrinterVariables 14.4, 4.5, CStr(Val(TextMeses.Text) * 30)
      End If
      PrinterVariables 4.7, 5.8, TextMeses.Text
      PrinterVariables 2.6, 6.1, Format(.Fields("Pagos"), "#,##0.00")
      PrinterVariables 8.1, 6.1, CStr(Day(MiFecha))
      PrinterVariables 16.5, 6.1, .Fields("Fecha")
      Codigo1 = "": Codigo2 = "": Codigo3 = "": Codigo4 = ""
 End If
End With
PrinterVariables 2.5, 16, NombreCiudad
PrinterVariables 5, 16, FechaStrg(FechaSistema)
PrinterVariables 2.5, 25, NombreCiudad
PrinterVariables 6, 25, FechaStrg(FechaSistema)

PrinterVariables 3, 18.8, TxtRazonSocial.Text
PrinterVariables 3, 19.15, TxtRUCS.Text
PrinterVariables 3, 21.8, TxtRazonSocial.Text
PrinterVariables 3, 22.15, TxtRUCS.Text
Cta = "": Cta_Sup = ""
With AdoConyugue.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     Cta = .Fields("Nombres")
     Cta_Sup = .Fields("Cedula")
     PrinterVariables 10, 18.8, Cta
     PrinterVariables 10, 19.15, Cta_Sup
     PrinterVariables 10, 21.8, Cta
     PrinterVariables 10, 22.15, Cta_Sup
 End If
End With
PosLinea = 26.5
With AdoGarantes.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Opcion = .Fields("Num")
     Do While Not .EOF
        If Opcion <> .Fields("Num") Then
           Opcion = .Fields("Num")
           PosLinea = PosLinea + 1.8
        End If
        If PosLinea > 27 Then
           Printer.NewPage
           PosLinea = 3.5
        End If
        Cta = .Fields("Nombres")
        Cta_Sup = .Fields("CI")
        If .Fields("GC") Then PCol = 3 Else PCol = 10
        PrinterVariables PCol, PosLinea, Cta
        PrinterVariables PCol, PosLinea + 0.35, Cta_Sup
       .MoveNext
     Loop
 End If
End With
RatonNormal
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Tabla Cliente"
If BoxMensaje = vbYes Then
RatonReloj
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     MiFecha = FechaStrg(.Fields("Fecha"))
 End If
End With
InicioX = 0.5: InicioY = 0
sSQL = "SELECT * FROM Pagare " _
     & "ORDER BY LN "
SelectAdodc AdoPagare, sSQL
DataAnchoCampos 1, AdoPagare, 8, TipoTimes, 1
Ancho(3) = 19
CantCampos = 3
Pagina = 1
Printer.FontBold = True
'Iniciamos la impresion
Printer.FontBold = False
If AdoTabla.Recordset.RecordCount > 0 Then
   PosLinea = 1
   Printer.FontSize = 14
   Printer.FontBold = True
   PrinterTexto 1, 0.3, Empresa
   Printer.FontSize = 12
   PrinterTexto 1, 1.1, "LIQUIDACION PRESTAMO DE AMORTIZACION GRADUAL"
   PrinterTexto 1, 1.6, "PAGOS MENSUALES"
   Printer.FontSize = 10
   PrinterTexto 1.5, 2.2, "Fecha de emisión:"
   PrinterTexto 1.5, 2.6, "Plazo:"
   PrinterTexto 1.5, 3, "Tasa:"
   PrinterTexto 1.5, 3.4, "Cliente:"
   PrinterTexto 1.5, 3.8, "Garantes:"
   Printer.FontBold = False
   PrinterTexto 6, 2.2, FechaStrg(FechaSistema)
   If Si_No Then
      PrinterVariables 6, 2.6, TextMeses.Text & " día(s)"
   Else
      PrinterVariables 6, 2.6, TextMeses.Text & " meses"
   End If
   PrinterVariables 6, 3, Round(CCur(TextInt.Text), 2) & "%"
   PrinterVariables 6, 3.4, TxtNombresS.Text
   PosLinea = 3.8
   With AdoGarantes.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Cta = .Fields("Nombres")
           If Cta <> "" Then
              PrinterVariables 6, PosLinea, Cta
              PosLinea = PosLinea + 0.4
           End If
          .MoveNext
        Loop
    End If
   End With
   PosLinea = PosLinea + 0.2
   Printer.FontBold = True
   Printer.FontSize = 12
   PrinterTexto 1.5, PosLinea, UCase(TextTP.Text)
   PosLinea = PosLinea + 0.7
   Printer.FontSize = 10
   PrinterTexto 1.5, PosLinea, "Valor a Financiar:"
   Total = CCur(TextMonto.Text)
   PrinterVariables 6, PosLinea, Total
   PosLinea = PosLinea + 0.5
End If
Printer.FontBold = True
PrinterTexto 1.5, PosLinea, "CUOTA No."
PrinterTexto 3.5, PosLinea, "VENCIMIENTO"
PrinterTexto 7, PosLinea, "DIVIDENDO"
PrinterTexto 10, PosLinea, "SALDO"
PosLinea = PosLinea + 0.4
Printer.FontBold = False
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Do While Not .EOF
        If .Fields("Mes_No") <> 0 Then
            PrinterFields 1.5, PosLinea, .Fields("Mes_No")
            PrinterTexto 3.6, PosLinea, FechaStrgCorta(.Fields("Fecha"))
            PrinterFields 6.5, PosLinea, .Fields("Pagos")
            PrinterFields 9, PosLinea, .Fields("Saldo")
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto - 1 Then
               Printer.NewPage
               PosLinea = 2.5
               Printer.FontBold = True
               PrinterTexto 1.5, PosLinea, "CUOTA No."
               PrinterTexto 3.5, PosLinea, "VENCIMIENTO"
               PrinterTexto 7, PosLinea, "DIVIDENDO"
               PrinterTexto 10, PosLinea, "SALDO"
               PosLinea = PosLinea + 0.4
               Printer.FontBold = False
            End If
        End If
       .MoveNext
     Loop
 End If
End With
RatonNormal
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Private Sub Command6_Click()
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Tabla Cliente"
If BoxMensaje = vbYes Then
RatonReloj
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     MiFecha = FechaStrg(.Fields("Fecha"))
 End If
End With
InicioX = 0.5: InicioY = 0
sSQL = "SELECT * FROM Pagare "
sSQL = sSQL & "ORDER BY LN "
SelectAdodc AdoPagare, sSQL
DataAnchoCampos 1, AdoPagare, 8, TipoTimes, 1
Ancho(3) = 19
CantCampos = 3
Pagina = 1
Printer.FontBold = True
'Iniciamos la impresion
Printer.FontBold = False
If AdoTabla.Recordset.RecordCount > 0 Then
   PosLinea = 1
   Printer.FontSize = 14
   Printer.FontBold = True
   PrinterTexto 1, 0.3, Empresa
   Printer.FontSize = 12
   PrinterTexto 1, 1.1, "LIQUIDACION PRESTAMO DE AMORTIZACION GRADUAL"
   PrinterTexto 1, 1.6, "PAGOS MENSUALES"
   Printer.FontSize = 9
   PrinterTexto 1.5, 2.2, "Fecha de emisión:"
   PrinterTexto 1.5, 2.6, "Plazo:"
   PrinterTexto 1.5, 3, "Tasa:"
   PrinterTexto 1.5, 3.4, "Comisión:"
   PrinterTexto 1.5, 3.8, "Cliente:"
   PrinterTexto 1.5, 4.2, "Debito de la Libreta No."
   PrinterTexto 1.5, 4.6, "Garantes:"
   Printer.FontBold = False
   PrinterTexto 6, 2.2, FechaStrg(FechaSistema)
   If Si_No Then
      PrinterVariables 6, 2.6, TextMeses.Text & " día(s)"
   Else
      PrinterVariables 6, 2.6, TextMeses.Text & " meses"
   End If
   PrinterVariables 6, 3, Round(Val(TextInt.Text), 2) & "%"
   PrinterVariables 6, 3.4, "1.20% Sobre Saldos"
   PrinterVariables 6, 3.8, TxtNombresS.Text
   PrinterVariables 6, 4.2, Cuenta_No
   PosLinea = 4.6
   With AdoGarantes.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Cta = .Fields("Nombres")
           If Cta <> "" Then
              PrinterVariables 6, PosLinea, Cta
              PosLinea = PosLinea + 0.4
           End If
          .MoveNext
        Loop
    End If
   End With
   PosLinea = PosLinea + 0.2
   Printer.FontBold = True
   Printer.FontSize = 12
   PrinterTexto 1.5, PosLinea, UCase(TextTP.Text)
   PosLinea = PosLinea + 0.7
   Printer.FontSize = 10
   PrinterTexto 1.5, PosLinea, "Valor a Financiar:"
   Total = CDbl(TextMonto.Text)
   PrinterVariables 6, PosLinea, Total
   PosLinea = PosLinea + 0.55
End If
Printer.FontSize = 9
Printer.FontBold = True
PrinterTexto 1.5, PosLinea, "CUOTA No."
PrinterTexto 3.5, PosLinea, "VENCIMIENTO"
PrinterTexto 6.5, PosLinea, "AMORTIZACION"
PrinterTexto 9.5, PosLinea, "INTERES"
PrinterTexto 12, PosLinea, "COMISION"
PrinterTexto 14.5, PosLinea, "DIVIDENDO"
PrinterTexto 17, PosLinea, "S A L D O"
PosLinea = PosLinea + 0.4
Printer.FontBold = False
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Valor = 0
     Do While Not .EOF
        If .Fields("Mes_No") <> 0 Then
            PrinterFields 1.5, PosLinea, .Fields("Mes_No")
            PrinterTexto 3.8, PosLinea, FechaStrgCorta(.Fields("Fecha"))
            PrinterFields 6.5, PosLinea, .Fields("Capital")
            PrinterFields 9, PosLinea, .Fields("Interes")
            PrinterFields 11.5, PosLinea, .Fields("Comision")
            PrinterFields 14, PosLinea, .Fields("Pagos")
            PrinterFields 16.5, PosLinea, .Fields("Saldo")
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto - 1 Then
               Printer.NewPage
               PosLinea = 2.5
               Printer.FontBold = True
               PrinterTexto 1.5, PosLinea, "CUOTA No."
               PrinterTexto 3.5, PosLinea, "VENCIMIENTO"
               PrinterTexto 6.5, PosLinea, "AMORTIZACION"
               PrinterTexto 9.5, PosLinea, "INTERES"
               PrinterTexto 12, PosLinea, "COMISION"
               PrinterTexto 14.5, PosLinea, "DIVIDENDO"
               PrinterTexto 17, PosLinea, "S A L D O"
               PosLinea = PosLinea + 0.4
               Printer.FontBold = False
            End If
        End If
       .MoveNext
     Loop
 End If
End With
PosLinea = PosLinea + 0.1
Printer.Line (1.5, PosLinea)-(18.5, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Cadena = "Declaro que he revisado la presente liquidación, el valor de los impuestos y las comisiones por Servicios " _
       & "Fiancieros del crédito que he solicitado, así como la tabla de Amortización. Todo lo cual acepto de manera " _
       & "expresa, el valor total del crédito es de USD " & Round(Total, 2) & ", la tasa nominal del crédito es del " _
       & Round(Val(TextInt.Text), 2) & "% " _
       & "Anual, también acepto que el monto expresado será el mínimo, para prepago mas los intereses y comisión diferida " _
       & "por vencer de la siguiente cuota al igual que para la renovación del crédito si los pagos se encuentran al día."
Printer.FontSize = 9
PrinterLineas 1.5, PosLinea, Cadena, 17
PosLinea = PosLinea + 1.5
If PosLinea > LimiteAlto - 1 Then
   Printer.NewPage
   PosLinea = 2.5
End If
Printer.Line (1.5, PosLinea)-(8, PosLinea), QBColor(0)
Printer.Line (9, PosLinea)-(16, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterVariables 1.5, PosLinea, TxtNombresS.Text
PrinterVariables 10, PosLinea, "AUTORIZADO POR"
PosLinea = PosLinea + 0.1
RatonNormal
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Private Sub DCTipoPrestamo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyE And CtrlDown Then
     Codigo = SinEspaciosIzq(DCTipoPrestamo.Text)
     Contrato_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 2)
     Cuenta_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 3)
     Titulo = "ELIMINACION"
     Mensajes = "Seguro de Eliminar " & Chr(13) _
              & "Crédito No. " & Contrato_No & Chr(13) _
              & "Tipo = " & Codigo & Chr(13) _
              & "Cuenta No. " & Cuenta_No & Chr(13)
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * FROM Prestamos " _
             & "WHERE T = 'N' " _
             & "AND TP = '" & Codigo & "' " _
             & "AND Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        ConectarAdoExecute sSQL
        sSQL = "DELETE * FROM Garantes " _
             & "WHERE TP = '" & Codigo & "' " _
             & "AND Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        ConectarAdoExecute sSQL
        sSQL = "SELECT TP & '  ' & Credito_No & '  ' & Cuenta_No As TipoP " _
             & "FROM Prestamos " _
             & "WHERE T = 'N' " _
             & "ORDER BY TP,Credito_No "
        SelectDBCombo DCTipoPrestamo, AdoTipoPrest, sSQL, "TipoP", False
        DCTipoPrestamo.SetFocus
     End If
  End If
End Sub

Private Sub Form_Activate()
   Trans_No = 51
   MBoxFechaI.Text = FechaSistema
   sSQL = "SELECT TP & '  ' & Credito_No & '  ' & Cuenta_No As TipoP " _
        & "FROM Prestamos " _
        & "WHERE T = 'N' " _
        & "ORDER BY TP,Credito_No "
   SelectDBCombo DCTipoPrestamo, AdoTipoPrest, sSQL, "TipoP", False
   MiFecha = BuscarFecha(FechaSistema)
   TipoDoc = CompDiario
   IniciarAsientosDe DGAsiento, AdoAsiento
   RatonNormal
   DCTipoPrestamo.SetFocus
End Sub

Private Sub Form_Load()
  'CentrarForm Aprobacion
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoTabla
   ConectarAdodc AdoPagare
   ConectarAdodc AdoGarantes
   ConectarAdodc AdoConyugue
   ConectarAdodc AdoTipoPrest
   ConectarAdodc AdoAsiento
   TextImpuesto.Text = "0.00"
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub TextImpuesto_GotFocus()
   MarcarTexto TextImpuesto
End Sub

Private Sub TextImpuesto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextImpuesto_LostFocus()
  TextoValido TextImpuesto, True
  FechaValida MBoxFechaI
  Codigo = SinEspaciosIzq(DCTipoPrestamo.Text)
  Contrato_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 2)
  Cuenta_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 3)
  ListarCuenta Cuenta_No
  TipoDoc = SinEspaciosIzq(TextTP.Text)
  GenerarTablaPrestamo MBoxFechaI.Text, AdoTabla, DGTabla, TextInt, TextMeses, TextMonto, Si_No, Codigo
  IniciarAsientosDe DGAsiento, AdoAsiento
  Total = Round(CCur(TextMonto.Text), 2)
  Numero = Round(CCur(TextMeses.Text), 2)
  Interes = Round(CCur(TextInt.Text) / 100, 4)
  Trans_No = 51
  sSQL = "SELECT * " _
       & "FROM Asiento_De_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY T_No "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
    Do While Not .EOF
       Select Case .Fields("T_No")
         Case 1
              Debe = Total
              InsertarAsientos AdoAsiento, Cta_Prestamos, 0, Debe, 0
         Case 2
              Numero = Val(TextMeses.Text)
              If Si_No Then
                 Haber = Round((Total * 0.01 * Numero) / 360, 2)
              Else
                 If Numero > 12 Then Numero = 12
                 Haber = Round((Total * 0.01 * Numero * 30) / 360, 2)
              End If
              InsertarAsientos AdoAsiento, .Fields("Codigo"), 0, 0, Haber
         Case 3
              Numero = Val(TextMeses.Text)
              If Si_No Then
                 If Numero <= 30 Then
                    Haber = Round(Total * 0.03, 2)
                 Else
                    Haber = Round(Total * 0.04, 2)
                 End If
              Else
                 If Numero <= 12 Then
                    Haber = Round(Total * 0.03, 2)
                 Else
                    Haber = Round(Total * 0.05, 2)
                 End If
              End If
              InsertarAsientos AdoAsiento, .Fields("Codigo"), 0, 0, Haber
         Case 4
              Haber = CDbl(TextImpuesto.Text)
              InsertarAsientos AdoAsiento, Cta_Impuestos, 0, 0, Haber
         Case 5
              If Si_No Then
                 Haber = Round(((Total * Interes) / 360) * (NoDias + 3), 2)
                 InsertarAsientos AdoAsiento, Cta_Provision, 0, 0, Haber
              End If
       End Select
      .MoveNext
    Loop
  End With
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  TotalLibreta = Round(Debe - Haber, 2)
  InsertarAsientos AdoAsiento, Cta_Libretas, 0, 0, TotalLibreta
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  Cadena = "(" & NumEmpresa & ") Por préstamo No " & Contrato_No & ", Otorgado al Sr.(A) " & TxtNombresS.Text
  If Si_No Then
     Cadena = Cadena & ", Cta No. " & Cuenta_No & ", Plazo " & TextMeses.Text & " día(s), Taza " & TextInt.Text & "%"
  Else
     Cadena = Cadena & ", Cta No. " & Cuenta_No & ", Plazo " & TextMeses.Text & " mes(es), Taza " & TextInt.Text & "%"
  End If
  TextConcepto.Text = Cadena
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  TxtNombresS.SetFocus
End Sub

Private Sub TxtRazonSocial_LostFocus()
  TextoValido TxtRazonSocial, , True
End Sub
