VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{64AED23E-31A2-4023-8C7D-E628B15843D8}#1.0#0"; "Code39X.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FAbonos 
   Caption         =   "CIERRE DE CAJA"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin Code39X.Code39Clt Code39Clt1 
      Left            =   315
      Top             =   6720
      _ExtentX        =   1905
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command10 
      Caption         =   "S.R.I"
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
      Left            =   6510
      Picture         =   "FAbonos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   945
      Width           =   960
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
      Height          =   855
      Left            =   8610
      Picture         =   "FAbonos.frx":0882
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Reactivar"
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
      Left            =   5460
      Picture         =   "FAbonos.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Asiento"
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
      Left            =   4410
      Picture         =   "FAbonos.frx":158E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "D&iario"
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
      Left            =   3360
      Picture         =   "FAbonos.frx":1E58
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cuadre de Caja"
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
      Left            =   2310
      Picture         =   "FAbonos.frx":2162
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Left            =   1260
      Picture         =   "FAbonos.frx":25A4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Diario &Caja"
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
      Left            =   210
      Picture         =   "FAbonos.frx":29E6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   945
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   105
      TabIndex        =   14
      Top             =   525
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&1.- VENTAS"
      TabPicture(0)   =   "FAbonos.frx":2E28
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelAbonos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "AdoVentas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DGVentas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&2.- ABONOS"
      TabPicture(1)   =   "FAbonos.frx":2E44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGCxC"
      Tab(1).Control(1)=   "AdoCxC"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "LabelCheque"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&3.- INVENTARIO"
      TabPicture(2)   =   "FAbonos.frx":2E60
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGInv"
      Tab(2).Control(1)=   "DGProductos"
      Tab(2).Control(2)=   "DGCierres"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&4.- CONTABILIDAD"
      TabPicture(3)   =   "FAbonos.frx":2E7C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGAsiento"
      Tab(3).Control(1)=   "DGAsiento1"
      Tab(3).Control(2)=   "Label15"
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(4)=   "LblDiferencia1"
      Tab(3).Control(5)=   "LabelDebe1"
      Tab(3).Control(6)=   "LabelHaber1"
      Tab(3).Control(7)=   "LblConcepto1"
      Tab(3).Control(8)=   "LblConcepto"
      Tab(3).Control(9)=   "LabelHaber"
      Tab(3).Control(10)=   "LabelDebe"
      Tab(3).Control(11)=   "LblDiferencia"
      Tab(3).Control(12)=   "Label1"
      Tab(3).Control(13)=   "Label11"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "&5.- ANULADAS"
      TabPicture(4)   =   "FAbonos.frx":2E98
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command9"
      Tab(4).Control(1)=   "DGFactAnul"
      Tab(4).Control(2)=   "Label3"
      Tab(4).Control(3)=   "LblTotAnuladas"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "&6.- REPORTE DE AUDITORIA"
      TabPicture(5)   =   "FAbonos.frx":2EB4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DGSRI"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "AdoSRI"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label9"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label12"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label14"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label18"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "LblConIVA"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "LblSinIVA"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "LblDescuento"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "LblIVA"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label7"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "LblServicio"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label16"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "LblTotalFacturado"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).ControlCount=   14
      Begin VB.CommandButton Command8 
         Caption         =   "I.E.S.S."
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
         Left            =   7455
         Picture         =   "FAbonos.frx":2ED0
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   420
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "A&nuladas"
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
         Left            =   -67545
         Picture         =   "FAbonos.frx":379A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   420
         Width           =   960
      End
      Begin MSDataGridLib.DataGrid DGInv 
         Bindings        =   "FAbonos.frx":4064
         Height          =   2010
         Left            =   -73005
         TabIndex        =   16
         Top             =   1365
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   3545
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "FAbonos.frx":4079
         Height          =   1380
         Left            =   -74895
         TabIndex        =   17
         Top             =   1680
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   2434
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGCxC 
         Bindings        =   "FAbonos.frx":4092
         Height          =   4110
         Left            =   -74895
         TabIndex        =   18
         ToolTipText     =   "<Ctrl+P> Protestar Cheques"
         Top             =   1365
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7250
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSDataGridLib.DataGrid DGVentas 
         Bindings        =   "FAbonos.frx":40A7
         Height          =   4215
         Left            =   105
         TabIndex        =   19
         Top             =   1365
         Width           =   14160
         _ExtentX        =   24977
         _ExtentY        =   7435
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGFactAnul 
         Bindings        =   "FAbonos.frx":40BF
         Height          =   4110
         Left            =   -74895
         TabIndex        =   20
         Top             =   1365
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7250
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGSRI 
         Bindings        =   "FAbonos.frx":40D9
         Height          =   3375
         Left            =   -74895
         TabIndex        =   21
         Top             =   1365
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5953
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc AdoVentas 
         Height          =   330
         Left            =   9555
         Top             =   945
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
         Caption         =   "Ventas"
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
      Begin MSAdodcLib.Adodc AdoSRI 
         Height          =   330
         Left            =   -65445
         Top             =   945
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
         Caption         =   "SRI"
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
      Begin MSDataGridLib.DataGrid DGProductos 
         Bindings        =   "FAbonos.frx":40EE
         Height          =   2115
         Left            =   -73005
         TabIndex        =   22
         Top             =   3360
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   3731
         _Version        =   393216
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
      Begin MSAdodcLib.Adodc AdoCxC 
         Height          =   330
         Left            =   -65445
         Top             =   945
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
         Caption         =   "CxC"
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
      Begin MSDataGridLib.DataGrid DGAsiento1 
         Bindings        =   "FAbonos.frx":4109
         Height          =   1695
         Left            =   -74895
         TabIndex        =   47
         Top             =   3675
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   2990
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGCierres 
         Bindings        =   "FAbonos.frx":4123
         Height          =   4320
         Left            =   -74895
         TabIndex        =   54
         Top             =   1365
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   7620
         _Version        =   393216
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES "
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
         Left            =   -67965
         TabIndex        =   53
         Top             =   5355
         Width           =   1065
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia "
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
         Left            =   -70800
         TabIndex        =   52
         Top             =   5355
         Width           =   1065
      End
      Begin VB.Label LblDiferencia1 
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
         Left            =   -69750
         TabIndex        =   51
         Top             =   5355
         Width           =   1800
      End
      Begin VB.Label LabelDebe1 
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
         Left            =   -66915
         TabIndex        =   50
         Top             =   5355
         Width           =   1800
      End
      Begin VB.Label LabelHaber1 
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
         Left            =   -65130
         TabIndex        =   49
         Top             =   5355
         Width           =   1800
      End
      Begin VB.Label LblConcepto1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   48
         Top             =   3360
         Width           =   11040
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL"
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
         Left            =   -65445
         TabIndex        =   46
         Top             =   420
         Width           =   750
      End
      Begin VB.Label LabelCheque 
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
         Left            =   -64710
         TabIndex        =   45
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CON I.V.A."
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
         Left            =   -74895
         TabIndex        =   44
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SIN I.V.A."
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
         Left            =   -73110
         TabIndex        =   43
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESCUENTO"
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
         Left            =   -71325
         TabIndex        =   42
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL  I.V.A."
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
         Left            =   -69540
         TabIndex        =   41
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label LblConIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -74895
         TabIndex        =   40
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label LblSinIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -73110
         TabIndex        =   39
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label LblDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -71325
         TabIndex        =   38
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label LblIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -69540
         TabIndex        =   37
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label LabelAbonos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10290
         TabIndex        =   36
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL"
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
         TabIndex        =   35
         Top             =   420
         Width           =   750
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   34
         Top             =   1365
         Width           =   11040
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
         Height          =   330
         Left            =   -65130
         TabIndex        =   33
         Top             =   3045
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
         Height          =   330
         Left            =   -66915
         TabIndex        =   32
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label LblDiferencia 
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
         Left            =   -69750
         TabIndex        =   31
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia "
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
         Left            =   -70800
         TabIndex        =   30
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES "
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
         Left            =   -67965
         TabIndex        =   29
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Facturas Anuladas "
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
         Left            =   -74895
         TabIndex        =   28
         Top             =   5460
         Width           =   2325
      End
      Begin VB.Label LblTotAnuladas 
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
         Left            =   -72585
         TabIndex        =   27
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL  SERVICIO"
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
         Left            =   -67755
         TabIndex        =   26
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label LblServicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -67755
         TabIndex        =   25
         Top             =   5460
         Width           =   1800
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " T O T A L"
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
         Left            =   -65970
         TabIndex        =   24
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Label LblTotalFacturado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -65970
         TabIndex        =   23
         Top             =   5460
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   210
      Top             =   4830
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoSQL 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoVentaAct 
      Height          =   330
      Left            =   210
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
      Caption         =   "VentaAct"
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
   Begin MSAdodcLib.Adodc AdoInv1 
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
      Caption         =   "Inv1"
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
   Begin MSAdodcLib.Adodc AdoFactAnul 
      Height          =   330
      Left            =   210
      Top             =   4515
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
      Caption         =   "FactAnul"
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   210
      Top             =   2625
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
      Caption         =   "Productos"
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
   Begin VB.CheckBox CheqOrdDep 
      Caption         =   "Ordenar Por Depósito"
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
      Left            =   12180
      TabIndex        =   13
      Top             =   105
      Width           =   2220
   End
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "FAbonos.frx":413C
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   5775
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin VB.CheckBox CheqCajero 
      Caption         =   "Por Cajero"
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
      Left            =   4410
      TabIndex        =   11
      Top             =   105
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3045
      TabIndex        =   2
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
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
      Left            =   1785
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoCxC1 
      Height          =   330
      Left            =   210
      Top             =   5460
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
      Caption         =   "CxC1"
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
   Begin MSAdodcLib.Adodc AdoAsiento1 
      Height          =   330
      Left            =   210
      Top             =   5145
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
      Caption         =   "Asiento1"
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
   Begin MSAdodcLib.Adodc AdoCierres 
      Height          =   330
      Left            =   210
      Top             =   5775
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
      Caption         =   "Cierres"
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
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Periodo de Cierre"
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
      Width           =   1695
   End
End
Attribute VB_Name = "FAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ErrorInventario As String
Dim ErrorFacturas As String
Dim CtasProc() As CtasAsiento
Dim ContCtas As Integer
Dim Combos As String
Dim NumTrans As Long
Dim FormaCierre As Boolean

Private Sub CheqCajero_Click()
  If CheqCajero.value = 1 Then DCBenef.Visible = True Else DCBenef.Visible = False
End Sub

'Cierre diario de Caja y asientos contables
Private Sub Command1_Click()
  RatonReloj
  Presentar_Inventario = False
  TextoImprimio = ""
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  If Inv_Promedio Then
     FAbonos.Caption = "CIERRE DE CAJA INVENTARIO PRECIO PROMEDIO"
  Else
     FAbonos.Caption = "CIERRE DE CAJA INVENTARIO ULTIMO PRECIO"
  End If
  MayorizarInv.Show 1
  RatonReloj
  Actualizar_Ejecutivo_Facturas FechaIni, FechaFin, False
  RatonReloj
  Actualizar_CI_RUC_SRI MBFechaI, MBFechaF
  RatonReloj
  
  sSQL = "SELECT Fecha " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TC IN('C','P','OP') " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "UNION " _
       & "SELECT Fecha " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TP IN('C','P','OP') " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "GROUP BY Fecha " _
       & "ORDER BY Fecha "
  SelectDataGrid DGCierres, AdoCierres, sSQL
  
  RatonReloj
  DGCierres.Caption = "Dias Cierres"
  sSQL = "UPDATE Detalle_Factura " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET X = 'P' " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "SET DF.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET Cta_Venta = CP.Cta_Ventas " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "SET DF.Cta_Venta = CP.Cta_Ventas "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.Total_IVA <> 0 " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET Cta_Venta = CP.Cta_Ventas_0 " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "SET DF.Cta_Venta = CP.Cta_Ventas_0 "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.Total_IVA = 0 " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv "
  ConectarAdoExecute sSQL
  
  RatonReloj
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET Gramaje = (DF.Cantidad * (CP.Gramaje/1000)) " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "SET DF.Gramaje = (DF.Cantidad * (CP.Gramaje/1000)) "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND CP.TC = 'P' " _
       & "AND CP.Gramaje > 0 " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND CP.INV <> 0 " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv "
  ConectarAdoExecute sSQL
    
  sSQL = "SELECT * " _
       & "FROM Detalle_Factura " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND X = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Verifique el Producto de la(s) siguiente(s) Factura(s): " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio _
                        & .Fields("Fecha") & " " & .Fields("TC") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Producto: [" & .Fields("Codigo") & "]" & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  
  RatonReloj
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Detalle_Factura " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ConectarAdoExecute sSQL
    
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ConectarAdoExecute sSQL
    
  RatonReloj
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET X = 'P' " _
          & "FROM Detalle_Factura As DF,Clientes As C "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Clientes As C " _
          & "SET DF.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.CodigoC = C.Codigo "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET X = 'P' " _
          & "FROM Facturas As DF,Clientes As C "
  Else
     sSQL = "UPDATE Facturas As DF,Clientes As C " _
          & "SET DF.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT DF.TC IN('C','P','OP') " _
       & "AND DF.CodigoC = C.Codigo "
  ConectarAdoExecute sSQL
    
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = 'P' " _
          & "FROM Trans_Abonos As DF,Clientes As C "
  Else
     sSQL = "UPDATE Trans_Abonos As DF,Clientes As C " _
          & "SET DF.X = 'P' "
  End If
  sSQL = sSQL _
       & "WHERE DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT DF.TP IN('C','P','OP') " _
       & "AND DF.CodigoC = C.Codigo "
  ConectarAdoExecute sSQL
    
  RatonReloj
  sSQL = "UPDATE Facturas " _
       & "SET CodigoC = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT TC IN('C','P','OP') " _
       & "AND X = '.' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Detalle_Factura " _
       & "SET CodigoC = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT TC IN('C','P','OP') " _
       & "AND X = '.' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Trans_Abonos " _
       & "SET CodigoC = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT TP IN('C','P','OP') " _
       & "AND X = '.' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND NOT TC IN('C','P','OP') " _
       & "AND CodigoC = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Factura(s) sin Codigo de Cliente: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Fecha") & " " & .Fields("TC") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Cta: " & .Fields("Cta_CxP") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Detalle_Factura " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND CodigoC = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Detalle de Factura(s) sin codigo de Cliente: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Fecha") & " " & .Fields("TC") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Código Producto: " & .Fields("Codigo") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND CodigoC = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Abonos de Factura(s) sin codigo de Cliente: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Fecha") & " " & .Fields("TP") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Código Producto: " & .Fields("CodigoC") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
      
  RatonReloj
  
 'Eliminar sub ctas
  sSQL = "DELETE * " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '.' " _
       & "AND TP = '.' " _
       & "AND Numero = 0 "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET Base_Imponible = F.SubTotal,Porc = (Abono/F.SubTotal)*100 " _
          & "FROM Trans_Abonos As TA,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
          & "SET TA.Base_Imponible = F.SubTotal,Porc = (TA.Abono/F.SubTotal)*100 "
  End If
  sSQL = sSQL _
       & "WHERE TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TA.Banco,1,16) = 'RETENCION FUENTE' " _
       & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND F.SubTotal > 0 " _
       & "AND TA.Base_Imponible <= 0 " _
       & "AND TA.TP = F.TC " _
       & "AND TA.Item = F.Item " _
       & "AND TA.Periodo = F.Periodo " _
       & "AND TA.Factura = F.Factura " _
       & "AND TA.CodigoC = F.CodigoC "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET Base_Imponible = F.IVA,Porc = (Abono/F.IVA)*100 " _
          & "FROM Trans_Abonos As TA,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
          & "SET TA.Base_Imponible = F.IVA,Porc = (TA.Abono/F.IVA)*100 "
  End If
  sSQL = sSQL _
       & "WHERE TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TA.Banco,1,13) = 'RETENCION IVA' " _
       & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TA.Base_Imponible <= 0 " _
       & "AND F.IVA > 0 " _
       & "AND TA.TP = F.TC " _
       & "AND TA.Item = F.Item " _
       & "AND TA.Periodo = F.Periodo " _
       & "AND TA.Factura = F.Factura " _
       & "AND TA.CodigoC = F.CodigoC "
  ConectarAdoExecute sSQL
  
 'Verificacion de Cuentas Contables en Facturas
  RatonReloj
  ErrorFacturas = ""
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET X = 'X' " _
          & "FROM Facturas As T,Catalogo_Cuentas As C "
  Else
     sSQL = "UPDATE Facturas As T,Catalogo_Cuentas As C " _
          & "SET X = 'X' "
  End If
  sSQL = sSQL _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT T.TC IN('C','P','OP') " _
       & "AND T.Cta_CxP = C.Codigo " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = C.Periodo "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND X = '.' " _
       & "AND NOT TC IN('C','P','OP') " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Verifique las siguiente(s) Factura(s) de CxC no constan en el Catálogo: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Fecha") & " " & .Fields("TC") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Cta: " & .Fields("Cta_CxP") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  
 'Verificacion de Cuentas Contables en Abonos en CxC
  RatonReloj
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' "
  ConectarAdoExecute sSQL

  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = 'X' " _
          & "FROM Trans_Abonos As T,Catalogo_Cuentas As C "
  Else
     sSQL = "UPDATE Trans_Abonos As T,Catalogo_Cuentas As C " _
          & "SET X = 'X' "
  End If
  sSQL = sSQL _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Cta_CxP = C.Codigo " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = C.Periodo "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND X = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Abonos de las Facturas en CxC no existe en el Catálogo de Cuentas: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Autorizacion") & " " & .Fields("Fecha") & " " & .Fields("TP") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Cta: " & .Fields("Cta_CxP") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
 'Verificacion de Cuentas Contables en Abonos en Cta
  RatonReloj
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' "
  ConectarAdoExecute sSQL
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = 'X' " _
          & "FROM Trans_Abonos As T,Catalogo_Cuentas As C "
  Else
     sSQL = "UPDATE Trans_Abonos As T,Catalogo_Cuentas As C " _
          & "SET X = 'X' "
  End If
  sSQL = sSQL _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Cta = C.Codigo " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = C.Periodo "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND X = '.' " _
       & "ORDER BY Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Abonos de las Facturas no constan en el Catálogo: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & .Fields("Autorizacion") & " " & .Fields("Fecha") & " " & .Fields("TP") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Cta: " & .Fields("Cta") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT TC,Serie,Factura,Autorizacion,Total_MN,ROUND(Con_IVA+Sin_IVA+IVA-Descuento-Descuento2,2,0) As SubTotales " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Total_MN-ROUND(Con_IVA+Sin_IVA+IVA-Descuento-Descuento2,2,0) <> 0 " _
       & "AND NOT TC IN('C','P','OP') " _
       & "ORDER BY TC,Serie,Factura,Autorizacion "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextoImprimio = TextoImprimio & "Verifique las siguientes Facturas/Notas de Venta: " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio & "Autorizacion: " & .Fields("Autorizacion") & " " & .Fields("TC") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Diferencia: " & .Fields("Total_MN") - .Fields("SubTotales") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
  ErrorInventario = ""
  
  RatonReloj
  Grabar_Asientos_Facturacion Normal
  RatonNormal
  If Redondear(Debe - Haber, 2) <> 0 Then
     TextoImprimio = TextoImprimio & "Las Transacciones no cuadran, verifique las facturas emitidas o los abonos del día." & vbCrLf
     Command1.SetFocus
  Else
     If Command2.Enabled Then Command2.SetFocus Else Command5.SetFocus
  End If
  If Len(TextoImprimio) > 1 Then FInfoError.Show
  Trans_No = 96
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY DEBE DESC,HABER "
  SelectDataGrid DGAsiento, AdoAsiento, SQL2
  Trans_No = 97
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY DEBE DESC,HABER "
  SelectDataGrid DGAsiento1, AdoAsiento1, SQL2
End Sub

Private Sub Command10_Click()
  DGSRI.Visible = False
  If MBFechaI.Text = MBFechaF.Text Then
     SQLMsg3 = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI.Text
  Else
     SQLMsg3 = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI.Text & " al " & MBFechaF.Text
  End If
  MensajeEncabData = "RESUMEN DE FACTURAS EMITIDAS"
  With AdoSRI.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       SQLMsg2 = "Facturas desde No. " & .Fields("Secuencial")
      .MoveLast
       SQLMsg2 = SQLMsg2 & " Hasta la No. " & .Fields("Secuencial")
   End If
  End With
  SQLMsg1 = "TIPO DE DOCUMENTO:  NOTAS DE VENTA"
  ImprimirAdo_SRI AdoSRI, 7
  DGSRI.Visible = True
End Sub

'''Private Sub Command12_Click()
'''Dim Fecha_a_Numero
'''Dim TBaseGravada As Currency
'''Dim TBaseCero As Currency
'''Dim TBaseSubTotal As Currency
'''Dim TTotal_IVA As Currency
'''
'''  FechaValida MBFechaI
'''  FechaValida MBFechaF
'''
'''  Trans_No = 96
'''  RatonReloj
'''  ProgBar.Min = 0
'''  ProgBar.Value = 0
'''  FechaIni = BuscarFecha(MBFechaI)
'''  FechaFin = BuscarFecha(MBFechaF)
'''  Numero = CLng(Val(Replace(Format$(MBFechaF, "yyyy/MM/dd"), "/", "")))
'''  SQL1 = "DELETE * " _
'''       & "FROM Trans_Ventas " _
'''       & "WHERE TP = 'CD' " _
'''       & "AND Numero = " & Numero & " " _
'''       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  ConectarAdoExecute SQL1
'''
'''  SQL1 = "DELETE * " _
'''       & "FROM Trans_Air " _
'''       & "WHERE TP = 'CD' " _
'''       & "AND Numero = " & Numero & " " _
'''       & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  ConectarAdoExecute SQL1
'''
''' 'MsgBox Numero
'''  Codigo1 = Mid$(SerieFactura, 1, 3)
'''  Codigo2 = Mid$(SerieFactura, 4, 3)
'''
'''  sSQL = "SELECT * " _
'''       & "FROM Clientes " _
'''       & "WHERE Codigo <> '.' " _
'''       & "ORDER BY Codigo "
'''  SelectAdodc AdoClientes, sSQL
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Desc_0 = (SELECT SUM(Total_Desc) " _
'''         & "              FROM Detalle_Factura As DF " _
'''         & "              WHERE DF.Total_IVA = 0 " _
'''         & "              AND DF.TC = Facturas.TC " _
'''         & "              AND DF.Item = Facturas.Item " _
'''         & "              AND DF.Periodo = Facturas.Periodo " _
'''         & "              AND DF.Fecha = Facturas.Fecha " _
'''         & "              AND DF.Factura = Facturas.Factura " _
'''         & "              AND DF.CodigoC = Facturas.CodigoC " _
'''         & "              AND DF.Serie = Facturas.Serie " _
'''         & "              AND DF.Autorizacion = Facturas.Autorizacion) " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TC NOT IN ('C','P','OP') " _
'''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''    ConectarAdoExecute sSQL
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Desc_X = (SELECT SUM(Total_Desc) " _
'''         & "              FROM Detalle_Factura As DF " _
'''         & "              WHERE DF.Total_IVA > 0 " _
'''         & "              AND DF.TC = Facturas.TC " _
'''         & "              AND DF.Item = Facturas.Item " _
'''         & "              AND DF.Periodo = Facturas.Periodo " _
'''         & "              AND DF.Fecha = Facturas.Fecha " _
'''         & "              AND DF.Factura = Facturas.Factura " _
'''         & "              AND DF.CodigoC = Facturas.CodigoC " _
'''         & "              AND DF.Serie = Facturas.Serie " _
'''         & "              AND DF.Autorizacion = Facturas.Autorizacion) " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TC NOT IN ('C','P','OP') " _
'''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''    ConectarAdoExecute sSQL
'''
'''  sSQL = "UPDATE Facturas " _
'''       & "SET Desc_0 = 0 " _
'''       & "WHERE Desc_0 IS NULL "
'''  ConectarAdoExecute sSQL
'''
'''  sSQL = "UPDATE Facturas " _
'''       & "SET Desc_X = 0 " _
'''       & "WHERE Desc_X IS NULL "
'''  ConectarAdoExecute sSQL
'''
'''  TBaseGravada = 0
'''  TBaseCero = 0
'''  TBaseSubTotal = 0
'''  TBaseDescuento = 0
'''  TTotal_IVA = 0
'''  TTotal = 0
'''  Cantidad = 0
'''  CodigoCliente = "9999999999"
'''  If Val(Codigo1) < 1 Then Codigo1 = "001"
'''  If Val(Codigo2) < 1 Then Codigo2 = "001"
'''  NumTrans = Maximo_De("Trans_Ventas", "ID")
'''  sSQL = "SELECT F.RUC_CI,F.TB," _
'''       & "COUNT(F.Factura) As CantFact," _
'''       & "SUM(F.Con_IVA-F.Desc_X) As BaseGravada," _
'''       & "SUM(F.Sin_IVA-F.Desc_0) As BaseCero," _
'''       & "SUM(F.SubTotal) As BaseSubTotal," _
'''       & "SUM(F.Descuento) As BaseDescuento," _
'''       & "SUM(F.IVA) As Total_IVA," _
'''       & "SUM(F.Total_MN) As Total " _
'''       & "FROM Facturas F,Clientes C,Accesos As A " _
'''       & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND F.TC NOT IN ('C','P','OP') " _
'''       & "AND F.T <> 'A' " _
'''       & "AND F.Item = '" & NumEmpresa & "' " _
'''       & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND F.CodigoC = C.Codigo " _
'''       & "AND F.Cod_Ejec = A.Codigo " _
'''       & "GROUP BY F.RUC_CI,F.TB " _
'''       & "ORDER BY F.RUC_CI,F.TB "
'''  SelectAdodc AdoAux, sSQL
'''  With AdoAux.Recordset
'''   If .RecordCount > 0 Then
'''       ProgBar.Max = .RecordCount + 5
'''       Do While Not .EOF
'''          CodigoCliente = .Fields("RUC_CI")
'''          Factura_No = .Fields("CantFact")
'''          If .Fields("RUC_CI") = "9999999999999" Then
'''              CodigoCliente = "9999999999"
'''          Else
'''             If AdoClientes.Recordset.RecordCount > 0 Then
'''                AdoClientes.Recordset.MoveFirst
'''                AdoClientes.Recordset.Find ("CI_RUC_SRI = '" & CodigoCliente & "' ")
'''                If AdoClientes.Recordset.EOF Then
'''                   CodigoCliente = "9999999999"
'''                Else
'''                   CodigoCliente = AdoClientes.Recordset.Fields("Codigo")
'''                End If
'''             End If
'''          End If
'''         'MsgBox CodigoCliente & vbCrLf & Factura_No & vbCrLf & .Fields("RUC_CI")
'''         'Insertamos Retenciones en la Fuente e IVA
'''         NumTrans = Maximo_De("Trans_Air", "ID")
'''         sSQL = "SELECT Banco,Porc,SUM(Base_Imponible) As BaseImp,SUM(Abono) As Abonos " _
'''              & "FROM Trans_Abonos " _
'''              & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''              & "AND T <> 'A' " _
'''              & "AND Mid$(Banco,1,16) = 'RETENCION FUENTE' " _
'''              & "AND Item = '" & NumEmpresa & "' " _
'''              & "AND Periodo = '" & Periodo_Contable & "' " _
'''              & "AND CodigoC = '" & CodigoCliente & "' " _
'''              & "GROUP BY Banco,Porc " _
'''              & "ORDER BY Banco,Porc "
'''         SelectAdodc AdoAux, sSQL
'''         If AdoAux.Recordset.RecordCount > 0 Then
'''            Total = 0
'''            Do While Not AdoAux.Recordset.EOF
'''               TipoDoc = SinEspaciosDer(AdoAux.Recordset.Fields("Banco"))
'''               Porc = AdoAux.Recordset.Fields("Porc")
'''               Total = AdoAux.Recordset.Fields("BaseImp")
'''               Insertar_Ventas_Air CodigoCliente, Total, Porc, AdoAux.Recordset.Fields("Abonos"), Factura_No, 0, "001", "001", String(10, "9"), Ninguno
'''               NumTrans = NumTrans + 1
'''               AdoAux.Recordset.MoveNext
'''            Loop
'''         End If
'''         Total_RetIVAB = 0
'''         Total_RetIVAS = 0
'''         PorcIVAB = 0
'''         PorcIVAS = 0
'''        'Calculamos los totales de las Retenciones por IVA Bienes
'''         sSQL = "SELECT Porc,SUM(Base_Imponible) As BaseImp,SUM(Abono) As Abonos " _
'''              & "FROM Trans_Abonos " _
'''              & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''              & "AND T <> 'A' " _
'''              & "AND Banco = 'RETENCION IVA BIENES' " _
'''              & "AND Item = '" & NumEmpresa & "' " _
'''              & "AND Periodo = '" & Periodo_Contable & "' " _
'''              & "AND CodigoC = '" & CodigoCliente & "' " _
'''              & "GROUP BY Porc " _
'''              & "ORDER BY Porc "
'''         SelectAdodc AdoAux, sSQL
'''         If AdoAux.Recordset.RecordCount > 0 Then
'''            Do While Not AdoAux.Recordset.EOF
'''               PorcIVAB = AdoAux.Recordset.Fields("Porc")
'''               Total_RetIVAB = Total_RetIVAB + AdoAux.Recordset.Fields("Abonos")
'''               AdoAux.Recordset.MoveNext
'''            Loop
'''         End If
'''        'Calculamos los totales de las Retenciones por IVA Servicios
'''         sSQL = "SELECT Porc,SUM(Base_Imponible) As BaseImp,SUM(Abono) As Abonos " _
'''              & "FROM Trans_Abonos " _
'''              & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''              & "AND T <> 'A' " _
'''              & "AND Banco = 'RETENCION IVA SERVICIO' " _
'''              & "AND Item = '" & NumEmpresa & "' " _
'''              & "AND Periodo = '" & Periodo_Contable & "' " _
'''              & "AND CodigoC = '" & CodigoCliente & "' " _
'''              & "GROUP BY Porc " _
'''              & "ORDER BY Porc "
'''         SelectAdodc AdoAux, sSQL
'''         If AdoAux.Recordset.RecordCount > 0 Then
'''            Do While Not AdoAux.Recordset.EOF
'''               PorcIVAS = AdoAux.Recordset.Fields("Porc")
'''               Total_RetIVAS = Total_RetIVAS + AdoAux.Recordset.Fields("Abonos")
'''               AdoAux.Recordset.MoveNext
'''            Loop
'''         End If
'''        'Insertamos el encabezado de las ventas por Cliente
'''         Total_RetIVAB = Redondear(Total_RetIVAB, 2)
'''         Total_RetIVAS = Redondear(Total_RetIVAS, 2)
'''         Insertar_Ventas CodigoCliente, .Fields("CantFact"), .Fields("BaseCero"), .Fields("BaseGravada"), .Fields("Total_IVA"), .Fields("BaseSubTotal"), "18"
'''        'Sumatoria de Totales
'''         TBaseGravada = TBaseGravada + .Fields("BaseGravada")
'''         TBaseCero = TBaseCero + .Fields("BaseCero")
'''         TBaseSubTotal = TBaseSubTotal + .Fields("BaseSubTotal")
'''         TTotal_IVA = TTotal_IVA + .Fields("Total_IVA")
'''         Cantidad = Cantidad + .Fields("CantFact")
'''         NumTrans = NumTrans + 1
'''         ProgBar.Value = ProgBar.Value + 1
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
''' 'MsgBox TBaseCero & vbCrLf & TBaseGravada & vbCrLf & TBaseSubTotal & vbCrLf & TTotal_IVA & vbCrLf & Cantidad
'''  sSQL = "SELECT C.CI_RUC_SRI,C.TD_SRI,TA.Cheque,SUM(TA.Abono) As Total_NC,COUNT(TA.Abono) As Cantidad " _
'''       & "FROM Trans_Abonos AS TA, Clientes As C " _
'''       & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND TA.T <> 'A' " _
'''       & "AND TA.Banco = 'NOTA DE CREDITO' " _
'''       & "AND TA.Item = '" & NumEmpresa & "' " _
'''       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TA.CodigoC = C.Codigo " _
'''       & "GROUP BY C.CI_RUC_SRI,C.TD_SRI,TA.Cheque " _
'''       & "ORDER BY C.CI_RUC_SRI,C.TD_SRI,TA.Cheque DESC "
'''  SelectAdodc AdoAux, sSQL
'''  Total = 0
'''  Cantidad = 0
'''  TBaseCero = 0
'''  TTotal_IVA = 0
'''  TBaseGravada = 0
'''  TBaseSubTotal = 0
'''  Total_RetIVAS = 0
'''  Total_RetIVAB = 0
'''  With AdoAux.Recordset
'''   If .RecordCount > 0 Then
'''       CodigoCli = .Fields("CI_RUC_SRI")
'''       Do While Not .EOF
'''          If CodigoCli <> .Fields("CI_RUC_SRI") Then
'''             If TTotal_IVA > 0 Then
'''                TBaseGravada = Total
'''             Else
'''                TBaseCero = Total
'''             End If
'''             If CodigoCli = "9999999999999" Then
'''                CodigoCliente = "9999999999"
'''             Else
'''                If AdoClientes.Recordset.RecordCount > 0 Then
'''                   AdoClientes.Recordset.MoveFirst
'''                   AdoClientes.Recordset.Find ("CI_RUC_SRI = '" & CodigoCli & "' ")
'''                   If AdoClientes.Recordset.EOF Then
'''                      CodigoCliente = "9999999999"
'''                   Else
'''                      CodigoCliente = AdoClientes.Recordset.Fields("Codigo")
'''                   End If
'''                End If
'''             End If
'''             Insertar_Ventas CodigoCliente, CLng(Cantidad), TBaseCero, TBaseGravada, TTotal_IVA, TBaseSubTotal, "4"
'''             CodigoCli = .Fields("CI_RUC_SRI")
'''             Total = 0
'''             Cantidad = 0
'''             TBaseCero = 0
'''             TBaseGravada = 0
'''             TBaseSubTotal = 0
'''             TTotal_IVA = 0
'''          End If
'''          Select Case .Fields("Cheque")
'''            Case "VENTAS"
'''                 Cantidad = Cantidad + .Fields("Cantidad")
'''                 Total = Total + .Fields("Total_NC")
'''                 TBaseSubTotal = TBaseSubTotal + .Fields("Total_NC")
'''            Case "I.V.A."
'''              TTotal_IVA = TTotal_IVA + .Fields("Total_NC")
'''          End Select
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  If TTotal_IVA > 0 Then
'''     TBaseGravada = Total
'''  Else
'''     TBaseCero = Total
'''  End If
'''  If CodigoCli = "9999999999999" Then
'''     CodigoCliente = "9999999999"
'''  Else
'''     If AdoClientes.Recordset.RecordCount > 0 Then
'''        AdoClientes.Recordset.MoveFirst
'''        AdoClientes.Recordset.Find ("CI_RUC_SRI = '" & CodigoCli & "' ")
'''        If AdoClientes.Recordset.EOF Then
'''           CodigoCliente = "9999999999"
'''        Else
'''           CodigoCliente = AdoClientes.Recordset.Fields("Codigo")
'''        End If
'''     End If
'''  End If
'''  Insertar_Ventas CodigoCliente, CLng(Cantidad), TBaseCero, TBaseGravada, TTotal_IVA, TBaseSubTotal, "4"
'''  NumTrans = NumTrans + 1
'''  ProgBar.Value = ProgBar.Max
'''  RatonNormal
'''  MsgBox "Proceso de Asignación del AT en Ventas Terminado"
'''End Sub

'Grabacion de los comprobantes contables
Private Sub Command2_Click()
   FechaValida MBFechaI
   FechaValida MBFechaF
   FechaTexto = MBFechaF.Text
   FechaComp = FechaTexto
   Nombre_Cajero = Ninguno
   If CheqCajero.value = 1 Then
      Nombre_Cajero = Mid$(DCBenef.Text, 1, Len(DCBenef.Text) - Len(SinEspaciosDer(DCBenef.Text)) - 1)
   End If
   If MBFechaI.Text = MBFechaF.Text Then
      Cadena = "Cierre de Caja del " & MBFechaI.Text
   Else
      Cadena = "Cierre de Caja del " & MBFechaI.Text & " al " & MBFechaF.Text
   End If
  'Verificamos partida doble de los dos asientos
   Debe = 0: Haber = 0
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .Fields("DEBE")
           Haber = Haber + .Fields("HABER")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   With AdoAsiento1.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .Fields("DEBE")
           Haber = Haber + .Fields("HABER")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   LabelDebe.Caption = Format$(Debe, "#,##0.00")
   LabelHaber.Caption = Format$(Haber, "#,##0.00")
   LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
   If ((NuevoDiario) And Redondear(Debe - Haber, 2) = 0) Then
       FechaTexto = MBFechaF
       FechaComp = FechaTexto
       NumComp = ReadSetDataNum("Diario", True, False)
       Mensajes = "Esta seguro de Grabar el Cierre de Caja"
       Titulo = "Pregunta de grabación"
       If BoxMensaje = vbYes Then
          RatonReloj
          FechaTexto = MBFechaF.Text
          FechaIni = BuscarFecha(MBFechaF.Text)
          DiarioCaja = NumComp
          FechaTexto = MBFechaF.Text
          If FormaCierre Then
             Imprimir_Diario_Caja AdoVentas, AdoCxC, AdoInv, AdoProductos, MBFechaI, MBFechaF
          Else
             Imprimir_Diario_Caja_Resumen AdoVentas, AdoCxC, AdoInv, AdoProductos, MBFechaI, MBFechaF
          End If
         'Grabacion del Comprobante de CxC
          If AdoAsiento1.Recordset.RecordCount > 0 Then
             Trans_No = 97
             NumComp = ReadSetDataNum("Diario", True, True)
             Co.T = Normal
             Co.TP = CompDiario
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             If MBFechaI.Text = MBFechaF.Text Then
                Co.Concepto = "Cierre de Caja de Cuentas por Cobrar del " & MBFechaI.Text & ", Diario No. " & NumComp
             Else
                Co.Concepto = "Cierre de Caja de Cuentas por Cobrar del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. " & NumComp
             End If
             Co.CodigoB = Ninguno
             Co.Efectivo = 0
             Co.Monto_Total = Debe
             Co.T_No = Trans_No
             Co.Usuario = CodigoUsuario
             Co.Item = NumEmpresa
             GrabarComprobante Co, AdoAsiento1, Code39Clt1
             Control_Procesos Normal, Co.Concepto
             ImprimirComprobantesDe False, Co
             IniciarAsientosDe DGAsiento1, AdoAsiento1
          End If
         'Grabacion del Comprobante de Abonos
          If AdoAsiento.Recordset.RecordCount > 0 Then
             Trans_No = 96
             NumComp = ReadSetDataNum("Diario", True, True)
             Co.T = Normal
             Co.TP = CompDiario
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             If MBFechaI.Text = MBFechaF.Text Then
                Co.Concepto = "Cierre de Caja de Abonos del " & MBFechaI.Text & ", Diario No. " & NumComp
             Else
                Co.Concepto = "Cierre de Caja de Abonos del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. " & NumComp
             End If
             Co.CodigoB = Ninguno
             Co.Efectivo = 0
             Co.Monto_Total = Debe
             Co.T_No = Trans_No
             Co.Usuario = CodigoUsuario
             Co.Item = NumEmpresa
             GrabarComprobante Co, AdoAsiento, Code39Clt1
             Control_Procesos Normal, Co.Concepto
             ImprimirComprobantesDe False, Co
             IniciarAsientosDe DGAsiento, AdoAsiento
            'Los Asientos de SubModulos
             sSQL = "UPDATE Trans_SubCtas " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '.' " _
                  & "AND Numero = 0 " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
             ConectarAdoExecute sSQL
          End If
          LabelDebe.Caption = Format$(0, "#,##0.00")
          LabelHaber.Caption = Format$(0, "#,##0.00")
          RatonNormal
          Mifecha = BuscarFecha(FechaSistema)
          FechaIni = BuscarFecha(MBFechaI.Text)
          FechaFin = BuscarFecha(MBFechaF.Text)
          sSQL = "UPDATE Trans_Abonos " _
               & "SET C = " & Val(adTrue) & " " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha <= #" & FechaFin & "# "
          ConectarAdoExecute sSQL
         'MsgBox sSQL
          sSQL = "UPDATE Facturas " _
               & "SET C = " & Val(adTrue) & " " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha <= #" & FechaFin & "# "
          ConectarAdoExecute sSQL
          CierreDelDia
        End If
   Else
       RatonNormal
       MsgBox "Ya esta cerrado este día o " _
            & "no hay datos que procesar"
   End If
End Sub

Private Sub Command3_Click()
  Unload FAbonos
End Sub

Private Sub Command4_Click()
   RatonReloj
   FCuadreCaja.Show 1
'''   GrabarAsientosFacturacion Procesado
'''   LabelDebe.Caption = Format$(Debe, "#,##0.00")
'''   LabelHaber.Caption = Format$(Haber, "#,##0.00")
'''   RatonNormal
'''   If Debe <> Haber Then
'''      MsgBox "Las Transacciones no cuadran," & Chr(13) & "verifique de que anulo las facturas correctas"
'''      Command1.SetFocus
'''   Else
'''      Command2.SetFocus
'''   End If
End Sub

Private Sub Command5_Click()
  Nombre_Cajero = Ninguno
  If CheqCajero.value = 1 Then
     Nombre_Cajero = Mid$(DCBenef.Text, 1, Len(DCBenef.Text) - Len(SinEspaciosDer(DCBenef.Text)) - 1)
  End If
  'MsgBox FormaCierre
  If FormaCierre Then
     Imprimir_Diario_Caja AdoVentas, AdoCxC, AdoInv, AdoProductos, MBFechaI, MBFechaF
  Else
     Imprimir_Diario_Caja_Resumen AdoVentas, AdoCxC, AdoInv, AdoProductos, MBFechaI, MBFechaF
  End If
End Sub

Private Sub Command6_Click()
  DGAsiento.Visible = False
  MensajeEncabData = "RESUMEN DE VENTAS"
  SQLMsg1 = "Corte del " & MBFechaI.Text & " al " & MBFechaF.Text
  sSQL = "SELECT CODIGO,CUENTA,PARCIAL_ME,DEBE,HABER " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectAdodc AdoAsiento, sSQL
  
  ImprimirResumenAsientoCaja AdoAsiento
  
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGAsiento, AdoAsiento, sSQL
  DGAsiento.Visible = True
End Sub

Private Sub Command7_Click()
  If ClaveContador Then
     FechaValida MBFechaI
     FechaValida MBFechaF
     FechaIni = BuscarFecha(MBFechaI)
     FechaFin = BuscarFecha(MBFechaF)
     sSQL = "UPDATE Trans_Abonos " _
          & "SET C = " & Val(adFalse) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     ConectarAdoExecute sSQL

     sSQL = "UPDATE Facturas " _
          & "SET C = " & Val(adFalse) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     ConectarAdoExecute sSQL
     Trans_No = 97
     IniciarAsientosDe DGAsiento1, AdoAsiento1
     Trans_No = 96
     IniciarAsientosDe DGAsiento, AdoAsiento
     RatonNormal
     LabelDebe.Caption = Format$(0, "#,##0.00")
     LabelHaber.Caption = Format$(0, "#,##0.00")
     CierreDelDia
     MBFechaI.SetFocus
  End If
End Sub

Private Sub Command8_Click()
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CI_RUCC As String
Dim NombreC As String
   RatonReloj
    sSQL = "SELECT * " _
         & "FROM Clientes " _
         & "WHERE Codigo <> '.' " _
         & "ORDER BY Cliente "
    SelectAdodc AdoClientes, sSQL
   
   FechaIni = BuscarFecha(MBFechaI.Text)
   FechaFin = BuscarFecha(MBFechaF.Text)
   RutaGeneraFile = Left(RutaSysBases, 2) & "\SYSBASES\ARCHIVO_" & Replace(MBFechaI, "/", "-") & ".txt"
   NumFile = FreeFile
   Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.

   sSQL = "SELECT DF.Factura,DF.Fecha,DF.Cantidad,DF.Precio,DF.Precio2,CP.Producto," _
        & "C.Cliente,C.CI_RUC,DF.CodigoB,CP.Codigo_IESS,CP.Marca " _
        & "FROM Detalle_Factura As DF,Clientes As C,Catalogo_Productos As CP " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND DF.T <> 'A' " _
        & "AND DF.CodigoC = C.Codigo " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CP.Periodo " _
        & "ORDER BY DF.Fecha,DF.Factura "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           CI_RUCC = .Fields("CI_RUC")
           NombreC = .Fields("Cliente")
           Producto = .Fields("Producto") & " (" & .Fields("Marca") & ")"
           If AdoClientes.Recordset.RecordCount > 0 Then
              AdoClientes.Recordset.MoveFirst
              AdoClientes.Recordset.Find ("Codigo = '" & .Fields("CodigoB") & "' ")
              If Not AdoClientes.Recordset.EOF Then
                 CI_RUCC = AdoClientes.Recordset.Fields("CI_RUC")
                 NombreC = AdoClientes.Recordset.Fields("Cliente")
              End If
           End If
           Print #NumFile, Format$(Val(.Fields("CI_RUC")), "0000000000");
           Print #NumFile, .Fields("Cliente") & String(80 - Len(.Fields("Cliente")), " ");
           Print #NumFile, CI_RUCC;
           Print #NumFile, NombreC & String(64 - Len(NombreC), " ");
           Print #NumFile, Trim$(.Fields("Fecha"));
           Print #NumFile, .Fields("Codigo_IESS") & String(40 - Len(.Fields("Codigo_IESS")), " ");
           Print #NumFile, "      ";
           Producto = Mid$(Producto, 1, 80)
           Producto = Replace(Producto, "/", " ")
           Producto = Trim$(Producto)
           Print #NumFile, Producto & String(80 - Len(Producto), " ");
           Cadena = Format$(.Fields("Cantidad"), "0.00")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(13 - Len(Cadena), "0") & Cadena;
           Cadena = Format$(.Fields("Precio"), "0.00")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(18 - Len(Cadena), "0") & Cadena;
           Cadena = Format$(.Fields("Precio2"), "0.00")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(15 - Len(Cadena), "0") & Cadena;
           Print #NumFile, Format$(.Fields("Factura"), "000000000")
          .MoveNext
        Loop
    End If
   End With
   Close #NumFile
   RatonNormal
   MsgBox "ARCHIVO GENERADO EN:" & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Private Sub Command9_Click()
  MensajeEncabData = "FACTURAS ANULADAS"
  If MBFechaI.Text = MBFechaF.Text Then
     SQLMsg3 = "Diario de Caja del " & MBFechaI
  Else
     SQLMsg3 = "Diario de Caja del " & MBFechaI & " al " & MBFechaF
  End If
  ImprimirAdo AdoFactAnul, True, 1, 8
End Sub

Private Sub DCBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGAsiento.Visible = False
     GenerarDataTexto FAbonos, AdoAsiento
     DGAsiento.Visible = True
  End If
End Sub

Private Sub DGAsiento1_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGAsiento1.Visible = False
     GenerarDataTexto FAbonos, AdoAsiento1
     DGAsiento1.Visible = True
  End If
End Sub

Private Sub DGCierres_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGCierres.Visible = False
     GenerarDataTexto FAbonos, AdoCierres
     DGCierres.Visible = True
  End If
End Sub

Private Sub DGCxC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGCxC.Visible = False
     GenerarDataTexto FAbonos, AdoCxC
     DGCxC.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     TextoBanco = DGCxC.Columns(4)
     TextoCheque = DGCxC.Columns(5)
     Mifecha = DGCxC.Columns(1)
     Factura_No = Val(DGCxC.Columns(3))
     Valor = Val(DGCxC.Columns(6))
     Cta = DGCxC.Columns(8)
     If TextoBanco <> "EFECTIVO MN" Then
        Mensajes = "Cheque del: " & TextoBanco & " No. " & TextoCheque & vbCrLf _
                 & "Fecha del Cheque: " & Mifecha & vbCrLf _
                 & "Factura No. " & Factura_No & vbCrLf _
                 & "Valor USD " & Format$(Valor, "#,##0.00")
        Titulo = "CHEQUES PROTESTADOS"
        If BoxMensaje = vbYes Then
           sSQL = "UPDATE Trans_Abonos " _
                & "SET Protestado = " & Val(adTrue) & " " _
                & "WHERE Fecha = #" & BuscarFecha(Mifecha) & "# " _
                & "AND Cta = '" & Cta & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND Banco = '" & TextoBanco & "' " _
                & "AND Cheque = '" & TextoCheque & "' "
           ConectarAdoExecute sSQL
        End If
     Else
        MsgBox "No se puede protestar Abonos en Efectivo"
     End If
  End If
End Sub

Private Sub DGFactAnul_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGFactAnul.Visible = False
     GenerarDataTexto FAbonos, AdoFactAnul
     DGFactAnul.Visible = True
  End If
End Sub

Private Sub DGInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGInv.Visible = False
     GenerarDataTexto FAbonos, AdoInv
     DGInv.Visible = True
  End If
End Sub

Private Sub DGProductos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGSRI.Visible = False
     GenerarDataTexto FAbonos, AdoProductos
     DGSRI.Visible = True
  End If
End Sub

Private Sub DGSRI_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGSRI.Visible = False
     GenerarDataTexto FAbonos, AdoSRI
     DGSRI.Visible = True
  End If
End Sub

Private Sub DGVentas_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGVentas.Visible = False
     GenerarDataTexto FAbonos, AdoVentas
     DGVentas.Visible = True
  End If
End Sub

Private Sub Form_Activate()
   FAbonos.WindowState = vbMaximized
   sSQL = "UPDATE Accesos " _
        & "SET Ok = " & Val(adFalse) & " "
   ConectarAdoExecute sSQL
   
   If SQL_Server Then
      sSQL = "UPDATE Accesos " _
           & "SET Ok = " & Val(adTrue) & " " _
           & "FROM Accesos As A,Facturas As X "
   Else
      sSQL = "UPDATE Accesos As A,Facturas As X " _
           & "SET Ok = " & Val(adTrue) & " "
   End If
   sSQL = sSQL & "WHERE A.Codigo = X.CodigoU "
   ConectarAdoExecute sSQL
   
   If SQL_Server Then
      sSQL = "UPDATE Accesos " _
           & "SET Ok = " & Val(adTrue) & " " _
           & "FROM Accesos As A,Trans_Abonos As X "
   Else
      sSQL = "UPDATE Accesos As A,Trans_Abonos As X " _
           & "SET Ok = " & Val(adTrue) & " "
   End If
   sSQL = sSQL & "WHERE A.Codigo = X.CodigoU "
   ConectarAdoExecute sSQL

   sSQL = "SELECT (Nombre_Completo & ' - ' & Codigo) As Cajero " _
        & "FROM Accesos " _
        & "WHERE Ok <> " & Val(adFalse) & " " _
        & "ORDER BY Nombre_Completo "
   SelectDBCombo DCBenef, AdoClientes, sSQL, "Cajero"
   
   FormaCierre = Leer_Campo_Empresa("Cierre_Vertical")
   

   SSTab1.Height = MDI_Y_Max - 900
   SSTab1.width = MDI_X_Max - 100
   SSTab1.Tab = 5
   DGSRI.width = SSTab1.width - 200
   DGSRI.Height = SSTab1.Height - DGSRI.Top - 1000
   SSTab1.Tab = 4
   DGFactAnul.width = SSTab1.width - 200
   DGFactAnul.Height = SSTab1.Height - DGFactAnul.Top - 550
   SSTab1.Tab = 3
   DGAsiento.width = SSTab1.width - 200
   DGAsiento.Height = (SSTab1.Height / 2) - DGAsiento.Top
   
   Label1.Top = DGAsiento.Top + DGAsiento.Height + 10
   Label11.Top = DGAsiento.Top + DGAsiento.Height + 10
   LblDiferencia.Top = DGAsiento.Top + DGAsiento.Height + 10
   LabelDebe.Top = DGAsiento.Top + DGAsiento.Height + 10
   LabelHaber.Top = DGAsiento.Top + DGAsiento.Height + 10
   
   LblConcepto1.width = SSTab1.width - 200
   LblConcepto1.Top = Label1.Top + Label1.Height + 10
   
   DGAsiento1.width = SSTab1.width - 200
   DGAsiento1.Top = LblConcepto1.Top + LblConcepto1.Height + 10
   DGAsiento1.Height = SSTab1.Height - LblConcepto1.Top - LblConcepto1.Height - Label13.Height - Label13.Height
   SSTab1.Tab = 2
   DGInv.width = SSTab1.width - DGInv.Left - 200
   DGInv.Height = (SSTab1.Height / 2) - DGInv.Top
   
   DGProductos.Top = DGInv.Top + DGInv.Height
   DGProductos.width = SSTab1.width - DGProductos.Left - 200
   DGProductos.Height = SSTab1.Height - DGProductos.Top - 200
   DGCierres.Height = SSTab1.Height - DGCierres.Top - 200
   SSTab1.Tab = 1
   DGCxC.width = SSTab1.width - 200
   DGCxC.Height = SSTab1.Height - DGCxC.Top - 100
   SSTab1.Tab = 0
   DGVentas.width = SSTab1.width - 200
   DGVentas.Height = SSTab1.Height - DGVentas.Top - 100
   
'''   AdoSRI.Width = SSTab1.Width - AdoSRI.Left - 200
'''   AdoCxC.Width = SSTab1.Width - AdoCxC.Left - 200
'''   AdoVentas.Width = SSTab1.Width - AdoVentas.Left - 200
   Label7.Top = SSTab1.Height - SSTab1.Top - 340
   Label9.Top = SSTab1.Height - SSTab1.Top - 340
   Label12.Top = SSTab1.Height - SSTab1.Top - 340
   Label14.Top = SSTab1.Height - SSTab1.Top - 340
   Label16.Top = SSTab1.Height - SSTab1.Top - 340
   Label18.Top = SSTab1.Height - SSTab1.Top - 340
   
   LblConIVA.Top = SSTab1.Height - SSTab1.Top
   LblSinIVA.Top = SSTab1.Height - SSTab1.Top
   LblDescuento.Top = SSTab1.Height - SSTab1.Top
   LblIVA.Top = SSTab1.Height - SSTab1.Top
   LblServicio.Top = SSTab1.Height - SSTab1.Top
   LblTotalFacturado.Top = SSTab1.Height - SSTab1.Top

   LblConcepto.width = SSTab1.width - 200
   Label13.Top = DGAsiento1.Top + DGAsiento1.Height + 10
   Label15.Top = DGAsiento1.Top + DGAsiento1.Height + 10
   LblDiferencia1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
   LabelDebe1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
   LabelHaber1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
   Label3.Top = SSTab1.Height - SSTab1.Top + 10
   
   LblTotAnuladas.Top = SSTab1.Height - SSTab1.Top
   
   Select Case Modulo
     Case "CONTABILIDAD": Command2.Enabled = False
     Case "CAJACREDITO": Command2.Enabled = False
   End Select
   If Inv_Promedio Then
      FAbonos.Caption = "CIERRE DE CAJA INVENTARIO PRECIO PROMEDIO"
   Else
      FAbonos.Caption = "CIERRE DE CAJA INVENTARIO ULTIMO PRECIO"
   End If
   NuevoDiario = False
   
  'IniciarAsientosDe DGAsiento, AdoAsiento
   
   Mifecha = BuscarFecha(FechaSistema)
   RatonNormal
   CierreDelDia
End Sub

Private Sub Form_Deactivate()
  FAbonos.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoSRI
   ConectarAdodc AdoSQL
   ConectarAdodc AdoCxC
   ConectarAdodc AdoCxC1
   ConectarAdodc AdoInv
   ConectarAdodc AdoInv1
   ConectarAdodc AdoCierres
   ConectarAdodc AdoVentas
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoAsiento1
   ConectarAdodc AdoClientes
   ConectarAdodc AdoVentaAct
   ConectarAdodc AdoFactAnul
   ConectarAdodc AdoProductos
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF = MBFechaI
  'LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI) & " al " & FechaStrgDias(MBFechaF)
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyF10 Then
     FechaIni = BuscarFecha(MBFechaI.Text)
     FechaFin = BuscarFecha(MBFechaF.Text)
     sSQL = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND (Total_MN * 0.12) <> IVA " _
          & "AND F.TC NOT IN ('C','P') " _
          & "AND F.Periodo = '" & Periodo_Contable & "' "
    SelectAdodc AdoVentas, sSQL
  End If
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
 'LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI.Text) & " al " & FechaStrgDias(MBFechaF.Text)
End Sub

Public Sub Grabar_Asientos_Facturacion(TipoConsulta As String)
Dim VentasDia As Boolean
Dim Ctas_Catalogo As String
Dim ErrorTemp As String
Dim Total_Vaucher As Currency
Dim ContSC As Integer

   Ctas_Catalogo = ""
   Beneficiario = Ninguno
   DGCxC.Visible = False
   DGInv.Visible = False
   DGVentas.Visible = False
   DGAsiento.Visible = False
   FechaValida MBFechaI
   FechaValida MBFechaF
   ErrorInventario = ""
   Total_Vaucher = 0
   VentasDia = False
   RatonReloj
   FechaIni = BuscarFecha(MBFechaI)
   FechaFin = BuscarFecha(MBFechaF)
   Fecha_Vence = MBFechaF
  'MsgBox sSQL
   FAbonos.Caption = "Verificando Cuentas involucradas"
   sSQL = "SELECT Codigo_Inv,Cta_Inventario,Cta_Ventas,Cta_Ventas_0," _
        & "Cta_Costo_Venta,Cta_Venta_Anticipada,Cta_Ventas_Ant,COUNT(DF.Codigo) " _
        & "FROM Catalogo_Productos As CP,Detalle_Factura As DF " _
        & "WHERE CP.Item = '" & NumEmpresa & "' " _
        & "AND CP.Periodo = '" & Periodo_Contable & "' " _
        & "AND CP.TC = 'P' " _
        & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND CP.Codigo_Inv = DF.Codigo " _
        & "AND CP.Item = DF.Item " _
        & "AND CP.Periodo = DF.Periodo " _
        & "GROUP BY Codigo_Inv,Cta_Inventario,Cta_Ventas,Cta_Ventas_0," _
        & "Cta_Costo_Venta,Cta_Venta_Anticipada,Cta_Ventas_Ant " _
        & "ORDER BY Cta_Inventario,Codigo_Inv "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           ErrorTemp = ""
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Inventario"))
           If .Fields("Cta_Inventario") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Inv: " & .Fields("Cta_Inventario") & vbCrLf
           End If
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Ventas"))
           If .Fields("Cta_Ventas") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Ventas: " & .Fields("Cta_Ventas") & vbCrLf
           End If
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Ventas_0"))
           If .Fields("Cta_Ventas_0") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Ventas 0: " & .Fields("Cta_Ventas_0") & vbCrLf
           End If
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Costo_Venta"))
           If .Fields("Cta_Costo_Venta") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Costo Ventas: " & .Fields("Cta_Costo_Venta") & vbCrLf
           End If
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Venta_Anticipada"))
           If .Fields("Cta_Venta_Anticipada") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Venta Anticipada: " & .Fields("Cta_Venta_Anticipada") & vbCrLf
           End If
           Codigo = Leer_Cta_Catalogo(.Fields("Cta_Ventas_Ant"))
           If .Fields("Cta_Ventas_Ant") <> "0" And Codigo = Ninguno Then
               ErrorTemp = ErrorTemp & "Cta Venta Año Anterior: " & .Fields("Cta_Ventas_Ant") & vbCrLf
           End If
           If ErrorTemp <> "" Then
              TextoImprimio = TextoImprimio & vbCrLf & "Verifique las cuentas asignadas del Cod. Inv: " & .Fields("Codigo_Inv") & vbCrLf & ErrorTemp
           End If
          .MoveNext
        Loop
    End If
   End With
   
   RatonReloj
   Combos = Ninguno
   FechaIni = BuscarFecha(MBFechaI)
   FechaFin = BuscarFecha(MBFechaF)
   FechaFinal = BuscarFecha("31/12/" & FechaAnio(MBFechaF))
   FAbonos.Caption = "Verificando Detalle de Productos"
   If SQL_Server Then
      sSQL = "UPDATE Detalle_Factura " _
           & "SET T = F.T " _
           & "FROM Detalle_Factura As TA,Facturas As F "
   Else
      sSQL = "UPDATE Detalle_Factura As TA,Facturas As F " _
           & "SET TA.T = F.T "
   End If
   sSQL = sSQL & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC NOT IN ('C','P','OP') " _
        & "AND TA.Factura = F.Factura " _
        & "AND TA.Serie = F.Serie " _
        & "AND TA.Autorizacion = F.Autorizacion " _
        & "AND TA.Periodo = F.Periodo " _
        & "AND TA.Item = F.Item " _
        & "AND TA.TC = F.TC "
   ConectarAdoExecute sSQL
   ContCtas = 0
   Codigo1 = Ninguno
   sSQL = "SELECT Codigo " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND CxC <> '.' " _
        & "AND TL <> " & Val(adFalse) & " " _
        & "ORDER BY TL DESC,Codigo "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Codigo1 = .Fields("Codigo")
      ' Facturas
        sSQL = "UPDATE Facturas " _
             & "SET Cod_CxC = '" & Codigo1 & "' " _
             & "WHERE Cod_CxC = '" & Ninguno & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        ConectarAdoExecute sSQL
      ' Detalle Facturas
        sSQL = "UPDATE Detalle_Factura " _
             & "SET CodigoL = '" & Codigo1 & "' " _
             & "WHERE CodigoL = '" & Ninguno & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        ConectarAdoExecute sSQL
    End If
   End With
   
   sSQL = "SELECT Cta_CxP " _
        & "FROM Facturas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta_CxP <> '.' " _
        & "GROUP BY Cta_CxP "
   SelectAdodc AdoAux, sSQL
   ContCtas = AdoAux.Recordset.RecordCount
   
  'Presentamos las Ventas si manejamos una sola cuenta
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(Cta_Venta) > 1 " _
        & "AND TL <> " & Val(adFalse) & " " _
        & "ORDER BY TL DESC,Codigo "
   SelectAdodc AdoCxC1, sSQL
   If AdoCxC1.Recordset.RecordCount > 0 Then UnaSolaCtaVenta = True

   ContCtas = ContCtas + AdoCxC1.Recordset.RecordCount
   
   sSQL = "SELECT Cta " _
        & "FROM Trans_Abonos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta <> '.' " _
        & "GROUP BY Cta "
   SelectAdodc AdoSQL, sSQL
   ContCtas = ContCtas + AdoSQL.Recordset.RecordCount
   
   sSQL = "SELECT Cta_Inventario,Cta_Costo_Venta,Cta_Ventas,Cta_Ventas_0,Cta_Ventas_Ant,Cta_Venta_Anticipada " _
        & "FROM Catalogo_Productos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "GROUP BY Cta_Inventario,Cta_Costo_Venta,Cta_Ventas,Cta_Ventas_0,Cta_Ventas_Ant,Cta_Venta_Anticipada "
   SelectAdodc AdoSQL, sSQL
   
   ContCtas = ContCtas + (AdoSQL.Recordset.RecordCount * 6) + 3
   
   ReDim CtasProc(ContCtas) As CtasAsiento
   For IE = 0 To ContCtas - 1
       CtasProc(IE).Cta = "0"
       CtasProc(IE).Valor = 0
   Next IE
  'Cuentas de CxC
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta_CxP")
          .MoveNext
        Loop
    End If
   End With
   With AdoSQL.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta_Ventas")
           SetearCtasCierre .Fields("Cta_Ventas_0")
           SetearCtasCierre .Fields("Cta_Ventas_Ant")
           SetearCtasCierre .Fields("Cta_Inventario")
           SetearCtasCierre .Fields("Cta_Costo_Venta")
           SetearCtasCierre .Fields("Cta_Venta_Anticipada")
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT Cta " _
        & "FROM Trans_Abonos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta <> '.' " _
        & "GROUP BY Cta "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta")
          .MoveNext
        Loop
    End If
   End With
   SetearCtasCierre Cta_IVA
   SetearCtasCierre Cta_Desc
   SetearCtasCierre Cta_Desc2
   SetearCtasCierre Cta_Servicio
   SetearCtasCierre Cta_Tarjetas
   SetearCtasCierre Cta_Caja_Vaucher
  'Cuentas de Ventas
   With AdoCxC1.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .Fields("Cta_Venta")
          .MoveNext
        Loop
    End If
   End With
   Cta = Leer_Seteos_Ctas("Cta_CxP_NC")
   SetearCtasCierre Cta
   Cta = Leer_Seteos_Ctas("Cta_Gasto_Bancario")
   SetearCtasCierre Cta
   Total = 0
   Select Case TipoConsulta
     Case Procesado: NuevoDiario = False
     Case Normal:    NuevoDiario = True
   End Select
  'TextoImprimio
   Cadena = ""
   For IE = 0 To ContCtas - 1
      'Cadena = Cadena & CtasProc(IE).Cta & vbCrLf
       Cta = CtasProc(IE).Cta
       If Len(Cta) > 1 Then
          LeerCta Cta
          If Codigo = Ninguno Then Cadena = Cadena & CtasProc(IE).Cta & vbCrLf
       End If
   Next IE
   
   If Cadena <> "" Then TextoImprimio = TextoImprimio & "Estas Cuentas no constan en el catalogo: " & vbCrLf & Cadena
 ' ================================
 ' Iniciamos los asientos contables
 ' ================================
   ContSC = 0
   RatonReloj
   Trans_No = 97
   IniciarAsientosDe DGAsiento1, AdoAsiento1
   Trans_No = 96
   IniciarAsientosDe DGAsiento, AdoAsiento
 ' Ventas Anticipadas
   FAbonos.Caption = "Totalizando Ventas Anticipadas"
    sSQL = "SELECT P.Cta,P.TP,P.Fecha,AVG(P.Pagos) As Valor_Ed " _
        & "FROM Prestamos As P,Detalle_Factura As F " _
        & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & "" & FechaFin & "# " _
        & "AND P.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.T <> 'A' " _
        & "AND P.Pagos > 0 " _
        & "AND P.Cuenta_No = F.CodigoC " _
        & "AND P.Item = F.Item "
    If CheqCajero.value = 1 Then
       sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
       Beneficiario = SinEspaciosIzq(DCBenef.Text)
    End If
    sSQL = sSQL & "GROUP BY P.Cta,P.TP,P.Fecha "
   SelectAdodc AdoVentaAct, sSQL
''  'CxP Credito
''   sSQL = "SELECT TA.TP,TA.Fecha,C.Cliente,TA.Serie,TA.Autorizacion,TA.Factura,TA.Banco,TA.Cheque,TA.Abono,TA.Comprobante,TA.Cta,TA.Cta_CxP,CC.TC As TCS,TA.CodigoC " _
''        & "FROM Trans_Abonos As TA,Clientes C,Catalogo_Cuentas As CC " _
''        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''        & "AND TA.TP NOT IN ('C','P','OP') " _
''        & "AND TA.T <> 'A' " _
''        & "AND TA.Item = '" & NumEmpresa & "' " _
''        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
''        & "AND TA.CodigoC = C.Codigo " _
''        & "AND TA.Item = CC.Item " _
''        & "AND TA.Periodo = CC.Periodo " _
''        & "AND TA.Cta_CxP = CC.Codigo "
''   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TA.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
''   If CheqOrdDep.value = 1 Then
''      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,TA.Banco,C.Cliente,TA.Factura "
''   Else
''      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,C.Cliente,TA.Banco,TA.Factura "
''   End If
''   'MsgBox CheqOrdDep.Value
''   SelectDataGrid DGCxC, AdoCxC, sSQL
''   With AdoCxC.Recordset
''
''    If .RecordCount > 0 Then
''        Do While Not .EOF
''           InsValorCta .Fields("Cta"), .Fields("Abono")
''           InsValorCta .Fields("Cta_CxP"), -.Fields("Abono")
''          'Verificamos si es cta de submodulos
''           Select Case .Fields("TCS")
''             Case "C", "P"
''                  SetAdoAddNew "Asiento_SC"
''                  SetAdoFields "Codigo", .Fields("CodigoC")
''                  SetAdoFields "Beneficiario", .Fields("Cliente")
''                  SetAdoFields "TM", "1"
''                  SetAdoFields "DH", "2"
''                  SetAdoFields "Valor", Redondear(.Fields("Abono"), 2)
''                  SetAdoFields "FECHA_V", .Fields("Fecha")
''                  SetAdoFields "TC", .Fields("TCS")
''                  SetAdoFields "Cta", .Fields("Cta")
''                  SetAdoFields "T_No", Trans_No
''                  SetAdoFields "SC_No", ContSC
''                  SetAdoFields "Item", NumEmpresa
''                  SetAdoFields "CodigoU", CodigoUsuario
''                  SetAdoUpdate
''                  ContSC = ContSC + 1
''           End Select
''           Total = Total + Redondear(.Fields("Abono"), 2)
''          .MoveNext
''        Loop
''    End If
''   End With
''   LabelCheque.Caption = Format$(Total, "#,##0.00")
''
   Total = 0
  'Asientos de CxC Cheque
   FAbonos.Caption = "Totalizando Abonos"
   sSQL = "SELECT TA.TP,TA.Fecha,C.Cliente,TA.Serie,TA.Autorizacion,TA.Factura,TA.Banco,TA.Cheque,TA.Abono,TA.Comprobante,TA.Cta,TA.Cta_CxP,CC.TC As TCS,TA.CodigoC " _
        & "FROM Trans_Abonos As TA,Clientes C,Catalogo_Cuentas As CC " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.TP NOT IN ('C','P','OP') " _
        & "AND TA.T <> 'A' " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.CodigoC = C.Codigo " _
        & "AND TA.Item = CC.Item " _
        & "AND TA.Periodo = CC.Periodo " _
        & "AND TA.Cta = CC.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TA.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   If CheqOrdDep.value = 1 Then
      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,TA.Banco,C.Cliente,TA.Factura "
   Else
      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,C.Cliente,TA.Banco,TA.Factura "
   End If
   'MsgBox CheqOrdDep.Value
   SelectDataGrid DGCxC, AdoCxC, sSQL
   With AdoCxC.Recordset
        
    If .RecordCount > 0 Then
        Do While Not .EOF
           InsValorCta .Fields("Cta"), .Fields("Abono")
           InsValorCta .Fields("Cta_CxP"), -.Fields("Abono")
          'Verificamos si es cta de submodulos
           Select Case .Fields("TCS")
             Case "C", "P"
                  SetAdoAddNew "Asiento_SC"
                  SetAdoFields "Codigo", .Fields("CodigoC")
                  SetAdoFields "Beneficiario", .Fields("Cliente")
                  SetAdoFields "DH", "1"
                  SetAdoFields "Valor", Redondear(.Fields("Abono"), 2)
                  SetAdoFields "FECHA_V", .Fields("Fecha")
                  SetAdoFields "TC", .Fields("TCS")
                  SetAdoFields "Cta", .Fields("Cta")
                  SetAdoFields "TM", "1"
                  SetAdoFields "T_No", Trans_No
                  SetAdoFields "SC_No", ContSC
                  SetAdoFields "Item", NumEmpresa
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
                  ContSC = ContSC + 1
           End Select
           Total = Total + Redondear(.Fields("Abono"), 2)
          .MoveNext
        Loop
    End If
   End With
   LabelCheque.Caption = Format$(Total, "#,##0.00")
   Total = 0
   For IE = 0 To ContCtas - 1
    If CtasProc(IE).Valor >= 0 Then
       InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, CtasProc(IE).Valor, 0
    Else
       InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, 0, -CtasProc(IE).Valor
    End If
   Next IE
  FAbonos.Caption = "Procesando Asientos Tarjetas Crédito"
  sSQL = "SELECT TA.Cta, CC.Cuenta, SUM(TA.Abono) As Total_TJ " _
       & "FROM Trans_Abonos As TA,Catalogo_Cuentas As CC " _
       & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND CC.TC = 'TJ' " _
       & "AND TA.Cta = CC.Codigo " _
       & "AND TA.Periodo = CC.Periodo " _
       & "AND TA.Item = CC.Item " _
       & "GROUP BY TA.Cta, CC.Cuenta " _
       & "ORDER BY TA.Cta, CC.Cuenta "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Total_Vaucher = 0
       Do While Not .EOF
          Total_Vaucher = Total_Vaucher + .Fields("Total_TJ")
         .MoveNext
       Loop
       If Total_Vaucher > 0 Then
          InsertarAsientos AdoAsiento, Cta_Tarjetas, 0, Total_Vaucher, 0
          InsertarAsientos AdoAsiento, Cta_Caja_Vaucher, 0, 0, Total_Vaucher
          'MsgBox Cta_Caja_Vaucher
       End If
   End If
  End With
  'Enceramos para realizar la segunda parte del cierre
   Cadena = ""
   For IE = 0 To ContCtas - 1
       If CtasProc(IE).Valor <> 0 Then Cadena = Cadena & CtasProc(IE).Cta & " = " & CtasProc(IE).Valor & vbCrLf
       CtasProc(IE).Valor = 0
   Next IE
   
   'MsgBox Cadena
   Trans_No = 97
  'Asientos de CxC Efectivo
   FAbonos.Caption = "Totalizando Ventas"
   sSQL = "SELECT F.TC,F.Fecha,C.Cliente,F.Serie,F.Autorizacion,F.Factura,F.IVA As Total_IVA,F.Descuento,F.Descuento2,F.Servicio,F.Total_MN,F.Saldo_MN,F.Cta_CxP,A.Nombre_Completo As Ejecutivo " _
        & "FROM Facturas F,Clientes C,Accesos As A " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC NOT IN ('C','P','OP') " _
        & "AND F.T <> 'A' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.Cod_Ejec = A.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL & "ORDER BY F.Fecha,F.TC,F.Cta_CxP,C.Cliente,F.Factura "
   SelectDataGrid DGVentas, AdoVentas, sSQL
   With AdoVentas.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           InsValorCta Cta_IVA, -.Fields("Total_IVA")
           InsValorCta Cta_Desc, .Fields("Descuento")
           InsValorCta Cta_Desc2, .Fields("Descuento2")
           InsValorCta Cta_Servicio, -.Fields("Servicio")
           InsValorCta .Fields("Cta_CxP"), .Fields("Total_MN")
           Total = Total + Redondear(.Fields("Total_MN"), 2)
          .MoveNext
        Loop
    End If
   End With
   LabelAbonos.Caption = Format$(Total, "#,##0.00")
   FAbonos.Caption = "Totalizando Costeo de Inventario"
  'Abrimos espacios para el asiento
   sSQL = "SELECT * " _
        & "FROM Asiento_K " _
        & "WHERE CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' "
   SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|."
   SelectDataGrid DGInv, AdoInv, sSQL, SQLDec
   
  'Asiento de Salida de Inventario y Ventas del dia de una sola cuenta
     sSQL = "SELECT DF.CodigoL,DF.TC,DF.Codigo,DF.Cantidad,DF.Precio,DF.Total,DF.CodBodega,DF.Ticket,DF.Fecha," _
          & "A.Cta_Inventario,A.Cta_Costo_Venta,A.Cta_Ventas,A.Cta_Ventas_0,A.Cta_Ventas_Ant," _
          & "A.Cta_Venta_Anticipada,A.Unidad,A.Producto,A.PVP " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As A " _
          & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.T <> '" & Anulado & "' " _
          & "AND DF.TC NOT IN ('C','P','OP') "
     If CheqCajero.value = 1 Then sSQL = sSQL & "AND DF.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
     sSQL = sSQL _
          & "AND DF.Codigo = A.Codigo_Inv " _
          & "AND DF.Item = A.Item " _
          & "AND DF.Periodo = A.Periodo " _
          & "ORDER BY DF.Codigo,DF.Fecha,DF.Precio "
     SelectAdodc AdoAux, sSQL
     Total = 0
     TotalIngreso = 0
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Entrada = 0
          Combos = Ninguno
          Precio = .Fields("Precio")
          CodigoInv = .Fields("Codigo")
          Producto = .Fields("Producto")
          Unidad = .Fields("Unidad")
          Cta_Inventario = .Fields("Cta_Inventario")
          Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
          Cta_Ventas = .Fields("Cta_Ventas")
          Cta_Ventas_0 = .Fields("Cta_Ventas_0")
          Cta_Ventas_Ant = .Fields("Cta_Ventas_Ant")
          TipoProc = .Fields("TC")
          Cta_Provision = .Fields("Cta_Venta_Anticipada")
          Cod_Bodega = .Fields("CodBodega")
          Total = 0
          If Cod_Bodega = "" Then Cod_Bodega = Ninguno
          If CodigoInv = "" Then CodigoInv = Ninguno
          Do While Not .EOF
             If AdoCxC1.Recordset.RecordCount > 0 Then
               'Cuando Usamos una sola cuenta de ventas
                AdoCxC1.Recordset.MoveFirst
                AdoCxC1.Recordset.Find ("Codigo = '" & .Fields("CodigoL") & "' ")
                If Not AdoCxC1.Recordset.EOF Then
                   If Len(AdoCxC1.Recordset.Fields("Cta_Venta")) > 1 Then
                      Cta_Ventas = AdoCxC1.Recordset.Fields("Cta_Venta")
                      Cta_Ventas_0 = AdoCxC1.Recordset.Fields("Cta_Venta")
                      If TipoProc = "NV" Then Cta_Ventas = Cta_Ventas_0
                      If IsNumeric(.Fields("Ticket")) And .Fields("Ticket") <> Year(.Fields("Fecha")) Then Cta_Ventas = Cta_Ventas_Ant
                      Total = .Fields("Total")
                      InsValorCta Cta_Ventas, -Total
                   End If
                End If
                TotalIngreso = TotalIngreso + .Fields("Total")
             Else
               'cuando usamos una cuenta de venta por producto
                  If Cta_Ventas <> .Fields("Cta_Ventas") Then
                    'Cta_Provision = "."
                     If Len(Cta_Provision) > 1 Then
                        SaldoDisp = 0: SaldoCont = 0
                        If AdoVentaAct.Recordset.RecordCount > 0 Then
                           AdoVentaAct.Recordset.MoveFirst
                           AdoVentaAct.Recordset.Find ("Cta = '" & Cta_Ventas & "' ")
                           If Not AdoVentaAct.Recordset.EOF Then
                              Contador = 0
                              Do While Not AdoVentaAct.Recordset.EOF
                                 If AdoVentaAct.Recordset.Fields("Cta") = Cta_Ventas Then
                                    FechaTexto = AdoVentaAct.Recordset.Fields("Fecha")
                                    FechaTexto1 = "31/12/" & FechaAnio(MBFechaF.Text)
                                    Select Case AdoVentaAct.Recordset.Fields("TP")
                                      Case "MENS": Documento = DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1
                                      Case "QUNC": Documento = (DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1) * 2
                                      Case "SEMA": Documento = DatePart("ww", FechaTexto1) - DatePart("ww", FechaTexto) - 1
                                    End Select
                                    'MsgBox "Fecha: " & FechaTexto & vbCrLf & FechaTexto1 & vbCrLf & "Valor Ed = " & AdoVentaAct.Recordset.Fields("Valor_Ed") & vbCrLf & Documento
                                    SaldoCont = SaldoCont + Redondear(AdoVentaAct.Recordset.Fields("Valor_Ed") * Documento, 2)
                                    Contador = Contador + 1
                                 End If
                                 AdoVentaAct.Recordset.MoveNext
                              Loop
                           End If
                        End If
                        'MsgBox "Total(" & Contador & ") = " & Total & vbCrLf & SaldoCont
                        InsValorCta Cta_Provision, -SaldoCont
                        If SaldoCont > 0 Then Total = Total - SaldoCont
                     End If
                     If TipoProc = "NV" Then Cta_Ventas = Cta_Ventas_0
                     If IsNumeric(.Fields("Ticket")) And .Fields("Ticket") <> Year(.Fields("Fecha")) Then Cta_Ventas = Cta_Ventas_Ant
                     'MsgBox Cta_Ventas & vbCrLf & Total
                     InsValorCta Cta_Ventas, -Total
                     Cta_Ventas = .Fields("Cta_Ventas")
                     Cta_Ventas_0 = .Fields("Cta_Ventas_0")
                     Cta_Ventas_Ant = .Fields("Cta_Ventas_Ant")
                     Cta_Provision = .Fields("Cta_Venta_Anticipada")
                     Total = 0
                  End If
             End If
             If Cta_Inventario <> .Fields("Cta_Inventario") Or _
                CodigoInv <> .Fields("Codigo") Or _
                Cod_Bodega <> .Fields("CodBodega") Then
                EgresosArtInv
                InsValorCta Cta_Inventario, -ValorTotal
                InsValorCta Cta_Costo_Ventas, ValorTotal
                Combos = Ninguno
                Codigo = .Fields("Codigo")
                Precio = .Fields("Precio")
                CodigoInv = .Fields("Codigo")
                Producto = .Fields("Producto")
                Unidad = .Fields("Unidad")
                Cta_Inventario = .Fields("Cta_Inventario")
                Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
                Cod_Bodega = .Fields("CodBodega")
                If Cod_Bodega = "" Then Cod_Bodega = Ninguno
                If Codigo = "" Then Codigo = Ninguno
                If CodigoInv = "" Then CodigoInv = Ninguno
                Entrada = 0
             End If
             'Total = Total + (.Fields("Cantidad") * .Fields("Precio"))
             If AdoCxC1.Recordset.RecordCount <= 0 Then
                Total = Total + .Fields("Total")
                Entrada = Entrada + .Fields("Cantidad")
                TipoProc = .Fields("TC")
                TotalIngreso = TotalIngreso + .Fields("Total")
             End If
            .MoveNext
          Loop
          
          If Cta_Inventario <> Ninguno Then EgresosArtInv
          InsValorCta Cta_Inventario, -ValorTotal
          InsValorCta Cta_Costo_Ventas, ValorTotal
          'Cta_Provision = "."
          If Len(Cta_Provision) > 1 Then
             SaldoDisp = 0: SaldoCont = 0
             If AdoVentaAct.Recordset.RecordCount > 0 Then
                AdoVentaAct.Recordset.MoveFirst
                AdoVentaAct.Recordset.Find ("Cta = '" & Cta_Ventas & "' ")
                If Not AdoVentaAct.Recordset.EOF Then
                   Contador = 0
                   Do While Not AdoVentaAct.Recordset.EOF
                      If AdoVentaAct.Recordset.Fields("Cta") = Cta_Ventas Then
                         FechaTexto = AdoVentaAct.Recordset.Fields("Fecha")
                         FechaTexto1 = "31/12/" & FechaAnio(MBFechaF.Text)
                         Select Case AdoVentaAct.Recordset.Fields("TP")
                           Case "MENS": Documento = DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1
                           Case "QUNC": Documento = (DatePart("m", FechaTexto1) - DatePart("m", FechaTexto) - 1) * 2
                           Case "SEMA": Documento = DatePart("ww", FechaTexto1) - DatePart("ww", FechaTexto) - 1
                         End Select
                         'MsgBox "Fecha: " & FechaTexto & vbCrLf & FechaTexto1 & vbCrLf & "Valor Ed = " & AdoVentaAct.Recordset.Fields("Valor_Ed") & vbCrLf & Documento
                         SaldoCont = SaldoCont + Redondear(AdoVentaAct.Recordset.Fields("Valor_Ed") * Documento, 2)
                         'MsgBox Cta_Ventas & vbCrLf & SaldoCont
                         Contador = Contador + 1
                      End If
                      AdoVentaAct.Recordset.MoveNext
                   Loop
                End If
             End If
             'MsgBox "Total(" & Contador & ") = " & Total & vbCrLf & SaldoCont
             InsValorCta Cta_Provision, -SaldoCont
             If SaldoCont > 0 Then Total = Total - SaldoCont
          End If
          
          If AdoCxC1.Recordset.RecordCount <= 0 Then
             If TipoProc = "NV" Then Cta_Ventas = Cta_Ventas_0
             InsValorCta Cta_Ventas, -Total
          End If
      End If
    End With
 
  FAbonos.Caption = "Procesando Asientos Contables"
''  sSQL = "SELECT A.Codigo_Inv,DF.TC,DF.Codigo,A.Cta_Inventario,A.Cta_Costo_Venta,A.Cta_Ventas,A.Cta_Ventas_0," _
''       & "A.Cta_Venta_Anticipada,DF.Cantidad,DF.Precio,DF.Total,A.Unidad,A.Producto,A.PVP " _
''       & "FROM Detalle_Factura As DF,Catalogo_Productos As A " _
''       & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''       & "AND A.Periodo = '" & Periodo_Contable & "' " _
''       & "AND DF.Item = '" & NumEmpresa & "' " _
''       & "AND DF.T <> '" & Anulado & "' " _
''       & "AND DF.TC NOT IN ('C','P','OP') " _
''       & "AND DF.Codigo = A.Codigo_Inv " _
''       & "AND DF.Periodo = A.Periodo " _
''       & "AND DF.Item = A.Item "
''  If CheqCajero.value = 1 Then sSQL = sSQL & "AND DF.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
''  sSQL = sSQL & "ORDER BY DF.Codigo,A.Codigo_Inv,DF.Fecha,DF.Precio "
''  SelectAdodc AdoAux, sSQL
''  Total = 0
''  TotalIngreso = 0
''  'MsgBox "."
''  Entrada = 0
''  With AdoAux.Recordset
''   If .RecordCount > 0 Then
''       CodigoInv = .Fields("Codigo_Inv")
''       Cta_Inventario = .Fields("Cta_Inventario")
''       Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
''       Producto = .Fields("Producto")
''       Do While Not .EOF
''          If CodigoInv <> .Fields("Codigo_Inv") Then
''             EgresosArtInv
''             'MsgBox ValorTotal
''             InsValorCta Cta_Inventario, -ValorTotal
''             InsValorCta Cta_Costo_Ventas, ValorTotal
''             Entrada = 0
''             CodigoInv = .Fields("Codigo_Inv")
''             Cta_Inventario = .Fields("Cta_Inventario")
''             Cta_Costo_Ventas = .Fields("Cta_Costo_Venta")
''             Producto = .Fields("Producto")
''          End If
''          Entrada = Entrada + .Fields("Cantidad")
''          'MsgBox CodigoInv & vbCrLf & Entrada
''         .MoveNext
''       Loop
''       EgresosArtInv
''       InsValorCta Cta_Inventario, -ValorTotal
''       InsValorCta Cta_Costo_Ventas, ValorTotal
''   End If
''  End With
  'TextoImprimio
  If ErrorInventario <> "" Then
     TextoImprimio = TextoImprimio _
                   & "Warning: Falta de Ingresar Entrada Inicial de los siguientes producto(s):" & vbCrLf _
                   & ErrorInventario & vbCrLf
  End If
  For IE = 0 To ContCtas - 1
   If CtasProc(IE).Valor >= 0 Then
      InsertarAsientos AdoAsiento1, CtasProc(IE).Cta, 0, CtasProc(IE).Valor, 0
   Else
      InsertarAsientos AdoAsiento1, CtasProc(IE).Cta, 0, 0, -CtasProc(IE).Valor
   End If
  Next IE
  
   Cadena = ""
   For IE = 0 To ContCtas - 1
       If CtasProc(IE).Valor <> 0 Then Cadena = Cadena & CtasProc(IE).Cta & " = " & CtasProc(IE).Valor & vbCrLf
       CtasProc(IE).Valor = 0
   Next IE
   'MsgBox Cadena
   sSQL = "SELECT * " _
        & "FROM Asiento_K " _
        & "WHERE CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "ORDER BY CODIGO_INV "
   SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|."
   SelectDataGrid DGInv, AdoInv, sSQL, SQLDec
  
  Debe = 0
  Haber = 0
  Trans_No = 96
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY DEBE DESC,HABER "
  SelectDataGrid DGAsiento, AdoAsiento, SQL2
 'Verificacion SubTotal
  Debe = 0: Haber = 0: Ln_No = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .Fields("A_No") = Ln_No
          Ln_No = Ln_No + 1
         .MoveNext
       Loop
      .UpdateBatch
   End If
  End With
  LabelDebe.Caption = Format$(Debe, "#,##0.00")
  LabelHaber.Caption = Format$(Haber, "#,##0.00")
  LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
  
 
  Debe = 0
  Haber = 0
  Trans_No = 97
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY DEBE DESC,HABER "
  SelectDataGrid DGAsiento1, AdoAsiento1, SQL2
  With AdoAsiento1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
          'MsgBox .Fields("DEBE") & vbCrLf & .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
'  LabelVentas.Caption = Format$(TotalIngreso, "#,##0.00")
  LabelDebe1.Caption = Format$(Debe, "#,##0.00")
  LabelHaber1.Caption = Format$(Haber, "#,##0.00")
  LblDiferencia1.Caption = Format$(Debe - Haber, "#,##0.00")
  If MBFechaI.Text = MBFechaF.Text Then
     LblConcepto.Caption = "Cierre Diario de Caja de Abonos del " & MBFechaI.Text & ", Diario No. ?"
     LblConcepto1.Caption = "Cierre Diario de Caja de CxC del " & MBFechaI.Text & ", Diario No. ?"
  Else
     LblConcepto.Caption = "Cierre Diario de Caja de Abonos del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. ?"
     LblConcepto1.Caption = "Cierre Diario de Caja de CxC del " & MBFechaI.Text & " al " & MBFechaF.Text & ", Diario No. ?"
  End If
 'Listado de Facturas anuladas
  Total = 0
  sSQL = "SELECT F.T,F.TC,F.Fecha,C.Cliente,F.Factura,F.IVA As Total_IVA,F.Total_MN,F.Cta_CxP " _
       & "FROM Facturas F,Clientes C " _
       & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND F.T = 'A' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.TC NOT IN ('C','P','OP') " _
       & "AND F.CodigoC = C.Codigo "
  If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
  sSQL = sSQL & "ORDER BY F.TC,F.Fecha,F.Cta_CxP,C.Cliente,F.Factura "
  SelectDataGrid DGFactAnul, AdoFactAnul, sSQL
  With AdoFactAnul.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Total_MN")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LblTotAnuladas.Caption = Format$(Total, "#,##0.00")
 'Actualiza Retenciones del S.R.I.
  sSQL = "UPDATE Facturas " _
       & "SET Ret_IVA = 0 " _
       & "FROM Facturas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET Serie_R = (TA.Serie_R), Retencion = TA.Cheque, Ret_Fuente = TA.Abono " _
          & "FROM Facturas As F,Trans_Abonos As TA "
  Else
     sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
          & "SET F.Serie_R = (TA.Serie_R), F.Retencion = TA.Cheque, F.Ret_Fuente = TA.Abono "
  End If
  sSQL = sSQL _
       & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TA.Banco,1,16) = 'RETENCION FUENTE' " _
       & "AND TA.TP = F.TC " _
       & "AND TA.Item = F.Item " _
       & "AND TA.Periodo = F.Periodo " _
       & "AND TA.Factura = F.Factura " _
       & "AND TA.CodigoC = F.CodigoC "
  ConectarAdoExecute sSQL
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET Ret_IVA = TA.Abono " _
          & "FROM Facturas As F,Trans_Abonos As TA "
  Else
     sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
          & "SET F.Ret_IVA = TA.Abono "
  End If
  sSQL = sSQL _
       & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TA.Banco,1,15) = 'RETENCION IVA S' " _
       & "AND TA.TP = F.TC " _
       & "AND TA.Item = F.Item " _
       & "AND TA.Periodo = F.Periodo " _
       & "AND TA.Factura = F.Factura " _
       & "AND TA.CodigoC = F.CodigoC "
  ConectarAdoExecute sSQL
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET Ret_IVA = Ret_IVA + TA.Abono " _
          & "FROM Facturas As F,Trans_Abonos As TA "
  Else
     sSQL = "UPDATE Trans_Abonos As TA,Facturas As F " _
          & "SET F.Ret_IVA = F.Ret_IVA + TA.Abono "
  End If
  sSQL = sSQL _
       & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TA.Banco,1,15) = 'RETENCION IVA B' " _
       & "AND TA.TP = F.TC " _
       & "AND TA.Item = F.Item " _
       & "AND TA.Periodo = F.Periodo " _
       & "AND TA.Factura = F.Factura " _
       & "AND TA.CodigoC = F.CodigoC "
  ConectarAdoExecute sSQL
  
 'REPORTES DE AUDITORIA TRANSACCIONALES (S.R.I.)
  If MBFechaI = MBFechaF Then
     DGSRI.Caption = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI
  Else
     DGSRI.Caption = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI & " al " & MBFechaF
  End If
'  sSQL = "SELECT F.T,F.Factura,F.Fecha,C.Cliente,C.CI_RUC,F.Con_IVA,F.Sin_IVA,F.Descuento,F.IVA As Total_IVA," _
'       & "F.Total_MN As TOTAL,Serie_R,Retencion,Ret_Fuente,Ret_IVA "
  Codigo = CStr(Porc_IVA * 100)
  sSQL = "SELECT F.TC,F.T,F.RUC_CI,F.Razon_Social,F.Fecha,F.Hora,A.Nombre_Completo As Usuario," _
       & "F.Autorizacion,F.Serie,F.Factura As Secuencial,F.Con_IVA As Base_" & Codigo & "," _
       & "F.Sin_IVA As Base_0,F.Descuento,F.Descuento2,F.SubTotal,F.IVA As IVA_" & Codigo & ",F.Servicio,F.Total_MN As TOTAL," _
       & "Serie_R,Retencion,Ret_Fuente,Ret_IVA " _
       & "FROM Facturas F, Clientes C, Accesos As A " _
       & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND F.TC NOT IN ('C','P','OP') " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = C.Codigo " _
       & "AND F.CodigoU = A.Codigo "
  If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
  sSQL = sSQL & "ORDER BY F.Factura,F.TC,F.Fecha,F.Cta_CxP,C.Cliente "
  SelectDataGrid DGSRI, AdoSRI, sSQL
  Total_Con_IVA = 0
  Total_Sin_IVA = 0
  Total_Desc = 0
  Total_Desc2 = 0
  Total_IVA = 0
  Total = 0
  With AdoSRI.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If .Fields("T") <> Anulado Then
              Total_Con_IVA = Total_Con_IVA + .Fields("Base_" & Codigo)
              Total_Sin_IVA = Total_Sin_IVA + .Fields("Base_0")
              Total_Desc = Total_Desc + .Fields("Descuento")
              Total_Desc2 = Total_Desc2 + .Fields("Descuento2")
              Total_IVA = Total_IVA + .Fields("IVA_" & Codigo)
              Total_Servicio = Total_Servicio + .Fields("Servicio")
              Total = Total + .Fields("TOTAL")
          End If
         .MoveNext
       Loop
   End If
  End With
  LblConIVA.Caption = Format$(Total_Con_IVA, "#,##0.00")
  LblSinIVA.Caption = Format$(Total_Sin_IVA, "#,##0.00")
  LblDescuento.Caption = Format$(Total_Desc + Total_Desc2, "#,##0.00")
  LblIVA.Caption = Format$(Total_IVA, "#,##0.00")
  LblServicio.Caption = Format$(Total_Servicio, "#,##0.00")
  LblTotalFacturado.Caption = Format$(Total, "#,##0.00")
  'Fecha_Vence
  'SerieFactura
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoInv = .Fields("CODIGO_INV")
          Valor_Prom = Redondear(.Fields("VALOR_UNIT"), Dec_Costo)
          sSQL = "UPDATE Detalle_Factura " _
               & "SET Costo = " & Valor_Prom & " " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo = '" & CodigoInv & "' "
          ConectarAdoExecute sSQL
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT CP.Codigo_Inv,CP.Producto,SUM(DF.Cantidad) As CANTIDADES,SUM(DF.Total+DF.Total_IVA) As TOTALES,Cta_Ventas,Cta_Ventas_0  " _
       & "FROM Detalle_Factura DF,Catalogo_Productos CP " _
       & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.T <> '" & Anulado & "' " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv " _
       & "GROUP BY CP.Codigo_Inv,CP.Producto,Cta_Ventas,Cta_Ventas_0 " _
       & "UNION " _
       & "SELECT '-x-' As Codigo_Inv,'TOTAL DE VENTAS' As Producto,SUM(DF.Cantidad) As CANTIDADES,SUM(DF.Total+DF.Total_IVA) As TOTALES,'' As 'V12','' As 'V0' " _
       & "FROM Detalle_Factura DF,Catalogo_Productos CP " _
       & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.T <> '" & Anulado & "' " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv " _
       & "ORDER BY CP.Codigo_Inv,CP.Producto "
  SelectDataGrid DGProductos, AdoProductos, sSQL
'    & "GROUP BY DF.Fecha "
'''  Cadena = ""
'''   For IE = 0 To ContCtas - 1
'''       If CtasProc(IE).Valor <> 0 Then
'''          Cadena = Cadena & CtasProc(IE).Cta & "  -  " & CtasProc(IE).Valor & vbCrLf
'''       End If
'''   Next IE
  'MsgBox Cadena
  DGVentas.Visible = True
  DGCxC.Visible = True
  DGInv.Visible = True
  DGAsiento.Visible = True
  FAbonos.Caption = "CIERRE DEL DIARIO DE CAJA"
 'MsgBox TextoImprimio
End Sub

Public Sub EgresosArtInv()
 'Fecha <= #" & BuscarFecha(MBFechaF.Text) & "#
 If Len(Cta_Inventario) > 1 Then
    If CodigoInv <> Ninguno Then
       ValorUnit = 0: Total_Desc = 0: Saldo = 0
       sSQL = "SELECT TOP 1 Codigo_Inv,Costo As V_Unit,Existencia,Total,T " _
            & "FROM Trans_Kardex " _
            & "WHERE Fecha <= #" & BuscarFecha(MBFechaF) & "# " _
            & "AND Codigo_Inv = '" & CodigoInv & "' " _
            & "AND T <> 'A' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "ORDER BY Fecha DESC,Entrada,Salida DESC,TP DESC, Numero DESC,Kardex DESC "
       SelectData AdoSQL, sSQL
       'MsgBox sSQL
       With AdoSQL.Recordset
        If .RecordCount > 0 Then
            Cantidad = .Fields("Existencia")
            ValorUnit = Redondear(.Fields("V_Unit"), Dec_Costo)
            SaldoAnterior = Redondear(.Fields("Total"), 2)
        End If
       End With
       Precio = ValorUnit
       'If CodigoInv = "01.03.004" Then MsgBox ValorUnit
       If ValorUnit <= 0 Then
          ErrorInventario = ErrorInventario & CodigoInv & " - " _
                          & Cta_Inventario & " - " & Cta_Costo_Ventas _
                          & vbCrLf & Space(10)
       Else
          ValorTotal = Redondear(ValorUnit * Entrada, 2)
         'Llenamos el ultimo saldo del kardex
         'If Entrada > 0 And Precio > 0 Then
          Cantidad = Cantidad - Entrada
          Saldo = Redondear(SaldoAnterior - ValorTotal, 2)
          SetAddNew AdoInv
          SetFields AdoInv, "DH", 2
          SetFields AdoInv, "CODIGO_INV", CodigoInv
          SetFields AdoInv, "P_DESC", 0
          SetFields AdoInv, "PRODUCTO", Producto
          SetFields AdoInv, "CANT_ES", Entrada
          SetFields AdoInv, "VALOR_UNIT", ValorUnit
          SetFields AdoInv, "VALOR_TOTAL", ValorTotal
          SetFields AdoInv, "CTA_INVENTARIO", Cta_Inventario
          SetFields AdoInv, "CONTRA_CTA", Cta_Costo_Ventas
          SetFields AdoInv, "CANTIDAD", 0
          SetFields AdoInv, "SALDO", 0
          SetFields AdoInv, "CodigoU", CodigoUsuario
          SetFields AdoInv, "Codigo_B", Ninguno
          SetFields AdoInv, "ORDEN", "F" & Format$(Day(MBFechaF), "00") & Format$(Month(MBFechaF), "00") & CStr(Year(MBFechaF))
          SetFields AdoInv, "CodBod", Cod_Bodega
          SetFields AdoInv, "T_No", Trans_No
          SetFields AdoInv, "Item", NumEmpresa
          SetUpdate AdoInv
       End If
    End If
 End If
End Sub

Public Sub SetearCtasCierre(CtaFields As String)
  Si_No = True
  For IE = 0 To ContCtas - 1
      If CtaFields = CtasProc(IE).Cta Then Si_No = False
  Next IE
  If Si_No Then
     IE = 0
     While IE < ContCtas
        If CtasProc(IE).Cta = "0" Then
           CtasProc(IE).Cta = CtaFields
           IE = ContCtas + 1
        End If
        IE = IE + 1
     Wend
  End If
End Sub

Public Sub InsValorCta(NCta As String, _
                       NValor As Currency)
  For IE = 0 To ContCtas - 1
      If CtasProc(IE).Cta = NCta Then
         CtasProc(IE).Valor = CtasProc(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

Public Sub CierreDelDia()
  sSQL = "SELECT Fecha,Factura " _
       & "FROM Trans_Abonos " _
       & "WHERE C = " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND T <> 'A' " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Fecha,Factura " _
       & "UNION " _
       & "SELECT Fecha,Factura " _
       & "FROM Facturas " _
       & "WHERE C = " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "AND T <> 'A' " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "GROUP BY Fecha,Factura " _
       & "ORDER BY Fecha "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       MsgBox "Cierre del día: " & .Fields("Fecha") & "(" & .Fields("Factura") & ")" & vbCrLf
       MBFechaI = .Fields("Fecha")
       MBFechaF = .Fields("Fecha")
       MarcarTexto MBFechaI
       MBFechaI.SetFocus
   End If
  End With
End Sub

Public Sub Actualizar_Ejecutivo_Facturas(FechaIni As String, FechaFin As String, Opcion As Boolean)
    RatonReloj
    sSQL = "UPDATE Facturas " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Detalle_Factura " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Trans_Abonos " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
    ConectarAdoExecute sSQL
    
    RatonReloj
    If SQL_Server Then
       sSQL = "UPDATE Facturas " _
            & "SET X = 'R' " _
            & "FROM Facturas As F,Clientes C "
    Else
       sSQL = "UPDATE Facturas As F,Clientes C " _
            & "SET F.X = 'R' "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND F.Cod_Ejec = C.Codigo "
    ConectarAdoExecute sSQL
  
    If SQL_Server Then
       sSQL = "UPDATE Detalle_Factura " _
            & "SET X = 'R' " _
            & "FROM Detalle_Factura As F,Clientes C "
    Else
       sSQL = "UPDATE Detalle_Factura As F,Clientes C " _
            & "SET F.X = 'R' "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND F.Cod_Ejec = C.Codigo "
    ConectarAdoExecute sSQL
  
    If SQL_Server Then
       sSQL = "UPDATE Trans_Abonos " _
            & "SET X = 'R' " _
            & "FROM Trans_Abonos As F,Clientes C "
    Else
       sSQL = "UPDATE Trans_Abonos As F,Clientes C " _
            & "SET F.X = 'R' "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND F.Cod_Ejec = C.Codigo "
    ConectarAdoExecute sSQL
  
    RatonReloj
    sSQL = "UPDATE Facturas " _
         & "SET Cod_Ejec = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND X = '.' "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Detalle_Factura " _
         & "SET Cod_Ejec = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND X = '.' "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Trans_Abonos " _
         & "SET Cod_Ejec = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND X = '.' "
    ConectarAdoExecute sSQL
  
    sSQL = "UPDATE Facturas " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
    ConectarAdoExecute sSQL
    
    RatonReloj
    If SQL_Server Then
       sSQL = "UPDATE Facturas " _
            & "SET X = 'R' " _
            & "FROM Facturas As F,Accesos As A "
    Else
       sSQL = "UPDATE Facturas As F,Accesos As A " _
            & "SET F.X = 'R' "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND F.Cod_Ejec = A.Codigo "
    ConectarAdoExecute sSQL
  
   'Catalogo_Rol_Pagos
    sSQL = "SELECT C.Cliente As Ejecutivo_Venta,F.TC,F.Serie,F.Autorizacion,F.Fecha,F.Factura " _
         & "FROM Facturas F,Clientes C " _
         & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
         & "AND F.TC NOT IN ('C','P','OP') " _
         & "AND F.T <> 'A' " _
         & "AND F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.X = '.' " _
         & "AND F.Cod_Ejec = C.Codigo "
    If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
    sSQL = sSQL & "ORDER BY C.Cliente,F.TC,F.Serie,F.Autorizacion,F.Fecha,F.Factura "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       TextoImprimio = TextoImprimio & "Asigne al Rol de Pagos, al Ejecutivo de Venta de la(s) siguiente(s) Factura(s): " & vbCrLf
       Do While Not .EOF
          TextoImprimio = TextoImprimio _
                        & .Fields("Fecha") & " " & .Fields("TC") & " " & .Fields("Autorizacion") _
                        & " No. " & .Fields("Serie") & "-" & Format$(.Fields("Factura"), "00000000") _
                        & ", Ejecutivo de Venta: " & .Fields("Ejecutivo_Venta") & vbCrLf
         .MoveNext
       Loop
   End If
  End With
End Sub
