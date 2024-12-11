VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CierreEjercicio 
   Caption         =   "BALANCE DE COMPROBACION"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheqRenumerar 
      Caption         =   "Renumerar Comprobantes"
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
      Left            =   2940
      TabIndex        =   24
      Top             =   945
      Width           =   1590
   End
   Begin VB.CheckBox CheqSinConc 
      Caption         =   "Sin Concilicacion"
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
      Left            =   1260
      TabIndex        =   20
      Top             =   945
      Width           =   1485
   End
   Begin VB.CheckBox CheqDetalle 
      Caption         =   "Detalle Auxiliar"
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
      Left            =   105
      TabIndex        =   18
      Top             =   945
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   7350
      TabIndex        =   15
      Top             =   105
      Width           =   4215
      Begin MSMask.MaskEdBox MBoxCtaI 
         Height          =   330
         Left            =   2415
         TabIndex        =   16
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
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta &Utilidad/Perdida"
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
         TabIndex        =   17
         Top             =   210
         Width           =   2325
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2415
      Top             =   7455
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cierreej.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cierreej.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cierreej.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cierreej.frx":1046
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cierreej.frx":1360
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5475
      Left            =   105
      TabIndex        =   5
      Top             =   1470
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&CONTABILIZACION"
      TabPicture(0)   =   "Cierreej.frx":1C3A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SUB &MODULOS"
      TabPicture(1)   =   "Cierreej.frx":1C56
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGBanco"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "IN&VENTARIO"
      TabPicture(2)   =   "Cierreej.frx":1C72
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGInv"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cheques Girados y No Cobrados"
      TabPicture(3)   =   "Cierreej.frx":1C8E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGCheques"
      Tab(3).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGInv 
         Bindings        =   "Cierreej.frx":1CAA
         Height          =   4950
         Left            =   -74895
         TabIndex        =   8
         Top             =   420
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         BorderStyle     =   0
         ForeColor       =   8388608
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
      Begin MSDataGridLib.DataGrid DGBanco 
         Bindings        =   "Cierreej.frx":1CBF
         Height          =   4950
         Left            =   -74895
         TabIndex        =   7
         Top             =   420
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16744576
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
      Begin MSDataGridLib.DataGrid DGBalance 
         Bindings        =   "Cierreej.frx":1CD6
         Height          =   4950
         Left            =   105
         TabIndex        =   6
         Top             =   420
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
      Begin MSDataGridLib.DataGrid DGCheques 
         Bindings        =   "Cierreej.frx":1CEC
         Height          =   4950
         Left            =   -74895
         TabIndex        =   19
         Top             =   420
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   8731
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
         BorderStyle     =   0
         ForeColor       =   8388608
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
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   2
      Top             =   7455
      Width           =   330
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1575
      TabIndex        =   4
      Top             =   6300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   315
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoEstRes 
      Height          =   330
      Left            =   315
      Top             =   4725
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
      Caption         =   "EstRes"
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
   Begin MSAdodcLib.Adodc AdoRet 
      Height          =   330
      Left            =   315
      Top             =   4410
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
      Caption         =   "Ret"
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet 
      Height          =   330
      Left            =   315
      Top             =   4095
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
      Caption         =   "SubCtaDet"
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
   Begin MSAdodcLib.Adodc AdoIngKar 
      Height          =   330
      Left            =   315
      Top             =   3780
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
      Caption         =   "IngKar"
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
   Begin MSAdodcLib.Adodc AdoBalGenCon 
      Height          =   330
      Left            =   315
      Top             =   3465
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
      Caption         =   "BalGenCon"
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
      Top             =   3150
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
   Begin MSAdodcLib.Adodc AdoFechaBal 
      Height          =   330
      Left            =   315
      Top             =   2835
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
      Caption         =   "FechaBal"
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
   Begin MSAdodcLib.Adodc AdoBalGen 
      Height          =   330
      Left            =   315
      Top             =   2520
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
      Caption         =   "BalGen"
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
      Top             =   2205
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
      Left            =   315
      Top             =   5040
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
      Left            =   315
      Top             =   5670
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   315
      Top             =   5355
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
   Begin MSAdodcLib.Adodc AdoCheques 
      Height          =   330
      Left            =   315
      Top             =   5985
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
      Caption         =   "Cheques"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1588
      ButtonWidth     =   2540
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Cierre del Ejercicio"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Procesar Cierre"
            Key             =   "Procesar"
            Object.ToolTipText     =   "Procesar Cierre del Ejercicio"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar Cierre"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Cierre del Ejercicio"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualizar Cierre"
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar Cierre del Ejercicio"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Asiento"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Asiento Contable"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   11655
         TabIndex        =   21
         Top             =   105
         Width           =   4110
         Begin VB.Label LabelTotInv 
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
            Height          =   330
            Left            =   2100
            TabIndex        =   23
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TOTAL INVENTARIO"
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
            TabIndex        =   22
            Top             =   210
            Width           =   2010
         End
      End
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Haber"
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
      TabIndex        =   14
      Top             =   7455
      Width           =   1170
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Debe"
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
      Left            =   3045
      TabIndex        =   13
      Top             =   7455
      Width           =   1170
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
      Left            =   4200
      TabIndex        =   12
      Top             =   7455
      Width           =   1800
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
      Left            =   7140
      TabIndex        =   11
      Top             =   7455
      Width           =   1800
   End
   Begin VB.Label Label11 
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
      Left            =   8925
      TabIndex        =   10
      Top             =   7455
      Width           =   1065
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
      Left            =   9975
      TabIndex        =   9
      Top             =   7455
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Cierre del Ejercicio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   4725
      TabIndex        =   1
      Top             =   945
      Width           =   11145
   End
End
Attribute VB_Name = "CierreEjercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ErrorSubCtas As Boolean
Dim SiTieneSubModulo As Boolean

'''Public Sub Procesar_Saldos_Facturas()
'''  sSQL = "DELETE * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND TP = 'CCXP' "
'''  Ejecutar_SQL_SP sSQL
'''  TipoSubCta = Ninguno
'''  'FechaValida MBoxFechaI
'''  'FechaValida MBoxFechaF
'''  DGBanco.Visible = False
'''  'FechaInicial = BuscarFecha(MBoxFechaI.Text)
'''  'FechaFinal = BuscarFecha(MBoxFechaF.Text)
'''  CodigoB = Ninguno
'''  With AdoSubCta.Recordset
'''    If CheqIndiv.value = 1 Then
'''       If .RecordCount > 0 Then
'''        .MoveFirst
'''        .Find ("Nombre_Cta Like '" & DCCtas.Text & "'")
'''        If Not .EOF Then CodigoB = .Fields("Codigo")
'''       End If
'''    End If
'''  End With
'''  Cta_Sup = Ninguno
'''  If CheqCta.value = 1 Then Cta_Sup = SinEspaciosIzq(DCCtas.Text)
'''  If OpcCxC.value Then
'''     SQLMsg1 = "SALDO DE CUENTAS POR COBRAR"
'''     TipoCta = "C"
'''  ElseIf OpcCxP.value Then
'''     SQLMsg1 = "SALDO DE CUENTAS POR PAGAR"
'''     TipoCta = "P"
'''  End If
'''  Saldo_Facturas_CxCxP MBoxFechaI.Text, MBoxFechaF.Text, PictProcMod
'''  If OpcP.value Then TipoDoc = Pendiente Else TipoDoc = Cancelado
'''End Sub

'''Public Sub Stock_Invent_Cierre(TipoDeBodega As String)
'''  RatonReloj
'''  Contador = 0
'''
'''  sSQL = "UPDATE Catalogo_Productos " _
'''       & "SET Stock_Anterior=0,Entradas=0,Salidas=0," _
'''       & "Promedio=0,Stock_Actual=0,Valor_Total=0 " _
'''       & "WHERE Item = '" & NumEmpresa & "' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "SELECT K.CodBodega,P.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero," _
'''       & "Entrada,Salida,Existencia,K.Valor_Total,K.Valor_Unitario,K.Stock_Bod,K.Total_Bod " _
'''       & "FROM Trans_Kardex As K,Catalogo_Productos As P " _
'''       & "WHERE K.Fecha <= #" & FechaFin & "# " _
'''       & "AND K.Item = '" & NumEmpresa & "' " _
'''       & "AND K.Codigo_Inv = P.Codigo_Inv " _
'''       & "AND K.Item = P.Item " _
'''       & "AND K.CodBodega = '" & TipoDeBodega & "' " _
'''       & "ORDER BY K.CodBodega,P.Codigo_Inv,K.Fecha,K.TP,K.Numero,K.ID "
'''  'MsgBox sSQL
'''  Select_Adodc AdoIngKar, sSQL
'''  Entrada = 0: Salida = 0: Saldo = 0
'''  Total = 0: Total_ME = 0
'''  No_Desde = 0: Precio = 0
'''  With AdoIngKar.Recordset
'''   If .RecordCount > 0 Then
'''       I = CFechaLong(FechaInicial)
'''       J = CFechaLong(FechaFinal)
'''       Codigo = .Fields("Codigo_Inv")
'''       Entrada = 0: Salida = 0
'''       Contador = 0
'''       RatonReloj
'''       Do While Not .EOF
'''          CierreEjercicio.Caption = "INVENTARIO(" & TipoDeBodega & "): " & Format(No_Desde / .RecordCount, "00%")
'''          'MsgBox Contador
'''          'Contador = Contador + 1
'''          No_Desde = No_Desde + 1
'''          If Codigo <> .Fields("Codigo_Inv") Then
'''             PFil = Saldo + Salida - Entrada
'''             PCol = Saldo
'''             CalculosTotalesInventario AdoInv, Codigo, Entrada, Salida, PFil, PCol, Total
'''             Codigo = .Fields("Codigo_Inv")
'''             Entrada = 0: Salida = 0: Total = 0
'''             Contador = 0: Precio = 0
'''          End If
'''          K = CFechaLong(.Fields("Fecha"))
'''          If I <= K And K <= J Then
'''             Entrada = Entrada + .Fields("Entrada")
'''             Salida = Salida + .Fields("Salida")
'''          End If
'''          If .Fields("Entrada") <> 0 Then
'''              Total = Total + .Fields("Valor_Total")
'''              Total_ME = Total_ME + .Fields("Valor_Total")
'''          Else
'''              Total = Total - .Fields("Valor_Total")
'''              Total_ME = Total_ME - .Fields("Valor_Total")
'''          End If
'''          Precio = Precio + .Fields("Valor_Unitario")
'''          Contador = Contador + 1
'''          'If CheqBod.Value = 1 Then
'''             Saldo = .Fields("Stock_Bod")
'''          'Else
'''          '   Saldo = .Fields("Existencia")
'''          'End If
'''         .MoveNext
'''       Loop
'''       PFil = Saldo + Salida - Entrada
'''       PCol = Saldo
'''       CalculosTotalesInventario AdoInv, Codigo, Entrada, Salida, PFil, PCol, Total
'''   End If
'''  End With
'''  sSQL = "UPDATE Catalogo_Productos " _
'''       & "SET Valor_Total = Promedio * Stock_Actual "
'''  Ejecutar_SQL_SP sSQL
'''  RatonNormal
'''End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Label5.Caption = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaFinal)
  If CheqDetalle.value = 1 Then
     DetalleComp = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaFinal)
  Else
     DetalleComp = Ninguno
  End If
 'MsgBox Button.key
  Select Case Button.key
    Case "Salir":      Unload CierreEjercicio
    Case "Procesar":   Procesar_Cierre_Ejercicio
    Case "Grabar":     Grabar_Cierre_Ejercicio
    Case "Actualizar": Actualizar_Cierre_Ejercicio
    Case "Imprimir":   Imprimir
  End Select
End Sub

Public Sub Renumerar_Codigos(Renumerar As Boolean)
Dim AdoNum As ADODB.Recordset
Dim AdoComp As ADODB.Recordset
Dim MesNo As Byte
Dim CompNo As Byte
Dim Numero1 As Long
Dim Numero2 As Long
Dim TipoCodigo As String
Dim ListaTP(5) As String
Dim ListaCodigos(5) As String
   If Renumerar And Periodo_Contable = Ninguno Then
      ListaCodigos(0) = "Diario"
      ListaCodigos(1) = "Egresos"
      ListaCodigos(2) = "Ingresos"
      ListaCodigos(3) = "NotaCredito"
      ListaCodigos(4) = "NotaDebito"
     
      ListaTP(0) = "CD"
      ListaTP(1) = "CE"
      ListaTP(2) = "CI"
      ListaTP(3) = "NC"
      ListaTP(4) = "ND"
     
     'CLng(MesComp & "000001")
      For CompNo = 0 To 4
          TipoCodigo = ListaCodigos(CompNo)
          sSQL = "SELECT Concepto, Numero, ID " _
               & "FROM Codigos " _
               & "WHERE Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Concepto = '" & TipoCodigo & "' "
          Select_AdoDB AdoNum, sSQL
          If AdoNum.RecordCount > 0 Then
             Numero1 = AdoNum.fields("Numero")
             sSQL = "SELECT Concepto, Numero, ID " _
                  & "FROM Comprobantes " _
                  & "WHERE Periodo = '" & Periodo_Contable & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Month(Fecha) = " & TipoCodigo & "' "
          Select_AdoDB AdoNum, sSQL

             
          End If
          AdoNum.Close
          For MesNo = 1 To 12
              TipoCodigo = Format(MesNo, "00") & TipoCodigo
          
          Next MesNo
      Next CompNo
      
   End If
End Sub
 
Public Sub InsertarTotales(AdoG As Adodc, cod, Cta, Anal, Parc, Tot)
   With AdoG.Recordset
        .AddNew
        .fields("Codigo") = cod
        .fields("Cuenta") = Cta
        .fields("Analitico") = Anal
        .fields("Parcial") = Parc
        .fields("Total") = Tot
        .fields("Item") = NumEmpresa
        .Update
   End With
End Sub

Public Sub InsertarTotalesCon(AdoG As Adodc, cod, CtaDG, Cta, Sal_ME, Sal_MN, Tot)
   With AdoG.Recordset
        .AddNew
        .fields("Codigo") = cod
        .fields("Cuenta") = Cta
        .fields("Saldo_ME") = Sal_ME
        .fields("Saldo_MN") = Sal_MN
        .fields("Total") = Tot
        .fields("DG") = CtaDG
        .fields("Item") = NumEmpresa
        .Update
   End With
End Sub

Private Sub Procesar_Cierre_Ejercicio()
Dim OpcDH1 As Byte

Dim TDebito As Currency
Dim TCredito As Currency
Dim ValorSubModulo As Currency
Dim SumaCheqDebe As Currency
Dim SumaCheqHaber As Currency
Dim TempBaValorDH As Currency
Dim TempBaOpcDH As String
Dim Fecha_V1 As String
Dim FechaEmision As String
Dim Factura_No1 As Long

    Ln_No = 1
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Codigo,DEBE DESC,HABER "
    Select_Adodc_Grid DGBalance, AdoCtas, SQL2
    
   'Leemos la Cta de Utilidad/Perdida
    Codigo = Leer_Cta_Catalogo(CambioCodigoCta(MBoxCtaI.Text))
    If Cuenta = Ninguno Then MsgBox "Cuenta no asignada en el Catalogo de Cuentas"
    
    RatonReloj
    DGBalance.Visible = False
    Evaluar = True
    ErrorSubCtas = False
    Ln_No = 1
    TipoCta = Ninguno
    Fecha_Vence = FechaFinal
    Fecha_V1 = FechaFinal
    FechaEmision = FechaFinal
    
   'Variables Generales de Entrada
    CodigoCli = Ninguno
    Procesar_Cierre_Fiscal_SP Codigo, CheqSinConc.value

'''
'''
''' 'Asiento de SubCtas de los saldos con fecha de vencimiento
'''  RatonReloj
'''  TipoSubCta = Ninguno
'''  DGBanco.Visible = False
'''
'''  sSQL = "SELECT * " _
'''       & "FROM Asiento_SC " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Select_Adodc AdoBanco, sSQL
'''
'''  sSQL = "SELECT TS.Codigo, C.Cliente, TS.Serie, TS.Factura, TS.TC, TS.Cta, MIN(TS.Fecha_V) As Fecha_Venc, MIN(TS.Fecha_E) As Fecha_Emis, " _
'''       & "SUM(TS.Debitos) As TDebitos, SUM(TS.Creditos) As TCreditos " _
'''       & "FROM Clientes As C, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
'''       & "WHERE TS.Item = '" & NumEmpresa & "' " _
'''       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TS.Fecha <= #" & FechaFin & "# " _
'''       & "AND TS.TC IN ('C','P') " _
'''       & "AND TS.Codigo = C.Codigo " _
'''       & "AND TS.Cta = CC.Codigo " _
'''       & "AND TS.Item = CC.Item " _
'''       & "AND TS.Periodo = CC.Periodo " _
'''       & "GROUP BY TS.Codigo, C.Cliente, TS.Serie, TS.Factura, TS.TC, TS.Cta " _
'''       & "HAVING (SUM(TS.Debitos)-SUM(TS.Creditos)) <> 0 " _
'''       & "ORDER BY TS.TC, TS.Cta, C.Cliente, TS.Serie, TS.Factura, SUM(TS.Debitos) DESC, SUM(TS.Creditos) DESC "
'''  Select_Adodc AdoTrans, sSQL
''' 'MsgBox sSQL
'''  Contador = 0
'''  With AdoTrans.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''       Progreso_Barra.Incremento = 0
'''       Progreso_Barra.Valor_Maximo = .RecordCount
'''       Progreso_Barra.Mensaje_Box = "SUBMODULOS"
'''       Progreso_Esperar
'''       Do While Not .EOF
'''          Contador = Contador + 1
'''          Saldo = 0
'''          Saldo_ME = 0
'''          Valor = 0
'''          OpcTM = 1
'''          FechaEmision = .fields("Fecha_Emis")
'''          SerieFactura = .fields("Serie")
'''          Debitos = .fields("TDebitos")
'''          Creditos = .fields("TCreditos")
'''          Codigo = .fields("Codigo")
'''          Beneficiario = .fields("Cliente")
'''          SubCtaGen = .fields("Cta")
'''          Mifecha = .fields("Fecha_Venc")
'''          Factura_No = .fields("Factura")
'''          TipoSubCta = .fields("TC")
'''          Progreso_Barra.Mensaje_Box = "Submódulos de: " & SubCtaGen & " => " & Beneficiario
'''          Select Case .fields("TC")
'''            Case "C"
'''                 Saldo = Debitos - Creditos
'''                 ValorDH = Saldo
'''                 OpcDH = 1
'''                 If ValorDH < 0 Then
'''                    ValorDH = -ValorDH
'''                    OpcDH = 2
'''                 End If
'''            Case "P"
'''                 Saldo = Creditos - Debitos
'''                 ValorDH = Saldo
'''                 OpcDH = 2
'''                 If ValorDH < 0 Then
'''                    ValorDH = -ValorDH
'''                    OpcDH = 1
'''                 End If
'''          End Select
'''         'If SubCtaGen = "1.1.04.03.01" And Codigo = "GRUP2" And Factura_No = 18553 Then MsgBox ValorDH & " .."
'''          If ValorDH > 0 Then
'''             OpcDH1 = OpcDH
'''             Factura_No1 = Factura_No
'''             ValorSubModulo = ValorDH
'''             sSQL = "SELECT TS.Factura, TS.Fecha_V, TS.Debitos, TS.Creditos, C.Codigo " _
'''                  & "FROM Trans_SubCtas As TS, Clientes As C " _
'''                  & "WHERE TS.Item = '" & NumEmpresa & "' " _
'''                  & "AND TS.Periodo = '" & Periodo_Contable & "' " _
'''                  & "AND TS.Fecha <= #" & FechaFin & "# " _
'''                  & "AND TS.Fecha_V > #" & FechaFin & "# " _
'''                  & "AND TS.Codigo = '" & Codigo & "' " _
'''                  & "AND TS.Cta = '" & SubCtaGen & "' " _
'''                  & "AND TS.TC = '" & TipoSubCta & "' " _
'''                  & "AND TS.Serie = '" & SerieFactura & "' " _
'''                  & "AND TS.Factura = " & Factura_No & " " _
'''                  & "AND TS.T <> 'A' " _
'''                  & "AND TS.Codigo = C.Codigo "
'''             Select_Adodc AdoRet, sSQL
'''             If AdoRet.Recordset.RecordCount > 0 Then
'''               'MsgBox "Tiene prestamos superiores: " & SubCtaGen & vbCrLf & Codigo & vbCrLf & AdoRet.Recordset.RecordCount
'''                Do While Not AdoRet.Recordset.EOF
'''                   If AdoRet.Recordset.fields("Debitos") > 0 Then
'''                      ValorDH = AdoRet.Recordset.fields("Debitos")
'''                      If ValorDH = Debitos Then ValorDH = 0
'''                      OpcDH = 1
'''                   Else
'''                      ValorDH = AdoRet.Recordset.fields("Creditos")
'''                      If ValorDH = Creditos Then ValorDH = 0
'''                      OpcDH = 2
'''                   End If
'''                   Fecha_V1 = AdoRet.Recordset.fields("Fecha_V")
'''                   Factura_No = AdoRet.Recordset.fields("Factura")
'''                  'If SubCtaGen = "2.1.03.01.03" And Factura_No = 79498 Then MsgBox ValorSubModulo & " .."
'''                  'If SubCtaGen = "1.1.04.03.01" And Codigo = "GRUP2" And Factura_No = 18553 Then MsgBox ValorDH & " .."
'''                   If ValorDH > 0 Then
'''                      SetAddNew AdoBanco
'''                      SetFields AdoBanco, "FECHA_E", FechaEmision
'''                      SetFields AdoBanco, "FECHA_V", Fecha_V1
'''                      SetFields AdoBanco, "TC", TipoSubCta
'''                      SetFields AdoBanco, "Serie", SerieFactura
'''                      SetFields AdoBanco, "Factura", Factura_No
'''                      SetFields AdoBanco, "Codigo", Codigo
'''                      SetFields AdoBanco, "Beneficiario", Beneficiario
'''                      SetFields AdoBanco, "Cta", SubCtaGen
'''                      SetFields AdoBanco, "DH", OpcDH
'''                      SetFields AdoBanco, "Valor", ValorDH
'''                      SetFields AdoBanco, "TM", OpcTM
'''                      SetFields AdoBanco, "Item", NumEmpresa
'''                      SetFields AdoBanco, "T_No", Trans_No
'''                      SetFields AdoBanco, "CodigoU", CodigoUsuario
'''                      SetUpdate AdoBanco
'''                      ValorSubModulo = ValorSubModulo - ValorDH
'''                   End If
'''                   AdoRet.Recordset.MoveNext
'''                Loop
'''             End If
'''            'If SubCtaGen = "2.1.03.01.03" And ValorSubModulo < 0 Then MsgBox ValorSubModulo & " .."
'''            'If ValorSubModulo < 0 Then MsgBox "Negativo: " & SubCtaGen & vbCrLf & ValorSubModulo
'''             If ValorSubModulo <> 0 Then
'''               'MsgBox SubCtaGen & vbCrLf & Codigo & vbCrLf & " Valor Submodulo = " & ValorSubModulo
'''                OpcDH = OpcDH1
'''                Factura_No = Factura_No1
'''                ValorDH = ValorSubModulo
'''                Select Case TipoSubCta
'''                  Case "C"
'''                       If ValorDH < 0 Then
'''                          OpcDH = 2
'''                          ValorDH = -ValorDH
'''                       End If
'''                  Case "P"
'''                       If ValorDH < 0 Then
'''                          OpcDH = 1
'''                          ValorDH = -ValorDH
'''                       End If
'''                End Select
'''                SetAddNew AdoBanco
'''                SetFields AdoBanco, "FECHA_E", FechaEmision
'''                SetFields AdoBanco, "FECHA_V", Fecha_V1
'''                SetFields AdoBanco, "TC", TipoSubCta
'''                SetFields AdoBanco, "Serie", SerieFactura
'''                SetFields AdoBanco, "Factura", Factura_No
'''                SetFields AdoBanco, "Codigo", Codigo
'''                SetFields AdoBanco, "Beneficiario", Beneficiario
'''                SetFields AdoBanco, "Cta", SubCtaGen
'''                SetFields AdoBanco, "DH", OpcDH
'''                SetFields AdoBanco, "Valor", ValorDH
'''                SetFields AdoBanco, "TM", OpcTM
'''                SetFields AdoBanco, "Item", NumEmpresa
'''                SetFields AdoBanco, "T_No", Trans_No
'''                SetFields AdoBanco, "CodigoU", CodigoUsuario
'''                SetUpdate AdoBanco
'''             End If
'''          End If
'''          Progreso_Esperar
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''
  
  sSQL = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY TC,Cta,Beneficiario,FECHA_E,FECHA_V,Factura "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY CodBod,CODIGO_INV "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|VALOR_TOTAL 2|."
  Select_Adodc_Grid DGInv, AdoInv, sSQL
  
  sSQL = "SELECT C.Cliente,T.Fecha_Efec,T.Cheq_Dep,T.Haber As Monto,T.TP,T.Numero,CC.Cuenta " _
       & "FROM Transacciones As T,Catalogo_Cuentas As CC,Clientes As C " _
       & "WHERE T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Fecha <= #" & FechaFin & "# " _
       & "AND T.C = " & Val(adFalse) & " " _
       & "AND CC.TC = 'BA' " _
       & "AND T.Haber > 0 " _
       & "AND T.T <> 'A' " _
       & "AND T.Item = CC.Item " _
       & "AND T.Periodo = CC.Periodo " _
       & "AND T.Cta = CC.Codigo " _
       & "AND T.Codigo_C = C.Codigo " _
       & "ORDER BY T.Cta,C.Cliente,T.Fecha_Efec,T.Cheq_Dep "
  Select_Adodc_Grid DGCheques, AdoCheques, sSQL
  
 'Fin de subCtas
  SumaDebe = 0: SumaHaber = 0
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY A_No "
  Select_Adodc_Grid DGBalance, AdoCtas, SQL2
  DGBalance.Visible = False
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .fields("DEBE")
          SumaHaber = SumaHaber + .fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelTotInv.Caption = Format(Total, "#,##0.00")
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  CierreEjercicio.Caption = "CIERRE DEL EJERCICIO"
  RatonNormal
  DGBanco.Visible = True
  DGBalance.Visible = True
  Progreso_Final
  If Round(SumaDebe - SumaHaber) <> 0 Then
     MsgBox "Usuario: " & NombreUsuario & vbCrLf & vbCrLf _
          & "No se puede Cerrar el Ejercicio Contable," & vbCrLf & vbCrLf _
          & "Revise el Catalogo de Cuentas."
  End If
End Sub

Private Sub Imprimir()
  RatonReloj
  DGBalance.Visible = False
  MensajeEncabado = "DIARIO PRELIMINAR DE CIERRE "
  SQLMsg1 = "AL " & FechaStrg(FechaFin)
  ImprimirAdodc AdoCtas, 1, 9
  DGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command3_Click()
  Unload CierreEjercicio
End Sub

Private Sub Grabar_Cierre_Ejercicio()
Dim SiTieneDatos As Boolean
Dim FechaFact As String
Control_Procesos Normal, "Cierre del Periodo al " & FechaFinal
Progreso_Barra.Incremento = 0
Progreso_Barra.Valor_Maximo = 100
Progreso_Barra.Mensaje_Box = "CERRANDO Y GRABANDO ASIENTO INICIAL"
Progreso_Esperar
FechaFact = Format(Day(FechaFinal), "00") & "/" & Format(Month(FechaFinal), "00") & "/" & Format(Year(FechaFinal) - 7, "0000")
'MsgBox FechaFact
Diferencia = SumaDebe - SumaHaber
Si_No = False
sSQL = "SELECT * " _
     & "FROM Asiento " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND CodigoU = '" & CodigoUsuario & "' " _
     & "AND T_No = " & Trans_No & " " _
     & "ORDER BY A_No "
Select_Adodc_Grid DGBalance, AdoCtas, sSQL
If AdoCtas.Recordset.RecordCount > 0 Then Si_No = True
If Diferencia = 0 Then
  'MsgBox FechaFin
   CierreEjercicio.Caption = "Cerrando cuentas"
  'Obtenemos el numero de comprobantes
   RatonReloj
  'Averiguamos si ya se cerro el periodo
   sSQL = "SELECT Periodo " _
        & "FROM Comprobantes " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & FechaFin & "' "
   Select_Adodc AdoTrans, sSQL
   If AdoTrans.Recordset.RecordCount > 0 Then
      MsgBox "No se puede volver a cerrar el mismo PERIODO" & vbCrLf & vbCrLf _
           & TextoBusqueda
   Else
     'Cerramos las Subcuentas
      sSQL = "UPDATE Trans_SubCtas " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
            
      sSQL = "UPDATE Trans_Rol_Horas " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Rol_de_Pagos " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha_H <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Entrada_Salida " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Gastos_Caja " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Ventas " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Compras " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar

      sSQL = "UPDATE Trans_Air " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Exportaciones " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Importaciones " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Trans_Kardex " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
                  
      sSQL = "UPDATE Comprobantes " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      
      sSQL = "UPDATE Transacciones " _
           & "SET Periodo = '" & FechaFinal & "' " _
           & "WHERE Fecha <= #" & FechaFin & "# " _
           & "AND Periodo = '.' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Bodegas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Cuentas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_CxCxP", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_SubCtas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Lineas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Marcas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Productos", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Recetas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Rol_Cuentas", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Rol_Pagos", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Catalogo_Rol_Rubros", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Clientes_Facturacion", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Codigos", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      CopiarAdoTablaPeriodo "Ctas_Proceso", FechaFinal
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
      Progreso_Esperar
      If CheqRenumerar.value <> 0 Then
         For I = 1 To 12
             Numero = CLng(Format(I, "00") & "000001")
             'MsgBox Numero
             sSQL = "SELECT Concepto, Numero, ID " _
                  & "FROM Codigos " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '.' " _
                  & "AND (Concepto LIKE '" & Format(I, "00") & "%') " _
                  & "ORDER BY Concepto "
             Select_Adodc AdoAux, sSQL
             With AdoAux.Recordset
              If .RecordCount > 0 Then
                  Do While Not .EOF
                    .fields("Numero") = Numero
                    .MoveNext
                  Loop
                 .UpdateBatch
              End If
             End With
         Next I
''''         sSQL = "DELETE * " _
''''              & "FROM Codigos " _
''''              & "WHERE Item = '" & NumEmpresa & "' " _
''''              & "AND Periodo = '" & Ninguno & "' " _
''''              & "AND NOT (Concepto LIKE '%_SERIE_%') "
''''         Ejecutar_SQL_SP sSQL
''''
''''         sSQL = "INSERT INTO Codigos (Item, Concepto, Numero, Periodo, X) " _
''''              & "SELECT '" & NumEmpresa & "', Concepto, Numero, Periodo, X " _
''''              & "FROM Codigos " _
''''              & "WHERE Item = '000' " _
''''              & "AND Periodo = '.' " _
''''              & "ORDER BY Concepto, Numero "
''''         Ejecutar_SQL_SP sSQL
      Else
         CopiarAdoTablaPeriodo "Codigos", FechaFinal
      End If
      Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
      Progreso_Esperar
      CodigoBenef = Ninguno
      Asiento = 0
      sSQL = "SELECT * " _
           & "FROM Asiento " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' " _
           & "AND T_No = " & Trans_No & " " _
           & "ORDER BY A_No "
      Select_Adodc_Grid DGBalance, AdoCtas, sSQL
      If AdoCtas.Recordset.RecordCount > 0 Then Si_No = True
      Progreso_Esperar
      If Si_No Then
         NumComp = 1
         RatonReloj
         Co.TP = CompDiario
         Co.T = Normal
         Co.Fecha = FechaTexto
         Co.Numero = NumComp
         Co.Monto_Total = SumaDebe
         Co.Concepto = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaTexto)
         Co.CodigoB = Ninguno
         Co.Efectivo = SumaDebe
         Co.Cotizacion = 0
         Co.Item = NumEmpresa
         Co.Usuario = CodigoUsuario
         Co.T_No = Trans_No
         GrabarComprobante Co
         Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
         Progreso_Esperar
         RatonNormal
         ImprimirComprobantesDe False, Co
      End If
      
      sSQL = "UPDATE Catalogo_Cuentas " _
           & "SET Procesado = " & Val(adTrue) & " " _
           & "WHERE Periodo = '" & FechaFinal & "' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
      
      Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
      Progreso_Esperar
      RatonNormal
      MsgBox "Fin del Cierre del Ejercicio"
      Unload CierreEjercicio
      RatonReloj
      Control_Procesos Normal, "Mayorizar Cuentas por cierre"
      'Mayorizacion.Show
   End If
Else
   MsgBox "No se puede cerrar el ejercicio Contable"
End If
End Sub

Private Sub Actualizar_Cierre_Ejercicio()
Dim SiTieneDatos As Boolean
Dim FechaFact As String
If ClaveContador Then
    Control_Procesos Normal, "Actualizar Cierre del Periodo al " & FechaFinal
    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = 100
    Progreso_Barra.Mensaje_Box = "CERRANDO Y GRABANDO ASIENTO INICIAL"
    Progreso_Esperar
    Diferencia = SumaDebe - SumaHaber
    Si_No = False
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY A_No "
    Select_Adodc_Grid DGBalance, AdoCtas, sSQL
    If AdoCtas.Recordset.RecordCount > 0 Then Si_No = True
    Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
    Progreso_Esperar
    If (Diferencia = 0) And Si_No Then
       NumComp = 1
       CambioPeriodo.Show 1
       'Periodo_Contable = Ninguno
       FechaTexto = FechaFinal
       FechaFin = Format(Day(FechaFinal), "00") & "/" & Format(Month(FechaFinal), "00") & "/" & Format(Year(FechaFinal), "0000")
       CierreEjercicio.Caption = "Cerrando Cuentas"
      'Averiguamos si hay que actualizar el asiento inicial
       sSQL = "SELECT * " _
            & "FROM Comprobantes " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Numero = " & NumComp & " " _
            & "AND TP = 'CD' "
       Select_Adodc AdoTrans, sSQL
       If AdoTrans.Recordset.RecordCount > 0 Then
          Mensajes = "USTED AL VOLVER A CERRAR EL MISMO PERIODO" & vbCrLf & vbCrLf _
                   & "SE ACTUALIZARA EL ASIENTO DE CIERRE CON" & vbCrLf & vbCrLf _
                   & "LOS NUEVOS VALORES QUE SE PRESENTAN EN PANTALLA" & vbCrLf & vbCrLf _
                   & "Y LUEGO VOLVERA AUTOMATICAMENTE AL PERIODO ACTUAL" & vbCrLf & vbCrLf _
                   & "PARA LUEGO MAYORIZAR AUTOMATICAMENTE EL PERIODO." & vbCrLf & vbCrLf & vbCrLf _
                   & "Esta seguro de Grabar Cierre del Periodo"
          Titulo = "Pregunta de Grabación"
          If BoxMensaje = vbYes Then
             RatonReloj
             Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
             Progreso_Esperar
             CodigoBenef = Ninguno
             Asiento = 0
             RatonReloj
             Co.TP = CompDiario
             Co.T = Normal
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             Co.Monto_Total = SumaDebe
             Co.Concepto = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaTexto)
             Co.CodigoB = Ninguno
             Co.Efectivo = SumaDebe
             Co.Cotizacion = 0
             Co.Item = NumEmpresa
             Co.Usuario = CodigoUsuario
             Co.T_No = Trans_No
            'Eliminamos el Cierre anterior
             EliminarComprobantes Co
             Progreso_Barra.Incremento = Progreso_Barra.Incremento + 30
             Progreso_Esperar
            'Procedemos a grabar el comprobante
             GrabarComprobante Co
             Progreso_Barra.Incremento = Progreso_Barra.Incremento + 30
             Progreso_Esperar
             RatonNormal
             ImprimirComprobantesDe False, Co
             Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
             Progreso_Esperar
             PonerDirEmpresa
             Ver_Grafico_FormPict
             Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
             Progreso_Esperar
             RatonNormal
             MsgBox "Fin del Cierre del Ejercicio"
             Unload CierreEjercicio
             RatonReloj
             Control_Procesos Normal, "Mayorizar por actualizar cierre"
             Mayorizar_Cuentas_SP
             
             
             'CLng(MesComp & "000001")
          End If
       Else
          MsgBox "No se puede actualizar el cierre del periodo " & Periodo_Contable
       End If
    Else
       MsgBox "No se puede cerrar el ejercicio Contable " & Periodo_Contable
    End If
End If
End Sub

Private Sub DGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGBalance.Visible = False
     GenerarDataTexto CierreEjercicio, AdoCtas
     DGBalance.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyA Then
     DGBalance.Visible = False
     Codigo1 = DGBalance.Columns(0)
     Mensajes = "CONCILIAR CUENTA DE AHORROS" & vbCrLf & vbCrLf _
              & Codigo1 & " - " & DGBalance.Columns(1)
     Titulo = "Pregunta de Actualizacion"
     If BoxMensaje = vbYes Then
        If Len(Codigo1) > 2 Then
           sSQL = "UPDATE Transacciones " _
                & "SET C = " & Val(adTrue) & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Cta = '" & Codigo1 & "' "
           Ejecutar_SQL_SP sSQL
           MsgBox "Proceso realizado, Vuelva a General el cierre"
        End If
     End If
     DGBalance.Visible = True
  End If
End Sub

Private Sub DGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGBanco.Visible = False
     GenerarDataTexto CierreEjercicio, AdoBanco
     DGBanco.Visible = True
  End If
End Sub

Private Sub DGCheques_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGCheques.Visible = False
     GenerarDataTexto CierreEjercicio, AdoCheques
     DGCheques.Visible = True
  End If
End Sub

Private Sub DGInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGInv.Visible = False
     GenerarDataTexto CierreEjercicio, AdoInv
     DGInv.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  SSTab1.Height = MDI_Y_Max - DGBalance.Top - 1800
  SSTab1.width = MDI_X_Max - DGBalance.Left - 10
  
  SSTab1.Tab = 3
  DGCheques.Height = SSTab1.Height - DGCheques.Top - 140
  DGCheques.width = SSTab1.width - DGCheques.Left - 140
  SSTab1.Tab = 2
  DGInv.Height = SSTab1.Height - DGInv.Top - 140
  DGInv.width = SSTab1.width - DGInv.Left - 140
  SSTab1.Tab = 1
  DGBanco.Height = SSTab1.Height - DGBanco.Top - 140
  DGBanco.width = SSTab1.width - DGBanco.Left - 140
  SSTab1.Tab = 0
  DGBalance.Height = SSTab1.Height - DGBalance.Top - 140
  DGBalance.width = SSTab1.width - DGBalance.Left - 140
  
  Label6.Top = SSTab1.Top + SSTab1.Height + 30
  Label9.Top = SSTab1.Top + SSTab1.Height + 30
  Label11.Top = SSTab1.Top + SSTab1.Height + 30
  
  LabelTotSaldo.Top = SSTab1.Top + SSTab1.Height + 30
  LabelTotDebe.Top = SSTab1.Top + SSTab1.Height + 30
  LabelTotHaber.Top = SSTab1.Top + SSTab1.Height + 30

  DGBalance.Visible = False
  TipoDoc = CompDiario
  Trans_No = 1
  IniciarAsientosDe DGBalance, AdoCtas
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Cta = 'I==============>' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Cta >= '4' "
  Ejecutar_SQL_SP sSQL
  
  FormatoMaskCta MBoxCtaI
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Detalle = 'Balance' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc AdoTrans, sSQL
  If AdoTrans.Recordset.RecordCount > 0 Then
     FechaInicial = AdoTrans.Recordset.fields("Fecha_Inicial")
     FechaFinal = AdoTrans.Recordset.fields("Fecha_Final")
     FechaIni = BuscarFecha(FechaInicial)
     FechaFin = BuscarFecha(FechaFinal)
     FechaTexto = FechaFinal
     Label5.Caption = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaFinal)
     If CheqDetalle.value = 1 Then
        DetalleComp = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaFinal)
     Else
        DetalleComp = Ninguno
     End If
  Else
     Label5.Caption = "NO SE PUEDE CERRAR EL EJERCICIO NO HAY FECHA DE CIERRE"
  End If
  If Periodo_Contable = Ninguno Then
     Toolbar1.buttons("Grabar").Enabled = True
     Toolbar1.buttons("Actualizar").Enabled = False
  Else
     Toolbar1.buttons("Grabar").Enabled = False
     Toolbar1.buttons("Actualizar").Enabled = True
  End If
  If Bloquear_Control Then
     Toolbar1.buttons("Procesar").Enabled = False
     Toolbar1.buttons("Grabar").Enabled = False
     Toolbar1.buttons("Actualizar").Enabled = False
     Toolbar1.buttons("Imprimir").Enabled = False
  End If
  RatonReloj
  SiTieneSubModulo = False
  sSQL = "SELECT Item, COUNT(Item) As ContSM " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> 'A' " _
       & "AND TC IN ('C','P') " _
       & "GROUP BY Item "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     If AdoAux.Recordset.fields("ContSM") > 0 Then SiTieneSubModulo = True
  End If
  DGBalance.Visible = True
  RatonNormal
  MBoxCtaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm CierreEjercicio
  ConectarAdodc AdoInv
  ConectarAdodc AdoRet
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoTrans
  ConectarAdodc AdoEstRes
  ConectarAdodc AdoBalGen
  ConectarAdodc AdoIngKar
  ConectarAdodc AdoCheques
  ConectarAdodc AdoFechaBal
  ConectarAdodc AdoBalGenCon
  ConectarAdodc AdoSubCtaDet
  RatonNormal
End Sub

Private Sub MBoxCtaI_GotFocus()
  MarcarTexto MBoxCtaI
End Sub

Private Sub MBoxCtaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCtaI_LostFocus()
    Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
    If MidStrg(Codigo1, 1, 1) = "3" Then
       Codigo = Leer_Cta_Catalogo(Codigo1)
       If Cuenta = Ninguno Then
          MsgBox "Cuenta no asignada en el Catalogo de Cuentas"
          Toolbar1.buttons("Procesar").Enabled = False
          MBoxCtaI.SetFocus
       Else
          If TipoCta <> "D" Then
             MsgBox "Existe Cuenta de cierre pero no es de detalle, no se puede hacer el proceso"
             Toolbar1.buttons("Procesar").Enabled = False
             MBoxCtaI.SetFocus
          Else
             Toolbar1.buttons("Procesar").Enabled = True
          End If
       End If
    Else
       MsgBox "Advertencia: Solo se admiten Cuentas Patrimoniales de tipo '3'"
       Toolbar1.buttons("Procesar").Enabled = False
       MBoxCtaI.SetFocus
    End If
    
    Toolbar1.Refresh
End Sub

