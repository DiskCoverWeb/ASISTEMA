VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CierreEjercicio 
   Caption         =   "BALANCE DE COMPROBACION"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5790
      Left            =   105
      TabIndex        =   10
      Top             =   1575
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   10213
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&ASIENTO CONTABLE"
      TabPicture(0)   =   "Cierreej.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelTotDebe"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LabelTotHaber"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelTotSaldo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "AdoCtas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DGBalance"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "ASIENTO DE SUBCUENTAS DE &BLOQUE"
      TabPicture(1)   =   "Cierreej.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "AdoBanco"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DGBanco"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin MSDataGridLib.DataGrid DGBanco 
         Bindings        =   "Cierreej.frx":0038
         Height          =   4425
         Left            =   -74895
         TabIndex        =   24
         Top             =   420
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   7805
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGBalance 
         Bindings        =   "Cierreej.frx":004F
         Height          =   4425
         Left            =   105
         TabIndex        =   23
         Top             =   420
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   7805
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
      Begin MSAdodcLib.Adodc AdoBanco 
         Height          =   330
         Left            =   -74895
         Top             =   4830
         Width           =   11250
         _ExtentX        =   19844
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
         Left            =   105
         Top             =   4830
         Width           =   11250
         _ExtentX        =   19844
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   -73425
         TabIndex        =   21
         Top             =   5265
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         TabIndex        =   20
         Top             =   5265
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   -65760
         TabIndex        =   19
         Top             =   5265
         Width           =   2010
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Debe - Haber:"
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
         Left            =   -67545
         TabIndex        =   11
         Top             =   5265
         Width           =   1800
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Haber:"
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
         Left            =   -71220
         TabIndex        =   18
         Top             =   5265
         Width           =   1485
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Debe:"
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
         TabIndex        =   22
         Top             =   5265
         Width           =   1485
      End
      Begin VB.Label LabelTotSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   9240
         TabIndex        =   12
         Top             =   5265
         Width           =   2010
      End
      Begin VB.Label LabelTotHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   5250
         TabIndex        =   14
         Top             =   5265
         Width           =   1905
      End
      Begin VB.Label LabelTotDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
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
         Left            =   1575
         TabIndex        =   16
         Top             =   5265
         Width           =   1905
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Debe:"
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
         Top             =   5265
         Width           =   1485
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Haber:"
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
         Left            =   3780
         TabIndex        =   15
         Top             =   5265
         Width           =   1485
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Debe - Haber:"
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
         Left            =   7455
         TabIndex        =   13
         Top             =   5265
         Width           =   1800
      End
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   1155
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
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
      Left            =   10500
      Picture         =   "Cierreej.frx":0065
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
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
      Picture         =   "Cierreej.frx":02E7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
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
      Height          =   960
      Left            =   8190
      Picture         =   "Cierreej.frx":0951
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
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
      Left            =   7035
      Picture         =   "Cierreej.frx":0D93
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   525
      Width           =   1065
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   420
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   420
      TabIndex        =   7
      Top             =   4935
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxCtaI 
      Height          =   330
      Left            =   2310
      TabIndex        =   2
      Top             =   630
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
   Begin MSAdodcLib.Adodc AdoEstRes 
      Height          =   330
      Left            =   210
      Top             =   4200
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
      Left            =   210
      Top             =   3255
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
      Left            =   210
      Top             =   2940
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   210
      Top             =   2625
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
   Begin MSAdodcLib.Adodc AdoFechaBal 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Left            =   210
      Top             =   1995
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
      Left            =   210
      Top             =   1680
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
      TabIndex        =   1
      Top             =   630
      Width           =   2220
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
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   11460
   End
End
Attribute VB_Name = "CierreEjercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Public Sub Refreshado()
  ConectarAdodc AdoRet
  ConectarAdodc AdoCta
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoTrans
  ConectarAdodc AdoEstRes
  ConectarAdodc AdoBalGen
  ConectarAdodc AdoIngKar
  ConectarAdodc AdoFechaBal
  ConectarAdodc AdoBalGenCon
  ConectarAdodc AdoSubCtaDet
End Sub

Public Sub InsertarTotales(AdoG As Adodc, Cod, Cta, Anal, Parc, Tot)
   With AdoG.Recordset
        .AddNew
        .Fields("Codigo") = Cod
        .Fields("Cuenta") = Cta
        .Fields("Analitico") = Anal
        .Fields("Parcial") = Parc
        .Fields("Total") = Tot
        .Fields("Item") = NumEmpresa
        .Update
   End With
End Sub

Public Sub InsertarTotalesCon(AdoG As Adodc, Cod, CtaDG, Cta, Sal_ME, Sal_MN, Tot)
   With AdoG.Recordset
        .AddNew
        .Fields("Codigo") = Cod
        .Fields("Cuenta") = Cta
        .Fields("Saldo_ME") = Sal_ME
        .Fields("Saldo_MN") = Sal_MN
        .Fields("Total") = Tot
        .Fields("DG") = CtaDG
        .Fields("Item") = NumEmpresa
        .Update
   End With
End Sub

Private Sub Command1_Click()
  RatonReloj
  Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
  SQL1 = "DELETE * FROM Asiento_SC_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  ConectarAdoExecute SQL1
  SQL1 = "DELETE * FROM Asiento_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  ConectarAdoExecute SQL1
  If OpcCoop Then
     SQL2 = "SELECT CODIGO,CUENTA,DEBE,HABER,ME "
  Else
     SQL2 = "SELECT CODIGO,CUENTA,PARCIAL_ME,DEBE,HABER,ME "
  End If
  SQL2 = SQL2 & "FROM Asiento_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  SelectAdodc AdoCtas, SQL2, False
  
  sSQL = "SELECT * FROM Catalogo_Cuentas "
  sSQL = sSQL & "WHERE Codigo = '" & Codigo1 & "' " _
       & "WHERE Item = " & NumEmpresa & " "
  SelectAdodc AdoCta, sSQL, False
  If AdoCta.Recordset.RecordCount > 0 Then
     Sub_Cuenta = AdoCta.Recordset.Fields("Cuenta")
     Moneda_E = AdoCta.Recordset.Fields("ME")
     
     SumaDebe = 0: SumaHaber = 0
     sSQL = "SELECT ME,DG, Codigo, Cuenta, Saldo_Total, Saldo_Total_ME "
     sSQL = sSQL & "FROM Catalogo "
     sSQL = sSQL & "ORDER BY Codigo "
     SelectAdodc AdoTrans, sSQL, False
     DBGBalance.Visible = False
     With AdoTrans.Recordset
          SetProgBar ProgBarra, .RecordCount
          Do While Not .EOF
             Debe = 0: Haber = 0
             TipoCta = .Fields("DG")
             Moneda_US = .Fields("ME")
             Codigo = .Fields("Codigo")
             Cuenta = .Fields("Cuenta")
             Total = Round(.Fields("Saldo_Total"))
             Total_ME = Round(.Fields("Saldo_Total_ME"))
             If OpcCoop Then
                If Moneda_US Then
                   Select Case Mid(Codigo, 1, 1)
                   Case "1": If Total_ME >= 0 Then Debe = Total_ME Else Haber = -Total_ME
                   Case "2": If Total_ME >= 0 Then Haber = Total_ME Else Debe = -Total_ME
                   Case "3": If Total_ME >= 0 Then Haber = Total_ME Else Debe = -Total_ME
                   End Select
                Else
                   Select Case Mid(Codigo, 1, 1)
                   Case "1": If Total >= 0 Then Debe = Total Else Haber = -Total
                   Case "2": If Total >= 0 Then Haber = Total Else Debe = -Total
                   Case "3": If Total >= 0 Then Haber = Total Else Debe = -Total
                   End Select
                End If
             Else
                Select Case Mid(Codigo, 1, 1)
                Case "1": If Total >= 0 Then Debe = Total Else Haber = -Total
                Case "2": If Total >= 0 Then Haber = Total Else Debe = -Total
                Case "3": If Total >= 0 Then Haber = Total Else Debe = -Total
                End Select
             End If
             If (Debe - Haber) <> 0 And TipoCta = "D" And Codigo < "4" Then
                AdoCtas.Recordset.AddNew
                AdoCtas.Recordset.Fields("ME") = Moneda_US
                AdoCtas.Recordset.Fields("CODIGO") = Codigo
                AdoCtas.Recordset.Fields("CUENTA") = Cuenta
                If OpcCoop = False Then
                   AdoCtas.Recordset.Fields("PARCIAL_ME") = Total_ME
                End If
                AdoCtas.Recordset.Fields("DEBE") = Debe
                AdoCtas.Recordset.Fields("HABER") = Haber
                AdoCtas.Recordset.Update
                SumaDebe = SumaDebe + Debe
                SumaHaber = SumaHaber + Haber
             End If
             Select Case .Fields("Codigo")
               Case "4"
                    TotalIngreso = Total
                    TotalIngreso_ME = Total_ME
               Case "5"
                    TotalEgreso = Total
                    TotalEgreso_ME = Total_ME
             End Select
             IncProgBar ProgBarra
            .MoveNext
          Loop
     End With
     If OpcCoop Then
        SQL2 = "SELECT CODIGO,CUENTA,DEBE,HABER,ME FROM Asiento_"
     Else
        SQL2 = "SELECT CODIGO,CUENTA,PARCIAL_ME,DEBE,HABER,ME FROM Asiento_"
     End If
     SQL2 = SQL2 & CodigoUsuario & " "
     SelectAdodc AdoCtas, SQL2
     With AdoCtas.Recordset
       SumaDebe = 0: SumaHaber = 0
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
     End With
     AdoCtas.Recordset.AddNew
     AdoCtas.Recordset.Fields("ME") = Moneda_E
     AdoCtas.Recordset.Fields("CODIGO") = Codigo1
     AdoCtas.Recordset.Fields("CUENTA") = Sub_Cuenta
     If OpcCoop = False Then
        AdoCtas.Recordset.Fields("PARCIAL_ME") = TotalIngreso_ME - TotalEgreso_ME
     End If
     SumaSaldo = SumaDebe - SumaHaber
     Debe = 0: Haber = 0
     If SumaSaldo >= 0 Then Haber = SumaSaldo Else Debe = -SumaSaldo
     AdoCtas.Recordset.Fields("DEBE") = Debe
     AdoCtas.Recordset.Fields("HABER") = Haber
     AdoCtas.Recordset.Update
     SelectDataGrid DGBalance, AdoCtas, SQL2
     SumaDebe = SumaDebe + Debe
     SumaHaber = SumaHaber + Haber
     DBGBalance.Visible = True
  Else
     MsgBox "La cuenta de Utilidad o perdida no esta definida," & vbCrLf _
            & "por lo tanto no se relizara El Cierre del Ejercicio."
  End If
 'Asiento de SubCtas
  RatonReloj
  TipoSubCta = Ninguno
  DBGBanco.Visible = False
  sSQL = "DELETE * FROM Saldo_SubCtas_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * FROM Saldo_SubCtas_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  SelectAdodc AdoBanco, sSQL
  FechaFinal = BuscarFecha(FechaFin)
  sSQL = "SELECT TC,Codigo,Beneficiario "
  sSQL = sSQL & "FROM Catalogo_SubCtas "
  sSQL = sSQL & "ORDER BY TC,Codigo "
  SelectAdodc AdoTrans, sSQL
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Saldo = 0: Saldo_ME = 0: Valor = 0
          Codigo = .Fields("Codigo")
          Cuenta = .Fields("Beneficiario")
          TipoSubCta = .Fields("TC")
          CierreEjercicio.Caption = "SALDO DE SUBCUENTA: " & Cuenta & "."
          Select Case TipoSubCta
          Case "C", "P"
          sSQL = "SELECT Cta,Factura,Fecha,TP,Numero,(Debitos-Creditos) As Valor_MN," _
               & "(Debitos_ME-Creditos_ME) As Valor_ME,ID " _
               & "FROM Trans_SubCtas " _
               & "WHERE Fecha <= #" & FechaFinal & "# " _
               & "AND Codigo = '" & Codigo & "' " _
               & "AND TC = '" & TipoSubCta & "' " _
               & "AND T = '" & Normal & "' " _
               & "ORDER BY Cta,Codigo,Factura,Fecha,TP,Numero,Debitos DESC,Creditos,ID "
          SelectAdodc AdoSubCtaDet, sSQL, False
          If AdoSubCtaDet.Recordset.RecordCount > 0 Then
             Cta = AdoSubCtaDet.Recordset.Fields("Cta")
             Saldo = 0: Saldo_ME = 0: Valor = 0
             Do While Not AdoSubCtaDet.Recordset.EOF
                If Cta <> AdoSubCtaDet.Recordset.Fields("Cta") Then
                   If ((Round(Saldo) <> 0) Or (Round(Saldo_ME) <> 0)) Then
                      AdoBanco.Recordset.AddNew
                      AdoBanco.Recordset.Fields("TC") = TipoSubCta
                      AdoBanco.Recordset.Fields("Cta") = Cta
                      AdoBanco.Recordset.Fields("Codigo") = Codigo
                      AdoBanco.Recordset.Fields("Cuenta") = Cuenta
                      AdoBanco.Recordset.Fields("Saldo_ME") = Saldo_ME
                      AdoBanco.Recordset.Fields("Saldo_MN") = Saldo
                      AdoBanco.Recordset.Fields("Saldo") = Valor
                      AdoBanco.Recordset.Fields("Item") = NumEmpresa
                      AdoBanco.Recordset.Update
                   End If
                   Cta = AdoSubCtaDet.Recordset.Fields("Cta")
                   Saldo = 0: Saldo_ME = 0: Valor = 0
                End If
                Saldo_ME = Saldo_ME + AdoSubCtaDet.Recordset.Fields("Valor_ME")
                Valor = Valor + AdoSubCtaDet.Recordset.Fields("Valor_MN")
                If AdoSubCtaDet.Recordset.Fields("Valor_ME") = 0 Then
                   Saldo = Saldo + AdoSubCtaDet.Recordset.Fields("Valor_MN")
                End If
                AdoSubCtaDet.Recordset.MoveNext
             Loop
             If ((Round(Saldo) <> 0) Or (Round(Saldo_ME) <> 0)) Then
                AdoBanco.Recordset.AddNew
                AdoBanco.Recordset.Fields("TC") = TipoSubCta
                AdoBanco.Recordset.Fields("Cta") = Cta
                AdoBanco.Recordset.Fields("Codigo") = Codigo
                AdoBanco.Recordset.Fields("Cuenta") = Cuenta
                AdoBanco.Recordset.Fields("Saldo_ME") = Saldo_ME
                AdoBanco.Recordset.Fields("Saldo_MN") = Saldo
                AdoBanco.Recordset.Fields("Saldo") = Valor
                AdoBanco.Recordset.Fields("Item") = NumEmpresa
                AdoBanco.Recordset.Update
             End If
          End If
          End Select
         .MoveNext
       Loop
   End If
  End With
  CierreEjercicio.Caption = "SALDO DE SUBCUENTAS DE BLOQUE"
  sSQL = "SELECT TC,Cta,Codigo,Cuenta,Saldo_ME,Saldo_MN,Saldo " _
       & "FROM Saldo_SubCtas_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " " _
       & "ORDER BY TC,Cta,Codigo "
  SelectDataGrid DGBanco, AdoBanco, sSQL
  SQL1 = "SELECT * FROM Asiento_SC_" & CodigoUsuario & " " _
       & "WHERE Item = " & NumEmpresa & " "
  SelectAdodc AdoSubCtaDet, SQL1
  With AdoSubCtaDet.Recordset
    If AdoBanco.Recordset.RecordCount > 0 Then
       AdoBanco.Recordset.MoveFirst
       Do While Not AdoBanco.Recordset.EOF
          OpcTM = 1
          If AdoBanco.Recordset.Fields("Saldo_ME") <> 0 Then OpcTM = 2
          ValorDH = AdoBanco.Recordset.Fields("Saldo_MN")
          Valor = AdoBanco.Recordset.Fields("Saldo_ME")
          SubCta = AdoBanco.Recordset.Fields("TC")
          Codigo = AdoBanco.Recordset.Fields("Codigo")
          Cuenta = AdoBanco.Recordset.Fields("Cuenta")
          SubCtaGen = AdoBanco.Recordset.Fields("Cta")
          If ValorDH >= 0 Then
             OpcDH = 1
          Else
             OpcDH = 2
             ValorDH = -ValorDH
          End If
          If Valor < 0 Then Valor = -Valor
         .AddNew
         .Fields("FECHA_V") = FechaFin
         .Fields("TC") = SubCta
         .Fields("Factura") = 0
         .Fields("Codigo") = Codigo
         .Fields("Beneficiario") = Cuenta
         .Fields("Cta") = SubCtaGen
         .Fields("DH") = OpcDH
         .Fields("Valor") = ValorDH
         .Fields("Valor_ME") = 0
         .Fields("TM") = OpcTM
         .Fields("Valor_ME") = Valor
         .Fields("Item") = NumEmpresa
         .Update
          AdoBanco.Recordset.MoveNext
       Loop
    End If
  End With
  DBGBanco.Visible = True
  RatonNormal
 'Fin de subCtas
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  CierreEjercicio.Caption = "CIERRE DEL EJERCICIO"
  RatonNormal
  If Round(SumaDebe - SumaHaber) <> 0 Then
     MsgBox "Usuario: " & NombreUsuario & vbCrLf & vbCrLf _
          & "No se puede Cerrar el Ejercicio Contable," & vbCrLf & vbCrLf _
          & "Revise el Catalogo de Cuentas."
  End If
End Sub

Private Sub Command2_Click()
  RatonReloj
  DBGBalance.Visible = False
  MensajeEncabado = "DIARIO PRELIMINAR DE CIERRE "
  SQLMsg1 = "AL " & FechaStrg(FechaFin)
  ImprimirAdodc AdoCtas, True, 1, 8
  DBGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command3_Click()
  Unload CierreEjercicio
End Sub

Private Sub Command4_Click()
Dim SiTieneDatos As Boolean
Diferencia = SumaDebe - SumaHaber
If (Diferencia = 0) And (AdoCtas.Recordset.RecordCount > 0) Then
  FechaTexto = FechaFin
  SiTieneDatos = False
  sSQL = "SELECT * FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & Codigo1 & "' " _
       & "AND Item = " & NumEmpresa & " "
  SelectAdodc AdoCtas, sSQL
  If AdoCtas.Recordset.RecordCount > 0 Then SiTieneDatos = True
  If SiTieneDatos Then
     RatonReloj
     CierreEjercicio.Caption = "Cerrando las Bases"
     RutaCierreEjercicio = RutaEmpresa & "\" & FechaDeCierre(FechaFin)
     If Dir(RutaCierreEjercicio, vbDirectory) = "" Then
        MkDir RutaCierreEjercicio
     Else
        If Dir(RutaCierreEjercicio & "\*.MDB", vbNormal) <> "" Then Kill RutaCierreEjercicio & "\*.*"
        RmDir RutaCierreEjercicio
        MkDir RutaCierreEjercicio
     End If
     CierreEjercicio.Caption = "Copiando bases requeridas"
     Dir1.Path = RutaEmpresa
     File1.FileName = Dir1.Path & "\*.MDB"
     For Ind = 0 To File1.ListCount - 1
         RutaOrigen = Dir1.Path & "\" & File1.List(Ind)
         RutaDestino = RutaCierreEjercicio & "\" & File1.List(Ind)
         CierreEjercicio.Caption = "Copiando " & RutaOrigen & " en " & RutaDestino
         FileCopy RutaOrigen, RutaDestino
     Next Ind
     RutaEmpresa = RutaCierreEjercicio
     Refreshado
     CierreEjercicio.Caption = "Cerrando cuentas"
     sSQL = "DELETE * FROM Bancos " _
          & "WHERE Fecha > #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Retenciones " _
          & "WHERE Fecha > #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Comprobantes " _
          & "WHERE Fecha > #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Transacciones " _
          & "WHERE Fecha > #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM TransaccionesSC " _
          & "WHERE Fecha > #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
    'Obtenemos el numero de comprobantes
     NumComp = 1
     sSQL = "SELECT * FROM Comprobantes " _
          & "WHERE TP = 'CD' " _
          & "ORDER BY Numero,Fecha "
     SelectAdodc AdoTrans, sSQL
     If AdoTrans.Recordset.RecordCount > 0 Then
        NumComp = AdoTrans.Recordset.Fields("Numero") - 1
        If NumComp <= 0 Then NumComp = 1
        If NumComp > 1 Then NumComp = 1
     End If
    'Borramos todos los movimientos del ejercicio anterior
     RutaEmpresa = RutaEmpresaOld
     Refreshado
     sSQL = "DELETE * FROM Bancos " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Retenciones " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Comprobantes " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Transacciones " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM TransaccionesSC " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# " _
          & "AND Mid(Cta,1,1) = '4' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM TransaccionesSC " _
          & "WHERE Fecha <= #" & BuscarFecha(FechaFin) & "# " _
          & "AND Mid(Cta,1,1) = '5' "
     ConectarAdoExecute sSQL
     RatonNormal
     CierreEjercicio.Caption = "Grabando Comprobante de Diario No. " & NumComp
     Co.T = Normal
     Co.TP = CompDiario
     Co.Fecha = FechaTexto
     Co.Numero = NumComp
     Co.Concepto = "Asiento de Cierre al " & FechaTexto
     Co.Beneficiario = Ninguno
     Co.RUC_CI = LimpiarRUC
     Co.Efectivo = 0
     Co.Monto_Total = SumaDebe
     If OpcCoop Then
        SQL2 = "SELECT CODIGO,CUENTA,DEBE,HABER,ME FROM Asiento_"
     Else
        SQL2 = "SELECT CODIGO,CUENTA,PARCIAL_ME,DEBE,HABER,ME FROM Asiento_"
     End If
     SQL1 = "SELECT * FROM Asiento_SC_" & CodigoUsuario & " "
     SQL2 = SQL2 & CodigoUsuario & " "
     SelectAdodc AdoSubCtaDet, SQL1
     SelectAdodc AdoCtas, SQL2
     GrabarComprobante Co, AdoCtas
     Codigo1 = CambioCodigoCta(MBoxCtaI.Text)
'     AdoRet.AdoBase.Close
'     AdoCta.AdoBase.Close
'     AdoCtas.AdoBase.Close
'     AdoBanco.AdoBase.Close
'     AdoTrans.AdoBase.Close
'     AdoEstRes.AdoBase.Close
'     AdoBalGen.AdoBase.Close
'     AdoIngKar.AdoBase.Close
'     AdoFechaBal.AdoBase.Close
'     AdoBalGenCon.AdoBase.Close
'     AdoSubCtaDet.AdoBase.Close
     CierreEjercicio.Caption = "Copiando bases requeridas"
     Dir1.Path = RutaEmpresa
     File1.FileName = Dir1.Path & "\*.MDB"
     For Ind = 0 To File1.ListCount - 1
       RutaOrigen = RutaEmpresa & "\" & File1.List(Ind)
       RutaDestino = RutaCierreEjercicio & "\" & File1.List(Ind)
       CierreEjercicio.Caption = "Compactando " & RutaOrigen
       CompactarDatabase RutaOrigen, False
       CierreEjercicio.Caption = "Compactando " & RutaDestino
       CompactarDatabase RutaDestino, False
     Next Ind
     MsgBox "Fin del Cierre del Ejercicio"
     Unload CierreEjercicio
  End If
Else
  MsgBox "No se puede cerrar el ejercicio Contable"
End If
End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  IniciarAsientosDe DGBalance, AdoCtas, AdoSubCtaDet, AdoBanco, AdoRet, AdoIngKar
  FormatoMaskCta MBoxCtaI
  sSQL = "SELECT * FROM FechaBalance "
  SelectAdodc AdoTrans, sSQL
  If AdoTrans.Recordset.RecordCount > 0 Then
     FechaIni = AdoTrans.Recordset.Fields("Fecha_Inicial")
     FechaFin = AdoTrans.Recordset.Fields("Fecha_Final")
     FechaInicial = AdoTrans.Recordset.Fields("Fecha_Inicial")
     FechaFinal = AdoTrans.Recordset.Fields("Fecha_Final")
  End If
  Label5.Caption = "CIERRE DEL EJERCICIO AL " & FechaDeCierre(FechaFin)
  MBoxCtaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm CierreEjercicio
  Refreshado
  RatonNormal
End Sub

