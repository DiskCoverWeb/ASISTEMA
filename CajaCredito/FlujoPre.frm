VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FlujoDePrestamos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flujo de Prestamos"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Y"
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
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6300
      Width           =   225
   End
   Begin VB.CommandButton Command5 
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
      Left            =   10290
      Picture         =   "FlujoPre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3675
      Width           =   960
   End
   Begin VB.CommandButton Command4 
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
      Left            =   10290
      Picture         =   "FlujoPre.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2625
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Caja Anterior"
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
      Left            =   10290
      Picture         =   "FlujoPre.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   525
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "G&rabar Cierre"
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
      Left            =   10290
      Picture         =   "FlujoPre.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1575
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   105
      TabIndex        =   10
      Top             =   525
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10583
      _Version        =   393216
      TabOrientation  =   3
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Saldo de Prestamos"
      TabPicture(0)   =   "FlujoPre.frx":1458
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGVencidos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DGCreditos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DGDebitos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Saldo de Prestamos Resumindo"
      TabPicture(1)   =   "FlujoPre.frx":1474
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGResumen"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGResumen 
         Bindings        =   "FlujoPre.frx":1490
         Height          =   5790
         Left            =   -74895
         TabIndex        =   14
         Top             =   105
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   10213
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
         Caption         =   "RESUMEN DE PRESTAMOS"
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
      Begin MSDataGridLib.DataGrid DGDebitos 
         Bindings        =   "FlujoPre.frx":14A9
         Height          =   1905
         Left            =   105
         TabIndex        =   11
         Top             =   105
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3360
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
         Caption         =   "PRESTAMOS VIGENTES"
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
      Begin MSDataGridLib.DataGrid DGCreditos 
         Bindings        =   "FlujoPre.frx":14C2
         Height          =   2010
         Left            =   105
         TabIndex        =   12
         Top             =   1995
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3545
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
         Caption         =   "PRESTAMOS CANCELADOS"
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
      Begin MSDataGridLib.DataGrid DGVencidos 
         Bindings        =   "FlujoPre.frx":14DC
         Height          =   1905
         Left            =   105
         TabIndex        =   13
         Top             =   3990
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3360
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
         Caption         =   "PRESTAMOS VENCIDOS"
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
   Begin MSDataListLib.DataCombo DCTP 
      Bindings        =   "FlujoPre.frx":14F6
      DataSource      =   "AdoTP"
      Height          =   315
      Left            =   7035
      TabIndex        =   9
      Top             =   105
      Width           =   2745
      _ExtentX        =   4842
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
   Begin VB.CheckBox CheckTP 
      Caption         =   "&Tipo Proc"
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
      Left            =   5775
      TabIndex        =   7
      Top             =   105
      Width           =   1170
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1365
      TabIndex        =   1
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   315
      Top             =   1995
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   315
      Top             =   1680
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
      Caption         =   "TP"
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
   Begin MSAdodcLib.Adodc AdoVencidos 
      Height          =   330
      Left            =   315
      Top             =   1365
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
      Caption         =   "Vencidos"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   315
      Top             =   1050
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
      Caption         =   "Creditos"
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
   Begin MSAdodcLib.Adodc AdoDebitos 
      Height          =   330
      Left            =   315
      Top             =   735
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
      Caption         =   "Debitos"
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
   Begin MSAdodcLib.Adodc AdoResumen 
      Height          =   330
      Left            =   315
      Top             =   2310
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
      Caption         =   "Resumen"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3885
      TabIndex        =   3
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FechaFinal"
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
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Inicial"
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
      Width           =   1275
   End
End
Attribute VB_Name = "FlujoDePrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ListarFlujoPrestamos()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "SELECT P.Fecha,P.TP,P.Cuenta_No,Cliente As Nombre_Cliente,P.Credito_No,P.Capital " _
       & "FROM Prestamos As P,Clientes_Datos_Extras As C,Clientes As Cl " _
       & "WHERE P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND P.T = 'P' "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND P.TP = '" & DCTP.Text & "' "
  sSQL = sSQL & "AND P.Cuenta_No = C.Cuenta_No " _
       & "AND Cl.Codigo = C.Codigo " _
       & "ORDER BY P.TP,P.Cuenta_No "
  SelectDataGrid DGCreditos, AdoCreditos, sSQL
  
  sSQL = "SELECT P.Fecha_C,P.TP,P.Cuenta_No,Cliente As Nombre_Cliente,P.Credito_No,P.Cuota_No,P.Capital " _
       & "FROM Trans_Prestamos As P,Clientes_Datos_Extras As C,Clientes AS Cl " _
       & "WHERE P.Fecha_C BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND P.T = 'C' "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND P.TP = '" & DCTP.Text & "' "
  sSQL = sSQL & "AND P.Cuenta_No = C.Cuenta_No " _
       & "AND C.Codigo = Cl.Codigo " _
       & "ORDER BY P.TP,P.Cuenta_No,P.Credito_No,P.Cuota_No "
  SelectDataGrid DGDebitos, AdoDebitos, sSQL
  
  sSQL = "SELECT P.Fecha,P.TP,P.Cuenta_No,Cliente As Nombre_Cliente, P.Credito_No," _
       & "P.Cuota_No,P.Capital " _
       & "FROM Trans_Prestamos As P,Clientes_Datos_Extras As C,Clientes As Cl " _
       & "WHERE P.Fecha_V BETWEEN #" & MBFechaI.Text & "# and #" & MBFechaF.Text & "# " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND P.V <> " & Val(adFalse) & " " _
       & "AND P.Cuenta_No = C.Cuenta_No " _
       & "AND C.Codigo = Cl.Codigo "
  If CheckTP.value = 1 Then sSQL = sSQL & "AND P.TP = '" & DCTP.Text & "' "
  sSQL = sSQL & "ORDER BY P.TP,P.Fecha,P.Cuota_No,Cliente,P.Credito_No "
  SelectDataGrid DGVencidos, AdoVencidos, sSQL
  SaldoAnterior = 0: SaldoActual = 0
'''  sSQL = "SELECT * " _
'''       & "FROM Saldo_Prestamos " _
'''       & "WHERE Fecha < #" & MiFecha & "# " _
'''       & "ORDER BY Fecha "
'''  SelectAdodc AdoCaja, sSQL
'''  With AdoCaja.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveLast
'''       SaldoAnterior = .Fields("Saldo_Anterior")
'''       SaldoActual = .Fields("Saldo_Actual")
'''   End If
'''  End With
  sSQL = "SELECT TP,T,Fecha,Fecha_V,Fecha_C,SUM(Capital) As Total_Cap " _
       & "FROM Trans_Prestamos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "GROUP BY TP,T,Fecha,Fecha_V,Fecha_C "
  SelectDataGrid DGResumen, AdoResumen, sSQL
  RatonNormal
  MBFechaI.SetFocus
End Sub

Public Sub SumarIngEgrPre(DtaDebitos As Adodc, DtaCreditos As Adodc, DtaVencidos As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCreditos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Haber = Haber + .Fields("Capital")
         .MoveNext
       Loop
   End If
  End With
  With DtaDebitos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("Capital")
         .MoveNext
       Loop
   End If
  End With
  With DtaVencidos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe_ME = Debe_ME + .Fields("Capital")
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
End Sub

Private Sub Command1_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  ListarFlujoPrestamos
  SumarIngEgrPre AdoDebitos, AdoCreditos, AdoVencidos
  RatonNormal
End Sub

Private Sub Command2_Click()
  sSQL = "SELECT * " _
       & "FROM Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCaja, sSQL
  sSQL = "SELECT Credito_No,TP,Cuenta_No,SUM(Capital) As Capt " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY Credito_No,TP,Cuenta_No "
  SelectAdodc AdoResumen, sSQL
  With AdoResumen.Recordset
       Do While Not .EOF
          Contrato_No = .Fields("Credito_No")
          Cuenta_No = .Fields("Cuenta_No")
          TipoDoc = .Fields("TP")
          Total = Round(.Fields("Capt"), 2)
          AdoCaja.Recordset.MoveFirst
          AdoCaja.Recordset.Find ("Credito_No like '" & Contrato_No & "' ")
          If Not AdoCaja.Recordset.EOF Then
             AdoCaja.Recordset.Fields("Capital") = Total
             AdoCaja.Recordset.Update
          End If
         .MoveNext
       Loop
  End With
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE FLUJO DE PRESTAMOS"
  ImprimirFlujoDePrestamos AdoCreditos, AdoDebitos
End Sub

Private Sub Command5_Click()
  Unload FlujoDePrestamos
End Sub

Private Sub Command6_Click()
  If MBFechaI.Text = FechaSistema Then
'''  sSQL = "DELETE * " _
'''       & "FROM Saldo_Prestamos " _
'''       & "WHERE Fecha = #" & BuscarFecha(FechaSistema) & "# "
'''  ConectarAdoExecute sSQL
'''  SaldoAnterior = Round(CDbl(LabelSaldoIni.Caption), 2)
'''  SaldoActual = Round(CDbl(LabelSaldo.Caption), 2)
'''  sSQL = "SELECT * " _
'''       & "FROM Saldo_Prestamos "
'''  SelectAdodc AdoCaja, sSQL
'''  With AdoCaja.Recordset
'''      .AddNew
'''      .Fields("Fecha") = FechaSistema
'''      .Fields("Saldo_Anterior") = SaldoAnterior
'''      .Fields("Saldo_Actual") = SaldoActual
'''      .Fields("Usuario") = NombreUsuario
'''      .Fields("Item") = NumEmpresa
'''      .Update
'''  End With
  Else
     MsgBox "No puede grabar de dias anterioeres"
  End If
End Sub

Private Sub DCTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then
     If CNivel(6) Then
        Command6.Enabled = False
     End If
  End If
  sSQL = "SELECT TP " _
       & "FROM Trans_Prestamos " _
       & "GROUP BY TP "
  SelectDBCombo DCTP, AdoTP, sSQL, "TP"
  ListarFlujoPrestamos
End Sub

Private Sub Form_Load()
CentrarForm FlujoDePrestamos
ConectarAdodc AdoTP
ConectarAdodc AdoCaja
ConectarAdodc AdoResumen
ConectarAdodc AdoDebitos
ConectarAdodc AdoCreditos
ConectarAdodc AdoVencidos
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub
