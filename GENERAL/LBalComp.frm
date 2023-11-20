VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form LibroBanco 
   Caption         =   "LIBRO BANCO"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   12780
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "LBalComp.frx":0000
      DataSource      =   "AdoAgencias"
      Height          =   345
      Left            =   3570
      TabIndex        =   8
      Top             =   735
      Width           =   4845
      _ExtentX        =   8546
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
   Begin MSComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   4410
      TabIndex        =   29
      Top             =   6195
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "LBalComp.frx":001A
      DataSource      =   "AdoUsuario"
      Height          =   345
      Left            =   3570
      TabIndex        =   6
      Top             =   420
      Width           =   4845
      _ExtentX        =   8546
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
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por &Usuario"
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
      Left            =   2205
      TabIndex        =   5
      Top             =   420
      Width           =   1380
   End
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "Agencia:"
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
      Left            =   2205
      TabIndex        =   7
      Top             =   735
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGBanco 
      Bindings        =   "LBalComp.frx":0033
      Height          =   3900
      Left            =   105
      TabIndex        =   28
      Top             =   1050
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   6879
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "LBalComp.frx":004A
      DataSource      =   "AdoBanco1"
      Height          =   345
      Left            =   2205
      TabIndex        =   4
      Top             =   105
      Width           =   6210
      _ExtentX        =   10954
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   6195
      Width           =   4320
      _ExtentX        =   7620
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
   Begin VB.CommandButton Command1 
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
      Left            =   10815
      Picture         =   "LBalComp.frx":0062
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
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
      Height          =   855
      Left            =   9660
      Picture         =   "LBalComp.frx":092C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Consultar"
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
      Left            =   8505
      Picture         =   "LBalComp.frx":11F6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   630
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
      Left            =   840
      TabIndex        =   1
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   420
      Top             =   1890
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   420
      Top             =   2205
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
   Begin MSAdodcLib.Adodc AdoBanco1 
      Height          =   330
      Left            =   420
      Top             =   2520
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
      Caption         =   "Banco1"
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
      Left            =   420
      Top             =   2835
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
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   420
      Top             =   3150
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
      Caption         =   "Agencias"
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   420
      Top             =   3465
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
      Caption         =   "Usuario"
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
   Begin VB.Label LabelTotSaldoME 
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
      Left            =   9660
      TabIndex        =   19
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo ME"
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
      Left            =   8610
      TabIndex        =   20
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label LabelTotHaberME 
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
      Left            =   7035
      TabIndex        =   21
      Top             =   6930
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Haber ME"
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
      TabIndex        =   22
      Top             =   6930
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
      Left            =   9660
      TabIndex        =   17
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      Left            =   8610
      TabIndex        =   16
      Top             =   6615
      Width           =   1065
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
      Left            =   7035
      TabIndex        =   15
      Top             =   6615
      Width           =   1590
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Haber MN"
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
      Top             =   6615
      Width           =   1065
   End
   Begin VB.Label LabelTotDebeME 
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
      Left            =   4305
      TabIndex        =   23
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe ME"
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
      TabIndex        =   24
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label LabelSaldoAntME 
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
      Left            =   1575
      TabIndex        =   27
      Top             =   6930
      Width           =   1695
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
      Left            =   4305
      TabIndex        =   13
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe MN"
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
      TabIndex        =   12
      Top             =   6615
      Width           =   1065
   End
   Begin VB.Label LabelSaldoAntMN 
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
      Left            =   1575
      TabIndex        =   25
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Ant. ME"
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
      TabIndex        =   18
      Top             =   6930
      Width           =   1485
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Ant. MN"
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
      TabIndex        =   26
      Top             =   6615
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      Top             =   630
      Width           =   750
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      Width           =   750
   End
End
Attribute VB_Name = "LibroBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload LibroBanco
End Sub

Private Sub Command2_Click()
  DGBanco.Visible = False
  Imprimir_Libro_Banco AdoBanco
  DGBanco.Visible = True
End Sub

Private Sub Command3_Click()
  RatonReloj
  Saldo = 0: Saldo_ME = 0
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  sSQL = "SELECT Cta,T.Fecha,T.TP,T.Numero,Cheq_Dep,Cliente,C.Concepto,Debe,Haber,Saldo,Parcial_ME,Saldo_ME,T.T,T.Item " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  If CheckAgencia.value = 1 Then
     Co.Item = SinEspaciosIzq(DCAgencia.Text)
     sSQL = sSQL & "AND T.Item = '" & Co.Item & "' "
  Else
     If Not ConSucursal Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  End If
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND C.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
  sSQL = sSQL _
       & "AND T.Cta = '" & Codigo1 & "' " _
       & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Fecha = T.Fecha " _
       & "AND C.Item = T.Item " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "AND C.Periodo = T.Periodo " _
       & "ORDER BY Cta,T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
  Select_Adodc_Grid DGBanco, AdoBanco, sSQL
  DGBanco.Visible = False
  Debe = 0: Haber = 0: Saldo = 0
  Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
       'SetProgBar ProgBarra, AdoBanco.Recordset.RecordCount
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("Debe")
          Haber = Haber + .Fields("Haber")
          Saldo = .Fields("Saldo")
          If .Fields("Parcial_ME") >= 0 Then
              Debe_ME = Debe_ME + .Fields("Parcial_ME")
          Else
              Haber_ME = Haber_ME - .Fields("Parcial_ME")
          End If
          Saldo_ME = .Fields("Saldo_ME")
          'IncProgBar ProgBarra
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGBanco.Visible = True
  SaldoAnterior = CalculosSaldoAnt(DCCtas.Text, Debe, Haber, Saldo)
  LabelSaldoAntMN.Caption = Format(SaldoAnterior, "#,##0.00")
  LabelSaldoAntME.Caption = Format(Saldo_ME - Debe_ME + Haber_ME, "#,##0.00")
  LabelTotSaldo.Caption = Format(Saldo, "#,##0.00")
  LabelTotSaldoME.Caption = Format(Saldo_ME, "#,##0.00")
  LabelTotDebe.Caption = Format(Debe, "#,##0.00")
  LabelTotHaber.Caption = Format(Haber, "#,##0.00")
  LabelTotDebeME.Caption = Format(Debe_ME, "#,##0.00")
  LabelTotHaberME.Caption = Format(Haber_ME, "#,##0.00")
  ProgBarra.value = ProgBarra.Max
  AdoCtas.Caption = Cadena
  RatonNormal
  LibroBanco.Caption = "LIBRO BANCO"
  DGBanco.SetFocus
End Sub

Private Sub DCCtas_LostFocus()
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & Codigo1 & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc AdoCta, sSQL
  With AdoCta.Recordset
   If .RecordCount > 0 Then Moneda_US = .Fields("ME")
  End With
End Sub

Private Sub DGBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
         DGBanco.Visible = False
         GenerarDataTexto LibroBanco, AdoBanco
         DGBanco.Visible = True
    Case vbKeyF10
         If ClaveContador Then
            Co.Fecha = DGBanco.Columns(1).Text
            Co.TP = DGBanco.Columns(2).Text
            Co.Numero = DGBanco.Columns(3).Text
            Co.Beneficiario = DGBanco.Columns(5).Text
            FechaComp = Co.Fecha
            NumeroComp = Co.Numero
            Mensajes = "Seguro que quiere Modificar" & vbCrLf & "El Comprobante: " & Co.TP & " No. " & Co.Numero & " Con Fecha: " & Co.Fecha & vbCrLf
            If Len(Co.Beneficiario) > 1 Then Mensajes = Mensajes & "De " & ULCase(Co.Beneficiario)
            Titulo = "Pregunta de Modificacion"
            If BoxMensaje = vbYes Then
               ModificarComp = True
               CopiarComp = False
               NuevoComp = False
               Trans_No = 1
               Unload LibroBanco
               FComprobantes.Show
            End If
         End If
  End Select
End Sub

Private Sub Form_Activate()
  If ConSucursal Then
     sSQL = "SELECT (Item & '  ' & Empresa) As NomEmpresa " _
          & "FROM Empresas " _
          & "WHERE Item IN (" & ListSucursales & ") " _
          & "ORDER BY Item, Empresa "
     SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NomEmpresa"
     CheckAgencia.value = 0
     DCAgencia.Visible = True
     CheckAgencia.Visible = True
  Else
     DCAgencia.Visible = False
     CheckAgencia.Visible = False
  End If
  sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = '" & CtaBancos & "' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoBanco1, sSQL, "Nombre_Cta", False
  sSQL = "UPDATE Comprobantes " _
       & "SET Cotizacion = 0.004 " _
       & "WHERE Cotizacion = 0 " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT (Item & '  ' & Empresa) As NomEmpresa " _
       & "FROM Empresas " _
       & "WHERE Item <> '000' " _
       & "ORDER BY Item,Empresa "
  SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NomEmpresa"
  sSQL = "SELECT (Nombre_Completo & '  ' & Codigo) As CodUsuario " _
       & "FROM Accesos " _
       & "WHERE Codigo <> '*' " _
       & "ORDER BY Nombre_Completo "
  SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "CodUsuario", False
  LibroBanco.Caption = "LIBRO BANCO"
  Co.Item = NumEmpresa
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCta
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoBanco1
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoAgencias
  
  DGBanco.Height = MDI_Y_Max - DGBanco.Top - 900
  DGBanco.width = MDI_X_Max - DGBanco.Left
  AdoBanco.Top = DGBanco.Top + DGBanco.Height + 10
  ProgBarra.Top = DGBanco.Top + DGBanco.Height + 10
  
  Label13.Top = AdoBanco.Top + AdoBanco.Height + 10
  Label6.Top = AdoBanco.Top + AdoBanco.Height + 10
  Label9.Top = AdoBanco.Top + AdoBanco.Height + 10
  Label11.Top = AdoBanco.Top + AdoBanco.Height + 10
  LabelSaldoAntMN.Top = AdoBanco.Top + AdoBanco.Height + 10
  LabelTotSaldo.Top = AdoBanco.Top + AdoBanco.Height + 10
  LabelTotDebe.Top = AdoBanco.Top + AdoBanco.Height + 10
  LabelTotHaber.Top = AdoBanco.Top + AdoBanco.Height + 10
  
  Label15.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  Label10.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  Label5.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  Label3.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  LabelSaldoAntME.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  LabelTotSaldoME.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  LabelTotDebeME.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  LabelTotHaberME.Top = LabelSaldoAntMN.Top + LabelSaldoAntMN.Height + 10
  
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  MBoxFechaF = UltimoDiaMes(MBoxFechaI)
End Sub
