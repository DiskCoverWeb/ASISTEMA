VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FMigraAT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Talon Resúmen Anexo Transaccional"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "FMigraAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4897.262
   ScaleMode       =   0  'User
   ScaleWidth      =   8507.028
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictVentas 
      Height          =   645
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   8775
      TabIndex        =   6
      Top             =   1785
      Width           =   8835
   End
   Begin VB.PictureBox PictAir 
      Height          =   645
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   8775
      TabIndex        =   5
      Top             =   2520
      Width           =   8835
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   960
      Left            =   105
      TabIndex        =   3
      Top             =   0
      Width           =   6525
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MIGRACION DE DATOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   645
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   6315
      End
   End
   Begin VB.PictureBox PictCompras 
      Height          =   645
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   8775
      TabIndex        =   2
      Top             =   1050
      Width           =   8835
   End
   Begin VB.CommandButton CmdGeneArch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Generar Archivo"
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
      Left            =   6720
      Picture         =   "FMigraAT.frx":0BC2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00C0C0C0&
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
      Left            =   7875
      Picture         =   "FMigraAT.frx":1258
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoVentas 
      Height          =   330
      Left            =   4410
      Top             =   1050
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
   Begin MSAdodcLib.Adodc AdoImpor 
      Height          =   330
      Left            =   210
      Top             =   1785
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Impor"
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
   Begin MSAdodcLib.Adodc AdoCompras 
      Height          =   330
      Left            =   210
      Top             =   1365
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
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
      Caption         =   "Compras"
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
   Begin MSAdodcLib.Adodc AdoExpor 
      Height          =   330
      Left            =   210
      Top             =   2100
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Expor"
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
   Begin MSAdodcLib.Adodc AdoCambio 
      Height          =   330
      Left            =   210
      Top             =   1050
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
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
      Caption         =   "Cambio"
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
   Begin MSAdodcLib.Adodc AdoRetFte 
      Height          =   330
      Left            =   2310
      Top             =   1785
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RetFte"
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
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   330
      Left            =   6300
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Empresa"
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
   Begin MSAdodcLib.Adodc AdoTC 
      Height          =   330
      Left            =   2310
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "TC"
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
   Begin MSAdodcLib.Adodc AdoRI2B 
      Height          =   330
      Left            =   2295
      Top             =   1380
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RI2B"
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
   Begin MSAdodcLib.Adodc AdoRI2S 
      Height          =   330
      Left            =   4410
      Top             =   1365
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RI2S"
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
   Begin MSAdodcLib.Adodc AdoRI 
      Height          =   330
      Left            =   210
      Top             =   2835
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RI"
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
   Begin MSAdodcLib.Adodc AdoRE 
      Height          =   330
      Left            =   2310
      Top             =   2100
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RE"
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
   Begin MSAdodcLib.Adodc AdoRT 
      Height          =   330
      Left            =   2310
      Top             =   2835
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "RT"
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
   Begin MSAdodcLib.Adodc AdoAnulados 
      Height          =   330
      Left            =   210
      Top             =   2520
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Anulados"
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
   Begin MSAdodcLib.Adodc AdoRRFF 
      Height          =   330
      Left            =   2310
      Top             =   1050
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RendFinF"
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
   Begin MSAdodcLib.Adodc AdoRRF 
      Height          =   330
      Left            =   4200
      Top             =   2520
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "RendFin"
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
   Begin MSAdodcLib.Adodc AdoFechas 
      Height          =   330
      Left            =   4200
      Top             =   2835
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Fechas"
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
End
Attribute VB_Name = "FMigraAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ProcCompras As Progreso_Barras
Dim ProcVentas As Progreso_Barras
Dim ProcAir As Progreso_Barras
Dim SumaAnulados As Integer

Public Sub GenerarArchivoATXML(Tipo As Boolean)
Dim CaptionOld As String
Dim ValorBool As String
Dim CodCampS As String
Dim FechaReg As String
Dim FechaEmi As String
Dim NumTrans As Long
Dim RetFte As Adodc
RatonReloj
FAConLineas = False
Cadena = " "
Contador = 0
NumTrans = 0
'Encabezados de las Compras
With AdoCompras.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     ProcCompras.Incremento = 0
     ProcCompras.Valor_Maximo = .RecordCount
     ProcCompras.Mensaje_Box = "MIGRANDO TRANSACCIONES DE COMPRAS"
     'Pict_Proceso PictCompras, ProcCompras
     Do While Not .EOF
        NumTrans = NumTrans + 1
        Periodo_Contable = .Fields("Periodo")
        Factura_No = CTNumero(.Fields("Secuencial"))
        Codigo1 = Mid$(.Fields("Serie"), 1, 3)
        Codigo2 = Mid$(.Fields("Serie"), 4, 3)
        If Val(Codigo1) <= 0 Then Codigo1 = "1"
        If Val(Codigo2) <= 0 Then Codigo2 = "1"
        Codigo1 = Format(Val(Codigo1), "000")
        Codigo2 = Format(Val(Codigo2), "000")
        FechaReg = .Fields("Fecha")
        FechaEmi = .Fields("FechaE")
        If CFechaLong(FechaEmi) > CFechaLong(FechaReg) Then FechaEmi = FechaReg
        SetAdoAddNew "Trans_Compras"
        SetAdoFields "IdProv", .Fields("Codigo")
        SetAdoFields "DevIva", .Fields("Dev")
        SetAdoFields "CodSustento", .Fields("IdenCT")
        SetAdoFields "TipoComprobante", Val(.Fields("TD"))
        SetAdoFields "Establecimiento", Codigo1
        SetAdoFields "PuntoEmision", Codigo2
        SetAdoFields "Secuencial", CTNumero(.Fields("Secuencial"))
        SetAdoFields "Autorizacion", .Fields("Autorizacion")
        SetAdoFields "FechaEmision", FechaEmi
        SetAdoFields "FechaRegistro", FechaReg
        SetAdoFields "FechaCaducidad", .Fields("FechaC")
        SetAdoFields "BaseImponible", .Fields("BImpotCero")
        SetAdoFields "BaseImpGrav", .Fields("Valor_Fact")
        SetAdoFields "PorcentajeIva", 2 ' .Fields("CodPorc")
        SetAdoFields "MontoIva", .Fields("MontoIVA1") + .Fields("MontoIVA2")
        SetAdoFields "BaseImpIce", .Fields("MontoICE")
        SetAdoFields "PorcentajeIce", Mid$(.Fields("CPorcICE"), 1, 1)
        SetAdoFields "MontoIce", .Fields("RetICE")
        SetAdoFields "MontoIvaBienes", .Fields("MontoIVA1")
        SetAdoFields "PorRetBienes", .Fields("PorRetIVA1")
        SetAdoFields "ValorRetBienes", .Fields("MontoRetIVA1")
        SetAdoFields "MontoIvaServicios", .Fields("MontoIVA2")
        SetAdoFields "PorRetServicios", .Fields("PorRetIVA2")
        SetAdoFields "ValorRetServicios", .Fields("MontoRetIVA2")
        SetAdoFields "DocModificado", "0"
        SetAdoFields "FechaEmiModificado", "01/01/1900"
        SetAdoFields "EstabModificado", "000"
        SetAdoFields "PtoEmiModificado", "000"
        SetAdoFields "SecModificado", "0000000"
        SetAdoFields "AutModificado", "0000000000"
      ' notas de debito y credito compras
        Select Case .Fields("TD")
          Case "04", "05"
               Codigo = .Fields("Cambio")
               If AdoCambio.Recordset.RecordCount > 0 Then
                  AdoCambio.Recordset.MoveFirst
                  AdoCambio.Recordset.Find ("CompMod = '" & Codigo & "' ")
                  If Not AdoCambio.Recordset.EOF Then
                     Cadena1 = Val(Mid$(Codigo, 1, 2))
                     Cadena2 = Val(Mid$(Codigo, 3, 7))
                     Codigo1 = Mid$(AdoCambio.Recordset.Fields("Serie"), 1, 3)
                     Codigo2 = Mid$(AdoCambio.Recordset.Fields("Serie"), 4, 3)
                     If Val(Codigo1) <= 0 Then Codigo1 = "1"
                     If Val(Codigo2) <= 0 Then Codigo2 = "1"
                     Codigo1 = Format(Val(Codigo1), "000")
                     Codigo2 = Format(Val(Codigo2), "000")
                     SetAdoFields "DocModificado", Cadena1
                     SetAdoFields "FechaEmiModificado", AdoCambio.Recordset.Fields("FechaE")
                     SetAdoFields "EstabModificado", Codigo1
                     SetAdoFields "PtoEmiModificado", Codigo2
                     SetAdoFields "SecModificado", Cadena2
                     SetAdoFields "AutModificado", AdoCambio.Recordset.Fields("Autorizacion")
                     AdoCambio.Recordset.MoveNext
                  End If
               End If
        End Select
      ' gasto electoral compras
        SetAdoFields "ContratoPartidoPolitico", 0
        SetAdoFields "MontoTituloOneroso", 0
        SetAdoFields "MontoTituloGratuito", 0
        Codigo = .Fields("Comprobante")
        If AdoRI2S.Recordset.RecordCount > 0 Then
           AdoRI2S.Recordset.MoveFirst
           AdoRI2S.Recordset.Find ("Comprobante = '" & Codigo & "' ")
           If Not AdoRI2S.Recordset.EOF Then
              SetAdoFields "ContratoPartidoPolitico", AdoRI2S.Recordset.Fields("Contrato")
              SetAdoFields "MontoTituloOneroso", AdoRI2S.Recordset.Fields("Titulo_Oneroso")
              SetAdoFields "MontoTituloGratuito", AdoRI2S.Recordset.Fields("Titulo_Gratuito")
              AdoRI2S.Recordset.MoveNext
           End If
        End If
        SetAdoFields "T", Normal
        SetAdoFields "TP", .Fields("TP")
        SetAdoFields "Numero", .Fields("Numero")
        SetAdoFields "Fecha", .Fields("Fecha")
        SetAdoFields "ID", NumTrans
        SetAdoFields "Porc_Bienes", "0"
        SetAdoFields "Porc_Servicios", "0"
        If .Fields("MontoIVA1") > 0 Then SetAdoFields "Porc_Bienes", CStr(Round(100 * .Fields("MontoRetIVA1") / .Fields("MontoIVA1")))
        If .Fields("MontoIVA2") > 0 Then SetAdoFields "Porc_Servicios", CStr(Round(100 * .Fields("MontoRetIVA2") / .Fields("MontoIVA2")))
        SetAdoFields "Cta_Servicio", .Fields("Cta")
        SetAdoFields "Cta_Bienes", .Fields("Cta1")
        SetAdoUpdate
        'Pict_Proceso PictCompras, ProcCompras
       .MoveNext
     Loop
     ProcCompras.Incremento = ProcCompras.Valor_Maximo
     'Pict_Proceso PictCompras, ProcCompras
 End If
End With
' -------------------------------------------------------------------------
' Ventas
NumTrans = 0
With AdoVentas.Recordset
 If .RecordCount > 0 Then
     ProcVentas.Incremento = 0
     ProcVentas.Valor_Maximo = .RecordCount
     ProcVentas.Mensaje_Box = "MIGRANDO TRANSACCIONES DE VENTAS"
     'Pict_Proceso PictVentas, ProcVentas
     Do While Not .EOF
        NumTrans = NumTrans + 1
        Periodo_Contable = .Fields("Periodo")
        Codigo1 = Mid$(.Fields("Serie"), 1, 3)
        Codigo2 = Mid$(.Fields("Serie"), 4, 3)
        If Val(Codigo1) <= 0 Then Codigo1 = "1"
        If Val(Codigo2) <= 0 Then Codigo2 = "1"
        Codigo1 = Format(Val(Codigo1), "000")
        Codigo2 = Format(Val(Codigo2), "000")
        SetAdoAddNew "Trans_Ventas"
        SetAdoFields "IdProv", .Fields("Codigo")
        SetAdoFields "TipoComprobante", Val(.Fields("TD"))
        SetAdoFields "FechaRegistro", UltimoDiaMes(.Fields("Fecha"))
        SetAdoFields "FechaEmision", UltimoDiaMes(.Fields("FechaE"))
        SetAdoFields "Establecimiento", Codigo1
        SetAdoFields "PuntoEmision", Codigo2
        SetAdoFields "Secuencial", CTNumero(.Fields("Secuencial"))
        SetAdoFields "NumeroComprobantes", .Fields("Numero")
        SetAdoFields "BaseImponible", .Fields("BImpotCero")
        SetAdoFields "IvaPresuntivo", .Fields("ConvInt")
        SetAdoFields "BaseImpGrav", .Fields("Valor_Fact")
        SetAdoFields "PorcentajeIva", 2
        SetAdoFields "MontoIva", .Fields("MontoIVA1") + .Fields("MontoIVA2")
       'SetAdoFields "montoIvaPresuntivo" & .Fields("MontoIVA1") + .Fields("MontoIVA2")
        SetAdoFields "BaseImpIce", .Fields("MontoICE")
        SetAdoFields "PorcentajeIce", Mid$(.Fields("CPorcICE"), 1, 1)
        SetAdoFields "MontoIce", .Fields("RetICE")
        SetAdoFields "MontoIvaBienes", .Fields("MontoIVA1")
        SetAdoFields "PorRetBienes", .Fields("PorRetIVA1")
        SetAdoFields "ValorRetBienes", .Fields("MontoRetIVA1")
        SetAdoFields "MontoIvaServicios", .Fields("MontoIVA2")
        SetAdoFields "PorRetServicios", .Fields("PorRetIVA2")
        SetAdoFields "ValorRetServicios", .Fields("MontoRetIVA2")
        SetAdoFields "RetPresuntiva", .Fields("ConvInt")
        SetAdoFields "Porc_Bienes", "0"
        SetAdoFields "Porc_Servicios", "0"
        If .Fields("MontoIVA1") > 0 Then SetAdoFields "Porc_Bienes", CStr(Round(100 * .Fields("MontoRetIVA1") / .Fields("MontoIVA1")))
        If .Fields("MontoIVA2") > 0 Then SetAdoFields "Porc_Servicios", CStr(Round(100 * .Fields("MontoRetIVA2") / .Fields("MontoIVA2")))
        SetAdoFields "Cta_Servicio", .Fields("Cta")
        SetAdoFields "Cta_Bienes", .Fields("Cta1")
        SetAdoFields "T", Normal
        SetAdoFields "TP", .Fields("TP")
        SetAdoFields "Numero", .Fields("Numero")
        SetAdoFields "Fecha", .Fields("Fecha")
        SetAdoFields "ID", NumTrans
        SetAdoUpdate
        'Pict_Proceso PictVentas, ProcVentas
       .MoveNext
     Loop
     ProcVentas.Incremento = ProcVentas.Valor_Maximo
     'Pict_Proceso PictVentas, ProcVentas
 End If
End With
' -------------------------------------------------------------------------
' Retencion en la Fuente de Compras/Ventas/Importaciones y Exportaciones
  NumTrans = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Retenciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND CodigoTR IN ('RF','RV','RI','RE') " _
       & "AND T <> 'A' " _
       & "ORDER BY Fecha,TT,CodigoTR "
  SelectAdodc AdoRetFte, sSQL
  With AdoRetFte.Recordset
   If .RecordCount > 0 Then
       ProcAir.Incremento = 0
       ProcAir.Valor_Maximo = .RecordCount
       ProcAir.Mensaje_Box = "MIGRANDO TRANSACCIONES AIR"
       'Pict_Proceso PictAir, ProcAir
       Do While Not .EOF
          NumTrans = NumTrans + 1
          Factura_No = CTNumero(.Fields("Secuencial"))
          Codigo3 = Mid$(.Fields("Serie"), 1, 3)
          Codigo4 = Mid$(.Fields("Serie"), 4, 3)
          If Val(Codigo3) <= 0 Then Codigo3 = "1"
          If Val(Codigo4) <= 0 Then Codigo4 = "1"
          Codigo3 = Format(Val(Codigo3), "000")
          Codigo4 = Format(Val(Codigo4), "000")
          Codigo = .Fields("Codigo")
          Mifecha = .Fields("Fecha")
          TipoDoc = .Fields("TP")
          Numero = .Fields("Numero")
          Codigo1 = "001"
          Codigo2 = "001"
          If AdoCompras.Recordset.RecordCount > 0 Then
             AdoCompras.Recordset.MoveFirst
             AdoCompras.Recordset.Find ("Codigo = '" & Codigo & "' ")
             If Not AdoCompras.Recordset.EOF Then
                AdoCompras.Recordset.Find ("Fecha = #" & BuscarFecha(Mifecha) & "# ")
                If Not AdoCompras.Recordset.EOF Then
                   AdoCompras.Recordset.Find ("TP = '" & TipoDoc & "' ")
                   If Not AdoCompras.Recordset.EOF Then
                      AdoCompras.Recordset.Find ("Numero = " & Numero & " ")
                      If Not AdoCompras.Recordset.EOF Then
                         Codigo1 = Mid$(AdoCompras.Recordset.Fields("Serie"), 1, 3)
                         Codigo2 = Mid$(AdoCompras.Recordset.Fields("Serie"), 4, 3)
                         If Val(Codigo1) <= 0 Then Codigo1 = "1"
                         If Val(Codigo2) <= 0 Then Codigo2 = "1"
                         Codigo1 = Format(Val(Codigo1), "000")
                         Codigo2 = Format(Val(Codigo2), "000")
                         Factura_No = Val(AdoCompras.Recordset.Fields("Secuencial"))
                      End If
                   End If
                End If
             End If
          End If
          SetAdoAddNew "Trans_Air"
          SetAdoFields "CodRet", .Fields("TD")
          SetAdoFields "BaseImp", .Fields("Valor_Fact")
          SetAdoFields "Porcentaje", .Fields("Porc")
          SetAdoFields "ValRet", .Fields("Valor_Ret")
          SetAdoFields "EstabRetencion", Codigo3
          SetAdoFields "PtoEmiRetencion", Codigo4
          SetAdoFields "SecRetencion", CTNumero(.Fields("Secuencial"))
          SetAdoFields "AutRetencion", .Fields("Autorizacion")
          SetAdoFields "IdProv", .Fields("Codigo")
          SetAdoFields "Cta_Retencion", .Fields("Cta")
          SetAdoFields "EstabFactura", Codigo1
          SetAdoFields "PuntoEmiFactura", Codigo2
          SetAdoFields "Factura_No", Factura_No
          SetAdoFields "T", Normal
          SetAdoFields "TP", .Fields("TP")
          SetAdoFields "Numero", .Fields("Numero")
          SetAdoFields "Fecha", .Fields("Fecha")
          SetAdoFields "ID", NumTrans
          Codigo3 = "C"
          If .Fields("TT") = "C" And .Fields("CodigoTR") = "RF" Then Codigo3 = "C"
          If .Fields("TT") = "V" And .Fields("CodigoTR") = "RV" Then Codigo3 = "V"
          If .Fields("TT") = "I" And .Fields("CodigoTR") = "RI" Then Codigo3 = "I"
          If .Fields("TT") = "E" And .Fields("CodigoTR") = "RE" Then Codigo3 = "E"
          SetAdoFields "Tipo_Trans", Codigo3
          SetAdoUpdate
         'SetAdoFields "<fechaEmiRet" & No_RegC & ">", .Fields("FechaE")
          'Pict_Proceso PictAir, ProcAir
         .MoveNext
       Loop
       ProcAir.Incremento = ProcAir.Valor_Maximo
       'Pict_Proceso PictAir, ProcAir
    End If
  End With

''''' datos importaciones
''''With AdoImpor.Recordset
''''If .RecordCount > 0 Then
''''   .MoveFirst
''''   SetAdoFields "<importaciones>"
''''   Do While Not .EOF
''''      SetAdoFields "<detalleImportaciones>"
''''        SetAdoFields "<codSustento>" & .Fields("SCT") & "</codSustento>"
''''        If .Fields("ConvInt") = "N" Then
''''           SetAdoFields "<importacionDe>2</importacionDe>"
''''        Else
''''           SetAdoFields "<importacionDe>1</importacionDe>"
''''        End If
''''        SetAdoFields "<fechaLiquidacion>" & .Fields("Fecha") & "</fechaLiquidacion>"
''''        SetAdoFields "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
''''        SetAdoFields "<distAduanero>" & Mid$(.Fields("Aduana"), 1, 3) & "</distAduanero>"
''''        SetAdoFields "<anio>" & Mid$(.Fields("Aduana"), 4, 4) & "</anio>"
''''        SetAdoFields "<regimen>" & Mid$(.Fields("Aduana"), 8, 2) & "</regimen>"
''''        SetAdoFields "<correlativo>" & Mid$(.Fields("Aduana"), 10, 6) & "</correlativo>"
''''        SetAdoFields "<verificador>" & Mid$(.Fields("Aduana"), 16, 1) & "</verificador>"
''''        SetAdoFields "<idFiscalProv>" & .Fields("CI_RUC") & "</idFiscalProv>"
''''        SetAdoFields "<valorCIF>" & .Fields("ValorCif") & "</valorCIF>"
''''        SetAdoFields "<razonSocialProv>" & .Fields("Cliente") & "</razonSocialProv>"
''''        Select Case .Fields("CTD")
''''        Case "R"
''''        SetAdoFields "<tipoSujeto>2</tipoSujeto>"
''''        Case Else
''''        SetAdoFields "<tipoSujeto>1</tipoSujeto>"
''''        End Select
''''        SetAdoFields "<baseImponible>" & .Fields("BImpotCero") & "</baseImponible>"
''''        SetAdoFields "<baseImpGrav>" & .Fields("Valor_Fact") & "</baseImpGrav>"
''''        SetAdoFields "<porcentajeIva>" & .Fields("CodPorc") & "</porcentajeIva>"
''''        SetAdoFields "<montoIva>" & .Fields("MontoIVA1") & "</montoIva>"
''''        SetAdoFields "<baseImpIce>" & .Fields("MontoICE") & "</baseImpIce>"
''''        SetAdoFields "<porcentajeIce>" & Mid$(.Fields("CPorcICE"), 1, 1) & "</porcentajeIce>"
''''        SetAdoFields "<montoIce>" & Format(.Fields("RetICE"), "#0.00") & "</montoIce>"
''''        ' retencion en la fuente del iva en importaciones
''''        Codigo = .Fields("Comprobante")
''''                 If AdoRI.Recordset.RecordCount > 0 Then
''''                    'AdoRI.Recordset.MoveFirst
''''                     AdoRI.Recordset.Find ("Comprobante = '" & Codigo & "' ")
''''                    If Not AdoRI.Recordset.EOF Then
''''                       SetAdoFields "<air>"
''''                       SetAdoFields "<detalleAir>"
''''                       SetAdoFields "<codRetAir>" & AdoRI.Recordset.Fields("TRet") & "</codRetAir>"
''''                       SetAdoFields "<baseImpAir>" & AdoRI.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
''''                       SetAdoFields "<porcentajeAir>" & AdoRI.Recordset.Fields("CodPorc") & "</porcentajeAir>"
''''                       SetAdoFields "<valRetAir>" & AdoRI.Recordset.Fields("Valor_Ret") & "</valRetAir>"
''''                       SetAdoFields "</detalleAir>"
''''                       SetAdoFields "</air>"
''''                       ' AdoRI.Recordset.MoveLast
'''''                       AdoRI.Recordset.MoveNext
''''                    Else
''''                       SetAdoFields "<air/>"
''''                    End If
''''                 Else
''''                   SetAdoFields "<air/>"
''''                 End If
''''        SetAdoFields "</detalleImportaciones>"
''''    .MoveNext
''''   Loop
''''        SetAdoFields "</importaciones>"
''''Else
''''SetAdoFields "<importaciones/>"
''''End If
''''End With
''''' datos exportaciones
''''With AdoExpor.Recordset
''''If .RecordCount > 0 Then
''''   .MoveFirst
''''   SetAdoFields "<exportaciones>"
''''   Do While Not .EOF
''''      SetAdoFields "<detalleExportaciones>"
''''        If .Fields("ConvInt") = "N" Then
''''           SetAdoFields "<exportacionDe>2</exportacionDe>"
''''        Else
''''           SetAdoFields "<exportacionDe>1</exportacionDe>"
''''        End If
''''        SetAdoFields "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
''''        SetAdoFields "<distAduanero>" & Mid$(.Fields("Aduana"), 1, 3) & "</distAduanero>"
''''        SetAdoFields "<anio>" & Mid$(.Fields("Aduana"), 4, 4) & "</anio>"
''''        SetAdoFields "<regimen>" & Mid$(.Fields("Aduana"), 8, 2) & "</regimen>"
''''        SetAdoFields "<correlativo>" & Mid$(.Fields("Aduana"), 10, 6) & "</correlativo>"
''''        SetAdoFields "<verificador>" & Mid$(.Fields("Aduana"), 16, 1) & "</verificador>"
''''        SetAdoFields "<fechaEmbarque>" & .Fields("FechaC") & "</fechaEmbarque>"
''''        SetAdoFields "<idFiscalCliente>" & .Fields("CI_RUC") & "</idFiscalCliente>"
''''        Select Case .Fields("CTD")
''''        Case "R"
''''        SetAdoFields "<tipoSujeto>2</tipoSujeto>"
''''        Case Else
''''        SetAdoFields "<tipoSujeto>1</tipoSujeto>"
''''        End Select
''''        SetAdoFields "<valorFOB>" & .Fields("ValorCif") & "</valorFOB>"
''''        SetAdoFields "<razonSocial>" & .Fields("Cliente") & "</razonSocial>"
''''        SetAdoFields "<devIva>" & .Fields("Dev") & "</devIva>"
''''        SetAdoFields "<facturaExportacion>1</facturaExportacion>"
''''        SetAdoFields "<valorFOBComprobante>" & .Fields("ValorCif") & "</valorFOBComprobante>"
''''        SetAdoFields "<establecimiento>" & Mid$(.Fields("Serie"), 1, 3) & "</establecimiento>"
''''        SetAdoFields "<puntoEmision>" & Mid$(.Fields("Serie"), 4, 3) & "</puntoEmision>"
''''        SetAdoFields "<secuencial>" & .Fields("Secuencial") & "</secuencial>"
''''        SetAdoFields "<fechaRegistro>" & .Fields("Fecha") & "</fechaRegistro>"
''''        SetAdoFields "<autorizacion>" & .Fields("autorizacion") & "</autorizacion>"
''''        SetAdoFields "<fechaEmision>" & .Fields("FechaE") & "</fechaEmision>"
''''        ' retencion en la fuente del iva en importaciones
''''        Codigo = .Fields("Comprobante")
''''                 If AdoRI.Recordset.RecordCount > 0 Then
''''                    'AdoRI.Recordset.MoveFirst
''''                     AdoRI.Recordset.Find ("Comprobante = '" & Codigo & "' ")
''''                    If Not AdoRI.Recordset.EOF Then
''''                       SetAdoFields "<air>"
''''                       SetAdoFields "<detalleAir>"
''''                       SetAdoFields "<codRetAir>" & AdoRE.Recordset.Fields("TRet") & "</codRetAir>"
''''                       SetAdoFields "<baseImpAir>" & AdoRE.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
''''                       SetAdoFields "<porcentajeAir>" & AdoRE.Recordset.Fields("Porc") & "</porcentajeAir>"
''''                       SetAdoFields "<valRetAir>" & AdoRE.Recordset.Fields("Valor_Ret") & "</valRetAir>"
''''                       SetAdoFields "</detalleAir>"
''''                       SetAdoFields "</air>"
'''''                       AdoRE.Recordset.MoveLast
'''''                       AdoRE.Recordset.MoveNext
''''                    Else
''''                       SetAdoFields "<air/>"
''''                    End If
''''                 Else
''''                    SetAdoFields "<air/>"
''''                 End If
''''        SetAdoFields "</detalleExportaciones>"
''''    .MoveNext
''''   Loop
''''        SetAdoFields "</exportaciones>"
''''Else
''''   SetAdoFields "<exportaciones/>"
''''End If
''''End With
'''''SetAdoFields "<recap/>"
''''' detalle recap
''''With AdoTC.Recordset  ' detalle recap
''''If .RecordCount > 0 Then
''''   .MoveFirst
''''   SetAdoFields "<recap>"
''''   Do While Not .EOF
''''      SetAdoFields "<detalleRecap>"
''''        Select Case .Fields("CTD")
''''        Case "R"
''''        SetAdoFields "<establecimientoRecap>10</establecimientoRecap>"
''''        Case Else
''''        SetAdoFields "<establecimientoRecap>11</establecimientoRecap>"
''''        End Select
''''        SetAdoFields "<identificacionRecap>" & .Fields("CI_RUC") & "</identificacionRecap>"
''''        SetAdoFields "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
''''        SetAdoFields "<numeroRecap>" & .Fields("Aduana") & "</numeroRecap>"
''''        SetAdoFields "<fechaPago>" & .Fields("Fecha") & "</fechaPago>"
''''        SetAdoFields "<tarjetaCredito>" & .Fields("IdenCT") & "</tarjetaCredito>"
''''        SetAdoFields "<fechaEmisionRecap>" & .Fields("FechaE") & "</fechaEmisionRecap>"
''''        SetAdoFields "<consumoCero>" & .Fields("BImpotCero") & "</consumoCero>"
''''        SetAdoFields "<consumoGravado>" & .Fields("Valor_Fact") & "</consumoGravado>"
''''        SetAdoFields "<totalConsumo>" & .Fields("Valor_Fact") + .Fields("BImpotCero") & "</totalConsumo>"
''''        SetAdoFields "<montoIva>" & .Fields("MontoIVA1") + .Fields("MontoIVA2") & "</montoIva>"
''''        SetAdoFields "<comision>" & .Fields("Comision") & "</comision>"
''''        SetAdoFields "<numeroVouchers>" & .Fields("SerieEx") & "</numeroVouchers>"
''''        SetAdoFields "<montoIvaBienes>" & .Fields("MontoIVA1") & "</montoIvaBienes>"
''''        SetAdoFields "<porRetBienes>" & .Fields("PorRetIVA1") & "</porRetBienes>"
''''        SetAdoFields "<valorRetBienes>" & .Fields("MontoRetIVA1") & "</valorRetBienes>"
''''        SetAdoFields "<montoIvaServicios>" & .Fields("MontoIVA2") & "</montoIvaServicios>"
''''        SetAdoFields "<porRetServicios>" & .Fields("PorRetIVA2") & "</porRetServicios>"
''''        SetAdoFields "<valorRetServicios>" & .Fields("MontoRetIVA2") & "</valorRetServicios>"
''''    ' retencion en la fuente del iva en rendimientos financieros
''''        Codigo = .Fields("Comprobante")
''''                 If AdoRT.Recordset.RecordCount > 0 Then
''''                    'AdoRI.Recordset.MoveFirst
''''                     AdoRT.Recordset.Find ("Comprobante = '" & Codigo & "' ")
''''                    If Not AdoRI.Recordset.EOF Then
''''                       SetAdoFields "<air>"
''''                       SetAdoFields "<detalleAir>"
''''                       SetAdoFields "<codRetAir>" & AdoRT.Recordset.Fields("TD") & "</codRetAir>"
''''                       SetAdoFields "<baseImpAir>" & AdoRT.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
''''                       SetAdoFields "<porcentajeAir>" & AdoRT.Recordset.Fields("CodPorc") & "</porcentajeAir>"
''''                       SetAdoFields "<valRetAir>" & AdoRT.Recordset.Fields("Valor_Ret") & "</valRetAir>"
''''                       SetAdoFields "</detalleAir>"
''''                       SetAdoFields "</air>"
''''                       SetAdoFields "<establecimiento>" & Mid$(AdoRT.Recordset.Fields("Serie"), 1, 3) & "</establecimiento>"
''''                       SetAdoFields "<puntoEmision>" & Mid$(AdoRT.Recordset.Fields("Serie"), 4, 3) & "</puntoEmision>"
''''                       SetAdoFields "<secuencial>" & AdoRT.Recordset.Fields("Secuencial") & "</secuencial>"
''''                       SetAdoFields "<fechaRegistro>" & AdoRT.Recordset.Fields("Fecha") & "</fechaRegistro>"
''''                       SetAdoFields "<autorizacion>" & AdoRT.Recordset.Fields("Autorizacion") & "</autorizacion>"
''''                       SetAdoFields "<fechaEmision>" & AdoRT.Recordset.Fields("FechaE") & "</fechaEmision>"
'''''                       AdoRI.Recordset.MoveLast
''''                       AdoRT.Recordset.MoveNext
''''                    Else
''''                      SetAdoFields "<air/>"
''''                    End If
''''                 End If
''''        SetAdoFields "</detalleRecap>"
''''    .MoveNext
''''   Loop
''''        SetAdoFields "</recap>"
''''Else
''''SetAdoFields "<recap/>"
''''End If
''''End With
''''
''''SetAdoFields "<fideicomisos/>"
''''' detalle dondos y fideicomisos
''''With AdoAnulados.Recordset    ' comprobantes anulados
''''If .RecordCount > 0 Then
''''    .MoveFirst
''''    SetAdoFields "<anulados>"
''''     Do While Not .EOF
''''      SetAdoFields "<detalleAnulados>"
''''      Cadena = Val(.Fields("TD"))
''''      If Cadena > 99 Then Cadena = 7
''''      SetAdoFields "<tipoComprobante>" & Cadena & "</tipoComprobante>"
''''      SetAdoFields "<establecimiento>" & Mid$(.Fields("Serie"), 1, 3) & "</establecimiento>"
''''      SetAdoFields "<puntoEmision>" & Mid$(.Fields("Serie"), 4, 3) & "</puntoEmision>"
''''      SetAdoFields "<secuencialInicio>" & .Fields("Secuencial") & "</secuencialInicio>"
''''      SetAdoFields "<secuencialFin>" & .Fields("Secuencial") & "</secuencialFin>"
''''      SetAdoFields "<autorizacion>" & .Fields("Autorizacion") & "</autorizacion>"
''''      SetAdoFields "<fechaAnulacion>" & .Fields("FechaA") & "</fechaAnulacion>"
''''      SetAdoFields "</detalleAnulados>"
''''      .MoveNext
''''    Loop
''''    SetAdoFields "</anulados>"
''''Else
''''    SetAdoFields "<anulados/>"
''''End If
''''End With
''''With AdoRRF.Recordset  ' rendimientos financieros
''''If .RecordCount > 0 Then
''''   .MoveFirst
''''   SetAdoFields "<rendFinancieros>"
''''   Do While Not .EOF
''''      SetAdoFields "<detalleRendFinancieros>"
''''        Select Case .Fields("CTD")
''''        Case "R"
''''        SetAdoFields "<retenido>12</retenido>"
''''        Case Else
''''        SetAdoFields "<retenido>13</retenido>"
''''        End Select
''''        SetAdoFields "<idRetenido>" & .Fields("CI_RUC") & "</idRetenido>"
''''        SetAdoFields "<tpCompb>" & .Fields("TTD") & "</tpCompb>"
''''        SetAdoFields "<tipoCompR>40</tipoCompR>"
''''        ' retencion en la fuente de rendimientos financieros
''''        Codigo = .Fields("Comprobante")
''''                 If AdoRRFF.Recordset.RecordCount > 0 Then
''''                    'AdoRI.Recordset.MoveFirst
''''                     AdoRRFF.Recordset.Find ("Comprobante = '" & Codigo & "' ")
''''                    If Not AdoRRFF.Recordset.EOF Then
''''                       SetAdoFields "<conRetT>" & AdoRRFF.Recordset.Fields("TD") & "</conRetT>"
''''                       SetAdoFields "<baseImponibleRetT>" & AdoRRFF.Recordset.Fields("Valor_Fact") & "</baseImponibleRetT>"
''''                       SetAdoFields "<codPorcRetT>" & AdoRRFF.Recordset.Fields("CodPorc") & "</codPorcRetT>"
''''                       SetAdoFields "<montoRetT>" & AdoRRFF.Recordset.Fields("Valor_Ret") & "</montoRetT>"
''''                       SetAdoFields "<serieRetT>" & AdoRRFF.Recordset.Fields("Serie") & "</serieRetT>"
''''                       SetAdoFields "<secuencialRetT>" & AdoRRFF.Recordset.Fields("Secuencial") & "</secuencialRetT>"
''''                       SetAdoFields "<autorizacionRetT>" & AdoRRFF.Recordset.Fields("Autorizacion") & "</autorizacionRetT>"
''''                       SetAdoFields "<fechaEmisionRetT>" & AdoRRFF.Recordset.Fields("FechaE") & "</fechaEmisionRetT>"
'''''                       AdoRI.Recordset.MoveLast
'''''                       AdoRI.Recordset.MoveNext
''''                    End If
''''                 End If
''''        SetAdoFields "</detalleRendFinancieros>"
''''    .MoveNext
''''   Loop
''''        SetAdoFields "</rendFinancieros>"
''''Else
''''SetAdoFields "<rendFinancieros/>"
''''End If
''''End With
''''Print #NumFile, "</iva>";
RatonNormal
'MsgBox "Proceso Terminado"
End Sub

Private Sub CmdGeneArch_Click()
  FechaIni = "01/01/2000"
  FechaFin = BuscarFecha(FechaSistema)
  Periodo_Contable = Ninguno
  sSQL = "SELECT * " _
       & "FROM Empresas " _
       & "WHERE Item <> '.' " _
       & "ORDER BY Item "
  SelectAdodc AdoEmpresa, sSQL
  With AdoEmpresa.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
         'MsgBox "..............."
          Label1.Caption = "MIGRACION DE DATOS DE "
          NumEmpresa = .Fields("Item")
          SerieFactura = .Fields("Serie_Factura")
          sSQL = "SELECT Fecha,COUNT(Fecha) As Cantidad " _
               & "FROM Trans_Retenciones " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "GROUP BY Fecha " _
               & "ORDER BY Fecha "
          SelectAdodc AdoFechas, sSQL
         'Comprobantes modificados con notas de debito en compras
          If AdoFechas.Recordset.RecordCount > 0 Then
             AdoFechas.Recordset.MoveFirst
             Label1.Caption = Label1.Caption & AdoFechas.Recordset.Fields("Fecha")
             If CFechaLong(AdoFechas.Recordset.Fields("Fecha")) < CFechaLong("01/01/2000") Then
                FechaIni = "01/01/2000"
             Else
                FechaIni = BuscarFecha(AdoFechas.Recordset.Fields("Fecha"))
             End If
             AdoFechas.Recordset.MoveLast
             FechaFin = BuscarFecha(AdoFechas.Recordset.Fields("Fecha"))
             Label1.Caption = Label1.Caption & " AL " & AdoFechas.Recordset.Fields("Fecha")
             'MsgBox FechaIni & " - " & FechaFin
             Label1.Caption = Label1.Caption & vbCrLf _
                            & .Fields("Empresa")
             Label1.Refresh
             sSQL = "DELETE * " _
                  & "FROM Trans_Air " _
                  & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             Conectar_Ado_Execute sSQL
             sSQL = "DELETE * " _
                  & "FROM Trans_Compras " _
                  & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             Conectar_Ado_Execute sSQL
             sSQL = "DELETE * " _
                  & "FROM Trans_Ventas " _
                  & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             Conectar_Ado_Execute sSQL
             sSQL = "DELETE * " _
                  & "FROM Trans_Importaciones " _
                  & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             Conectar_Ado_Execute sSQL
             sSQL = "DELETE * " _
                  & "FROM Trans_Exportaciones " _
                  & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' "
             Conectar_Ado_Execute sSQL
            'Empezamos la consulta de las Retenciones
             sSQL = "SELECT (TD&Secuencial) As CompMod,Fecha,Serie,FechaE,Autorizacion " _
                  & "FROM Trans_Retenciones " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND TT = 'C' " _
                  & "AND T <> 'A' " _
                  & "AND CodigoTR = 'IV' " _
                  & "GROUP BY Fecha,FechaE,TD,Secuencial,Serie,Autorizacion"
             SelectAdodc AdoCambio, sSQL ' comprobantes modificados con notas de debito en compras
             sSQL = "SELECT TD,Valor_Fact,CodPorc,Porc,Valor_Ret,Serie,Secuencial,Autorizacion, " _
                  & "FechaE,Fecha,Comprobante " _
                  & "FROM Trans_Retenciones " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND CodigoTR = 'RF' "
             SelectAdodc AdoRetFte, sSQL ' retencion en la fuente en compras
             sSQL = "SELECT T.*,C.CI_RUC " _
                  & "FROM Trans_Retenciones As T, Clientes As C " _
                  & "WHERE T.Item = '" & NumEmpresa & "' " _
                  & "AND T.TT = 'C' " _
                  & "AND T.T <> 'A' " _
                  & "AND CodigoTR = 'IV' " _
                  & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                  & "AND T.Codigo = C.Codigo " _
                  & "ORDER BY T.Codigo,T.Fecha,T.TP,T.Numero "
              SelectAdodc AdoCompras, sSQL ' retencion de iva compras
              sSQL = "SELECT Comprobante,Autorizacion As Contrato, Valor_Fact As Titulo_Oneroso, BImpotCero As Titulo_Gratuito " _
                   & "FROM Trans_Retenciones " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND CodigoTR = 'GE' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "ORDER BY Comprobante,Autorizacion "
              SelectAdodc AdoRI2S, sSQL  ' gasto electoral
              sSQL = "SELECT C.CI_RUC,T.* " _
                   & "FROM Trans_Retenciones As T, Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.TT = 'V' " _
                   & "AND T.T <> 'A' " _
                   & "AND CodigoTR = 'IV' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND T.Codigo = C.Codigo " _
                   & "ORDER BY T.Codigo,T.Fecha "
              SelectAdodc AdoVentas, sSQL  ' retencion de iva en ventas
              sSQL = "SELECT C.TD As CTD,C.Codigo,Sum(Valor_Fact) As Valor_Fact,T.TD As TRet,Porc,Sum(Valor_Ret)As Valor_Ret,month(T.Fecha) " _
                   & "FROM Trans_Retenciones As T,Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND T.CodigoTR = 'RV' " _
                   & "AND T.Codigo = C.Codigo " _
                   & "GROUP BY C.Codigo,C.TD,T.TD,Month(T.Fecha),Porc "
              SelectAdodc AdoRI2B, sSQL   ' retenciones en la fuente venta
              sSQL = "SELECT T.ConvInt,T.TD As TTD,T.Fecha,T.Aduana,C.TD As CTD,C.Codigo,C.Cliente,T.MontoIVA2 As ValorCif,T.SCT,T.BImpotCero,T.Valor_Fact,T.CodPorc,MontoIVA1,T.MontoICE,T.CPorcICE,T.RetICE,T.Comprobante " _
                   & "FROM Trans_Retenciones As T, Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.TT = 'I' " _
                   & "AND T.T <> 'A' " _
                   & "AND CodigoTR = 'IV' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND T.Codigo = C.Codigo " _
                   & "GROUP BY T.ConvInt,T.TD,T.Fecha,T.Aduana,C.TD,C.Codigo,C.Cliente,T.MontoIVA2,T.SCT,T.BImpotCero,T.Valor_Fact,T.CodPorc,MontoIVA1,T.MontoICE,T.CPorcICE,T.RetICE,T.Comprobante "
              SelectAdodc AdoImpor, sSQL  ' retencion de iva en importaciones
              sSQL = "SELECT TD,Valor_Fact,CodPorc,Valor_Ret,Fecha,Comprobante " _
                   & "FROM Trans_Retenciones " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND CodigoTR = 'RI' "
              SelectAdodc AdoRI, sSQL  ' retencion en la fuente de importaciones
              sSQL = "SELECT T.Dev,T.ConvInt,T.TD As TTD,T.Fecha,T.Aduana,C.TD As CTD,C.Codigo,C.Cliente,T.MontoIVA2 As ValorCif,T.SCT,T.BImpotCero,T.Valor_Fact,T.CodPorc,MontoIVA1,T.MontoICE,T.CPorcICE,T.RetICE,T.Comprobante " _
                   & "FROM Trans_Retenciones As T, Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.TT = 'E' " _
                   & "AND T.T <> 'A' " _
                   & "AND CodigoTR = 'IV' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND T.Codigo = C.Codigo " _
                   & "GROUP BY T.Dev,T.ConvInt,T.TD,T.Fecha,T.Aduana,C.TD,C.Codigo,C.Cliente,T.MontoIVA2,T.SCT,T.BImpotCero,T.Valor_Fact,T.CodPorc,MontoIVA1,T.MontoICE,T.CPorcICE,T.RetICE,T.Comprobante "
              SelectAdodc AdoExpor, sSQL 'retencion de iva en exportaciones
              sSQL = "SELECT TD,Valor_Fact,CodPorc,Valor_Ret,Fecha,Comprobante " _
                   & "FROM Trans_Retenciones " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND CodigoTR = 'RE' "
              SelectAdodc AdoRE, sSQL  'retencion en la fuente de exportaciones
              sSQL = "SELECT C.TD As CTD,C.Codigo,T.TD As TTD,Aduana,T.Fecha,FechaE,IdenCT,BImpotCero,Valor_Fact,MontoICE As Comision,SerieEx,MontoIVA1,PorRetIVA1,MontoRetIVA1,MontoIVA2,PorRetIVA2,MontoRetIVA2,T.Comprobante,T.Cambio " _
                   & "FROM Trans_Retenciones As T, Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND CodigoTR = 'TC' " _
                   & "AND T.T <> 'A' AND T.Codigo = C.Codigo "
              SelectAdodc AdoTC, sSQL ' tarjetas de credito
              sSQL = "SELECT TD,Valor_Fact,CodPorc,Valor_Ret,Serie,Secuencial,Autorizacion,FechaE,Fecha,Comprobante " _
                   & "FROM Trans_Retenciones " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND CodigoTR = 'RT' " _
                   & "AND T <> 'A' " 'tarjetas de credito retencion en la fuente
              SelectAdodc AdoRT, sSQL
              sSQL = "SELECT TD,Serie,Secuencial,Autorizacion,FechaA,Comprobante " _
                   & "FROM Trans_Retenciones " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND T = 'A' " 'comprobantes anulados
              SelectAdodc AdoAnulados, sSQL
              sSQL = "SELECT C.TD As CTD,C.Codigo,T.TD As TTD,T.Comprobante " _
                   & "FROM Trans_Retenciones As T, Clientes As C " _
                   & "WHERE T.Item = '" & NumEmpresa & "' " _
                   & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
                   & "AND CodigoTR = 'FF' " _
                   & "AND T.T <> 'A' AND T.Codigo = C.Codigo "
              SelectAdodc AdoRRF, sSQL ' rendimientos financieros
              GenerarArchivoATXML False
          End If
         .MoveNext
       Loop
   End If
  End With
  MsgBox "Proceso Terminado"
  Unload FMigraAT
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FMigraAT
   SumaAnulados = 0
  'Abriendo bases relacionadas
   ConectarAdodc AdoEmpresa ' datos de empresa
   ConectarAdodc AdoCompras  'facturas de compras
   ConectarAdodc AdoCambio  ' comprobantes modificados con notas de debito en compras
   ConectarAdodc AdoRetFte  ' retencion en la fuente en compras
   ConectarAdodc AdoRI2S  ' gasto electoral
   ConectarAdodc AdoVentas  ' facturas de ventas
   ConectarAdodc AdoRI2B  ' retenciones en la fuente de ventas
   ConectarAdodc AdoImpor  ' facturas de importaciones
   ConectarAdodc AdoRI  ' retenciones en la fuente de importaciones
   ConectarAdodc AdoExpor  ' facturas de exportaciones
   ConectarAdodc AdoRE    ' retenciones en la fuente de exportaciones
   ConectarAdodc AdoTC  ' tarjetas de credito
   ConectarAdodc AdoRT   ' retencion en la fuente de tarjetas de credito
   ConectarAdodc AdoAnulados  ' comprobantes anulados
   ' rendimientos financieros
   ConectarAdodc AdoFechas 'Fechas procesadas segun las retenciones
   ConectarAdodc AdoRRF   'rendimientos financieros
   ConectarAdodc AdoRRFF ' retencion en la fuente de rend financieros
End Sub

