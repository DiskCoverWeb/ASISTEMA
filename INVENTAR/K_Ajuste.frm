VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Kard_IngAjuste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS DE AJUSTES"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "K_Ajuste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11715
   Begin MSDataListLib.DataList DLArt 
      Height          =   1425
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2514
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoSQL 
      Height          =   330
      Left            =   150
      Top             =   3705
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "SQLs"
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
   Begin MSAdodcLib.Adodc AdoCtaObra 
      Height          =   330
      Left            =   150
      Top             =   3945
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "CtaObra"
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   150
      Top             =   4185
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Kardex"
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
      Left            =   150
      Top             =   4425
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   2250
      Top             =   4035
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc AdoIngArt 
      Height          =   330
      Left            =   2250
      Top             =   4275
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "IngArt"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   2250
      Top             =   4515
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   2205
      Top             =   3675
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc8"
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
      Left            =   7545
      Top             =   3840
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoCodBar 
      Height          =   330
      Left            =   7545
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "CodBar"
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
      Left            =   7545
      Top             =   4320
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   7545
      Top             =   4560
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Comp"
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
      Left            =   7545
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7545
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc2"
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
   Begin VB.TextBox TextTotal 
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
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "K_Ajuste.frx":0442
      Top             =   3045
      Width           =   1590
   End
   Begin VB.TextBox TextDesc 
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
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "K_Ajuste.frx":0444
      Top             =   2730
      Width           =   1590
   End
   Begin VB.CheckBox CheckCtaObra 
      Caption         =   "Contra Cuenta"
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
      Top             =   1050
      Width           =   2850
   End
   Begin VB.TextBox TextIVA 
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
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "K_Ajuste.frx":0446
      Top             =   5565
      Width           =   1590
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
      Left            =   1470
      Picture         =   "K_Ajuste.frx":0448
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5565
      Width           =   1170
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
      Height          =   330
      Left            =   4725
      MaxLength       =   50
      TabIndex        =   5
      Top             =   630
      Width           =   5685
   End
   Begin MSDBCtls.DBCombo DBCInv 
      Bindings        =   "K_Ajuste.frx":0D12
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   1785
      Width           =   7155
      _ExtentX        =   12621
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
      Left            =   210
      Picture         =   "K_Ajuste.frx":0D28
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5565
      Width           =   1170
   End
   Begin VB.TextBox TextVUnit 
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
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "K_Ajuste.frx":116A
      Top             =   2415
      Width           =   1590
   End
   Begin VB.TextBox TextEntrada 
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
      IMEMode         =   3  'DISABLE
      Left            =   8820
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   17
      Text            =   "K_Ajuste.frx":116C
      Top             =   2100
      Width           =   1590
   End
   Begin MSDBCtls.DBCombo DBCBenef 
      Bindings        =   "K_Ajuste.frx":116E
      DataSource      =   "DataBenef"
      Height          =   315
      Left            =   4725
      TabIndex        =   1
      Top             =   210
      Width           =   5685
      _ExtentX        =   10028
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   210
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSDBCtls.DBCombo DBCCtaObra 
      Bindings        =   "K_Ajuste.frx":1186
      DataSource      =   "DataCtaObra"
      Height          =   315
      Left            =   3045
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   7365
      _ExtentX        =   12991
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
   Begin MSDataGridLib.DataGrid DGCodBar 
      Bindings        =   "K_Ajuste.frx":11A0
      Height          =   1755
      Left            =   4410
      TabIndex        =   29
      Top             =   3675
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   3096
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
   Begin VB.Label Label8 
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
      Height          =   330
      Left            =   7350
      TabIndex        =   14
      Top             =   3045
      Width           =   1485
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESCUENTOS"
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
      TabIndex        =   27
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   8820
      TabIndex        =   24
      Top             =   5985
      Width           =   1590
   End
   Begin VB.Label Label10 
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
      Left            =   7560
      TabIndex        =   23
      Top             =   5985
      Width           =   1275
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A."
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
      Left            =   7560
      TabIndex        =   21
      Top             =   5565
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO"
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
      TabIndex        =   4
      Top             =   630
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Top             =   210
      Width           =   1380
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR"
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
      TabIndex        =   0
      Top             =   210
      Width           =   1590
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR UNIT."
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
      TabIndex        =   13
      Top             =   2415
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CANTIDAD"
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
      TabIndex        =   12
      Top             =   2100
      Width           =   1485
   End
   Begin VB.Label LabelUnidad 
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
      Left            =   8820
      TabIndex        =   16
      Top             =   1785
      Width           =   1590
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD"
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
      TabIndex        =   11
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE INVENTARIO"
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
      Top             =   1470
      Width           =   7155
   End
   Begin VB.Label LabelCodigo 
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
      Left            =   8820
      TabIndex        =   15
      Top             =   1470
      Width           =   1590
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      TabIndex        =   10
      Top             =   1470
      Width           =   1485
   End
End
Attribute VB_Name = "Kard_IngAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckCtaObra_Click()
 If CheckCtaObra.Value = 1 Then
    DCCtaObra.Visible = True
 Else
    DCCtaObra.Visible = False
 End If
End Sub

Private Sub Command1_Click()
  sSQL = "SELECT * FROM Asiento_K_" & CodigoUsuario & " "
  sSQL = sSQL & "ORDER BY CTA_INV,CODIGO_INV "
  SelectData AdoIngArt, sSQL
  GenerarArchivoPlano Kard_IngAjuste, AdoIngArt, "K" & NumComp & "_" & CodigoUsuario & ".TXT", True
  FechaTexto = MBoxFechaI.Text
  CodigoBenef = Ninguno
  CodigoBenef = SinEspaciosIzq(DCBenef.Text)
  FechaTexto = MBoxFechaI.Text
  With AdoIngArt.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       NumComp = ReadSetDataNum("Diario", True, True)
       Asiento = ReadSetDataNum("Kardex", True, True)
       sSQL = "SELECT * FROM Asiento_CBK_" & CodigoUsuario & " "
       sSQL = sSQL & "ORDER BY CODIGO_INV,CODIGO_B "
       SelectData AdoCodBar, sSQL
       If AdoCodBar.Recordset.RecordCount > 0 Then
          Do While Not AdoCodBar.Recordset.EOF
             Codigo1 = AdoCodBar.Recordset.Fields("CODIGO_B")
             Codigo2 = AdoCodBar.Recordset.Fields("CODIGO_INV")
             sSQL = "SELECT * FROM Codigos_Barra "
             sSQL = sSQL & "WHERE Codigo_Bar = '" & Codigo1 & "' "
             sSQL = sSQL & "AND Codigo_Inv = '" & Codigo2 & "' "
             SelectData AdoKardex, sSQL, False
             If AdoKardex.Recordset.RecordCount > 0 Then
                AdoKardex.Recordset.Edit
                AdoKardex.Recordset.Fields("T") = Cancelado
             Else
                AdoKardex.Recordset.AddNew
                AdoKardex.Recordset.Fields("T") = Pendiente
                AdoKardex.Recordset.Fields("Codigo_P") = CodigoBenef
                AdoKardex.Recordset.Fields("Codigo_Inv") = Codigo2
                AdoKardex.Recordset.Fields("Codigo_Bar") = Codigo1
                AdoKardex.Recordset.Fields("Fecha") = MBoxFechaI.Text
             End If
             AdoKardex.Recordset.Update
             AdoCodBar.Recordset.MoveNext
          Loop
       End If
      .MoveFirst
       CodigoInv = .Fields("Codigo_Inv")
       Cta_Inventario = .Fields("CTA_INV")
       Contra_Cta = .Fields("CTA")
       ValorDH = 0: Total = 0
      'Llenamos los datos ingresados al Kardex
       Do While Not .EOF
          If Cta_Inventario <> .Fields("CTA_INV") Or CodigoInv <> .Fields("Codigo_Inv") Then
             InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
             Cta_Inventario = .Fields("CTA_INV")
             'Contra_Cta = .Fields("CTA")
             CodigoInv = .Fields("Codigo_Inv")
             ValorDH = 0
          End If
          sSQL = "SELECT * FROM Kardex "
          SelectData AdoKardex, sSQL, False
          AdoKardex.Recordset.AddNew
          AdoKardex.Recordset.Fields("T") = Normal
          AdoKardex.Recordset.Fields("Codigo_Inv") = .Fields("Codigo_Inv")
          AdoKardex.Recordset.Fields("Codigo_P") = CodigoBenef
          AdoKardex.Recordset.Fields("Fecha") = FechaTexto
          AdoKardex.Recordset.Fields("TP") = CompDiario
          AdoKardex.Recordset.Fields("Numero") = NumComp
          AdoKardex.Recordset.Fields("Kardex") = Asiento
          AdoKardex.Recordset.Fields("Salida") = 0
          AdoKardex.Recordset.Fields("Descuento") = .Fields("P_DESC")
          AdoKardex.Recordset.Fields("Entrada") = .Fields("CANT_ES")
          AdoKardex.Recordset.Fields("Valor_Total") = .Fields("VALOR_TOTAL")
          AdoKardex.Recordset.Fields("Cantidad") = .Fields("CANTIDAD")
          AdoKardex.Recordset.Fields("Valor_Unitario") = .Fields("VALOR_UNIT")
          AdoKardex.Recordset.Fields("Saldo_Total") = .Fields("SALDO")
          AdoKardex.Recordset.Fields("Cta_Inv") = .Fields("CTA_INV")
          AdoKardex.Recordset.Fields("Cta") = .Fields("CTA")
          AdoKardex.Recordset.Fields("Cta_Obra") = Ninguno
          AdoKardex.Recordset.Fields("Orden_No") = Ninguno
          AdoKardex.Recordset.Update
          ValorDH = ValorDH + .Fields("VALOR_TOTAL")
          Total = Total + .Fields("VALOR_TOTAL")
         'Actualizar stock
          sSQL = "UPDATE Productos SET Stock = " & .Fields("CANTIDAD") & " "
          sSQL = sSQL & "WHERE Codigo_Inv = '" & .Fields("Codigo_Inv") & "' "
          UpdateData AdoKardex, sSQL
          AdoKardex.Recordset.AddNew
         .MoveNext
       Loop
      'Procesar Diario
       Cta_Aux = Contra_Cta
       If CheckCtaObra.Value = 1 Then Cta_Aux = SinEspaciosIzq(DCCtaObra.Text)
       InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
       InsertarAsientos AdoAsientos, Cta_IVA_Inventario, 0, Total_IVA, 0
       InsertarAsientos AdoAsientos, Cta_Aux, 0, 0, Total + Total_IVA
       CodigoBenef = SinEspaciosIzq(DCBenef.Text)
       With AdoSubCtas.Recordset
         If ValorDH <> 0 Then
           .AddNew
           .Fields("TC") = "P"
           .Fields("Codigo") = CodigoBenef
           .Fields("Beneficiario") = SinCodigoIzq(DCBenef.Text)
           .Fields("Cta") = Contra_Cta
           .Fields("DH") = OpcDH
           .Fields("Valor") = Total
           .Fields("Fecha_V") = FechaTexto
           .Fields("Valor_ME") = 0
           .Fields("Factura") = 0
           .Update
         End If
       End With
       Si_No = True
       Co.T = Normal
       Co.TP = CompDiario
       Co.Fecha = FechaTexto
       Co.Numero = NumComp
       Co.Concepto = TextConcepto.Text
       Co.Beneficiario = Ninguno
       Co.Efectivo = 0
       Co.Monto_Total = ValorDH
       GrabarComprobantes Co, AdoAsientos, AdoSubCtas, , AdoRet
       sSQL = "SELECT * FROM Beneficiarios "
       sSQL = sSQL & "WHERE Codigo = '" & CodigoBenef & "' "
       sSQL = sSQL & "AND TC = 'R' "
       SelectData AdoSQL, sSQL, False
       RatonNormal
       ImprimirComprobantesDe False, CompDiario, NumComp, NumEmpresa, AdoComp, AdoTrans, , AdoRet
       ImprimirNotaInventario AdoSQL, AdoIngArt, NumComp, Si_No, Ninguno
       IniciarAsientosDataDe AdoAsientos, AdoSubCtas, AdoBanco, AdoRet, AdoIngArt
   End If
  End With
  Unload Kard_IngAjuste
End Sub

Private Sub Command2_Click()
  Unload Kard_IngAjuste
End Sub

Private Sub DBCInv_LostFocus()
  Codigo = SinEspaciosIzq(DBCInv.Text)
  LongStrg = Len(Codigo)
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomProd "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE Mid(Codigo_Inv,1," & LongStrg & ") = '" & Codigo & "' "
  sSQL = sSQL & "AND TP = 'P' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBList DLArt, AdoArt, sSQL, "NomProd"
  DLArt.SetFocus
End Sub

Private Sub DGrid1_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoIngArt)
End Sub

Private Sub DLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: Command1.SetFocus
    Case vbKeyReturn: SiguienteControl
  End Select
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub TextCodigoBar_GotFocus()
  TextCodigoBar.Text = ""
End Sub

Private Sub TextCodigoBar_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     FrameBar.Visible = False
     TextEntrada.Text = LabelCantTotal.Caption
     TextEntrada.SetFocus
  End If
End Sub

Private Sub TextCodigoBar_LostFocus()
   TextoValido TextCodigoBar, , True
   With AdoCodBar.Recordset
     If TextCodigoBar.Text <> Ninguno Then
       .AddNew
       .Fields("CODIGO_INV") = Codigo
       .Fields("CODIGO_B") = TextCodigoBar.Text
       .Update
        TextCodigoBar.SetFocus
     End If
     LabelCantTotal.Caption = Format(Val(.RecordCount), "#,##0")
   End With
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextEntrada_Change()
  Entrada = Val(TextEntrada.Text)
  ValorTotal = ValorUnit * Entrada
  TextTotal.Text = Format(ValorTotal, "#,##0.00")
End Sub

Private Sub TextEntrada_GotFocus()
  MarcarTexto TextEntrada
  OpcDH = 1
  Codigo = SinEspaciosIzq(DLArt.Text)
  'MsgBox "|" & Codigo & "|"
  sSQL = "SELECT * FROM Productos " _
       & "WHERE Codigo_Inv = '" & Codigo & "' " _
       & "AND TP = 'P' "
  SelectData AdoSQL, sSQL, False
  Codigo = ""
  Cantidad = 0
  ValorUnitAnt = 0
  SaldoAnterior = 0
  Saldo = 0
  Producto = Ninguno
  With AdoSQL.Recordset
   If .RecordCount > 0 Then
       Unidad = .Fields("Unidad")
       Codigo = .Fields("Codigo_Inv")
       Producto = .Fields("Producto")
       Cta_Inventario = .Fields("Cta_Inv")
       Contra_Cta = .Fields("Cta")
       Contra_Cta1 = .Fields("Cta1")
   End If
  End With
  If Codigo <> "" Then
     sSQL = "SELECT * FROM Kardex "
     sSQL = sSQL & "WHERE Codigo_Inv = '" & Codigo & "' "
     sSQL = sSQL & "ORDER BY Fecha,TP,Numero,Kardex "
     SelectData AdoSQL, sSQL, False
     With AdoSQL.Recordset
      If .RecordCount > 0 Then
         .MoveLast
          Cantidad = .Fields("Cantidad")
          ValorUnitAnt = .Fields("Valor_Unitario")
          SaldoAnterior = .Fields("Saldo_Total")
         'If OpcE.Value Then
          TextVUnit.Text = ValorUnitAnt
      End If
     End With
  End If
  LabelCodigo.Caption = Codigo
  LabelUnidad.Caption = Unidad
  'TextEntrada.Text = ""
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyB Then
     sSQL = "SELECT * FROM Asiento_CBK_" & CodigoUsuario & " "
     sSQL = sSQL & "WHERE CODIGO_INV = '" & Codigo & "' "
     SelectDataGrid DGCodBar, AdoCodBar, sSQL
     DGCodBar.Visible = True
     FrameBar.Visible = True
     TextCodigoBar.SetFocus
  End If
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True
  Entrada = Val(TextEntrada.Text)
End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  CTAsientoContable
  IniciarAsientosDataDe AdoAsientos, AdoSubCtas, AdoBanco, AdoRet, AdoIngArt
  sSQL = "SELECT Codigo_Inv & Space(10) & Producto As NomInv "
  sSQL = sSQL & "FROM Productos "
  sSQL = sSQL & "WHERE TP = 'I' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  SelectDBCombo DBCInv, AdoInv, sSQL, "NomInv", False
  SelectDataGrid DGrid1, AdoIngArt, "Asiento_K_" & CodigoUsuario & " "
  Total_IVA = 0
  Label11.Visible = False
  TextIVA.Visible = False
  Label3.Caption = " RESPONSABLE:"
  sSQL = "SELECT Codigo & Space(5) & Beneficiario As Nombre_Cta "
  sSQL = sSQL & "FROM Beneficiarios "
  sSQL = sSQL & "WHERE TC = 'R' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DCBenef, AdoBenef, sSQL, "Nombre_Cta", False
  sSQL = "SELECT Codigo & Space(4) & Cuenta As Nomb_Cta "
  sSQL = sSQL & "FROM Catalogo "
  sSQL = sSQL & "WHERE DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBCombo DCCtaObra, AdoCtaObra, sSQL, "Nomb_Cta", False
  FechaValida MBoxFechaI
  DCBenef.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Kard_IngAjuste
  ConectarAdodc AdoInv
  ConectarAdodc AdoArt
  ConectarAdodc AdoSQL
  ConectarAdodc AdoBenef
  ConectarAdodc AdoKardex
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoIngArt
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoSubCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoComp
  ConectarAdodc AdoRet
  ConectarAdodc AdoBanco
  ConectarAdodc AdoCodBar
End Sub

Private Sub TextIVA_GotFocus()
  TextIVA.Text = ""
  Sumatoria = TotalInventario(AdoIngArt)
End Sub

Private Sub TextIVA_LostFocus()
 TextoValido TextIVA, True
 Sumatoria = TotalInventario(AdoIngArt)
 Total_IVA = Val(TextIVA.Text) * Sumatoria / 100
 TextIVA.Text = Format(Total_IVA, "#,##0.00")
 Label1.Caption = Format(Sumatoria + Total_IVA, "#,##0.00")
End Sub

Private Sub TextTotal_LostFocus()
   Total_Desc = Val(TextDesc.Text) / 100
   ValorUnit = Val(TextVUnit.Text) - (Val(TextVUnit.Text) * Total_Desc)
   Entrada = Val(TextEntrada.Text)
   If OpcDH = 1 Then
      ValorUnit = Val(TextVUnit.Text)
   Else
      ValorUnit = ValorUnitAnt
   End If
   ValorTotal = ValorUnit * Entrada
   TextVUnit.Text = Format(ValorUnit, "#,##0.0000")
  'llenamos el ultimo saldo del kardex
   sSQL = "SELECT * FROM Asiento_K_" & CodigoUsuario & " "
   sSQL = sSQL & "WHERE Codigo_Inv = '" & Codigo & "' "
   SelectData AdoSQL, sSQL, False
   If Entrada > 0 And ValorUnit > 0 And AdoSQL.Recordset.RecordCount <= 0 Then
      With AdoIngArt.Recordset
          .AddNew
          .Fields("DH") = OpcDH
          .Fields("CODIGO_INV") = Codigo
          .Fields("PRODUCTO") = Producto
          .Fields("CANT_ES") = Entrada
          .Fields("P_DESC1") = 0
          .Fields("P_DESC") = Total_Desc
          .Fields("VALOR_UNIT") = Round(ValorUnit, 4)
          .Fields("VALOR_TOTAL") = Round(ValorTotal)
           If OpcDH = 1 Then
              Cantidad = Cantidad + Entrada
              Saldo = Round(SaldoAnterior + ValorTotal)
           Else
              Cantidad = Cantidad - Entrada
              Saldo = Round(SaldoAnterior - ValorTotal)
           End If
          .Fields("CTA_INV") = Cta_Inventario
          .Fields("CTA") = Contra_Cta
          .Fields("CANTIDAD") = Cantidad
          .Fields("SALDO") = Saldo
          .Fields("UNIDAD") = LabelUnidad.Caption
          .Update
      End With
      TextTotal.Text = Format(ValorTotal, "#,##0.00")
      Sumatoria = TotalInventario(AdoIngArt)
      Label1.Caption = Format(Sumatoria, "#,##0.00")
      Total_IVA = 0
      DLArt.SetFocus
   End If
End Sub

Private Sub TextVUnit_Change()
  ValorUnit = Val(TextVUnit.Text)
  ValorTotal = ValorUnit * Entrada
  TextTotal.Text = Format(ValorTotal, "#,##0.00")
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True
  ValorTotal = ValorUnit * Entrada
  TextTotal.Text = Format(ValorTotal, "#,##0.00")
End Sub

Public Function TotalInventario(DtaInv As Data)
Dim TotalInvs As Double
  TotalInvs = 0
  With DtaInv.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           TotalInvs = TotalInvs + .Fields("VALOR_TOTAL")
          .MoveNext
        Loop
    End If
  End With
  TotalInventario = TotalInvs
End Function
