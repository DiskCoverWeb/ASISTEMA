VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Kard_Ing_DUI 
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS/EGRESOS DE MARCADERIA"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "KardDui.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OpcI 
      Caption         =   "Ingreso"
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
      Left            =   12705
      TabIndex        =   6
      Top             =   105
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.OptionButton OpcE 
      Caption         =   "Egreso"
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
      Left            =   12705
      TabIndex        =   7
      Top             =   420
      Width           =   960
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   2730
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   7350
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "KardDui.frx":0442
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   7035
      TabIndex        =   13
      Top             =   1155
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "FAMILIA"
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
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "KardDui.frx":0458
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   1470
      TabIndex        =   11
      Top             =   1155
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "BODEGA"
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
   Begin MSDataGridLib.DataGrid DGKardex 
      Bindings        =   "KardDui.frx":0470
      Height          =   5265
      Left            =   105
      TabIndex        =   16
      Top             =   1995
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   9287
      _Version        =   393216
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   255
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
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "KardDui.frx":0488
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   7035
      TabIndex        =   5
      Top             =   420
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoCtaObra 
      Height          =   330
      Left            =   4305
      Top             =   4305
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   2310
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   4305
      Top             =   4620
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
      Left            =   2310
      Top             =   4935
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   315
      Top             =   4935
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   4305
      Top             =   4935
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
   Begin VB.TextBox TextOrden 
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
      Left            =   105
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "KardDui.frx":049F
      Top             =   1155
      Width           =   1275
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
      Height          =   750
      Left            =   1365
      Picture         =   "KardDui.frx":04AB
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7350
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
      Left            =   1470
      MaxLength       =   100
      TabIndex        =   15
      Top             =   1575
      Width           =   11145
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
      Height          =   750
      Left            =   105
      Picture         =   "KardDui.frx":0D75
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7350
      Width           =   1170
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   315
      Top             =   4620
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   315
      Top             =   4305
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
      Caption         =   "Bodega"
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
      Left            =   2310
      Top             =   4305
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
   Begin MSAdodcLib.Adodc AdoIVA 
      Height          =   330
      Left            =   315
      Top             =   5250
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
      Caption         =   "IVA"
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
   Begin MSDataListLib.DataCombo DCCtaObra 
      Bindings        =   "KardDui.frx":11B7
      DataSource      =   "AdoCtaObra"
      Height          =   315
      Left            =   1470
      TabIndex        =   3
      Top             =   420
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Contra cuenta"
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE CONTRA CUENTA"
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
      TabIndex        =   2
      Top             =   105
      Width           =   5475
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GRUPO DE MERCADERIA"
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
      Left            =   7035
      TabIndex        =   12
      Top             =   840
      Width           =   5580
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGAS"
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
      Left            =   1470
      TabIndex        =   10
      Top             =   840
      Width           =   5475
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
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
      TabIndex        =   8
      Top             =   840
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
      Left            =   105
      TabIndex        =   14
      Top             =   1575
      Width           =   1380
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
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO"
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
      TabIndex        =   4
      Top             =   105
      Width           =   5580
   End
End
Attribute VB_Name = "Kard_Ing_DUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  Total_RetCta = 0
  RatonNormal
  Mensajes = "Seguro de Grabar?"
  Titulo = "GRABACION DEL COMPROBANTE"
  If BoxMensaje = vbYes Then
  RatonReloj
  FechaTexto = MBFechaI.Text
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Cta "
  Select_Adodc AdoArt, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, 4
  RatonReloj
  Si_No = True
  CodigoBenef = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Cliente Like '" & Cadena & "' ")
       If Not .EOF Then CodigoBenef = .fields("Codigo")
   End If
  End With
  Valor = 0
  TotalInventario
  
  FechaTexto = MBFechaI.Text
  TextoValido TextOrden, , True
  'MsgBox CodigoBenef
  If CodigoBenef <> Ninguno Then
  With AdoKardex.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       NumComp = ReadSetDataNum("Diario", True, True)
      .MoveFirst
       CodigoInv = .fields("Codigo_Inv")
       Cta_Inventario = .fields("CTA_INVENTARIO")
       Contra_Cta = .fields("CONTRA_CTA")
       Total = 0: ValorDH = 0
      'Llenamos los datos ingresados al Kardex
       Do While Not .EOF
          SetAdoAddNew "Trans_Kardex"
          SetAdoFields "T", Normal
          SetAdoFields "Codigo_Inv", .fields("Codigo_Inv")
          SetAdoFields "Codigo_P", CodigoBenef
          SetAdoFields "Fecha", FechaTexto
          SetAdoFields "TP", CompDiario
          SetAdoFields "Numero", NumComp
          SetAdoFields "Descuento", .fields("P_DESC")
          SetAdoFields "Descuento1", .fields("P_DESC1")
          If .fields("DH") = 1 Then
              SetAdoFields "Entrada", .fields("CANT_ES")
          Else
              SetAdoFields "Salida", .fields("CANT_ES")
              Si_No = False
          End If
          SetAdoFields "Valor_Total", .fields("VALOR_TOTAL")
          SetAdoFields "Producto", .fields("PRODUCTO")
          SetAdoFields "Existencia", .fields("CANTIDAD")
          SetAdoFields "Valor_Unitario", .fields("VALOR_UNIT")
          SetAdoFields "Total", .fields("SALDO")
          SetAdoFields "Cta_Inv", .fields("CTA_INVENTARIO")
          SetAdoFields "Contra_Cta", .fields("CONTRA_CTA")
          SetAdoFields "Orden_No", TextOrden.Text
          SetAdoFields "CodBodega", .fields("CodBod")
          SetAdoFields "Codigo_Barra", .fields("COD_BAR")
          If Inv_Promedio Then
             Cantidad = .fields("CANTIDAD")
             Saldo = .fields("SALDO")
             If Cantidad <= 0 Then Cantidad = 1
             SetAdoFields "Costo", Saldo / Cantidad
          End If
          SetAdoFields "Item", NumEmpresa
          SetAdoUpdate
          Total = Total + .fields("VALOR_TOTAL")
          ValorDH = ValorDH + .fields("VALOR_TOTAL")
         .MoveNext
       Loop
      'Procesar Diario
       If OpcI.value Then
          InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
          InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
       Else
          InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
          InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
          Si_No = False
       End If
   End If
  End With
'''  If Total <> 0 Then
'''     SetAdoAddNew "Facturas"
'''     SetAdoFields "T", "P"
'''     SetAdoFields "TC", "P"
'''     SetAdoFields "CodigoC", CodigoBenef
'''     SetAdoFields "Fecha", FechaTexto
'''     SetAdoFields "Fecha_C", FechaTexto
'''     SetAdoFields "Fecha_V", FechaTexto
'''     SetAdoFields "SubTotal", Total - Total_IVA
'''     SetAdoFields "Total_MN", Total
'''     SetAdoFields "Saldo_MN", Total - Total_RetCta + CCur(TxtDifxDec.Text)
'''     SetAdoFields "Cta_CxP", Contra_Cta
'''     SetAdoFields "Cod_Ejec", CodigoUsuario
'''     SetAdoFields "Item", NumEmpresa
'''     SetAdoFields "CodigoU", CodigoUsuario
'''     SetAdoFields "Factura", Val(TextOrden.Text)
'''     SetAdoUpdate
'''  End If
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       Debe = 0: Haber = 0
       Do While Not .EOF
          Debe = Debe + .fields("DEBE")
          Haber = Haber + .fields("HABER")
         .MoveNext
       Loop
       If (Debe - Haber) <> 0 Then MsgBox "Verifique el comprobante, no cuadra por: " & Round(Debe - Haber, 2)
      .MoveFirst
       RatonReloj
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto.Text
          Co.CodigoB = CodigoBenef
          Co.Efectivo = 0
          Co.Monto_Total = Total - Total_RetCta
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
          Co.Item = NumEmpresa
          If TextOrden.Text <> Ninguno Then Co.Concepto = Co.Concepto & ", Orden No. " & TextOrden.Text
          Grabar_Comprobante Co
          Ln_No = 1
          'ImprimirComprobantesDe False, Co
          
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          sSQL = "DELETE * " _
               & "FROM Asiento_R " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND T_No = " & Trans_No & " " _
               & "AND CodigoU = '" & CodigoUsuario & "' "
          Ejecutar_SQL_SP sSQL
          sSQL = "SELECT * " _
               & "FROM Asiento_K " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "AND T_No = " & Trans_No & " "
          SQLDec = "VALOR_UNIT 4|,VALOR_TOTAL 4|,CANTIDAD 4|,SALDO 4|."
          Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
   End If
  End With
  Else
     MsgBox "Beneficiario no asignado"
  End If
  End If
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Kard_Ing_DUI
End Sub

Private Sub DCBenef_LostFocus()
  InvImp = False
  CodigoBenef = Ninguno
  CodigoCliente = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Beneficiario Like '" & Cadena & "' ")
       If Not .EOF Then
          CodigoBenef = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
       End If
   End If
  End With
  TextConcepto.Text = DCBenef.Text & " "
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBodega_LostFocus()
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega = '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .fields("CodBod")
   End If
  End With
End Sub

Private Sub DCCtaObra_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCCtaObra_LostFocus()
    Contra_Cta = SinEspaciosDer(DCCtaObra)
    ListarProveedorUsuario Contra_Cta
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
 CodigoInv = SinEspaciosIzq(DCTInv.Text)
 
 sSQL = "SELECT * " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' " _
      & "AND NOT Cta_Inventario  IN ('.','0') " _
      & "AND TC = 'P' " _
      & "ORDER BY Producto "
 Select_Adodc AdoInv, sSQL
 
 
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT 4|,VALOR_TOTAL 4|,CANTIDAD 4|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
 RatonNormal
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

'''Private Sub LstSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
'''  PresionoEnter KeyCode
'''  If KeyCode = vbKeyReturn Then
'''     FrmSubCta.Visible = False
'''     DCRet.SetFocus
'''  End If
'''End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub OpcE_Click()
'   Total_IVA = 0
'   Label3.Caption = " RESPONSABLE:"
'   If Empleados Then
'      ListarProveedorUsuario True
'   Else
'      ListarProveedorUsuario False
'   End If
   MBFechaI.SetFocus
End Sub

Private Sub OpcI_Click()
'   Total_IVA = 0
'   Label3.Caption = " PROVEEDOR:"
'   If Empleados Then
'      ListarProveedorUsuario False
'   Else
'      ListarProveedorUsuario True
'   End If
   MBFechaI.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
 MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

''Private Sub TextEntrada_GotFocus()
''  MarcarTexto TextEntrada
''  If OpcI.value Then OpcDH = 1 Else OpcDH = 2
''  If OpcE.value Then
''     If Precio = 0 Then
''        MsgBox "Warning: " & vbCrLf _
''             & "        Falta de Ingresar en este codigo" & vbCrLf _
''             & "        " & Codigo & ": La entrada inicial"
''     End If
''  End If
''End Sub

Private Sub Form_Activate()
' Consultamos las cuentas de la tabla
 Trans_No = 95
 Ln_No = 1
 TipoDoc = CompDiario
 FechaValida MBFechaI
 FechaComp = MBFechaI
 IniciarAsientosAdo AdoAsientos
 If Inv_Promedio Then
    Kard_Ing_DUI.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    Kard_Ing_DUI.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
 End If
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
 
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  
 sSQL = "SELECT Cuenta & '  ->  ' & Codigo As Nomb_Cta " _
      & "FROM Catalogo_Cuentas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC IN ('RP','C','P','HC','I','G') " _
      & "AND DG = 'D' " _
      & "ORDER BY Codigo "
 SelectDB_Combo DCCtaObra, AdoCtaObra, sSQL, "Nomb_Cta"
 
 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Bodega "
 SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
 
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT 4|,VALOR_TOTAL 4|,CANTIDAD 4|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
 RatonNormal
 MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoInv
  ConectarAdodc AdoRet
  ConectarAdodc AdoIVA
  ConectarAdodc AdoArt
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoAsientos
End Sub

Private Sub TextOrden_GotFocus()
 ' Stock_Actual_Inventario MBFechaI, CodigoInv
  MarcarTexto TextOrden
End Sub

Private Sub TextOrden_LostFocus()
  TextoValido TextOrden, , True
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .MoveFirst
        Do While Not .EOF
           Total = Total + .fields("VALOR_TOTAL")
           Total_IVA = Total_IVA + .fields("IVA")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Round(Total, 2)
  Total_IVA = Round(Total_IVA, 2)
End Sub

Public Sub ListarProveedorUsuario(Contra_Cta As String)
   Codigo = Leer_Cta_Catalogo(Contra_Cta)
   Select Case SubCta
     Case "C", "P"
          sSQL = "SELECT C.Cliente As Beneficiario, CP.Codigo " _
               & "FROM Clientes As C, Catalogo_CxCxP As CP " _
               & "WHERE CP.Cta = '" & Contra_Cta & "' " _
               & "AND CP.Item = '" & NumEmpresa & "' " _
               & "AND CP.Periodo = '" & Periodo_Contable & "' " _
               & "AND C.Codigo = CP.Codigo " _
               & "ORDER BY C.Cliente "
     Case Else
          sSQL = "SELECT Detalle As Beneficiario, Codigo " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE TC = '" & SubCta & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY Detalle "
   End Select
   SelectDB_Combo DCBenef, AdoBenef, sSQL, "Beneficiario"
End Sub
