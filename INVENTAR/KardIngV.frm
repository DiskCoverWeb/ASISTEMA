VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Kard_Ing_Ventas 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS/EGRESOS"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13050
   Icon            =   "KardIngV.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   13050
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "KardIngV.frx":0442
      DataSource      =   "AdoInv"
      Height          =   2325
      Left            =   105
      TabIndex        =   13
      Top             =   1470
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   4101
      _Version        =   393216
      Style           =   1
      BackColor       =   8454143
      ForeColor       =   8388608
      Text            =   "Productos"
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
   Begin VB.TextBox TxtPVP 
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
      Left            =   9870
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "KardIngV.frx":0457
      Top             =   2730
      Width           =   1170
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   3465
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   46
      Top             =   6825
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo DCCtaObra 
      Bindings        =   "KardIngV.frx":045B
      DataSource      =   "AdoCtaObra"
      Height          =   315
      Left            =   4725
      TabIndex        =   0
      Top             =   105
      Width           =   8205
      _ExtentX        =   14473
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
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "KardIngV.frx":0474
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   4725
      TabIndex        =   6
      Top             =   420
      Width           =   8205
      _ExtentX        =   14473
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
      Left            =   11130
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "KardIngV.frx":048B
      Top             =   3465
      Width           =   1800
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
      MaxLength       =   100
      TabIndex        =   8
      Top             =   735
      Width           =   8205
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
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
   Begin VB.TextBox TxtFactNo 
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
      Left            =   10920
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "KardIngV.frx":048D
      Top             =   1155
      Width           =   2010
   End
   Begin VB.TextBox TxtSubTotal 
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
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "KardIngV.frx":048F
      Top             =   7140
      Width           =   1590
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "KardIngV.frx":0491
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   12
      Top             =   1155
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DC"
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
   Begin VB.TextBox TxtCodBar 
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
      Left            =   6720
      MaxLength       =   25
      TabIndex        =   30
      Text            =   "."
      Top             =   3465
      Width           =   2640
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "KardIngV.frx":04A7
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   5040
      TabIndex        =   15
      Top             =   1995
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DC"
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
      Bindings        =   "KardIngV.frx":04BF
      Height          =   2850
      Left            =   105
      TabIndex        =   35
      Top             =   3885
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   5027
      _Version        =   393216
      BackColor       =   12648447
      BorderStyle     =   0
      ForeColor       =   192
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
      Left            =   9450
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "KardIngV.frx":04D7
      Top             =   3465
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
      Left            =   8610
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "KardIngV.frx":04DB
      Top             =   2730
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoCtaObra 
      Height          =   330
      Left            =   2310
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
      Left            =   4320
      Top             =   5100
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
      Left            =   4305
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
      Left            =   4320
      Top             =   4785
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
      Left            =   7140
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "KardIngV.frx":04DD
      Top             =   2730
      Width           =   1380
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
      Left            =   7350
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "KardIngV.frx":04DF
      Top             =   7140
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
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
      Left            =   11970
      Picture         =   "KardIngV.frx":04E1
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6825
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
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
      Left            =   10920
      Picture         =   "KardIngV.frx":0DAB
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6825
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   315
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
   Begin MSAdodcLib.Adodc AdoDiario 
      Height          =   330
      Left            =   2310
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
   Begin MSMask.MaskEdBox MBVence 
      Height          =   330
      Left            =   1365
      TabIndex        =   4
      Top             =   735
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
   Begin MSDataListLib.DataCombo DCMarca 
      Bindings        =   "KardIngV.frx":11ED
      DataSource      =   "AdoMarca"
      Height          =   315
      Left            =   8715
      TabIndex        =   16
      Top             =   1995
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   "DC"
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   330
      Left            =   4305
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
      Caption         =   "Marca"
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
   Begin VB.Label LabelTotalVenta 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
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
      Left            =   11130
      TabIndex        =   26
      Top             =   2730
      Width           =   1800
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11130
      TabIndex        =   25
      Top             =   2415
      Width           =   1800
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " P.V.P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   9870
      TabIndex        =   23
      Top             =   2415
      Width           =   1170
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA POR COBRAR"
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
      TabIndex        =   45
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
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
      Left            =   2625
      TabIndex        =   7
      Top             =   735
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
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
      Left            =   2625
      TabIndex        =   5
      Top             =   420
      Width           =   2115
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
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
      TabIndex        =   1
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label22 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5040
      TabIndex        =   9
      Top             =   1155
      Width           =   4425
   End
   Begin VB.Label Label21 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vencimiento:"
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
      TabIndex        =   3
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
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
      Height          =   645
      Left            =   105
      TabIndex        =   44
      Top             =   6825
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5040
      TabIndex        =   27
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5040
      TabIndex        =   17
      Top             =   2415
      Width           =   2010
   End
   Begin VB.Label LabelProducto 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5040
      TabIndex        =   14
      Top             =   1575
      Width           =   7890
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080FFFF&
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
      Height          =   330
      Left            =   9555
      TabIndex        =   10
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S U B T O T A L"
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
      TabIndex        =   43
      Top             =   6825
      Width           =   1590
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COD. DE BAR. / LOTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6720
      TabIndex        =   29
      Top             =   3150
      Width           =   2640
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GUIA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   7140
      TabIndex        =   19
      Top             =   2415
      Width           =   1380
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
      Left            =   9030
      TabIndex        =   39
      Top             =   7140
      Width           =   1590
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COSTO TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11130
      TabIndex        =   33
      Top             =   3150
      Width           =   1800
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR COSTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   9450
      TabIndex        =   31
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   8610
      TabIndex        =   21
      Top             =   2415
      Width           =   1170
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
      Left            =   5040
      TabIndex        =   28
      Top             =   3465
      Width           =   1590
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
      Left            =   5040
      TabIndex        =   18
      Top             =   2730
      Width           =   2010
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T O T A L"
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
      Left            =   9030
      TabIndex        =   38
      Top             =   6825
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
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
      TabIndex        =   36
      Top             =   6825
      Width           =   1590
   End
End
Attribute VB_Name = "Kard_Ing_Ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'Grabar Comprobante de Ingreso/Egreso de Inventario
Private Sub Command1_Click()
Dim GrabarVentas As Boolean
  RatonNormal
  Trans_No = 97
  GrabarVentas = False
  DetalleComp = Ninguno
  Ln_No = 1
  Asiento = 1
  If CodigoCliente = "" Then CodigoCliente = Ninguno
  Total_Factura = 0
  Si_No = False
  FechaValida MBFechaI
  FechaValida MBVence
  FechaTexto = MBFechaI
  TextoValido TextOrden, , True
 'TotalInventario
 'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
  Factura_No = Val(TxtFactNo)
  If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
  sSQL = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
' Cta de Inventario
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  With AdoKardex.Recordset
   If .RecordCount > 0 Then
       Mensajes = "Seguro de Grabar?"
       Titulo = "GRABACION DEL COMPROBANTE"
       If BoxMensaje = vbYes Then
          RatonReloj
          Si_No = True
          FechaComp = MBFechaI
          FechaTexto = MBFechaI
          NumComp = ReadSetDataNum("Diario", True, True)
          CodigoInv = .fields("Codigo_Inv")
          Cta_Inventario = .fields("CTA_INVENTARIO")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Cta_Inventario <> .fields("CTA_INVENTARIO") Then
                InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
                CodigoInv = .fields("Codigo_Inv")
                Cta_Inventario = .fields("CTA_INVENTARIO")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .fields("VALOR_TOTAL")
             If .fields("TOTAL_PVP") <> 0 Then GrabarVentas = True
            .MoveNext
          Loop
          InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
       End If
   End If
  End With
  
  
  OpcTM = 1
  OpcDH = 2
  NoCheque = Ninguno
  DetalleComp = Ninguno
' Contra Cuenta del Kardex
  If Si_No Then
     If GrabarVentas Then
     sSQL = "SELECT * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "ORDER BY CTA_COSTO "
     SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
     Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          SubCta = .fields("TC")
          Contra_Cta = .fields("CTA_COSTO")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Contra_Cta <> .fields("CTA_COSTO") Then
                InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
                SubCta = .fields("TC")
                Contra_Cta = .fields("CTA_COSTO")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .fields("VALOR_TOTAL")
            .MoveNext
          Loop
          InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
      End If
     End With
     
     Total = 0
     sSQL = "SELECT * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "ORDER BY CTA_VENTA "
     SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
     Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          SubCta = .fields("TC")
          Contra_Cta = .fields("CTA_VENTA")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Contra_Cta <> .fields("CTA_VENTA") Then
                InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
                SubCta = .fields("TC")
                Contra_Cta = .fields("CTA_VENTA")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .fields("TOTAL_PVP")
            .MoveNext
          Loop
          InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
      End If
     End With
     End If
     
     
     
     
     TotalInventario
   ' Insertamos El IVA de la compra
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
      End If
     End With
     
        Contra_Cta = SinEspaciosDer(DCCtaObra)
        Codigo = Leer_Cta_Catalogo(Contra_Cta)
        ValorDH = Debe - Haber
        If ValorDH < 0 Then ValorDH = -ValorDH
        Select Case SubCta
          Case "C", "P"
               Total_Factura = ValorDH
               SetAdoAddNew "Asiento_SC"
               SetAdoFields "TM", "1"
               SetAdoFields "Factura", Val(TxtFactNo)
               SetAdoFields "Codigo", CodigoCliente
               SetAdoFields "FECHA_V", MBVence
               SetAdoFields "Cta", Contra_Cta
               SetAdoFields "TC", SubCta
               SetAdoFields "T_No", Trans_No
               SetAdoFields "SC_No", Ln_No
               SetAdoFields "DH", "1"
               SetAdoFields "Valor", ValorDH
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoUpdate
        End Select
        InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
     End If
  'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
  Contador = 1
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY DEBE DESC,HABER,CODIGO "
  Select_Adodc AdoAsientos, sSQL
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       Debe = 0: Haber = 0
       Do While Not .EOF
          Debe = Debe + .fields("DEBE")
          Haber = Haber + .fields("HABER")
         .fields("A_No") = Contador
         .Update
         .MoveNext
          Contador = Contador + 1
       Loop
       If (Debe - Haber) <> 0 Then MsgBox "Verifique el comprobante, no cuadra por: " & Redondear(Debe - Haber, 2)
      .MoveFirst
     ' MsgBox (Debe - Haber)
          RatonReloj
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto
          Co.CodigoB = CodigoCliente
          Co.Efectivo = 0
          Co.Monto_Total = Total - Total_RetCta
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
          Co.Item = NumEmpresa
          Factura_No = Val(TxtFactNo)
          If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
          TxtFactNo = Factura_No
          If Len(TextOrden) > 1 Then Co.Concepto = Co.Concepto & ", Orden No. " & TextOrden
          If Val(TxtFactNo) > 0 Then Co.Concepto = Co.Concepto & ", Factura No. " & TxtFactNo
          
          Grabar_Comprobante Co
          'MsgBox "Hola"
          ImprimirComprobantesDe False, Co
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          RatonReloj
          Mayorizar_Inventario_SP
          IniciarAsientosAdo AdoAsientos
          RatonNormal
          DCCtaObra.SetFocus
   Else
       MsgBox "No existen Datos para procesar"
   End If
  End With
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Kard_Ing_Ventas
End Sub

Private Sub DCBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBenef_LostFocus()
  InvImp = False
  Label2.Caption = " LOCAL"
  CodigoCliente = Ninguno
  NombreCliente = DCBenef
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & NombreCliente & "' ")
       If Not .EOF Then
          CICliente = .fields("CI_RUC")
          CodigoBenef = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          TipoBenef = .fields("TD")
          Cod_Benef = .fields("TipoBenef")
          InvImp = .fields("Importaciones")
          If InvImp Then Label2.Caption = " IMPORTACION"
          Si_No = True
          If TipoDoc = "R" Then Si_No = False
          'MsgBox "...."
       Else
          Si_No = False
       End If
   End If
  End With
  Label3.Caption = CodigoCliente
  TextConcepto = DCBenef & " "
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaObra_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaObra_LostFocus()
' SubCta = .Fields("TC")
  Contra_Cta = SinEspaciosDer(DCCtaObra)
  Codigo = Leer_Cta_Catalogo(Contra_Cta)
  Contra_Cta = Codigo
  ListarProveedorUsuario SubCta
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: Command1.SetFocus
    Case vbKeyReturn: SiguienteControl
  End Select
End Sub

Private Sub DCInv_LostFocus()
  Codigos = Ninguno
  TxtPVP = "0.00"
  If Leer_Codigo_Inv(DCInv, FechaSistema) Then
     Si_No = DatInv.IVA
     Unidad = DatInv.Unidad
     CodigoInv = DatInv.Codigo_Inv
     Producto = DatInv.Producto
     Cta_Inventario = DatInv.Cta_Inventario
     Contra_Cta1 = DatInv.Cta_Costo_Venta
     Cta_Ventas = DatInv.Cta_Ventas
     
     'Precio = DatInv.PVP
     Precio = DatInv.Valor_Unit
     TxtPVP = Format(Precio, "#,##0.000000")
     TextVUnit = Format(DatInv.Valor_Unit, "#,##0." & String(Dec_Costo, "0"))
  Else
     MsgBox "No existe Productos asignados"
     DCTInv.SetFocus
  End If
  LabelCodigo.Caption = CodigoInv
  LabelUnidad.Caption = Unidad
  LabelProducto.Caption = Producto
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
  ListarProductos
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyK Then
     DCDiario.Visible = True
     DCDiario.SetFocus
  End If
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  FechaComp = MBFechaI
  Fecha_Vence = MBFechaI
End Sub

Private Sub MBVence_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBVence_LostFocus()
  FechaValida MBVence
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextEntrada_GotFocus()
  MarcarTexto TextEntrada
  OpcDH = 2
    If Precio = 0 Then
       MsgBox "Warning: " & vbCrLf _
            & "        Falta de Ingresar en este codigo" & vbCrLf _
            & "        " & Codigo & ": La entrada inicial"
    End If
  TextVUnit.Text = Format(ValorUnit, "#,##0.00000")
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True, , Dec_Cant
  Entrada = Val(TextEntrada)
  ValorTotal = Redondear(ValorUnit * Entrada, 2)
  TextTotal.Text = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub Form_Activate()
'Consultamos las cuentas de la tabla
 Trans_No = 97
 Total_IVA = 0
 IniciarAsientosAdo AdoAsientos
 If Inv_Promedio Then
    Kard_Ing_Ventas.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    Kard_Ing_Ventas.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
 End If
 TipoDoc = CompDiario
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
 ListarProductos
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
      & "AND TC IN ('C','P','RP') " _
      & "AND DG = 'D' " _
      & "ORDER BY Codigo "
 SelectDB_Combo DCCtaObra, AdoCtaObra, sSQL, "Nomb_Cta"
 
 Contra_Cta = SinEspaciosDer(DCCtaObra)
 Codigo = Leer_Cta_Catalogo(Contra_Cta)
 ListarProveedorUsuario SubCta
  
 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Bodega "
 SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
 
 sSQL = "SELECT * " _
      & "FROM Catalogo_Marcas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Marca "
 SelectDB_Combo DCMarca, AdoMarca, sSQL, "Marca"
 
 FechaValida MBFechaI
 FechaComp = MBFechaI
 RatonNormal
 
End Sub

Private Sub Form_Load()
  CodigoBenef = Ninguno
  CodigoCliente = Ninguno
  ConectarAdodc AdoInv
  ConectarAdodc AdoRet
  ConectarAdodc AdoIVA
  ConectarAdodc AdoArt
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoMarca
  ConectarAdodc AdoDiario
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoAsientos
End Sub

Private Sub TextIVA_GotFocus()
  TextIVA.Text = ""
  TotalInventario
End Sub

Private Sub TextIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextIVA_LostFocus()
  TextoValido TextIVA, True
  TotalInventario
End Sub

Private Sub TextOrden_GotFocus()
  SaldoAnterior = 0: ValorUnitAnt = 0: Contador = 0: ValorUnit = 0: Cantidad = 0
  Stock_Actual_Inventario MBFechaI, CodigoInv
  Precio = ValorUnit
  TextVUnit = Format(Precio, "#,##0." & String(Dec_Costo, "0"))
  MarcarTexto TextOrden
End Sub

Private Sub TextOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextOrden_LostFocus()
  TextoValido TextOrden, , True
End Sub

Private Sub TextTotal_GotFocus()
Dim DValorUnit As Double
Dim DValorTotal As Double
   TextoValido TextOrden
   Total_Desc = 0
   Cod_Bodega = Ninguno
   Cod_Marca = Ninguno
   With AdoBodega.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Bodega Like '" & DCBodega & "' ")
        If Not .EOF Then Cod_Bodega = .fields("CodBod")
    End If
   End With
   With AdoMarca.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Marca Like '" & DCMarca & "' ")
        If Not .EOF Then Cod_Marca = .fields("CodMar")
    End If
   End With
   Entrada = Val(CCur(TextEntrada))
   ValorUnit = Val(CDbl(TextVUnit))
   DValorUnit = ValorUnit
   DValorTotal = Redondear(DValorUnit * Entrada, 2)
   ValorUnit = Val(CDbl(DValorUnit))
   ValorTotal = Redondear(Val(CCur(DValorTotal)), 2)
   TextTotal = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub TextTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTotal_LostFocus()
   FechaTexto = MBFechaI
   Entrada = Val(CCur(TextEntrada))
   Factura_No = Val(TxtFactNo)
   If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
   TxtFactNo = Factura_No
 ' Llenamos el ultimo saldo del kardex
   TextVUnit = Format(ValorUnit, "#,##0." & String(Dec_Costo, "0"))
   SaldoAnterior = Redondear(SaldoAnterior, 4)
   SubTotal_IVA = 0
   
   Cantidad = Cantidad - Entrada
   Saldo = SaldoAnterior - ValorTotal
   OpcDH = "2"
   Contra_Cta = SinEspaciosDer(DCCtaObra)
   Codigo = Leer_Cta_Catalogo(Contra_Cta)
   If Entrada > 0 And ValorUnit > 0 Then
      SetAddNew AdoKardex
      SetFields AdoKardex, "DH", OpcDH
      SetFields AdoKardex, "CODIGO_INV", CodigoInv
      SetFields AdoKardex, "PRODUCTO", Producto
      SetFields AdoKardex, "CANT_ES", Entrada
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "TOTAL_PVP", CCur(TxtPVP)
      SetFields AdoKardex, "VALOR_TOTAL", ValorTotal
      SetFields AdoKardex, "IVA", SubTotal_IVA
      SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inventario
      SetFields AdoKardex, "CONTRA_CTA", Contra_Cta
      SetFields AdoKardex, "CTA_COSTO", Contra_Cta1
      SetFields AdoKardex, "CTA_VENTA", Cta_Ventas
      SetFields AdoKardex, "CANTIDAD", Cantidad
      SetFields AdoKardex, "TOTAL_PVP", CCur(LabelTotalVenta.Caption)
      SetFields AdoKardex, "SALDO", Saldo
      SetFields AdoKardex, "UNIDAD", LabelUnidad.Caption
      SetFields AdoKardex, "CodBod", Cod_Bodega
      SetFields AdoKardex, "CodMar", Cod_Marca
      SetFields AdoKardex, "Item", NumEmpresa
      SetFields AdoKardex, "CodigoU", CodigoUsuario
      SetFields AdoKardex, "T_No", Trans_No
      SetFields AdoKardex, "SUBCTA", SubCtaGen
      SetFields AdoKardex, "TC", SubCta
      SetFields AdoKardex, "Codigo_B", CodigoCliente
      SetFields AdoKardex, "COD_BAR", TxtCodBar
      SetFields AdoKardex, "ORDEN", TextOrden
      SetUpdate AdoKardex
      TextTotal = Format(ValorTotal, "#,##0.00")
      TotalInventario
      DCInv.SetFocus
   End If
End Sub

Private Sub TextVUnit_GotFocus()
   MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True, , Dec_Costo
  ValorUnit = Val(CDbl(TextVUnit))
  
  ValorTotal = Redondear(ValorUnit * Entrada, 2)
  TextTotal = Format(ValorTotal, "#,##0.0000")
  TextVUnit = FormatDec(ValorUnit, Dec_Costo)
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Total = Total + .fields("TOTAL_PVP")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Redondear(Total, 2)
  Total_IVA = Redondear(Total_IVA, 2)
  TxtSubTotal.Text = Format(Total, "#,##0.00")
  TextIVA.Text = Format(Total_IVA, "#,##0.00")
  Label1.Caption = Format(Total + Total_IVA, "#,##0.00")
End Sub

Public Sub ListarProductos()
 CodigoInv = SinEspaciosIzq(DCTInv.Text)
 sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' " _
      & "AND LEN(Cta_Inventario) > 1 " _
      & "AND TC = 'P' " _
      & "ORDER BY Producto "
 SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
End Sub

Public Sub ListarProveedorUsuario(TipoSubCta As String)
  Select Case TipoSubCta
    Case "P", "C"
         'OpcIVA.Visible = True
         'OpcX.Visible = True
         sSQL = "SELECT C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo,'P' As TipoBenef " _
              & "FROM Clientes As C,Catalogo_CxCxP As CP " _
              & "WHERE CP.TC = '" & TipoSubCta & "' " _
              & "AND CP.Item = '" & NumEmpresa & "' " _
              & "AND CP.Periodo = '" & Periodo_Contable & "' " _
              & "AND CP.Cta = '" & Contra_Cta & "' " _
              & "AND C.Codigo = CP.Codigo " _
              & "GROUP BY C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo " _
              & "ORDER BY C.Cliente "
    Case Else
         sSQL = "SELECT Cliente,Codigo,CI_RUC,Grupo,Direccion,Telefono,TD,'.' As Cta,0 As Importaciones,'X' As TipoBenef " _
              & "FROM Clientes " _
              & "WHERE FA = " & Val(adFalse) & " " _
              & "ORDER BY Cliente "
  End Select
  SelectDB_Combo DCBenef, AdoBenef, sSQL, "Cliente"
  If AdoBenef.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Datos Asignados a Esta Cuenta"
     MBFechaI.SetFocus
  End If
End Sub

Private Sub TxtCodBar_GotFocus()
  MarcarTexto TxtCodBar
End Sub

Private Sub TxtCodBar_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodBar_LostFocus()
  TextoValido TxtCodBar, , True
End Sub

Private Sub TxtFactNo_GotFocus()
  MarcarTexto TxtFactNo
End Sub

Private Sub TxtFactNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFactNo_LostFocus()
  TextoValido TxtFactNo, , True
  Factura_No = Val(TxtFactNo)
  FechaTexto = MBFechaI
  If Factura_No < 0 Then Factura_No = 0
  TxtFactNo = Factura_No
End Sub

Private Sub TxtPVP_GotFocus()
  MarcarTexto TxtPVP
End Sub

Private Sub TxtPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPVP_LostFocus()
  TextoValido TxtPVP, True, , Dec_Costo
  LabelTotalVenta.Caption = Format(Val(TxtPVP) * Val(TextEntrada), "#,##0.00")
End Sub

