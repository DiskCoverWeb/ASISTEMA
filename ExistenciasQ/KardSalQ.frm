VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FSalidas_Ceros 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CONTROL DE INVENTARIO PARA EGRESOS DE QUIROFANO"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13050
   Icon            =   "KardSalQ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13050
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtFactura 
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
      Left            =   11865
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "KardSalQ.frx":0442
      Top             =   735
      Width           =   1065
   End
   Begin VB.TextBox TxtGestion 
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
      Left            =   9660
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "KardSalQ.frx":044E
      Top             =   4095
      Width           =   1380
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
      Left            =   1365
      MaxLength       =   100
      TabIndex        =   7
      Top             =   735
      Width           =   9150
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "KardSalQ.frx":0452
      DataSource      =   "AdoInv"
      Height          =   2910
      Left            =   105
      TabIndex        =   9
      Top             =   1470
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5133
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
      Left            =   8190
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "KardSalQ.frx":0467
      Top             =   4095
      Width           =   1380
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   3465
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   41
      Top             =   7770
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo DCCtaObra 
      Bindings        =   "KardSalQ.frx":046B
      DataSource      =   "AdoCtaObra"
      Height          =   315
      Left            =   4725
      TabIndex        =   1
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
      Bindings        =   "KardSalQ.frx":0484
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   4725
      TabIndex        =   5
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
      TabIndex        =   32
      Text            =   "KardSalQ.frx":049B
      Top             =   4095
      Width           =   1800
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1365
      TabIndex        =   3
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
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "KardSalQ.frx":049D
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   1155
      Width           =   6735
      _ExtentX        =   11880
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
      Left            =   9030
      MaxLength       =   25
      TabIndex        =   20
      Text            =   "."
      Top             =   2940
      Width           =   3900
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "KardSalQ.frx":04B3
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   8085
      TabIndex        =   12
      Top             =   1575
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "Bodegas"
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
      Bindings        =   "KardSalQ.frx":04CB
      Height          =   3060
      Left            =   105
      TabIndex        =   33
      Top             =   4515
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   5398
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
      Left            =   6930
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "KardSalQ.frx":04E3
      Top             =   4095
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoCtaObra 
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
      Left            =   6300
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
   Begin MSAdodcLib.Adodc AdoInv 
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
   Begin MSAdodcLib.Adodc AdoAsientos 
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
      Picture         =   "KardSalQ.frx":04E5
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7665
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
      Picture         =   "KardSalQ.frx":0DAF
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7665
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTInv 
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
      Left            =   4305
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
   Begin MSDataListLib.DataCombo DCMarca 
      Bindings        =   "KardSalQ.frx":11F1
      DataSource      =   "AdoMarca"
      Height          =   315
      Left            =   8085
      TabIndex        =   14
      Top             =   1890
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "Marcas"
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
      Left            =   6300
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
   Begin MSDataListLib.DataCombo DCDr 
      Bindings        =   "KardSalQ.frx":1208
      DataSource      =   "AdoDr"
      Height          =   315
      Left            =   8085
      TabIndex        =   16
      Top             =   2205
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "Dr"
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
   Begin MSAdodcLib.Adodc AdoDr 
      Height          =   330
      Left            =   8295
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
      Caption         =   "Dr"
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
   Begin MSDataListLib.DataCombo DCTratamiento 
      Bindings        =   "KardSalQ.frx":121C
      DataSource      =   "AdoTratamiento"
      Height          =   315
      Left            =   8085
      TabIndex        =   18
      Top             =   2520
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "Tratamiento"
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
   Begin MSAdodcLib.Adodc AdoTratamiento 
      Height          =   330
      Left            =   10395
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
      Caption         =   "Tratamiento"
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
      Left            =   11130
      TabIndex        =   24
      Top             =   3360
      Width           =   1800
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
      Left            =   10185
      TabIndex        =   23
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TRATAMI."
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
      Left            =   6930
      TabIndex        =   17
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FACTURA No."
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
      Left            =   10500
      TabIndex        =   40
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DOCTOR"
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
      Left            =   6930
      TabIndex        =   15
      Top             =   2205
      Width           =   1170
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GEST. 10%"
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
      Left            =   9660
      TabIndex        =   29
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MARCA"
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
      Left            =   6930
      TabIndex        =   13
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA"
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
      Left            =   6930
      TabIndex        =   11
      Top             =   1575
      Width           =   1170
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
      Left            =   8085
      TabIndex        =   22
      Top             =   3360
      Width           =   2115
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
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   1275
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
      TabIndex        =   4
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
      TabIndex        =   2
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COSTO"
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
      Left            =   8190
      TabIndex        =   27
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA POR COBRAR / PAGAR"
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
      Width           =   4635
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
      TabIndex        =   38
      Top             =   7665
      Width           =   2535
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
      Left            =   6930
      TabIndex        =   21
      Top             =   3360
      Width           =   1170
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
      Left            =   6930
      TabIndex        =   10
      Top             =   1155
      Width           =   6000
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
      Left            =   6930
      TabIndex        =   19
      Top             =   2940
      Width           =   2115
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
      Left            =   9135
      TabIndex        =   35
      Top             =   7980
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
      TabIndex        =   31
      Top             =   3780
      Width           =   1800
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
      Left            =   6930
      TabIndex        =   25
      Top             =   3780
      Width           =   1170
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
      Left            =   9135
      TabIndex        =   34
      Top             =   7665
      Width           =   1590
   End
End
Attribute VB_Name = "FSalidas_Ceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cod_Dr As String
Dim Cod_Tra As String

'Grabar Comprobante de Ingreso/Egreso de Inventario
Private Sub Command1_Click()
  RatonNormal
  Trans_No = 197
  DetalleComp = Ninguno
  Ln_No = 1
  Asiento = 1
  If CodigoCliente = "" Then CodigoCliente = Ninguno
  
  Cod_Dr = Ninguno
  With AdoDr.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCDr & "' ")
       If Not .EOF Then Cod_Dr = .Fields("Codigo")
   End If
  End With
  
  Cod_Tra = Ninguno
  With AdoTratamiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto = '" & DCTratamiento & "' ")
       If Not .EOF Then Cod_Tra = .Fields("Codigo_Inv")
   End If
  End With
    
  Total_Factura = 0
  Si_No = False
  FechaValida MBFechaI
  FechaTexto = MBFechaI
   
 'TotalInventario
 'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
  Factura_No = Val(TxtFactNo)
  If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
  sSQL = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
' Cta de Inventario
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
  SelectDataGrid DGKardex, AdoKardex, sSQL, SQLDec
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
          CodigoInv = .Fields("Codigo_Inv")
          Cta_Inventario = .Fields("CTA_INVENTARIO")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Cta_Inventario <> .Fields("CTA_INVENTARIO") Then
                InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
                CodigoInv = .Fields("Codigo_Inv")
                Cta_Inventario = .Fields("CTA_INVENTARIO")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .Fields("VALOR_TOTAL")
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
     sSQL = "SELECT * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "ORDER BY CTA_COSTO "
     SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
     SelectDataGrid DGKardex, AdoKardex, sSQL, SQLDec
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          SubCta = .Fields("TC")
          Contra_Cta = .Fields("CTA_COSTO")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Contra_Cta <> .Fields("CTA_COSTO") Then
                InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
                SubCta = .Fields("TC")
                Contra_Cta = .Fields("CTA_COSTO")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .Fields("VALOR_TOTAL")
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
     SelectDataGrid DGKardex, AdoKardex, sSQL, SQLDec
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          SubCta = .Fields("TC")
          Contra_Cta = .Fields("CTA_VENTA")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Contra_Cta <> .Fields("CTA_VENTA") Then
                InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
                SubCta = .Fields("TC")
                Contra_Cta = .Fields("CTA_VENTA")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .Fields("TOTAL_PVP")
            .MoveNext
          Loop
          InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
      End If
     End With
     
     TotalInventario
   ' Insertamos El IVA de la compra
     sSQL = "SELECT * " _
          & "FROM Asiento " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     SelectAdodc AdoAsientos, sSQL
     With AdoAsientos.Recordset
      If .RecordCount > 0 Then
          Debe = 0: Haber = 0
          Do While Not .EOF
             Debe = Debe + .Fields("DEBE")
             Haber = Haber + .Fields("HABER")
            .MoveNext
          Loop
      End If
     End With
     
        Contra_Cta = SinEspaciosDer(DCCtaObra)
        LeerCta Contra_Cta
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
  SelectAdodc AdoAsientos, sSQL
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       Debe = 0: Haber = 0
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .Fields("A_No") = Contador
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
          Co.CodigoDr = Cod_Dr
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
          
          GrabarComprobante Co
         'MsgBox TextOrden.Text
          ImprimirComprobantesDe False, Co
          Cadena = Generar_Salidas_Excel(Co)
          Abrir_Excel Cadena
          MsgBox "SE GENERO EL ARCHIVO EN: " & vbCrLf & Cadena
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, Format(NumComp, "000000000"), "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, Format(NumComp, "000000000"), "CD", FechaTexto, FechaTexto, Total
          MayorizarInv.Show
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
  Unload FSalidas_Ceros
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
          CodigoBenef = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
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
  Contra_Cta = SinEspaciosDer(DCCtaObra)
  LeerCta Contra_Cta
  Contra_Cta = Codigo
  ListarProveedorUsuario SubCta
End Sub

Private Sub DCDr_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCDr_LostFocus()
  Cod_Dr = Ninguno
  With AdoDr.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCDr & "' ")
       If Not .EOF Then Cod_Dr = .Fields("Codigo")
   End If
  End With
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
     Precio = DatInv.Costo
     ValorUnit = Precio
     TxtCodBar = DatInv.Codigo_Barra
     TxtPVP = Format(DatInv.Costo, "#,##0.00")
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

Private Sub DCTratamiento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTratamiento_LostFocus()
  Cod_Tra = Ninguno
  With AdoTratamiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto = '" & DCTratamiento & "' ")
       If Not .EOF Then Cod_Tra = .Fields("Codigo_Inv")
   End If
  End With
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
  Validar_Porc_IVA MBFechaI
  FechaComp = MBFechaI
  Fecha_Vence = MBFechaI
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
  TextEntrada = DatInv.Stock
  MarcarTexto TextEntrada
  OpcDH = 2
  If DatInv.Costo = 0 Then
     MsgBox "Warning: " & vbCrLf _
          & "        Falta de Ingresar en este codigo" & vbCrLf _
          & "        " & Codigo & ": La entrada inicial"
  End If
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True, , Dec_Cant
  Entrada = Val(TextEntrada)
  ValorTotal = Redondear(ValorUnit * Entrada, 2)
  TextTotal = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub Form_Activate()
'Consultamos las cuentas de la tabla
 Trans_No = 197
 Total_IVA = 0
 IniciarAsientosAdo AdoAsientos
 If Inv_Promedio Then
    FSalidas_Ceros.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    FSalidas_Ceros.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
 End If
 TipoDoc = CompDiario
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDBCombo DCTInv, AdoTInv, sSQL, "NomProd"
 ListarProductos
 sSQL = "SELECT Codigo_Inv, Producto " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'P' " _
      & "AND MID(Codigo_Inv,1,2) = 'QO' " _
      & "ORDER BY Codigo_Inv "
 SelectDBCombo DCTratamiento, AdoTratamiento, sSQL, "Producto"
 
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
 SelectDataGrid DGKardex, AdoKardex, sSQL, SQLDec
  
 sSQL = "SELECT Cuenta & '  ->  ' & Codigo As Nomb_Cta " _
      & "FROM Catalogo_Cuentas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Codigo = '7.2' " _
      & "AND DG = 'D' " _
      & "ORDER BY Codigo "
 SelectDBCombo DCCtaObra, AdoCtaObra, sSQL, "Nomb_Cta"
 
 Contra_Cta = SinEspaciosDer(DCCtaObra)
 LeerCta Contra_Cta
 ListarProveedorUsuario SubCta
  
 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Bodega "
 SelectDBCombo DCBodega, AdoBodega, sSQL, "Bodega"
 With AdoBodega.Recordset
   If Cod_Bodega <> Ninguno And .RecordCount > 0 Then
     .MoveFirst
     .Find ("CodBod = '" & Cod_Bodega & "' ")
      If Not .EOF Then DCBodega = .Fields("Bodega")
   End If
 End With
 
 sSQL = "SELECT Marca " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'P' " _
      & "AND Marca <> '" & Ninguno & "' " _
      & "GROUP BY Marca " _
      & "ORDER BY Marca "
  SelectDBCombo DCMarca, AdoMarca, sSQL, "Marca"
 
 FechaValida MBFechaI
 FechaComp = MBFechaI
 RatonNormal
 
End Sub

Private Sub Form_Load()
  CodigoBenef = Ninguno
  CodigoCliente = Ninguno
  ConectarAdodc AdoDr
  ConectarAdodc AdoInv
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoMarca
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoTratamiento
End Sub

Private Sub TextTotal_GotFocus()
Dim DValorUnit As Double
Dim DValorTotal As Double
   Total_Desc = 0
   Cod_Marca = Ninguno
   With AdoBodega.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Bodega Like '" & DCBodega & "' ")
        If Not .EOF Then Cod_Bodega = .Fields("CodBod")
    End If
   End With
''   With AdoMarca.Recordset
''    If .RecordCount > 0 Then
''       .MoveFirst
''       .Find ("Marca Like '" & DCMarca & "' ")
''        If Not .EOF Then Cod_Marca = .Fields("CodMar")
''    End If
''   End With
   Entrada = Val(CCur(TextEntrada))
   ValorUnit = Val(CDbl(TxtPVP))
   DValorUnit = ValorUnit
   DValorTotal = DValorUnit * Entrada
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
   LeerCta Contra_Cta
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
      SetFields AdoKardex, "TOTAL_PVP", ValorTotal
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
      SetFields AdoKardex, "ORDEN", TxtFactura 'TextOrden
      SetFields AdoKardex, "No_Refrendo", DCMarca
      SetFields AdoKardex, "Codigo_Dr", Cod_Dr
      SetFields AdoKardex, "Codigo_Tra", Cod_Tra
      SetFields AdoKardex, "P_DESC", TxtGestion
      SetUpdate AdoKardex
      TextTotal = Format(ValorTotal, "#,##0.00")
      TotalInventario
      DCInv.SetFocus
   End If
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Total = Total + .Fields("TOTAL_PVP")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Redondear(Total, 2)
  Total_IVA = Redondear(Total_IVA, 2)
  Label1.Caption = Format(Total + Total_IVA, "#,##0.00")
End Sub

Public Sub ListarProductos()
 CodigoInv = SinEspaciosIzq(DCTInv.Text)
 sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Mid$(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' " _
      & "AND LEN(Cta_Inventario) > 1 " _
      & "AND TC = 'P' " _
      & "ORDER BY Producto "
 SelectDBCombo DCInv, AdoInv, sSQL, "Producto"
End Sub

Public Sub ListarProveedorUsuario(TipoSubCta As String)

   sSQL = "SELECT Cliente,Codigo,CI_RUC " _
        & "FROM Clientes " _
        & "WHERE Asignar_Dr <> " & Val(adFalse) & " " _
        & "ORDER BY Cliente "
   SelectDBCombo DCDr, AdoDr, sSQL, "Cliente"
   
   sSQL = "SELECT C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo,'P' As TipoBenef " _
        & "FROM Clientes As C, Catalogo_CxCxP As CP " _
        & "WHERE CP.TC = '" & TipoSubCta & "' " _
        & "AND CP.Item = '" & NumEmpresa & "' " _
        & "AND CP.Periodo = '" & Periodo_Contable & "' " _
        & "AND CP.Cta = '" & Contra_Cta & "' " _
        & "AND C.Codigo = CP.Codigo " _
        & "GROUP BY C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo " _
        & "ORDER BY C.Cliente "
  SelectDBCombo DCBenef, AdoBenef, sSQL, "Cliente"
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

Private Sub TxtFactura_GotFocus()
  MarcarTexto TxtFactura
End Sub

Private Sub TxtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFactura_LostFocus()
   TxtFactura = Format(Val(TxtFactura), "000000000")
End Sub

Private Sub TxtGestion_GotFocus()
  MarcarTexto TxtGestion
End Sub

Private Sub TxtGestion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtGestion_LostFocus()
  TextoValido TxtGestion, True, , 2
End Sub

Private Sub TxtPVP_GotFocus()
  MarcarTexto TxtPVP
End Sub

Private Sub TxtPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPVP_LostFocus()
  TextoValido TxtPVP, True, , 2
  TextTotal = Format(Val(TxtPVP) * Val(TextEntrada), "#,##0.00")
End Sub

