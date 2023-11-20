VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FCyberNet 
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15465
   WindowState     =   1  'Minimized
   Begin VB.TextBox TxtEfectivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   510
      Left            =   13230
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   67
      Text            =   "CyberNet.frx":0000
      Top             =   4935
      Width           =   2115
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   540
      Left            =   13230
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   59
      Text            =   "CyberNet.frx":0007
      Top             =   2835
      Width           =   2115
   End
   Begin VB.Frame FrmCabina 
      BackColor       =   &H00C00000&
      Caption         =   "ABC"
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   6930
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox TxtValorCabina 
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
         Left            =   945
         MaxLength       =   11
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   525
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MBTiempo 
         Height          =   330
         Left            =   105
         TabIndex        =   38
         Top             =   525
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "0"
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "VALOR $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   945
         TabIndex        =   39
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         Caption         =   "Tiempo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   105
         TabIndex        =   37
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame FrmBenef 
      BackColor       =   &H00800000&
      Caption         =   "BUSCAR CLIENTE"
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
      Height          =   3060
      Left            =   1365
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   6735
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "CyberNet.frx":000E
         DataSource      =   "AdoBenef"
         Height          =   2715
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   4789
         _Version        =   393216
         Style           =   1
         Text            =   "Beneficiario"
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
   End
   Begin VB.Frame FrmGrupo 
      BackColor       =   &H00400000&
      Caption         =   "GRUPO DE FACTURACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   1470
      TabIndex        =   13
      Top             =   735
      Visible         =   0   'False
      Width           =   5895
      Begin MSDataListLib.DataList DLGrupo 
         Bindings        =   "CyberNet.frx":0025
         DataSource      =   "AdoGrupo"
         Height          =   2940
         Left            =   105
         TabIndex        =   14
         Top             =   315
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   5186
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox TextCant 
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   7665
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   47
      Text            =   "CyberNet.frx":003C
      Top             =   6405
      Width           =   1065
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   8715
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "CyberNet.frx":003E
      Top             =   6405
      Width           =   1485
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "CyberNet.frx":0043
      DataSource      =   "AdoBodega"
      Height          =   420
      Left            =   1470
      TabIndex        =   12
      Top             =   735
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   741
      _Version        =   393216
      BackColor       =   192
      ForeColor       =   16777215
      Text            =   "DC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextFacturaNo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   13230
      TabIndex        =   53
      Text            =   "0"
      Top             =   1260
      Width           =   2115
   End
   Begin VB.TextBox TextCotiza 
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
      Left            =   8085
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   420
      Width           =   1380
   End
   Begin VB.TextBox TextBenef 
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
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "CONSUMIDOR FINAL"
      ToolTipText     =   "Escriba el Nombre/C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   420
      Width           =   6735
   End
   Begin VB.OptionButton OpcMult 
      Caption         =   "(x)"
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
      Left            =   10290
      TabIndex        =   10
      Top             =   420
      Width           =   645
   End
   Begin VB.OptionButton OpcDiv 
      Caption         =   "(/)"
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
      TabIndex        =   9
      Top             =   420
      Value           =   -1  'True
      Width           =   645
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "CyberNet.frx":005B
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   45
      Top             =   6405
      Width           =   7470
      _ExtentX        =   13176
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
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "CyberNet.frx":0075
      Height          =   1695
      Left            =   105
      TabIndex        =   72
      Top             =   6825
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
      Height          =   330
      Left            =   420
      Top             =   7665
      Visible         =   0   'False
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
      Caption         =   "AsientoF"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   420
      Top             =   7980
      Visible         =   0   'False
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   4620
      Top             =   7980
      Visible         =   0   'False
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
      Caption         =   "Factura"
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
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   420
      Top             =   7665
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Articulo"
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   4620
      Top             =   7665
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   2520
      Top             =   7980
      Visible         =   0   'False
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
      Caption         =   "Grupo"
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
      Left            =   2520
      Top             =   7665
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoCyber 
      Height          =   330
      Left            =   6615
      Top             =   7665
      Visible         =   0   'False
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
      Caption         =   "Cyber"
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
   Begin VB.Label LblCambio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   69
      Top             =   5460
      Width           =   2115
   End
   Begin VB.Label LabelTotalME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   65
      Top             =   4410
      Width           =   2115
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   63
      Top             =   3885
      Width           =   2115
   End
   Begin VB.Label LabelIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   61
      Top             =   3360
      Width           =   2115
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAMBIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9870
      TabIndex        =   68
      Top             =   5460
      Width           =   3375
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &EFECTIVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   9870
      TabIndex        =   66
      Top             =   4935
      Width           =   3375
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Fact. (ME)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9870
      TabIndex        =   64
      Top             =   4410
      Width           =   3375
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9870
      TabIndex        =   62
      Top             =   3885
      Width           =   3375
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A. 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   9870
      TabIndex        =   60
      Top             =   3360
      Width           =   3375
   End
   Begin MSForms.CommandButton CButtonSalir 
      Height          =   645
      Left            =   13755
      TabIndex        =   71
      Top             =   6090
      Width           =   1590
      Caption         =   "Salir"
      PicturePosition =   327683
      Size            =   "2805;1138"
      Picture         =   "CyberNet.frx":008F
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonGrabar 
      Height          =   645
      Left            =   12075
      TabIndex        =   70
      Top             =   6090
      Width           =   1590
      Caption         =   "Grabar"
      PicturePosition =   327683
      Size            =   "2805;1138"
      Picture         =   "CyberNet.frx":0969
      Accelerator     =   71
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA:"
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
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total &Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   9870
      TabIndex        =   58
      Top             =   2835
      Width           =   3375
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   7
      Left            =   105
      TabIndex        =   33
      Top             =   5355
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   6
      Left            =   105
      TabIndex        =   31
      Top             =   4830
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   5
      Left            =   105
      TabIndex        =   29
      Top             =   4305
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      Picture         =   "CyberNet.frx":0C83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   4
      Left            =   105
      TabIndex        =   27
      Top             =   3780
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   3
      Left            =   105
      TabIndex        =   25
      Top             =   3255
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   2
      Left            =   105
      TabIndex        =   23
      Top             =   2730
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   1
      Left            =   105
      TabIndex        =   21
      Top             =   2205
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ToggleButton TButtonPCs 
      Height          =   540
      Index           =   0
      Left            =   105
      TabIndex        =   19
      Top             =   1680
      Width           =   1800
      BackColor       =   16761024
      ForeColor       =   16711680
      DisplayStyle    =   6
      MousePointer    =   1
      Size            =   "3175;952"
      Value           =   "0"
      Caption         =   "Abc"
      PicturePosition =   327683
      Picture         =   "CyberNet.frx":1329
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   7
      Left            =   1890
      TabIndex        =   32
      Top             =   5355
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   6
      Left            =   1890
      TabIndex        =   30
      Top             =   4830
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   5
      Left            =   1890
      TabIndex        =   28
      Top             =   4305
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   4
      Left            =   1890
      TabIndex        =   26
      Top             =   3780
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   3
      Left            =   1890
      TabIndex        =   24
      Top             =   3255
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   2
      Left            =   1890
      TabIndex        =   22
      Top             =   2730
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   1
      Left            =   1890
      TabIndex        =   20
      Top             =   2205
      Width           =   4950
      Size            =   "8731;952"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CButtonCabinas 
      Height          =   750
      Index           =   3
      Left            =   6930
      TabIndex        =   43
      Top             =   4200
      Width           =   2850
      Size            =   "5027;1323"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CButtonCabinas 
      Height          =   750
      Index           =   2
      Left            =   6930
      TabIndex        =   42
      Top             =   3360
      Width           =   2850
      Size            =   "5027;1323"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CButtonCabinas 
      Height          =   750
      Index           =   1
      Left            =   6930
      TabIndex        =   41
      Top             =   2520
      Width           =   2850
      Size            =   "5027;1323"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CButtonPCs 
      Height          =   540
      Index           =   0
      Left            =   1890
      TabIndex        =   18
      Top             =   1680
      Width           =   4950
      ForeColor       =   255
      BackColor       =   16761024
      Caption         =   "Abc"
      PicturePosition =   327683
      Size            =   "8731;952"
      Picture         =   "CyberNet.frx":19CF
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMPUTADORAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   17
      Top             =   1260
      Width           =   6735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CABINAS TELEFONICAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6930
      TabIndex        =   34
      Top             =   1260
      Width           =   2850
   End
   Begin MSForms.CommandButton CButtonCabinas 
      Height          =   750
      Index           =   0
      Left            =   6930
      TabIndex        =   35
      Top             =   1680
      Width           =   2850
      ForeColor       =   4194304
      BackColor       =   12648447
      Caption         =   "A"
      PicturePosition =   327683
      Size            =   "5027;1323"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label LabelVTotal 
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   10185
      TabIndex        =   51
      Top             =   6405
      Width           =   1800
   End
   Begin VB.Label LabelStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   8820
      TabIndex        =   16
      Top             =   735
      Width           =   2115
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock Bodega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6825
      TabIndex        =   15
      Top             =   735
      Width           =   2010
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL CLIENTE:"
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
      Left            =   1365
      TabIndex        =   2
      Top             =   105
      Width           =   6735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONVERSION"
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
      Left            =   9450
      TabIndex        =   8
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COTIZACION"
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
      Left            =   8085
      TabIndex        =   6
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Left            =   10185
      TabIndex        =   50
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
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
      Left            =   7665
      TabIndex        =   46
      Top             =   6090
      Width           =   1065
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio Unitario"
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
      Left            =   8715
      TabIndex        =   48
      Top             =   6090
      Width           =   1485
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &PRODUCTO"
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
      TabIndex        =   44
      Top             =   6090
      Width           =   7470
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   11
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LINEA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9870
      TabIndex        =   52
      Top             =   1260
      Width           =   3375
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   57
      Top             =   2310
      Width           =   2115
   End
   Begin VB.Label LabelSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   13230
      TabIndex        =   55
      Top             =   1785
      Width           =   2115
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9870
      TabIndex        =   56
      Top             =   2310
      Width           =   3375
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9870
      TabIndex        =   54
      Top             =   1785
      Width           =   3375
   End
End
Attribute VB_Name = "FCyberNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CButtonCabinas_Click(Index As Integer)
  FrmCabina.Top = CButtonCabinas(Index).Top
  FrmCabina.Left = CButtonCabinas(Index).Left
  FrmCabina.Caption = "Cabina No. " & Index + 1
  ICabina = Index
  FrmCabina.Visible = True
  MBTiempo.SetFocus
End Sub

Private Sub CButtonCabinas_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then TxtEfectivo.SetFocus
End Sub

Private Sub CButtonGrabar_Click()
  FechaValida MBFecha
  FechaTexto = MBFecha
  CalculosTotalesFactura AdoAsientoF
  If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then
     FA.Cod_CxC = "CXC" & TipoFactura
     Lineas_De_CxC FA
     FA.Factura = Numero_Factura(FA)
     If FA.TC = "PV" Then
        Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
                 & "TICKET No. " & FA.Factura
     ElseIf FA.TC = "NV" Then
        Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
                 & "NOTA DE VENTA No. " & FA.Factura
     Else
        Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
                 & "FACTURA No. " & FA.Factura
     End If
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = vbYes Then
        Moneda_US = False
        TextoFormaPago = PagoCont
        ProcGrabar
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
        Ln_No = 1
        Encerar_Facturas
        FA.Cod_CxC = "CXC" & TipoFactura
        Lineas_De_CxC FA
        FA.Factura = Numero_Factura(FA)
        TextFacturaNo.Text = Format(FA.Factura, "0000000")
        DCArticulo.SetFocus
     End If
  Else
     MsgBox "Error: El Efectivo no alcanza para grabar"
  End If
End Sub

Private Sub CButtonPCs_Click(Index As Integer)
  If TotalPCs(Index) > 0 Then
     With AdoAsientoF.Recordset
      If .RecordCount Then
         .MoveFirst
         .Find ("CODIGO = '99.02." & Format(Index + 1, "000") & "' ")
          If Not .EOF Then
            .Delete
            .Update
          End If
      End If
     End With
     SetAddNew AdoAsientoF
     SetFields AdoAsientoF, "CODIGO", "99.02." & Format(Index + 1, "000")
     SetFields AdoAsientoF, "CODIGO_L", CodigoL
     SetFields AdoAsientoF, "PRODUCTO", "PC No." & Format(Index + 1, "00") & ", T[" & TiempoPCs(Index) & "]"
     SetFields AdoAsientoF, "CANT", 1
     SetFields AdoAsientoF, "PRECIO", TotalPCs(Index)
     SetFields AdoAsientoF, "TOTAL", TotalPCs(Index)
     SetFields AdoAsientoF, "Total_IVA", 0
     SetFields AdoAsientoF, "RUTA", TiempoPCs(Index)
     SetFields AdoAsientoF, "Item", NumEmpresa
     SetFields AdoAsientoF, "CodigoU", CodigoUsuario
     SetFields AdoAsientoF, "A_No", Ln_No
     SetUpdate AdoAsientoF
     Ln_No = Ln_No + 1
  End If
  CalculosTotalesFactura AdoAsientoF
End Sub

Private Sub CButtonPCs_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then TxtEfectivo.SetFocus
End Sub

Private Sub CButtonSalir_Click()
  Cyber_Cabinas = False
  Unload FCyberNet
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCliente.Text & "'")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          TextBenef.Text = .Fields("Cliente")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          FrmBenef.Visible = False
          FA.CodigoC = CodigoCliente
          TextCotiza.SetFocus
       Else
          NombreCliente = DCCliente.Text
          FrmBenef.Visible = False
          FCyberNet.Visible = False
          Nuevo = True
          MsgBox "Cliente No existe"
          FClientesFlash.Show 1
          FCyberNet.Visible = True
          Listar_Clientes
       End If
   Else
       FrmBenef.Visible = False
       FComprobantes.Visible = False
       NombreCliente = DCCliente.Text
       Nuevo = True
       FClientesFlash.Show 1
       FComprobantes.Visible = True
       Listar_Clientes
   End If
  End With
End Sub

Private Sub MBTiempo_GotFocus()
  MarcarTexto MBTiempo
End Sub

Private Sub MBTiempo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TButtonPCs_Click(Index As Integer)
  sSQL = "UPDATE Catalogo_Cyber "
  If TButtonPCs(Index).Value Then
     sSQL = sSQL & "SET Inicio = '" & CDate(Time) & "', PC_Ocupaga = " & Val(adTrue) & " " _
          & "WHERE Inicio = '00:00:00' "
     TButtonPCs(Index).Caption = " PC Ocupado"
     TButtonPCs(Index).Picture = LoadPicture(RutaSistema & "\Iconos\candado1.ico")
  Else
     sSQL = sSQL & "SET Inicio = '00:00:00', PC_Ocupaga = " & Val(adFalse) & " " _
          & "WHERE Inicio <> '00:00:00' "
     TButtonPCs(Index).Caption = " PC Libre"
     TButtonPCs(Index).Picture = LoadPicture(RutaSistema & "\Iconos\candado2.ico")
     CButtonPCs(Index).Caption = " Equipo " & Index + 1 & " - [00:00:00] - USD 0.00"
  End If
  sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '99.02." & Format(Index + 1, "000") & "' "
  'MsgBox sSQL
  ConectarAdoExecute sSQL
End Sub

Private Sub TextBenef_GotFocus()
  MarcarTexto TextBenef
End Sub

Private Sub TextBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyB Then TextBenef.Text = "Ninguno"
End Sub

Private Sub TextBenef_LostFocus()
  TextoValido TextBenef, , True
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextBenef & "' ")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          TextBenef.Text = .Fields("Cliente")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          TextCotiza.SetFocus
       Else
         .MoveFirst
         .Find ("CI_RUC = '" & TextBenef & "' ")
          If Not .EOF Then
             CodigoBenef = .Fields("Codigo")
             TextBenef.Text = .Fields("Cliente")
             CodigoCliente = .Fields("Codigo")
             NombreCliente = .Fields("Cliente")
             Grupo_No = .Fields("Grupo")
             TipoDoc = .Fields("TD")
             TextCotiza.SetFocus
          Else
             DCCliente.Text = TextBenef.Text
             FrmBenef.Visible = True
             DCCliente.SetFocus
          End If
       End If
   Else
       DCCliente.Text = TextBenef.Text
       FrmBenef.Visible = True
       DCCliente.SetFocus
   End If
  End With
End Sub

Public Sub Listar_Clientes()
  sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCCliente, AdoBenef, sSQL, "Cliente"
End Sub

Public Sub CalculosTotalesFactura(DtaProd As Adodc)
Dim NumLn As Byte
  Total_ME = 0
  Si_No = False
  Total_Factura = 0: Total_Desc = 0
  Total_Con_IVA = 0: Total_Sin_IVA = 0
  Total_IVA = 0
  NumLn = 0
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY CODIGO "
  SQLDec = "PRECIO 4|,CORTE 5|."
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQL, SQLDec
  With DtaProd.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total_IVA = Total_IVA + .Fields("Total_IVA")
          Total_Desc = Total_Desc + .Fields("Total_Desc")
          If .Fields("Total_IVA") > 0 Then
             Total_Con_IVA = Total_Con_IVA + .Fields("TOTAL")
          Else
             Total_Sin_IVA = Total_Sin_IVA + .Fields("TOTAL")
          End If
          NumLn = NumLn + 1
         .MoveNext
       Loop
   End If
  End With
  Total_Con_IVA = Round(Total_Con_IVA, 2)
  Total_Sin_IVA = Round(Total_Sin_IVA, 2)
  Total_Desc = Round(Total_Desc, 2)
  Total_Servicio = Round((Total_Sin_IVA + Total_Con_IVA - Total_Desc) * Porc_Serv, 2)
  Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
  LabelSubTotal.Caption = Format(Total_Sin_IVA, "#,##0.00")
  LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0.00")
  LabelIVA.Caption = Format(Total_IVA, "#,##0.00")
  TxtDescuento = Format(Total_Desc, "#,##0.00")
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
  Total_FacturaME = 0
  If Val(TextCotiza) > 0 Then
     TotalDolar = Val(CCur(TextCotiza))
     If OpcDiv.Value Then
        Total_FacturaME = Round(Total_Factura / TotalDolar, 2)
     Else
        Total_FacturaME = Round(Total_Factura * TotalDolar, 2)
     End If
  End If
  LabelTotalME.Caption = Format(Total_FacturaME, "#,##0.00")
  Label7.Caption = " CAMBIO (" & NumLn & ")"
  TextCant.Text = ""
  LabelVTotal.Caption = ""
End Sub

Private Sub DCArticulo_GotFocus()
  CalculosTotalesFactura AdoAsientoF
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    Select Case KeyCode
      Case vbKeyEscape
           CalculosTotalesFactura AdoAsientoF
           TxtEfectivo = Format(Total_Factura, "#,##0.00")
           TxtEfectivo.SetFocus
      Case vbKeyD
           If CtrlDown Then
              CalculosTotalesFactura AdoAsientoF
              TxtEfectivo = Format(Total_Factura, "#,##0.00")
              TxtDescuento.SetFocus
           End If
      Case vbKeyF1
           With AdoArticulo.Recordset
            If .RecordCount Then
               .MoveFirst
               .Find ("Nom_Art = '" & DCArticulo & "' ")
                If Not .EOF Then MsgBox .Fields("Producto") & ":" & vbCrLf & .Fields("Ayuda")
            End If
           End With
      Case vbKeyReturn
           TextCant.SetFocus
    End Select
End Sub

Private Sub DCArticulo_LostFocus()
  Codigos = Ninguno
  If Leer_Codigo_Inv(DCArticulo.Text, FechaSistema, AdoArticulo, Cod_Bodega) Then DatosArticulos
End Sub

Private Sub DCBodega_LostFocus()
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .Fields("CodBod")
   End If
  End With
End Sub

Private Sub DGAsientoF_BeforeDelete(Cancel As Integer)
'''  Mensajes = "¿Realmente desea eliminar el campo " & vbCrLf & "(" _
'''           & AdoAsientoF.Recordset.Fields("CODIGO") & ") " _
'''           & AdoAsientoF.Recordset.Fields("PRODUCTO") & "   TOTAL -> " _
'''           & AdoAsientoF.Recordset.Fields("TOTAL") & "?"
'''  Titulo = "Confirmación de eliminación"
'''  If BoxMensaje = 6 Then Cancel = False Else Cancel = True
  Cancel = False
End Sub

Private Sub Form_Activate()
  FA.TC = TipoFactura
  Ln_No = 1
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cyber " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCyber, sSQL
  With AdoCyber.Recordset
   If .RecordCount Then
       Do While Not .EOF
          Codigo = .Fields("Codigo")
          I = Val(Mid(Codigo, Len(Codigo) - 1, 2)) - 1
          If .Fields("PC_Ocupaga") Then
              TButtonPCs(I).Value = True
              TButtonPCs(I).Caption = " PC Ocupado"
              TButtonPCs(I).Picture = LoadPicture(RutaSistema & "\Iconos\candado1.ico")
          Else
              TButtonPCs(I).Value = False
              TButtonPCs(I).Caption = " PC Libre"
              TButtonPCs(I).Picture = LoadPicture(RutaSistema & "\Iconos\candado2.ico")
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  ConectarAdoExecute sSQL
  CodigoCliente = "9999999999"
  NombreCliente = "CONSUMIDOR FINAL"
  DireccionCli = " SN"
  TextBenef = NombreCliente
  FA.CodigoC = CodigoCliente
  If TipoFactura = "PV" Then
     FCyberNet.Caption = "INGRESAR TICKET"
     Label1.Caption = " TICKET No."
     Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
  ElseIf TipoFactura = "NV" Then
     FCyberNet.Caption = "INGRESAR NOTA DE VENTA"
     Label1.Caption = " NOTA DE VENTA No."
     Label3.Caption = "I.V.A. 0.00%"
  Else
     FCyberNet.Caption = "INGRESAR FACTURA"
     Label1.Caption = " FACTURA No."
     Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
  End If
  FCyberNet.Caption = FCyberNet.Caption & " (" & TipoFactura & ")"
  
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  TextCant.Text = "0"
  TextVUnit.Text = "0"
  LabelVTotal.Caption = "0"
  Modificar = False
  Bandera = True
  Mifecha = BuscarFecha(FechaSistema)
   
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = 'CXC" & TipoFactura & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo "
  SelectAdodc AdoLinea, sSQL
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  With AdoLinea.Recordset
    If .RecordCount > 0 Then
        Cta_Cobrar = .Fields("CxC")
        CodigoL = .Fields("Codigo")
    Else
        MsgBox "Falta Organizar la CxC en Puntos de Venta." & vbCrLf _
             & "Salga de este proceso y llame al su técnico" & vbCrLf _
             & "o al Contador de su Organizacion."
    End If
  End With
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SQLDec = "PRECIO 4|,CORTE 5|."
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQL, SQLDec
  
  FA.Cod_CxC = "CXC" & TipoFactura
  Lineas_De_CxC FA
  FA.Factura = Numero_Factura(FA)
  TextFacturaNo = FA.Factura
  TextFacturaNo.Text = Format(FA.Factura, "0000000")
  
  VerSiExisteCta Cta_CajaG
  VerSiExisteCta Cta_CajaGE
  VerSiExisteCta Cta_CajaBA
  MBFecha.Text = FechaSistema
  sSQL = "SELECT * " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodBod "
  SelectDBCombo DCBodega, AdoBodega, sSQL, "Bodega"
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .Fields("CodBod")
   End If
  End With
  Listar_Clientes

  sSQL = "SELECT Producto & ' -> ' & Codigo_Inv As Nom_Art,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'I' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDBList DLGrupo, AdoGrupo, sSQL, "Nom_Art"
  
  sSQL = "SELECT Producto & ' -> ' & Codigo_Inv As Nom_Art,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND MID(Codigo_Inv,1,2) <> '99' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDBCombo DCArticulo, AdoArticulo, sSQL, "Nom_Art"
  DGAsientoF.Height = MDI_Y_Max - DGAsientoF.Top
  RatonNormal
  FCyberNet.WindowState = 2
  If AdoArticulo.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Productos de Venta"
     Unload FCyberNet
  End If
End Sub

Private Sub Form_Deactivate()
  FCyberNet.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoBenef
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoCyber
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFactura
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
  For I = 0 To 3
      CButtonCabinas(I).PicturePosition = fmPicturePositionLeftCenter
      CButtonCabinas(I).Picture = LoadPicture(RutaSistema & "\Iconos\cabina.ico")
      CButtonCabinas(I).BackColor = &HC0FFFF
      CButtonCabinas(I).ForeColor = &H400000
      CButtonCabinas(I).Font.Bold = True
      CButtonCabinas(I).Font.Size = 18
      CButtonCabinas(I).Caption = " Cabina " & I + 1 & " "
  Next I
  For I = 0 To 7
      CButtonPCs(I).PicturePosition = fmPicturePositionLeftCenter
      CButtonPCs(I).Picture = LoadPicture(RutaSistema & "\Iconos\pcs.ico")
      CButtonPCs(I).BackColor = &HFFC0C0
      CButtonPCs(I).ForeColor = &HFF0000
      CButtonPCs(I).Font.Bold = True
      CButtonPCs(I).Font.Size = 10
      TButtonPCs(I).Font.Size = 10
      CButtonPCs(I).Caption = " PC No. " & I + 1 & " - Tiempo - 00:00:00 "
  Next I
  Encerar_Facturas
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  FechaTexto1 = MBFecha.Text
End Sub

Private Sub TextCant_Change()
  Real1 = Val(TextCant) * Val(TextVUnit)
  LabelVTotal.Caption = Format(Real1, "#,##0.0000")
End Sub

Private Sub TextCant_GotFocus()
  If Val(TextVUnit) <= 0 Then TextVUnit = Format(Precio, "#,##0.0000")
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True, , 4
End Sub

Private Sub TextCotiza_GotFocus()
  TextoValido TextCotiza
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_Change()
   Real1 = Redondear(Val(TextCant) * Val(TextVUnit), 2)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextVUnit_GotFocus()
  'TextCant.Text = "0"
  MarcarTexto TextVUnit
  If Round(LabelStock.Caption, 2) <= 0 Then
     Mensajes = "Producto sin existencia" & vbCrLf _
              & "Quiere continuar?"
     Titulo = "PUNTO DE VENTA"
     If BoxMensaje <> vbYes Then DCArticulo.SetFocus
  End If
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
Dim Grabar_PV As Boolean
   TextoValido TextVUnit, True, , 4
   Grabar_PV = True
   'MsgBox TextVUnit
   If Cant_Item_PV > 0 And (AdoAsientoF.Recordset.RecordCount >= Cant_Item_PV) Then Grabar_PV = False
   'MsgBox Cant_Item_PV
   If Grabar_PV Then
      TextoValido TextCant, True
      LabelVTotal.Caption = Format(Real1, "#,##0.00")
      Real2 = 0: Real3 = 0
      Real1 = CCur(Val(TextCant)) * CCur(Val(TextVUnit))
      Real1 = Redondear(Real1, 2)
      With AdoAsientoF.Recordset
        If Real1 > 0 Then
           If BanIVA Then Real3 = (Real1 - Real2) * Porc_IVA Else Real3 = 0
           If EsNotaVenta Then Real3 = 0
           SetAddNew AdoAsientoF
           SetFields AdoAsientoF, "CODIGO", Codigos
           SetFields AdoAsientoF, "CODIGO_L", CodigoL
           SetFields AdoAsientoF, "PRODUCTO", Mid(Producto, 1, 150)
           SetFields AdoAsientoF, "CANT", CDbl(TextCant)
           SetFields AdoAsientoF, "PRECIO", CDbl(TextVUnit)
           SetFields AdoAsientoF, "TOTAL", Real1
           SetFields AdoAsientoF, "Total_IVA", Real3
           SetFields AdoAsientoF, "Item", NumEmpresa
           SetFields AdoAsientoF, "CodigoU", CodigoUsuario
           SetFields AdoAsientoF, "A_No", Ln_No
           SetUpdate AdoAsientoF
           CalculosTotalesFactura AdoAsientoF
           Ln_No = Ln_No + 1
           TextVUnit.Text = ""
        End If
      End With
   Else
      MsgBox "Ya no puede ingresar mas productos"
   End If
  DCArticulo.SetFocus
End Sub

Public Sub ProcGrabar()
 DGAsientoF.Visible = False
 'Seteamos los encabezados para las facturas
  CalculosTotalesFactura AdoAsientoF
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     FA.Cod_CxC = "CXC" & TipoFactura
     Lineas_De_CxC FA
     FA.Factura = Numero_Factura(FA)
     FA.Fecha_C = FA.Fecha
     FA.Fecha_V = FA.Fecha
     TextoFormaPago = PagoCred
     FA.T = Cancelado
    'Grabamos el numero de factura
     Grabar_Factura FA
     RatonNormal
     Evaluar = True
     FechaTexto = MBFecha
     If FA.TC <> "PV" Then
       'Abono de Factura
        TA.T = Normal
        TA.TP = FA.TC
        TA.Fecha = FA.Fecha
        TA.Cta = Cta_CajaG
        TA.Cta_CxP = Cta_Cobrar
        TA.Banco = "EFECTIVO MN"
        TA.Cheque = Ninguno
        TA.Factura = FA.Factura
        TA.Abono = FA.Total_MN
        TA.Serie = FA.Serie
        TA.Autorizacion = FA.Autorizacion
        Grabar_Abonos TA
     End If
     sSQL = "UPDATE Facturas " _
          & "SET Saldo_MN = 0, T = 'C' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Factura = " & FA.Factura & " " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND CodigoC = '" & FA.CodigoC & "' "
     ConectarAdoExecute sSQL
     'MsgBox TipoFactura
     If Grafico_PV Then
        Imprimir_Punto_Venta_Grafico AdoFactura, AdoAsientoF, FA
     Else
        Imprimir_Punto_Venta AdoFactura, AdoAsientoF, FA
     End If
     NombreCliente = "CONSUMIDOR FINAL"
     TextBenef = "CONSUMIDOR FINAL"
     TxtEfectivo = "0.00"
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub DatosArticulos()
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
       LabelStock.Caption = DatInv.Stock
       Codigos = DatInv.Codigo_Inv
       Producto = DatInv.Producto
       Cta_Ventas = DatInv.Cta_Ventas
       Precio = Redondear(DatInv.PVP, 2)
       BanIVA = DatInv.IVA
       'MsgBox DatInv.IVA
       TextVUnit = Format(Precio, "#,##0.0000")
       If EsNotaVenta Then BanIVA = False
       DCArticulo.Text = Producto & " -> " & Codigos
   End If
  End With
End Sub

''''Private Sub TimerPCs_Timer()
''''Dim MiTiempoFin As Single
''''  For I = 0 To 7
''''      TiempoPCs(I) = "00:00:00"
''''      TotalPCs(I) = 0
''''  Next I
''''  MiTiempoFin = CDbl(CDate(Time))
''''  sSQL = "SELECT * " _
''''       & "FROM Catalogo_Cyber " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND Periodo = '" & Periodo_Contable & "' " _
''''       & "ORDER BY Codigo "
''''  SelectAdodc AdoCyber, sSQL
''''  With AdoCyber.Recordset
''''   If .RecordCount Then
''''       Do While Not .EOF
''''          Codigo = .Fields("Codigo")
''''          Codigo = Mid(Codigo, Len(Codigo) - 1, 2)
''''          I = Val(Codigo) - 1
''''          MiTiempo = CDbl(CDate(.Fields("Inicio")))
''''         'MsgBox Codigo & vbCrLf & I
''''          If .Fields("PC_Ocupaga") Then
''''              TiempoPCs(I) = CDate(MiTiempoFin - MiTiempo)
''''              TotalPCs(I) = Minute(TiempoPCs(I)) * 0.65 / 60
''''              TotalPCs(I) = TotalPCs(I) + Hour(TiempoPCs(I)) * 0.65
''''              If TotalPCs(I) < 0.15 Then TotalPCs(I) = 0.15
''''              TotalPCs(I) = Format(TotalPCs(I), "#,##0.00")
''''              CButtonPCs(I).Caption = " Equipo " & I + 1 & " " & vbCrLf _
''''                                    & " [" & Format(TiempoPCs(I), FormatoTimes) & "] " & vbCrLf _
''''                                    & " USD " & TotalPCs(I)
''''          End If
''''       .MoveNext
''''     Loop
''''   End If
''''  End With
''''End Sub

Private Sub TxtDescuento_GotFocus()
  MarcarTexto TxtDescuento
End Sub

Private Sub TxtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento_LostFocus()
  TextoValido TxtDescuento
End Sub

Private Sub TxtEfectivo_GotFocus()
  MarcarTexto TxtEfectivo
End Sub

Private Sub TxtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_Change()
  If Val(TextCotiza) > 0 Then
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_FacturaME, "#,##0.00")
     FA.Total_MN = Total_FacturaME
  Else
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_Factura, "#,##0.00")
     FA.Total_MN = Total_Factura
  End If
  FA.Efectivo = Val(TxtEfectivo)
End Sub

Private Sub TxtEfectivo_LostFocus()
  TextoValido TxtEfectivo, True, , 2
  If Val(TextCotiza) > 0 Then
     LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_FacturaME, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - Total_FacturaME) >= 0 Then CButtonGrabar.SetFocus
  Else
     LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_Factura, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then CButtonGrabar.SetFocus
  End If
End Sub

Private Sub TxtValorCabina_GotFocus()
  MarcarTexto TxtValorCabina
End Sub

Private Sub TxtValorCabina_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorCabina_LostFocus()
   TextoValido TxtValorCabina, True
   If Val(TxtValorCabina) > 0 Then
      With AdoAsientoF.Recordset
       If .RecordCount Then
          .MoveFirst
          .Find ("CODIGO = '99.01." & Format(ICabina + 1, "000") & "' ")
           If Not .EOF Then
             .Delete
             .Update
           End If
       End If
      End With
      SetAddNew AdoAsientoF
      SetFields AdoAsientoF, "CODIGO", "99.01." & Format(ICabina + 1, "000")
      SetFields AdoAsientoF, "CODIGO_L", CodigoL
      SetFields AdoAsientoF, "PRODUCTO", "Cab No." & Format(ICabina + 1, "00") & ", T[" & MBTiempo & "]"
      SetFields AdoAsientoF, "CANT", 1
      SetFields AdoAsientoF, "PRECIO", TxtValorCabina
      SetFields AdoAsientoF, "TOTAL", TxtValorCabina
      SetFields AdoAsientoF, "Total_IVA", 0
      SetFields AdoAsientoF, "RUTA", MBTiempo
      SetFields AdoAsientoF, "Item", NumEmpresa
      SetFields AdoAsientoF, "CodigoU", CodigoUsuario
      SetFields AdoAsientoF, "A_No", Ln_No
      SetUpdate AdoAsientoF
      CalculosTotalesFactura AdoAsientoF
      Ln_No = Ln_No + 1
   End If
   FrmCabina.Visible = False
End Sub
